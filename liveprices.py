import os
import sys
import time
import traceback
import ctypes
from ctypes import ( wintypes, byref, WINFUNCTYPE, POINTER, c_void_p, c_uint, c_ulong, c_int, c_size_t, )
from PyQt5.QtCore import ( Qt, QTimer, QPoint, QRect, QEasingCurve, QPropertyAnimation, QParallelAnimationGroup, )
from PyQt5.QtGui import QColor, QFont, QFontDatabase, QIcon, QKeySequence
from PyQt5.QtWidgets import ( QApplication, QWidget, QLabel, QPushButton, QFrame, QVBoxLayout, QHBoxLayout, QMessageBox, QComboBox, QSpinBox, QGraphicsDropShadowEffect, QShortcut, QScrollArea, QStyledItemDelegate, QStyle, QTableWidget, QTableWidgetItem, QHeaderView, QAbstractItemView, QDialog, QDialogButtonBox, )

CONFIG_FILE = "config.txt"
DDE_SERVICE = "MT4"
REFRESH_INTERVAL_MS = 100
DEFAULT_ROWS = 12
MAX_ROWS = 20
STATUS_STALE_SECONDS = 10
DEFAULT_FONT_SIZE = 22          # point size of the bid/ask text; other elements derive from this
MIN_FONT_SIZE = 12
MAX_FONT_SIZE = 48
APP_USER_MODEL_ID = "livepriceapp"
ICON_FILE = "icon.ico"

def resource_path(name):
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, name)

def _fmt(price):
    if price is None or price == "":
        return ""
    try:
        p = float(price)
    except (TypeError, ValueError):
        return str(price)
    if p > 9999:
        decimals = 3
    elif p >= 999:
        decimals = 4
    elif p >= 99:
        decimals = 5
    else:
        decimals = 6
    return f"{p:.{decimals}f}"

def save_config(pairs=None, rows=None, font=None, is_darkmode=True):
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        if pairs:
            entries = ",".join(f"{n}:{s}" for n, s in pairs if n and s)
            f.write(f"SYMBOLS={entries}\n")
        if font:
            f.write(f"FONT={font.family()},{font.pointSize()}\n")
        f.write(f"IS_DARKMODE={is_darkmode}\n")
        if rows:
            f.write("ROWS=" + ",".join(rows) + "\n")

def load_config():
    if not os.path.exists(CONFIG_FILE):
        return None
    cfg = {}
    with open(CONFIG_FILE, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line or "=" not in line:
                continue
            key, val = line.split("=", 1)
            cfg[key] = val

    if "FONT" in cfg:
        family, _, size = cfg["FONT"].rpartition(",")
        try:
            pt = int(size)
        except ValueError:
            pt = DEFAULT_FONT_SIZE
        if pt < MIN_FONT_SIZE:          # legacy configs persisted an unused size of 10
            pt = DEFAULT_FONT_SIZE
        cfg["FONT"] = QFont(family.strip(), pt)
    cfg["IS_DARKMODE"] = cfg.get("IS_DARKMODE", "True") == "True"

    pairs = []
    for entry in cfg.get("SYMBOLS", "").split(","):
        entry = entry.strip()
        if not entry:
            continue
        if ":" in entry:
            name, sym = (s.strip() for s in entry.split(":", 1))
        else:
            name = sym = entry
        if sym:
            pairs.append((name or sym, sym))
    cfg["PAIRS"] = pairs
    rows_str = cfg.get("ROWS", "")
    cfg["ROWS"] = rows_str.split(",") if rows_str else []
    return cfg

_user32 = ctypes.WinDLL("user32", use_last_error=True)

_APPCMD_CLIENTONLY = 0x00000010
_CP_WINUNICODE = 1200
_CF_TEXT = 1
_DDE_FACK = 0x8000
_XCLASS_BOOL = 0x1000
_XCLASS_FLAGS = 0x4000
_XCLASS_NOTIFICATION = 0x8000
_XTYP_ADVDATA = 0x0010 | _XCLASS_FLAGS
_XTYP_ADVSTART = 0x0030 | _XCLASS_BOOL
_XTYP_ADVSTOP = 0x0040 | _XCLASS_NOTIFICATION

_DDECALLBACK = WINFUNCTYPE(
    c_void_p, c_uint, c_uint, c_void_p, c_void_p, c_void_p, c_void_p,
    c_size_t, c_size_t,
)
for _fn, _args, _res in (
    ("DdeInitializeW",        [POINTER(c_ulong), _DDECALLBACK, c_ulong, c_ulong], c_uint),
    ("DdeUninitialize",       [c_ulong],                                          wintypes.BOOL),
    ("DdeCreateStringHandleW",[c_ulong, ctypes.c_wchar_p, c_int],                 c_void_p),
    ("DdeFreeStringHandle",   [c_ulong, c_void_p],                                wintypes.BOOL),
    ("DdeQueryStringW",       [c_ulong, c_void_p, ctypes.c_wchar_p, c_ulong, c_int], c_ulong),
    ("DdeConnect",            [c_ulong, c_void_p, c_void_p, c_void_p],            c_void_p),
    ("DdeClientTransaction",  [c_void_p, c_ulong, c_void_p, c_void_p, c_uint, c_uint, c_ulong, POINTER(c_ulong)], c_void_p),
    ("DdeAccessData",         [c_void_p, POINTER(c_ulong)],                       c_void_p),
    ("DdeUnaccessData",       [c_void_p],                                         wintypes.BOOL),
    ("DdeGetLastError",       [c_ulong],                                          c_uint),
):
    _f = getattr(_user32, _fn)
    _f.argtypes = _args
    _f.restype = _res

class MT4DDESource:

    TOPICS = ("BID", "ASK", "HIGH", "LOW")

    def __init__(self, pairs):
        self.pairs = self._normalize(pairs)
        self.mt4_symbols = list(dict.fromkeys(s for _, s in self.pairs))
        self.cache = {}
        self.last_update_ts = None
        self._first_logged = False

        self._callback = _DDECALLBACK(self._on_dde_event)

        self._idInst = c_ulong(0)
        rc = _user32.DdeInitializeW(byref(self._idInst), self._callback,
                                    _APPCMD_CLIENTONLY, 0)
        if rc != 0:
            raise RuntimeError(
                f"DdeInitialize failed (code {rc}). "
                "Is MT4 running with DDE server enabled?"
            )

        self._service_hsz = self._make_hsz(DDE_SERVICE)
        self._topic_hsz = {t: self._make_hsz(t) for t in self.TOPICS}
        self._item_hsz = {s: self._make_hsz(s) for s in self.mt4_symbols}
        self._conv = {}

        try:
            for topic in self.TOPICS:
                self._conv[topic] = self._connect_with_retry(topic)
                for sym in self.mt4_symbols:
                    self._advise(self._conv[topic], self._item_hsz[sym], _XTYP_ADVSTART)
        except Exception:
            self.close()
            raise

    @staticmethod
    def _normalize(pairs):
        out = []
        for n, s in pairs:
            n = (n or "").strip()
            s = (s or "").strip()
            if n and s:
                out.append((n, s))
        return out

    def update_pairs(self, new_pairs):
        cleaned = self._normalize(new_pairs)
        new_mt4 = list(dict.fromkeys(s for _, s in cleaned))
        old_set, new_set = set(self.mt4_symbols), set(new_mt4)

        for sym in new_set - old_set:
            self._item_hsz[sym] = self._make_hsz(sym)
            for topic in self.TOPICS:
                hConv = self._conv.get(topic)
                if hConv:
                    self._advise(hConv, self._item_hsz[sym], _XTYP_ADVSTART)

        for sym in old_set - new_set:
            hsz = self._item_hsz.pop(sym, None)
            if hsz:
                for topic in self.TOPICS:
                    hConv = self._conv.get(topic)
                    if hConv:
                        self._advise(hConv, hsz, _XTYP_ADVSTOP)
                _user32.DdeFreeStringHandle(self._idInst.value, hsz)
            for topic in self.TOPICS:
                self.cache.pop((topic, sym), None)

        self.pairs = cleaned
        self.mt4_symbols = new_mt4
        self._first_logged = False

    def _make_hsz(self, s):
        return _user32.DdeCreateStringHandleW(self._idInst.value, s, _CP_WINUNICODE)

    def _hsz_to_str(self, hsz):
        if not hsz:
            return ""
        buf = ctypes.create_unicode_buffer(256)
        n = _user32.DdeQueryStringW(self._idInst.value, hsz, buf, 256, _CP_WINUNICODE)
        return buf.value if n else ""

    def _connect_with_retry(self, topic):
        for _ in range(5):
            hConv = _user32.DdeConnect(
                self._idInst, self._service_hsz, self._topic_hsz[topic], None
            )
            if hConv:
                return hConv
            time.sleep(0.4)
        err = _user32.DdeGetLastError(self._idInst.value)
        raise RuntimeError(
            f"DdeConnect failed for topic '{topic}' (DDE err 0x{err:04x}). "
            "Make sure MT4 is running and 'Enable DDE server' is on."
        )

    def _advise(self, hConv, item_hsz, xtype):
        pdwResult = c_ulong(0)
        _user32.DdeClientTransaction(
            None, 0, hConv, item_hsz, _CF_TEXT, xtype, 1000, byref(pdwResult)
        )

    def _on_dde_event(self, uType, uFmt, hConv, hsz1, hsz2, hData, dwData1, dwData2):
        if uType != _XTYP_ADVDATA:
            return 0
        topic = self._hsz_to_str(hsz1)
        item = self._hsz_to_str(hsz2)
        length = c_ulong(0)
        ptr = _user32.DdeAccessData(hData, byref(length))
        if ptr:
            try:
                raw = ctypes.string_at(ptr, length.value)
                self.cache[(topic, item)] = raw.decode("ascii", errors="ignore").rstrip("\x00").strip()
                self.last_update_ts = time.monotonic()
            finally:
                _user32.DdeUnaccessData(hData)
        return _DDE_FACK

    @staticmethod
    def _clean(v):
        if not v or v.upper().replace("\\", "/") in ("N/A", "NA"):
            return ""
        return v

    def read_rows(self):
        rows = []
        for name, sym in self.pairs:
            bid = self._clean(self.cache.get(("BID", sym), ""))
            ask = self._clean(self.cache.get(("ASK", sym), ""))
            high = self._clean(self.cache.get(("HIGH", sym), ""))
            low = self._clean(self.cache.get(("LOW", sym), ""))
            if not self._first_logged and (bid or ask):
                print(f"DDE [{name} -> {sym}] bid={bid!r} ask={ask!r} high={high!r} low={low!r}")
                self._first_logged = True
            rows.append((name, _fmt(bid), _fmt(ask), _fmt(low), _fmt(high)))
        return rows

    def status(self):
        """Return ('live' | 'stale' | 'down', human-readable detail)."""
        if not self._conv:
            return ("down", "No DDE conversation with MT4")
        if self.last_update_ts is None:
            return ("stale", "Connected to MT4 — waiting for first tick")
        age = time.monotonic() - self.last_update_ts
        if age <= STATUS_STALE_SECONDS:
            return ("live", f"Receiving ticks (last update {age:.1f}s ago)")
        return ("stale", f"No ticks for {age:.0f}s (market quiet or feed paused)")

    def close(self):
        try:
            if self._idInst.value:
                _user32.DdeUninitialize(self._idInst.value)
                self._idInst = c_ulong(0)
                self._conv.clear()
        except Exception:
            pass

class SymbolsConfigDialog(QDialog):

    LIGHT_QSS = """
        QDialog { background-color: #f5f5f5; color: #000; }
        QDialog QLabel { color: #000; background: transparent; }
        QDialog QTableWidget {
            background-color: #fff; color: #000; gridline-color: #ccc;
            selection-background-color: #d4af37; selection-color: #000;
        }
        QDialog QHeaderView::section {
            background-color: #e0e0e0; color: #000;
            padding: 4px; border: 1px solid #c0c0c0;
        }
        QDialog QTableWidget QTableCornerButton::section {
            background-color: #e0e0e0; border: 1px solid #c0c0c0;
        }
        QDialog QPushButton {
            background-color: #fff; color: #000;
            border: 1px solid #b0b0b0; padding: 5px 12px; border-radius: 3px;
        }
        QDialog QPushButton:hover { background-color: #ececec; }
        QDialog QPushButton:pressed { background-color: #d8d8d8; }
        QDialog QPushButton:disabled { color: #888; background-color: #f0f0f0; }
        QDialog QLineEdit {
            background-color: #ffffff; color: #000;
            selection-background-color: #d4af37; selection-color: #000;
            border: 1px solid #b0b0b0;
        }
        QDialog QTableWidget::item:selected {
            background-color: #d4af37; color: #000;
        }
    """

    def __init__(self, initial_pairs=None, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Live Price Symbols")
        self.resize(520, 420)
        self.setStyleSheet(self.LIGHT_QSS)
        self.pairs = []

        layout = QVBoxLayout(self)
        layout.addWidget(QLabel())

        self.table = QTableWidget(0, 2)
        self.table.setHorizontalHeaderLabels(["Name", "MT4 Symbol"])
        for col in (0, 1):
            self.table.horizontalHeader().setSectionResizeMode(col, QHeaderView.Stretch)
        self.table.verticalHeader().setVisible(False)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        layout.addWidget(self.table)

        self.add_btn = QPushButton("➕ Add row")
        self.rm_btn = QPushButton("➖ Remove selected")
        self.cap_label = QLabel()
        self.cap_label.setStyleSheet("color: #555; padding-left: 10px;")
        self.add_btn.clicked.connect(lambda: self._add_row("", ""))
        self.rm_btn.clicked.connect(self._remove_selected)

        btn_row = QHBoxLayout()
        btn_row.addWidget(self.add_btn)
        btn_row.addWidget(self.rm_btn)
        btn_row.addWidget(self.cap_label)
        btn_row.addStretch()
        layout.addLayout(btn_row)

        bb = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        bb.accepted.connect(self.accept)
        bb.rejected.connect(self.reject)
        layout.addWidget(bb)

        for n, s in (initial_pairs or [])[:MAX_ROWS]:
            self._add_row(n, s)
        if not self.table.rowCount():
            for _ in range(3):
                self._add_row("", "")
        self._update_cap()

    def _add_row(self, name, sym):
        if self.table.rowCount() >= MAX_ROWS:
            return
        r = self.table.rowCount()
        self.table.insertRow(r)
        self.table.setItem(r, 0, QTableWidgetItem(name))
        self.table.setItem(r, 1, QTableWidgetItem(sym))
        self._update_cap()

    def _remove_selected(self):
        rows = sorted({i.row() for i in self.table.selectedIndexes()}, reverse=True)
        for r in rows:
            self.table.removeRow(r)
        if not self.table.rowCount():
            self._add_row("", "")
        self._update_cap()

    def _update_cap(self):
        if not hasattr(self, "add_btn"):
            return
        n = self.table.rowCount()
        self.add_btn.setEnabled(n < MAX_ROWS)
        self.cap_label.setText(
            f"Maximum {MAX_ROWS} rows reached" if n >= MAX_ROWS else f"{n} / {MAX_ROWS}"
        )

    def accept(self):
        pairs, seen = [], set()
        for r in range(self.table.rowCount()):
            name = (self.table.item(r, 0).text() if self.table.item(r, 0) else "").strip()
            sym = (self.table.item(r, 1).text() if self.table.item(r, 1) else "").strip()
            if not (name or sym):
                continue
            if not sym:
                QMessageBox.warning(self, "Error", f"Row {r+1}: MT4 Symbol is required.")
                return
            name = name or sym
            if name in seen:
                QMessageBox.warning(self, "Error", f"Duplicate name '{name}'. Names must be unique.")
                return
            seen.add(name)
            pairs.append((name, sym))
        if not pairs:
            QMessageBox.warning(self, "Error", "Please add at least one symbol.")
            return
        self.pairs = pairs
        super().accept()

class PriceBox(QFrame):
    UP_STYLE   = "color: lime; font-weight: bold;"
    DOWN_STYLE = "color: red;  font-weight: bold;"

    def __init__(self, row_index, remove_callback, parent_widget):
        super().__init__()
        self.remove_callback = remove_callback
        self.parent_widget = parent_widget
        self.last_bid = 0.0
        self.last_ask = 0.0

        shadow = QGraphicsDropShadowEffect()
        shadow.setBlurRadius(25)
        shadow.setColor(QColor(212, 175, 55, 120))
        shadow.setOffset(2, 2)
        self.setGraphicsEffect(shadow)

        layout = QHBoxLayout(self)
        layout.setContentsMargins(10, 5, 10, 5)
        layout.setSpacing(12)

        self.symbol = self._mk_label(layout)
        self.bid    = self._mk_label(layout)
        self.ask    = self._mk_label(layout)
        self.low    = self._mk_label(layout)
        self.high   = self._mk_label(layout)

        col = QVBoxLayout()
        col.setContentsMargins(0, 0, 0, 0)
        col.setSpacing(0)
        self.up_btn = self._mk_arrow_btn("▲")
        self.down_btn = self._mk_arrow_btn("▼")
        col.addWidget(self.up_btn, alignment=Qt.AlignHCenter)
        col.addWidget(self.down_btn, alignment=Qt.AlignHCenter)
        layout.addLayout(col)
        self.up_btn.clicked.connect(lambda: parent_widget.request_move(self, -1))
        self.down_btn.clicked.connect(lambda: parent_widget.request_move(self, +1))

        self.remove_btn = QPushButton("✖")
        self.remove_btn.setCursor(Qt.PointingHandCursor)
        self.remove_btn.clicked.connect(self._on_remove)
        layout.addWidget(self.remove_btn)

        self.apply_theme()
        self.update_background(row_index)
        self.update_buttons()

    @staticmethod
    def _mk_label(layout):
        lbl = QLabel("")
        layout.addWidget(lbl, 1)
        return lbl

    @staticmethod
    def _mk_arrow_btn(glyph):
        btn = QPushButton(glyph)
        btn.setCursor(Qt.PointingHandCursor)
        return btn

    def _on_remove(self):
        name = self.symbol.text().strip()
        self.symbol.setText("")
        self.update_prices("", "", "", "")
        if self.remove_callback:
            self.remove_callback(self, name)

    def update_buttons(self):
        empty = not self.symbol.text().strip()
        self.remove_btn.setVisible(not empty)
        self.up_btn.setVisible(not empty)
        self.down_btn.setVisible(not empty)

    def update_prices(self, bid, ask, low, high):
        self._set_priced(self.bid, bid, "last_bid")
        self._set_priced(self.ask, ask, "last_ask")
        self.low.setText(_fmt(low))
        self.high.setText(_fmt(high))

    def _set_priced(self, lbl, value, prev_attr):
        if value in (None, ""):
            lbl.setText("")
            return
        try:
            num = float(value)
        except (TypeError, ValueError):
            lbl.setText(str(value))
            return
        prev = getattr(self, prev_attr)
        if num > prev:
            lbl.setStyleSheet(self.UP_STYLE)
        elif num < prev:
            lbl.setStyleSheet(self.DOWN_STYLE)
        lbl.setText(_fmt(num))
        setattr(self, prev_attr, num)

    def update_background(self, row_index):
        if self.parent_widget.is_darkmode:
            bg = "#000000"
        else:
            bg = "#f5f4e9" if row_index % 2 else "#fcf6dc"
        self.setStyleSheet(f"QFrame {{ background-color: {bg}; border-radius: 5px; }}")

    def apply_theme(self):
        pw = self.parent_widget
        dark = pw.is_darkmode
        fam = pw.current_font.family()
        base = pw.base_font_size()
        txt = "white" if dark else "black"

        self.symbol.setFont(QFont(fam, max(8, base - 2)))
        self.symbol.setStyleSheet(f"color: {txt};")
        for lbl in (self.bid, self.ask, self.low, self.high):
            lbl.setFont(QFont(fam, base))
            lbl.setStyleSheet(f"color: {txt};")

        btn_pt = max(8, base - 4)
        side = max(28, btn_pt + 10)
        arrow_color = "gray" if dark else "lightgray"
        for btn in (self.up_btn, self.down_btn):
            btn.setFont(QFont(fam, btn_pt))
            btn.setFixedSize(side, side)
            btn.setStyleSheet(f"color: {arrow_color}; background: transparent; border: none;")
        self.remove_btn.setFont(QFont(fam, btn_pt))
        self.remove_btn.setFixedSize(side, side)
        self.remove_btn.setStyleSheet("color: red; background: transparent; border: none;")

class _FontDelegate(QStyledItemDelegate):
    def paint(self, painter, option, index):
        painter.save()
        if option.state & QStyle.State_Selected:
            painter.fillRect(option.rect, option.palette.highlight())
        painter.setFont(QFont(index.data(), 12))
        painter.setPen(option.palette.text().color())
        painter.drawText(option.rect.adjusted(5, 0, 0, 0), Qt.AlignVCenter, index.data())
        painter.restore()

class FontChanger(QWidget):
    def __init__(self, main_window):
        super().__init__()
        self.main_window = main_window
        self.setWindowTitle("Font")
        self.resize(400, 280)

        layout = QVBoxLayout(self)
        layout.addWidget(QLabel("Font family"))
        self.combo = QComboBox()
        self.combo.setItemDelegate(_FontDelegate())
        self.combo.addItems(QFontDatabase().families())
        i = self.combo.findText(main_window.current_font.family())
        if i >= 0:
            self.combo.setCurrentIndex(i)
        layout.addWidget(self.combo)

        layout.addWidget(QLabel("Base size (bid / ask text)"))
        self.size_spin = QSpinBox()
        self.size_spin.setRange(MIN_FONT_SIZE, MAX_FONT_SIZE)
        self.size_spin.setValue(max(MIN_FONT_SIZE, min(MAX_FONT_SIZE, main_window.base_font_size())))
        layout.addWidget(self.size_spin)

        layout.addStretch()
        apply_btn = QPushButton("Apply")
        apply_btn.clicked.connect(self._apply)
        layout.addWidget(apply_btn)

    def _apply(self):
        font = QFont(self.combo.currentText(), self.size_spin.value())
        self.main_window.current_font = font
        self.main_window.apply_theme()
        self.main_window._save_state()
        self.close()

class MainWindow(QWidget):
    HEADERS = ("Symbol", "Bid", "Ask", "Low", "High")

    def __init__(self, pairs, saved_rows=None, is_darkmode=True, font=None):
        super().__init__()
        self.setWindowTitle("Live Prices")
        icon_path = resource_path(ICON_FILE)
        if os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))
        self.pairs = list(pairs)
        self.is_darkmode = is_darkmode
        self.current_font = font or QFont("Arial", DEFAULT_FONT_SIZE)
        self.is_fullscreen = False
        self._anim_group = None
        self.last_rows_dict = {}

        main = QVBoxLayout(self)
        main.setContentsMargins(5, 5, 5, 5)
        main.setSpacing(5)
        self._build_header(main)
        self._build_scroll(main)

        try:
            self.source = MT4DDESource(self.pairs)
        except Exception as e:
            QMessageBox.critical(
                self, "DDE Error",
                "Failed to connect to MT4 DDE server.\n\n"
                f"{e}\n\n"
                "Make sure MT4 is running and DDE server is enabled\n"
                "(Tools → Options → Server → Enable DDE server)."
            )
            raise

        n_initial = max(DEFAULT_ROWS, min(MAX_ROWS, len(self.pairs)))
        self.boxes = [self._make_box(i) for i in range(n_initial)]

        for i, name in enumerate(self._initial_order(saved_rows or [])):
            if i < len(self.boxes):
                self.boxes[i].symbol.setText(name)

        self.refresh_once()
        self.timer = QTimer()
        self.timer.timeout.connect(self.refresh_once)
        self.timer.start(REFRESH_INTERVAL_MS)

        self._wire_shortcuts()
        self.apply_theme()

    def _build_header(self, main_layout):
        self.header_frame = QFrame()
        hl = QHBoxLayout(self.header_frame)
        hl.setContentsMargins(10, 8, 10, 8)
        hl.setSpacing(12)
        self.header_symbol_lbl = None
        for h in self.HEADERS:
            lbl = QLabel(h)
            lbl.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
            hl.addWidget(lbl, 1)
            if h == "Symbol":
                self.header_symbol_lbl = lbl

        spacer = QFrame()
        spacer.setFixedWidth(5)
        hl.addWidget(spacer)

        self.status_dot = QLabel("●")
        self.status_dot.setFixedWidth(22)
        self.status_dot.setAlignment(Qt.AlignCenter)
        self.status_dot.setToolTip("Connecting…")
        self._status_state = None
        hl.addWidget(self.status_dot)

        self.mode_btn = QPushButton("🌓")
        self.mode_btn.clicked.connect(self.toggle_mode)
        hl.addWidget(self.mode_btn)

        main_layout.addWidget(self.header_frame)

    def _build_scroll(self, main_layout):
        self.scroll = QScrollArea()
        self.scroll.setWidgetResizable(True)
        self.scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.scroll.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        main_layout.addWidget(self.scroll, 1)

        self.rows_container = QWidget()
        self.rows_layout = QVBoxLayout(self.rows_container)
        self.rows_layout.setContentsMargins(0, 0, 0, 0)
        self.rows_layout.setSpacing(5)
        self.scroll.setWidget(self.rows_container)

    def _make_box(self, index):
        box = PriceBox(
            row_index=index,
            remove_callback=self.on_row_cleared,
            parent_widget=self,
        )
        self.rows_layout.addWidget(box)
        return box

    def _wire_shortcuts(self):
        for key, slot in (
            ("Ctrl+Shift+F1", self.toggle_fullscreen),
            ("Ctrl+Shift+F",  self.open_font_changer),
            ("Ctrl+Shift+S",  self.open_symbols_editor),
        ):
            QShortcut(QKeySequence(key), self).activated.connect(slot)

    def _initial_order(self, saved_rows):
        names = [n for n, _ in self.pairs]
        name_set = set(names)
        ordered = [r for r in saved_rows if r in name_set]
        ordered += [n for n in names if n not in ordered]
        return ordered

    def refresh_once(self):
        try:
            rows = self.source.read_rows()
        except Exception as e:
            print(f"refresh error: {e}")
            traceback.print_exc()
            return
        self.last_rows_dict = {n: (b, a, l, h) for n, b, a, l, h in rows}
        for box in self.boxes:
            name = box.symbol.text().strip()
            if name and name in self.last_rows_dict:
                box.update_prices(*self.last_rows_dict[name])
            elif not name:
                box.update_prices("", "", "", "")
        self._update_status_indicator()

    def _update_status_indicator(self):
        dot = getattr(self, "status_dot", None)
        if dot is None:
            return
        src = getattr(self, "source", None)
        state, detail = src.status() if src is not None else ("down", "Not connected")
        dot.setToolTip(detail)
        if state != self._status_state:
            self._status_state = state
            color = {"live": "#2ecc40", "stale": "#ff9500", "down": "#ff3b30"}.get(state, "gray")
            dot.setStyleSheet(f"color: {color};")

    def base_font_size(self):
        pt = self.current_font.pointSize()
        return pt if 8 <= pt <= 72 else DEFAULT_FONT_SIZE

    def apply_theme(self):
        if self.is_darkmode:
            self.setStyleSheet("background-color: black;")
            self.header_frame.setStyleSheet("background-color: #111;")
            header_color = "gold"
        else:
            self.setStyleSheet("background-color: white;")
            self.header_frame.setStyleSheet("background-color: white;")
            header_color = "black"

        fam = self.current_font.family()
        base = self.base_font_size()
        hdr_font = QFont(fam, max(8, base - 4))
        for lbl in self.header_frame.findChildren(QLabel):
            if lbl is self.status_dot:
                continue
            lbl.setFont(hdr_font)
            lbl.setStyleSheet(f"color: {header_color}; font-weight: bold;")

        dot_pt = max(8, base - 8)
        self.status_dot.setFont(QFont(fam, dot_pt))
        self.status_dot.setFixedWidth(max(22, dot_pt + 10))
        self._update_status_indicator()

        side = max(35, (base - 4) + 14)
        self.mode_btn.setFont(QFont(fam, max(8, base - 6)))
        self.mode_btn.setFixedSize(side, side)
        self.mode_btn.setStyleSheet(
            f"color: {'white' if self.is_darkmode else 'black'}; "
            f"border: 1px solid {'white' if self.is_darkmode else '#999'};"
        )

        for i, box in enumerate(self.boxes):
            box.update_background(i)
            box.apply_theme()
            box.update_buttons()

    def toggle_mode(self):
        self.is_darkmode = not self.is_darkmode
        self.apply_theme()

    def apply_font_to_widgets(self):
        # All fonts/sizes are derived from current_font inside apply_theme().
        self.apply_theme()

    def open_font_changer(self):
        self._font_window = FontChanger(self)
        self._font_window.show()

    def toggle_fullscreen(self):
        if self.is_fullscreen:
            self.showNormal()
        else:
            self.showFullScreen()
        self.is_fullscreen = not self.is_fullscreen

    def on_row_cleared(self, box, name=""):
        if name:
            new_pairs = [(n, s) for n, s in self.pairs if n != name]
            if len(new_pairs) != len(self.pairs):
                self.pairs = new_pairs
                try:
                    self.source.update_pairs(new_pairs)
                except Exception as e:
                    print(f"warning: update_pairs failed: {e}")
                self._save_state()

        if len(self.boxes) > max(DEFAULT_ROWS, len(self.pairs)):
            try:
                self.boxes.remove(box)
                self.rows_layout.removeWidget(box)
                box.deleteLater()
            except (ValueError, RuntimeError):
                pass

        self._reorder()
        self._refresh_button_visibility()
        self._apply_row_heights()

    def _reorder(self):
        active = [b for b in self.boxes if b.symbol.text().strip()]
        empty = [b for b in self.boxes if not b.symbol.text().strip()]
        ordered = active + empty
        for b in self.boxes:
            self.rows_layout.removeWidget(b)
        for i, b in enumerate(ordered):
            self.rows_layout.insertWidget(i, b)
            b.show()
        self.boxes = ordered

    def _refresh_button_visibility(self):
        for b in self.boxes:
            b.update_buttons()

    def request_move(self, box, direction):
        try:
            i = self.boxes.index(box)
        except ValueError:
            return
        j = i + direction
        if not (0 <= j < len(self.boxes)):
            return
        other = self.boxes[j]

        viewport = self.scroll.viewport()
        r1 = QRect(box.mapTo(viewport, QPoint(0, 0)), box.size())
        r2 = QRect(other.mapTo(viewport, QPoint(0, 0)), other.size())
        ghost1 = self._spawn_ghost(box, r1, viewport)
        ghost2 = self._spawn_ghost(other, r2, viewport)
        box.hide()
        other.hide()

        group = QParallelAnimationGroup(self)
        group.addAnimation(self._slide(ghost1, r1.topLeft(), r2.topLeft()))
        group.addAnimation(self._slide(ghost2, r2.topLeft(), r1.topLeft()))

        def finalize():
            self.boxes[i], self.boxes[j] = self.boxes[j], self.boxes[i]
            for b in self.boxes:
                self.rows_layout.removeWidget(b)
            for k, b in enumerate(self.boxes):
                self.rows_layout.insertWidget(k, b)
                b.show()
            ghost1.deleteLater()
            ghost2.deleteLater()
            self.scroll.ensureWidgetVisible(box)

        group.finished.connect(finalize)
        self._anim_group = group
        group.start()

    @staticmethod
    def _spawn_ghost(widget, rect, viewport):
        ghost = QLabel(viewport)
        ghost.setPixmap(widget.grab())
        ghost.setGeometry(rect)
        ghost.setAttribute(Qt.WA_TransparentForMouseEvents)
        ghost.raise_()
        ghost.show()
        return ghost

    @staticmethod
    def _slide(widget, from_pt, to_pt):
        anim = QPropertyAnimation(widget, b"pos")
        anim.setDuration(220)
        anim.setStartValue(from_pt)
        anim.setEndValue(to_pt)
        anim.setEasingCurve(QEasingCurve.OutCubic)
        return anim

    def open_symbols_editor(self):
        dlg = SymbolsConfigDialog(initial_pairs=self.pairs, parent=self)
        if dlg.exec_() != QDialog.Accepted:
            return
        try:
            self.source.update_pairs(dlg.pairs)
        except Exception as e:
            QMessageBox.critical(
                self, "DDE Error", f"Failed to update DDE subscriptions:\n\n{e}"
            )
            return
        self.pairs = list(dlg.pairs)
        self._refill_boxes()
        self._save_state()
        self.refresh_once()

    def _refill_boxes(self):
        new_names = [n for n, _ in self.pairs]
        new_set = set(new_names)
        for box in self.boxes:
            if box.symbol.text().strip() and box.symbol.text().strip() not in new_set:
                box.symbol.setText("")
                box.update_prices("", "", "", "")

        shown = {b.symbol.text().strip() for b in self.boxes if b.symbol.text().strip()}
        for name in (n for n in new_names if n not in shown):
            target = next((b for b in self.boxes if not b.symbol.text().strip()), None)
            if target is None:
                target = self._make_box(len(self.boxes))
                target.symbol.setFixedWidth(int(self.width() * 0.3))
                self.boxes.append(target)
                target.update_background(len(self.boxes) - 1)
                target.apply_theme()
            target.symbol.setText(name)

        desired = max(DEFAULT_ROWS, len(self.pairs))
        while len(self.boxes) > desired:
            target = next((b for b in reversed(self.boxes) if not b.symbol.text().strip()), None)
            if target is None:
                break
            self.boxes.remove(target)
            self.rows_layout.removeWidget(target)
            target.deleteLater()

        self._reorder()
        self._refresh_button_visibility()
        self._apply_row_heights()

    def resizeEvent(self, event):
        super().resizeEvent(event)
        sym_w = int(self.width() * 0.3)
        for box in self.boxes:
            box.symbol.setFixedWidth(sym_w)
        if self.header_symbol_lbl:
            self.header_symbol_lbl.setFixedWidth(sym_w)
        self._apply_row_heights()

    def _apply_row_heights(self):
        n = len(self.boxes)
        if n <= 0:
            return
        if n <= DEFAULT_ROWS:
            for box in self.boxes:
                box.setMinimumHeight(0)
                box.setMaximumHeight(16777215)
            return
        margins = self.rows_layout.contentsMargins()
        avail = (self.scroll.viewport().height()
                 - margins.top() - margins.bottom()
                 - max(0, n - 1) * self.rows_layout.spacing())
        if avail <= 0:
            return
        h = max(24, avail // n)
        for box in self.boxes:
            box.setFixedHeight(h)

    def _save_state(self):
        try:
            rows = [b.symbol.text().strip() for b in self.boxes]
            save_config(
                pairs=self.pairs, rows=rows,
                font=self.current_font, is_darkmode=self.is_darkmode,
            )
        except Exception as e:
            print(f"warning: could not save {CONFIG_FILE}: {e}")

    def closeEvent(self, event):
        try:
            self.timer.stop()
        except Exception:
            pass
        try:
            self.source.close()
        except Exception:
            pass
        self._save_state()
        super().closeEvent(event)

def main():
    try:
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(APP_USER_MODEL_ID)
    except Exception:
        pass

    app = QApplication(sys.argv)
    icon_path = resource_path(ICON_FILE)
    if os.path.exists(icon_path):
        app.setWindowIcon(QIcon(icon_path))
    cfg = load_config()

    if cfg and cfg.get("PAIRS"):
        pairs = cfg["PAIRS"]
        saved_rows = cfg.get("ROWS", [])
        is_darkmode = cfg.get("IS_DARKMODE", True)
        font = cfg.get("FONT", QFont("Arial", DEFAULT_FONT_SIZE))
    else:
        dlg = SymbolsConfigDialog()
        if dlg.exec_() != QDialog.Accepted:
            sys.exit()
        pairs = dlg.pairs
        saved_rows = []
        is_darkmode = True
        font = QFont("Arial", DEFAULT_FONT_SIZE)

    try:
        save_config(pairs=pairs, rows=saved_rows, font=font, is_darkmode=is_darkmode)
    except Exception as e:
        print(f"warning: could not write {CONFIG_FILE}: {e}")

    window = MainWindow(pairs, saved_rows=saved_rows, is_darkmode=is_darkmode, font=font)
    window.showMaximized()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()

import sys
import os
import time
import traceback
import ctypes
from ctypes import wintypes, byref, WINFUNCTYPE, POINTER, c_void_p, c_uint, c_ulong, c_int, c_size_t
from PyQt5.QtWidgets import (
    QApplication, QWidget, QLabel, QLineEdit, QPushButton,
    QVBoxLayout, QHBoxLayout, QMessageBox, QComboBox,
    QGridLayout, QGraphicsDropShadowEffect, QShortcut, QFrame,
    QListWidget, QListWidgetItem, QScrollArea, QStyledItemDelegate, QStyle, QShortcut,
    QTableWidget, QTableWidgetItem, QHeaderView, QAbstractItemView,
    QDialog, QDialogButtonBox
)
from PyQt5.QtGui import QColor, QKeySequence, QPixmap, QPainter, QPolygon, QBrush, QFont, QFontDatabase
from PyQt5.QtCore import (
    Qt, QTimer, QPoint, QEvent, QRect,
    QEasingCurve, QPropertyAnimation, QParallelAnimationGroup
)

# -------------------------------
# Config
# -------------------------------
CONFIG_FILE = "config.txt"
DDE_SERVICE = "MT4"
REFRESH_INTERVAL_MS = 100
MAX_BOXES = 12  # default visible row count; rows below this use natural height
MAX_ROWS = 20   # hard cap on pairs; rows shrink to fit when count > MAX_BOXES
# -------------------------------
# Config file handling
# -------------------------------
def save_config(pairs=None, rows=None, font=None, is_darkmode=True):
    """
    pairs: list of (display_name, mt4_symbol) tuples
    rows:  list of display names in the order the UI shows them
    """
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        if pairs:
            entries = [f"{n}:{s}" for n, s in pairs if n and s]
            f.write(f"SYMBOLS={','.join(entries)}\n")
        if font:
            f.write(f"FONT={font.family()},{font.pointSize()}\n")
        f.write(f"IS_DARKMODE={is_darkmode}\n")
        if rows:
            f.write(f"ROWS={','.join(rows)}\n")


def load_config():
    if not os.path.exists(CONFIG_FILE):
        return None
    config = {}
    with open(CONFIG_FILE, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line:
                continue
            key, val = line.split("=", 1)
            config[key] = val
    if "FONT" in config:
        family, size = config["FONT"].split(",")
        config["FONT"] = QFont(family, int(size))
    config["IS_DARKMODE"] = config.get("IS_DARKMODE","True") == "True"

    # PAIRS: each entry is "name:symbol". Legacy entries without ":" are
    # treated as the same value for both (display name == MT4 symbol).
    pairs = []
    if "SYMBOLS" in config:
        for entry in config["SYMBOLS"].split(","):
            entry = entry.strip()
            if not entry:
                continue
            if ":" in entry:
                name, sym = entry.split(":", 1)
                name, sym = name.strip(), sym.strip()
            else:
                name = sym = entry
            if sym:
                pairs.append((name or sym, sym))
    config["PAIRS"] = pairs
    config["ROWS"] = config["ROWS"].split(",") if "ROWS" in config else []
    return config


# -------------------------------
# Helpers
# -------------------------------
def _fmt(price):
    try:
        if price is None:
            return ""
        p = float(price)
        # Dynamic decimal places based on value
        if p > 9999:
            decimals = 3
        elif p >= 999:
            decimals = 4
        elif p >= 99:
            decimals = 5
        else:
            decimals = 6
        s = f"{p:.{decimals}f}"
        return s.rstrip('.') if '.' in s else s
    except Exception:
        return str(price) if price is not None else ""

# -------------------------------
# Arrow Helper (kept)
# -------------------------------
def create_arrow(color, direction="up"):
    pixmap = QPixmap(16, 16)
    pixmap.fill(Qt.transparent)
    painter = QPainter(pixmap)
    painter.setBrush(QBrush(QColor(color)))
    painter.setPen(Qt.NoPen)
    if direction == "up":
        points = [(8,0),(16,16),(0,16)]
    else:
        points = [(0,0),(16,0),(8,16)]
    polygon = QPolygon([QPoint(x,y) for x,y in points])
    painter.drawPolygon(polygon)
    painter.end()
    return pixmap






# -------------------------------
# Price Box
# -------------------------------
class PriceBox(QFrame):
    def __init__(self, symbol="", row_index=0, remove_callback=None, add_callback=None, parent_widget=None, header_symbol_lbl = None, header_frame = None):
        super().__init__()
        self.last_bid = 0.0
        self.last_ask = 0.0
        self.remove_callback = remove_callback
        self.add_callback = add_callback
        self.parent_widget = parent_widget  # MainWindow
        self.header_symbol_lbl = header_symbol_lbl
        self.header_frame = header_frame
        

        # shadow + style
        shadow = QGraphicsDropShadowEffect()
        shadow.setBlurRadius(25)
        shadow.setColor(QColor(212, 175, 55, 120))
        shadow.setOffset(2, 2)
        self.setGraphicsEffect(shadow)

        layout = QHBoxLayout(self)
        layout.setContentsMargins(10,5,10,5)
        layout.setSpacing(12)

        
        

        # Symbol
        self.symbol = QLabel(symbol)
        self.symbol.setStyleSheet("color: white; font-size: 20pt; ")
        #self.symbol.setFixedWidth(700)
        layout.addWidget(self.symbol, 1)

        # Bid
        self.bid = QLabel("")
        self.bid.setStyleSheet("color: white; font-size: 22pt;")
        self.bid_arrow = QLabel()
        bid_layout = QHBoxLayout()
        bid_layout.setContentsMargins(0,0,0,0)
        bid_layout.setSpacing(4)
        bid_layout.addWidget(self.bid)
        bid_layout.addWidget(self.bid_arrow)
        bid_frame = QFrame()
        bid_frame.setLayout(bid_layout)
        layout.addWidget(bid_frame,1)

        # Ask
        self.ask = QLabel("")
        self.ask.setStyleSheet("color: white; font-size: 22pt;")
        self.ask_arrow = QLabel()
        ask_layout = QHBoxLayout()
        ask_layout.setContentsMargins(0,0,0,0)
        ask_layout.setSpacing(4)
        ask_layout.addWidget(self.ask)
        ask_layout.addWidget(self.ask_arrow)
        ask_frame = QFrame()
        ask_frame.setLayout(ask_layout)
        layout.addWidget(ask_frame,1)

        # Low
        self.low = QLabel("")
        self.low.setStyleSheet("color: white; font-size: 22pt;")
        layout.addWidget(self.low,1)

        # High
        self.high = QLabel("")
        self.high.setStyleSheet("color: white; font-size: 22pt;")
        layout.addWidget(self.high,1)

        # Visual Up/Down arrows (stacked)
        self.arrow_col = QVBoxLayout()
        self.arrow_col.setContentsMargins(0,0,0,0)
        self.arrow_col.setSpacing(0)
        self.up_btn = QPushButton("▲")
        self.down_btn = QPushButton("▼")
        for btn in (self.up_btn, self.down_btn):
            btn.setStyleSheet("color: gray; font-size: 18pt; background: transparent; border: none;")
            btn.setFixedSize(28, 28)
            btn.setCursor(Qt.PointingHandCursor)
        self.arrow_col.addWidget(self.up_btn, alignment=Qt.AlignHCenter)
        self.arrow_col.addWidget(self.down_btn, alignment=Qt.AlignHCenter)
        layout.addLayout(self.arrow_col)
        
        
        # initialize button visibility
        

        # Remove (✖) button
        self.remove_btn = QPushButton("✖")
        self.remove_btn.setStyleSheet("color: red; font-size: 18pt; background: transparent; border: none;")
        self.remove_btn.setCursor(Qt.PointingHandCursor)
        self.remove_btn.clicked.connect(self.remove_self)
        layout.addWidget(self.remove_btn)

        # Add (➕) button
        self.add_btn = QPushButton("➕")
        self.add_btn.setStyleSheet("color: lime; font-size: 18pt; background: transparent; border: none;")
        self.add_btn.setCursor(Qt.PointingHandCursor)
        self.add_btn.clicked.connect(self.start_add)
        layout.addWidget(self.add_btn)

        # Input + dropdown for symbol search (hidden by default)
        self.input = QLineEdit()
        self.input.setStyleSheet("font-size: 18pt;")
        self.input.hide()
        layout.addWidget(self.input, 2)

        self.dropdown = QListWidget()
        self.dropdown.setWindowFlags(Qt.Popup)
        self.dropdown.setFocusPolicy(Qt.NoFocus)
        self.dropdown.hide()

        # connections
        self.input.textChanged.connect(self.update_dropdown)
        self.dropdown.itemClicked.connect(self.select_symbol)

        # arrow connections → ask parent to move row
        self.up_btn.clicked.connect(lambda: self.parent_widget.request_move(self, -1) if self.parent_widget else None)
        self.down_btn.clicked.connect(lambda: self.parent_widget.request_move(self, +1) if self.parent_widget else None)

        # initialize button visibility
        self.update_buttons(show_add=False)

    def update_buttons(self, show_add):
        """Show ✖ only on filled rows. The ➕ button is permanently hidden
        — adding new pairs goes through the Ctrl+Shift+S editor."""
        empty = (self.symbol.text().strip() == "")
        self.remove_btn.setVisible(not empty)
        self.add_btn.setVisible(False)

        # Hide the up/down reorder arrows when row has no symbol.
        self.up_btn.setVisible(not empty)
        self.down_btn.setVisible(not empty)
        
        
    def remove_self(self):
        """Notify the parent first (so it can read the name being removed),
        then clear this row's label/values."""
        name = self.symbol.text().strip()
        self.symbol.setText("")
        self.update_prices("", "", "", "")
        if self.remove_callback:
            self.remove_callback(self, name)

    def start_add(self):
        """Open input and show dropdown near it."""
        self.add_btn.hide()
        self.input.clear()
        self.input.show()
        self.input.setFocus()
        self.update_dropdown()

    def update_dropdown(self):
        """Filter available symbols (from Excel rows) not already used."""
        if not self.parent_widget:
            return
        all_syms = self.parent_widget.get_available_symbols_from_excel()
        used = {b.symbol.text().strip() for b in self.parent_widget.boxes if b.symbol.text().strip()}
        text = self.input.text().upper()
        matches = [s for s in all_syms if (text in s.upper()) and (s not in used)]

        self.dropdown.clear()
        if matches:
            for s in matches:
                QListWidgetItem(s, self.dropdown)
            # position dropdown under input
            pos = self.input.mapToGlobal(self.input.rect().bottomLeft())
            self.dropdown.move(pos)
            # set width to input width
            self.dropdown.setFixedWidth(self.input.width())
            self.dropdown.show()
        else:
            self.dropdown.hide()

    def select_symbol(self, item):
        """Set selected symbol for this row and start live updating."""
        self.symbol.setText(item.text())
        self.input.hide()
        self.dropdown.hide()
        if self.add_callback:
            self.add_callback(self)
        self.update_buttons(show_add=False)

    def update_prices(self, bid, ask, low, high):
        try:
            bid = float(bid)
            if bid > self.last_bid:
                self.bid.setStyleSheet("color: lime; font-size: 22pt; font-weight:bold;")
            elif bid < self.last_bid:
                self.bid.setStyleSheet("color: red; font-size: 22pt; font-weight:bold;")
            self.bid.setText(_fmt(bid))
            self.last_bid = bid
        except:
            self.bid.setText(str(bid))

        try:
            ask = float(ask)
            if ask > self.last_ask:
                self.ask.setStyleSheet("color: lime; font-size: 22pt; font-weight:bold;")
            elif ask < self.last_ask:
                self.ask.setStyleSheet("color: red; font-size: 22pt; font-weight:bold;")
            self.ask.setText(_fmt(ask))
            self.last_ask = ask
        except:
            self.ask.setText(str(ask))

        self.high.setText(_fmt("" if high == "" else high))
        self.low.setText(_fmt(low))

    # update backgroung and toggle mode for pricebox class
    def update_background(self, row_index):
        if self.parent_widget and self.parent_widget.is_darkmode:
            bg_color = "#2f3338" if row_index % 2 == 1 else "#22272b"
        else:
            bg_color = "#f5f4e9" if row_index % 2 == 1 else "#f7f4e9"
        self.setStyleSheet(f"QFrame {{ background-color: {bg_color}; border-radius: 5px;}}")
        
    def apply_theme(self):
        """Update text colors for labels based on current theme."""
        if self.parent_widget and self.parent_widget.is_darkmode:
            # Dark mode
            self.symbol.setStyleSheet("color: white; font-size: 20pt;")
            self.high.setStyleSheet("color: white; font-size: 22pt;")
            self.low.setStyleSheet("color: white; font-size: 22pt;")
            self.up_btn.setStyleSheet("color: gray; font-size: 18pt; background: transparent; border: none;")
            self.down_btn.setStyleSheet("color: gray; font-size: 18pt; background: transparent; border: none;")
        else:
            # Light mode
            self.symbol.setStyleSheet("color: black; font-size: 20pt;")
            self.high.setStyleSheet("color: black; font-size: 22pt;")
            self.low.setStyleSheet("color: black; font-size: 22pt;")
            self.up_btn.setStyleSheet("color: lightgray; font-size: 18pt; background: transparent; border: none;")
            self.down_btn.setStyleSheet("color: lightgray; font-size: 18pt; background: transparent; border: none;")


# -------------------------------
# MT4 DDE Live Source (DDEML, advise / hot-link mode)
# -------------------------------
# Excel uses DDE "advise" subscriptions, which is why it works while pywin32
# Request-mode does not — MT4 only populates symbol values for clients that
# have subscribed via XTYP_ADVSTART. pywin32's `dde` module doesn't expose
# the client-side advise API, so we drop down to the Win32 DDEML APIs via
# ctypes. MT4 then pushes data to us via XTYP_ADVDATA callbacks.
_user32 = ctypes.WinDLL("user32", use_last_error=True)

# DDEML constants
_APPCMD_CLIENTONLY = 0x00000010
_CP_WINUNICODE = 1200
_CF_TEXT = 1

_XCLASS_BOOL = 0x1000
_XCLASS_DATA = 0x2000
_XCLASS_FLAGS = 0x4000
_XCLASS_NOTIFICATION = 0x8000

_XTYP_ADVDATA = 0x0010 | _XCLASS_FLAGS
_XTYP_ADVSTART = 0x0030 | _XCLASS_BOOL
_XTYP_ADVSTOP = 0x0040 | _XCLASS_NOTIFICATION
_XTYP_DISCONNECT = 0x00C0 | _XCLASS_NOTIFICATION

_DDE_FACK = 0x8000
_TIMEOUT_ASYNC = 0xFFFFFFFF

# DDE callback signature: HDDEDATA (UINT, UINT, HCONV, HSZ, HSZ, HDDEDATA, ULONG_PTR, ULONG_PTR)
_DDECALLBACK = WINFUNCTYPE(
    c_void_p, c_uint, c_uint, c_void_p, c_void_p, c_void_p, c_void_p, c_size_t, c_size_t
)

_user32.DdeInitializeW.argtypes = [POINTER(c_ulong), _DDECALLBACK, c_ulong, c_ulong]
_user32.DdeInitializeW.restype = c_uint
_user32.DdeUninitialize.argtypes = [c_ulong]
_user32.DdeUninitialize.restype = wintypes.BOOL
_user32.DdeCreateStringHandleW.argtypes = [c_ulong, ctypes.c_wchar_p, c_int]
_user32.DdeCreateStringHandleW.restype = c_void_p
_user32.DdeFreeStringHandle.argtypes = [c_ulong, c_void_p]
_user32.DdeFreeStringHandle.restype = wintypes.BOOL
_user32.DdeQueryStringW.argtypes = [c_ulong, c_void_p, ctypes.c_wchar_p, c_ulong, c_int]
_user32.DdeQueryStringW.restype = c_ulong
_user32.DdeConnect.argtypes = [c_ulong, c_void_p, c_void_p, c_void_p]
_user32.DdeConnect.restype = c_void_p
_user32.DdeDisconnect.argtypes = [c_void_p]
_user32.DdeDisconnect.restype = wintypes.BOOL
_user32.DdeClientTransaction.argtypes = [
    c_void_p, c_ulong, c_void_p, c_void_p, c_uint, c_uint, c_ulong, POINTER(c_ulong)
]
_user32.DdeClientTransaction.restype = c_void_p
_user32.DdeAccessData.argtypes = [c_void_p, POINTER(c_ulong)]
_user32.DdeAccessData.restype = c_void_p
_user32.DdeUnaccessData.argtypes = [c_void_p]
_user32.DdeUnaccessData.restype = wintypes.BOOL
_user32.DdeGetLastError.argtypes = [c_ulong]
_user32.DdeGetLastError.restype = c_uint


class MT4DDESource:
    """
    Subscribes to MT4's DDE server in advise (hot-link) mode.
    MT4 pushes price updates to us via the DDE callback as ticks arrive.
    Latest values are kept in self.cache; read_rows() snapshots from there.
    """
    TOPICS = ("BID", "ASK", "HIGH", "LOW")

    def __init__(self, pairs):
        """
        pairs: iterable of (display_name, mt4_symbol) tuples. The display name
        is what the UI shows; the MT4 symbol is what we subscribe to over DDE.
        Multiple display names may map to the same MT4 symbol (we de-dupe
        the subscription list).
        """
        self.pairs = [(n.strip(), s.strip()) for n, s in pairs if n and s and n.strip() and s.strip()]
        self.mt4_symbols = list(dict.fromkeys(s for _, s in self.pairs))  # unique, order-preserving
        self.cache = {}  # (topic, mt4_symbol) -> latest value string
        self._first_logged = False

        # Keep a ref to the bound callback so it isn't GC'd while DDE holds the pointer.
        self._callback = _DDECALLBACK(self._dde_callback)

        self._idInst = c_ulong(0)
        rc = _user32.DdeInitializeW(byref(self._idInst), self._callback, _APPCMD_CLIENTONLY, 0)
        if rc != 0:
            raise RuntimeError(f"DdeInitialize failed (code {rc}). Is MT4 running with DDE server enabled?")

        self._service_hsz = self._make_hsz(DDE_SERVICE)
        self._topic_hsz = {t: self._make_hsz(t) for t in self.TOPICS}
        self._item_hsz = {s: self._make_hsz(s) for s in self.mt4_symbols}
        self._conv = {}  # topic -> HCONV

        try:
            for topic in self.TOPICS:
                # Retry: when a previous instance just exited, MT4 may still be
                # cleaning up the prior conversation and reject ours briefly.
                hConv = None
                for _attempt in range(5):
                    hConv = _user32.DdeConnect(
                        self._idInst, self._service_hsz, self._topic_hsz[topic], None
                    )
                    if hConv:
                        break
                    time.sleep(0.4)
                if not hConv:
                    err = _user32.DdeGetLastError(self._idInst.value)
                    raise RuntimeError(
                        f"DdeConnect failed for topic '{topic}' "
                        f"(DDE err 0x{err:04x}). "
                        "Make sure MT4 is running and 'Enable DDE server' is on. "
                        "If a previous run just exited, wait a few seconds and retry."
                    )
                self._conv[topic] = hConv
                # Start an advise loop for each unique MT4 symbol on this topic.
                for sym in self.mt4_symbols:
                    pdwResult = c_ulong(0)
                    _user32.DdeClientTransaction(
                        None, 0, hConv, self._item_hsz[sym],
                        _CF_TEXT, _XTYP_ADVSTART, 1000, byref(pdwResult)
                    )
        except Exception:
            self.close()
            raise

    def _make_hsz(self, s):
        return _user32.DdeCreateStringHandleW(self._idInst.value, s, _CP_WINUNICODE)

    def _hsz_to_str(self, hsz):
        if not hsz:
            return ""
        buf = ctypes.create_unicode_buffer(256)
        n = _user32.DdeQueryStringW(self._idInst.value, hsz, buf, 256, _CP_WINUNICODE)
        return buf.value if n else ""

    def _dde_callback(self, uType, uFmt, hConv, hsz1, hsz2, hData, dwData1, dwData2):
        # XTYP_ADVDATA: server is pushing fresh data for an item we subscribed to.
        if uType == _XTYP_ADVDATA:
            topic = self._hsz_to_str(hsz1)
            item = self._hsz_to_str(hsz2)
            length = c_ulong(0)
            ptr = _user32.DdeAccessData(hData, byref(length))
            if ptr:
                try:
                    raw = ctypes.string_at(ptr, length.value)
                    text = raw.decode("ascii", errors="ignore").rstrip("\x00").strip()
                    self.cache[(topic, item)] = text
                finally:
                    _user32.DdeUnaccessData(hData)
            return _DDE_FACK
        # Other transaction types (DISCONNECT etc.) — ignore.
        return 0

    @staticmethod
    def _clean(v):
        if not v:
            return ""
        if v.upper().replace("\\", "/") in ("N/A", "NA"):
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

    def close(self):
        # DdeUninitialize alone tears down all advise loops, conversations,
        # and string handles for this instance. Doing each step manually
        # during shutdown was racing with MT4 and leaving stale state that
        # blocked the next process's DdeConnect.
        try:
            if self._idInst.value:
                _user32.DdeUninitialize(self._idInst.value)
                self._idInst = c_ulong(0)
                self._conv.clear()
        except Exception:
            pass

    def update_pairs(self, new_pairs):
        """
        Apply a new list of (display_name, mt4_symbol) pairs to a running source.
        Adds advise subscriptions for newly-introduced MT4 symbols and stops
        subscriptions for symbols that are no longer referenced. Cached values
        for dropped symbols are cleared.
        """
        cleaned = [
            (n.strip(), s.strip())
            for n, s in new_pairs
            if n and s and n.strip() and s.strip()
        ]
        new_mt4_symbols = list(dict.fromkeys(s for _, s in cleaned))

        old_set = set(self.mt4_symbols)
        new_set = set(new_mt4_symbols)
        to_add = new_set - old_set
        to_remove = old_set - new_set

        for sym in to_add:
            self._item_hsz[sym] = self._make_hsz(sym)
            for topic in self.TOPICS:
                hConv = self._conv.get(topic)
                if not hConv:
                    continue
                pdwResult = c_ulong(0)
                _user32.DdeClientTransaction(
                    None, 0, hConv, self._item_hsz[sym],
                    _CF_TEXT, _XTYP_ADVSTART, 1000, byref(pdwResult)
                )

        for sym in to_remove:
            item_hsz = self._item_hsz.pop(sym, None)
            if item_hsz:
                for topic in self.TOPICS:
                    hConv = self._conv.get(topic)
                    if not hConv:
                        continue
                    pdwResult = c_ulong(0)
                    _user32.DdeClientTransaction(
                        None, 0, hConv, item_hsz,
                        _CF_TEXT, _XTYP_ADVSTOP, 500, byref(pdwResult)
                    )
                _user32.DdeFreeStringHandle(self._idInst.value, item_hsz)
            for topic in self.TOPICS:
                self.cache.pop((topic, sym), None)

        self.pairs = cleaned
        self.mt4_symbols = new_mt4_symbols
        self._first_logged = False


# -------------------------------
# Symbols Config Dialog (two-column editor)
# -------------------------------
class SymbolsConfigDialog(QDialog):
    """
    Two-column editor for the (display name, MT4 symbol) pairs.
    Used both at first launch (no config) and from the running app
    (Ctrl+Shift+S) to add / rename / remove rows.
    """
    def __init__(self, initial_pairs=None, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Live Price Symbols")
        self.resize(520, 420)
        self.pairs = []

        # Force a light theme regardless of the main app's dark/light mode.
        # Targets only this dialog and its children (selector starts with
        # QDialog so it doesn't leak to the parent window).
        self.setStyleSheet("""
            QDialog { background-color: #f5f5f5; color: #000000; }
            QDialog QLabel { color: #000000; background: transparent; }
            QDialog QTableWidget {
                background-color: #ffffff;
                color: #000000;
                gridline-color: #cccccc;
                selection-background-color: #d4af37;
                selection-color: #000000;
            }
            QDialog QHeaderView::section {
                background-color: #e0e0e0;
                color: #000000;
                padding: 4px;
                border: 1px solid #c0c0c0;
            }
            QDialog QTableWidget QTableCornerButton::section {
                background-color: #e0e0e0;
                border: 1px solid #c0c0c0;
            }
            QDialog QPushButton {
                background-color: #ffffff;
                color: #000000;
                border: 1px solid #b0b0b0;
                padding: 5px 12px;
                border-radius: 3px;
            }
            QDialog QPushButton:hover { background-color: #ececec; }
            QDialog QPushButton:pressed { background-color: #d8d8d8; }
            QDialog QLineEdit {
                background-color: #ffffff;
                color: #000000;
                border: 1px solid #b0b0b0;
            }
        """)

        layout = QVBoxLayout(self)
        info = QLabel(
            "Define each row of the live-price app.\n"
            " • Name — what shows in the Symbol column of the UI (e.g. GOLD).\n"
            " • MT4 Symbol — the exact ticker shown in MT4 Market Watch (e.g. XAUUSD).\n"
            "Multiple names can map to the same MT4 symbol."
        )
        info.setWordWrap(True)
        layout.addWidget(info)

        self.table = QTableWidget(0, 2)
        self.table.setHorizontalHeaderLabels(["Name", "MT4 Symbol"])
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.table.verticalHeader().setVisible(False)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        layout.addWidget(self.table)

        if initial_pairs:
            for n, s in initial_pairs[:MAX_ROWS]:
                self._add_row(n, s)
        else:
            for _ in range(3):
                self._add_row("", "")

        btn_row = QHBoxLayout()
        self.add_btn = QPushButton("➕ Add row")
        self.rm_btn = QPushButton("➖ Remove selected")
        self.add_btn.clicked.connect(lambda: self._add_row("", ""))
        self.rm_btn.clicked.connect(self._remove_selected)
        btn_row.addWidget(self.add_btn)
        btn_row.addWidget(self.rm_btn)
        self.cap_label = QLabel("")
        self.cap_label.setStyleSheet("color: #555; padding-left: 10px;")
        btn_row.addWidget(self.cap_label)
        btn_row.addStretch()
        layout.addLayout(btn_row)

        bb = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        bb.accepted.connect(self.accept)
        bb.rejected.connect(self.reject)
        layout.addWidget(bb)

        self._update_cap_state()

    def _add_row(self, name, sym):
        if self.table.rowCount() >= MAX_ROWS:
            return
        r = self.table.rowCount()
        self.table.insertRow(r)
        self.table.setItem(r, 0, QTableWidgetItem(name))
        self.table.setItem(r, 1, QTableWidgetItem(sym))
        self._update_cap_state()

    def _remove_selected(self):
        rows = sorted({i.row() for i in self.table.selectedIndexes()}, reverse=True)
        for r in rows:
            self.table.removeRow(r)
        if self.table.rowCount() == 0:
            self._add_row("", "")
        self._update_cap_state()

    def _update_cap_state(self):
        # Called from _add_row during the initial-pairs loop, which runs
        # before the button row is built — guard against that.
        if not hasattr(self, "add_btn"):
            return
        n = self.table.rowCount()
        self.add_btn.setEnabled(n < MAX_ROWS)
        if n >= MAX_ROWS:
            self.cap_label.setText(f"Maximum {MAX_ROWS} rows reached")
        else:
            self.cap_label.setText(f"{n} / {MAX_ROWS}")

    def accept(self):
        pairs = []
        seen_names = set()
        for r in range(self.table.rowCount()):
            name_item = self.table.item(r, 0)
            sym_item = self.table.item(r, 1)
            name = (name_item.text() if name_item else "").strip()
            sym = (sym_item.text() if sym_item else "").strip()
            if not sym and not name:
                continue
            if not sym:
                QMessageBox.warning(self, "Error", f"Row {r+1}: MT4 Symbol is required.")
                return
            if not name:
                name = sym
            if name in seen_names:
                QMessageBox.warning(self, "Error", f"Duplicate name '{name}'. Names must be unique.")
                return
            seen_names.add(name)
            pairs.append((name, sym))
        if not pairs:
            QMessageBox.warning(self, "Error", "Please add at least one symbol.")
            return
        if len(pairs) > MAX_ROWS:
            QMessageBox.warning(self, "Error", f"Maximum {MAX_ROWS} symbols allowed.")
            return
        self.pairs = pairs
        super().accept()


# -------------------------------
# Font Delegate for preview
# -------------------------------
class FontDelegate(QStyledItemDelegate):
    
    def paint(self, painter, option, index):
        font_name = index.data()
        painter.save()
        if option.state & QStyle.State_Selected:
            painter.fillRect(option.rect, option.palette.highlight())

        font = QFont(font_name, 12)
        painter.setFont(font)
        painter.setPen(option.palette.text().color())
        painter.drawText(option.rect.adjusted(5, 0, 0, 0), Qt.AlignVCenter, font_name)
        painter.restore()


# -------------------------------
# Font Changer Window
# -------------------------------
class FontChanger(QWidget):
    def __init__(self, main_window):
        super().__init__()
        self.setWindowTitle("Font Changer")
        self.resize(400, 250)
        self.main_window = main_window

        layout = QVBoxLayout()

        self.font_dropdown = QComboBox()
        self.font_dropdown.setItemDelegate(FontDelegate())
        self.populate_fonts()
        layout.addWidget(self.font_dropdown)

        self.apply_btn = QPushButton("Apply Font")
        self.apply_btn.clicked.connect(self.apply_font)
        layout.addWidget(self.apply_btn)

        self.setLayout(layout)

    def populate_fonts(self):
        font_db = QFontDatabase()
        fonts = font_db.families()
        self.font_dropdown.addItems(fonts)

    def apply_font(self):
        font_name = self.font_dropdown.currentText()
        font = QFont(font_name, 10)
        self.main_window.current_font = font
        self.main_window.apply_font_to_widgets()
        QApplication.setFont(font)
        self.close()            
            

# -------------------------------
# Main Window
# -------------------------------
class MainWindow(QWidget):
    def __init__(self, pairs):
        super().__init__()
        self.setWindowTitle("Live Prices")
        self.setStyleSheet("background-color: black;")
        self.pairs = list(pairs)  # [(display_name, mt4_symbol), ...]
        
        #default theme as dark
        self.is_darkmode = True
        
        # default font is arial
        self.current_font = QFont("Arial", 10)  


        main = QVBoxLayout(self)
        main.setContentsMargins(5,5,5,5)
        main.setSpacing(5)

        # Fixed header (kept)
        self.header_frame = QFrame()
        self.header_frame.setStyleSheet("background-color:#111;")
        hl = QHBoxLayout(self.header_frame)
        hl.setContentsMargins(10,8,10,8)
        hl.setSpacing(12)
        headers = ["Symbol","Bid","Ask","Low","High"]
        for i, h in enumerate(headers):
            lbl = QLabel(h)
            if self.is_darkmode:
                
                lbl.setStyleSheet("color: gold; font-size: 16pt; font-weight:bold;")
            else:
                lbl.setStyleSheet("color: black; font-size: 16pt; font-weight:bold;")
            lbl.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
            
            if h.lower() == "symbol":
                self.header_symbol_lbl = lbl
            
            hl.addWidget(lbl, 1)
        spacer = QFrame()
        spacer.setFixedWidth(5)  # space for ↑/↓ and ✖/➕
        hl.addWidget(spacer)
        main.addWidget(self.header_frame)
        

        # Scroll area for rows
        self.scroll = QScrollArea()
        self.scroll.setWidgetResizable(True)
        self.scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.scroll.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        main.addWidget(self.scroll, 1)

        self.rows_container = QWidget()
        self.scroll.setWidget(self.rows_container)
        self.rows_layout = QVBoxLayout(self.rows_container)
        self.rows_layout.setContentsMargins(0,0,0,0)
        self.rows_layout.setSpacing(5)

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

        # Boxes + state. Always create at least MAX_BOXES (default visible
        # rows). If config has more pairs than that, create extras up to
        # MAX_ROWS so all pairs render on first show.
        initial_count = max(MAX_BOXES, min(MAX_ROWS, len(self.pairs)))
        self.boxes = []
        for i in range(initial_count):
            box = PriceBox(
                row_index=i,
                remove_callback=self.on_row_cleared,
                add_callback=self.on_row_added,
                parent_widget=self
            )
            self.rows_layout.addWidget(box)
            self.boxes.append(box)

        self._anim_group = None  # keep reference to animations

        self.initial_fill_done = False
        self.last_rows_dict = {}  # symbol -> (bid, ask, low, high) strings

        self.refresh_once()
        self.timer = QTimer()
        self.timer.timeout.connect(self.refresh_once)
        self.timer.start(REFRESH_INTERVAL_MS)

        self.is_fullscreen = False
        shortcut = QShortcut(QKeySequence("Ctrl+Shift+F1"), self)
        shortcut.activated.connect(self.toggle_fullscreen)

        # Close dropdown/input when clicking elsewhere
        self.installEventFilter(self)
        
        
        #dark mode button
        self.mode_btn = QPushButton("🌓")
        self.mode_btn.setFixedSize(35, 35)
        self.mode_btn.clicked.connect(self.toggle_mode)
        self.mode_btn.setStyleSheet("color: white; font-size: 18pt; border: 1px solid white;")
        hl.addWidget(self.mode_btn)

        
        # Shortcut to open Font Changer
        self.font_shortcut = QShortcut(QKeySequence("Ctrl+Shift+F"), self)
        self.font_shortcut.activated.connect(self.open_font_changer)

        # Shortcut to open the symbols editor (add / rename / remove pairs).
        self.symbols_shortcut = QShortcut(QKeySequence("Ctrl+Shift+S"), self)
        self.symbols_shortcut.activated.connect(self.open_symbols_editor)


        # put it last ,
        self.apply_theme()
         
        

    # --- Symbols editor (Ctrl+Shift+S) ---
    def open_symbols_editor(self):
        """Open the symbols dialog pre-filled with current pairs and apply
        any changes live: subscribe to new MT4 symbols, drop removed ones,
        re-fill the rows in the UI, and persist to config."""
        dlg = SymbolsConfigDialog(initial_pairs=self.pairs, parent=self)
        if dlg.exec_() != QDialog.Accepted:
            return
        new_pairs = dlg.pairs
        try:
            self.source.update_pairs(new_pairs)
        except Exception as e:
            QMessageBox.critical(
                self, "DDE Error",
                f"Failed to update DDE subscriptions:\n\n{e}"
            )
            return
        self.pairs = list(new_pairs)
        self._refill_boxes()
        # Persist immediately so the change survives a crash before normal close.
        try:
            rows = [b.symbol.text().strip() for b in self.boxes]
            save_config(
                pairs=self.pairs, rows=rows,
                font=self.current_font, is_darkmode=self.is_darkmode,
            )
        except Exception as e:
            print(f"warning: could not save {CONFIG_FILE} after symbol edit: {e}")
        # Trigger an immediate refresh so live values populate quickly.
        self.refresh_once()

    def _refill_boxes(self):
        """Reconcile the on-screen boxes with self.pairs:
          - clear boxes whose name is no longer in pairs;
          - put any pair name that isn't shown into the first empty box,
            creating new boxes if we run out;
          - keep the order: existing rows in their current positions, new
            rows appended below."""
        new_names = [n for n, _ in self.pairs]
        new_names_set = set(new_names)

        # Drop names no longer present.
        for box in self.boxes:
            sym = box.symbol.text().strip()
            if sym and sym not in new_names_set:
                box.symbol.setText("")
                box.update_prices("", "", "", "")

        shown = {b.symbol.text().strip() for b in self.boxes if b.symbol.text().strip()}
        needed = [n for n in new_names if n not in shown]

        for name in needed:
            target = next((b for b in self.boxes if not b.symbol.text().strip()), None)
            if target is None:
                target = PriceBox(
                    row_index=len(self.boxes),
                    remove_callback=self.on_row_cleared,
                    add_callback=self.on_row_added,
                    parent_widget=self,
                )
                target.symbol.setFixedWidth(int(self.width() * 0.3))
                self.boxes.append(target)
                self.rows_layout.addWidget(target)
                target.update_background(len(self.boxes) - 1)
                target.apply_theme()
            target.symbol.setText(name)

        # If the dialog reduced the pair count, trim trailing empty boxes
        # back down to max(MAX_BOXES, len(pairs)) so heights expand again.
        desired = max(MAX_BOXES, len(self.pairs))
        while len(self.boxes) > desired:
            target = next((b for b in reversed(self.boxes) if not b.symbol.text().strip()), None)
            if target is None:
                break
            self.boxes.remove(target)
            self.rows_layout.removeWidget(target)
            target.deleteLater()

        self.reorder_boxes()
        self.update_add_buttons()
        self._apply_row_heights()

    # --- New helpers for +/search ---
    def get_available_symbols_from_excel(self):
        """Return list of symbols present in Excel (from last read)."""
        return list(self.last_rows_dict.keys())


    def resizeEvent(self, event):
        super().resizeEvent(event)
        for box in self.boxes:
            # Make the symbol label 30% of the MainWindow width
            box.symbol.setFixedWidth(int(self.width() * 0.3))
        if self.header_symbol_lbl:
            self.header_symbol_lbl.setFixedWidth(int(self.width() * 0.3))
        self._apply_row_heights()

    def _apply_row_heights(self):
        """When the visible row count exceeds MAX_BOXES (default 12), shrink
        every row so all rows fit in the scroll viewport without scrolling.
        Up to MAX_ROWS (20) rows are supported. At or below MAX_BOXES, rows
        use their natural height."""
        n = len(self.boxes)
        if n <= 0:
            return
        if n <= MAX_BOXES:
            # Restore natural sizing.
            for box in self.boxes:
                box.setMinimumHeight(0)
                box.setMaximumHeight(16777215)
            return
        spacing = self.rows_layout.spacing()
        margins = self.rows_layout.contentsMargins()
        avail = (
            self.scroll.viewport().height()
            - margins.top() - margins.bottom()
            - max(0, n - 1) * spacing
        )
        if avail <= 0:
            return
        h = max(24, avail // n)
        for box in self.boxes:
            box.setFixedHeight(h)
                

    def on_row_cleared(self, _box, name=""):
        """User clicked ✖. Remove the pair from the master list and DDE
        subscriptions. Drop the box itself only when we're above MAX_BOXES,
        so the remaining rows expand. At or below MAX_BOXES (default 12)
        the box stays as an empty placeholder."""
        _box.input.hide()
        _box.dropdown.hide()

        if name:
            new_pairs = [(n, s) for (n, s) in self.pairs if n != name]
            if len(new_pairs) != len(self.pairs):
                self.pairs = new_pairs
                try:
                    self.source.update_pairs(new_pairs)
                except Exception as e:
                    print(f"warning: source.update_pairs failed: {e}")
                self._save_current_config_safe()

        # Keep at least MAX_BOXES rows visible. Above that, dropping the
        # cleared box lets the rest expand back toward natural height.
        desired = max(MAX_BOXES, len(self.pairs))
        if len(self.boxes) > desired:
            try:
                self.boxes.remove(_box)
                self.rows_layout.removeWidget(_box)
                _box.deleteLater()
            except (ValueError, RuntimeError):
                pass

        self.reorder_boxes()
        self.update_add_buttons()
        self._apply_row_heights()

    def _save_current_config_safe(self):
        try:
            rows = [b.symbol.text().strip() for b in self.boxes]
            save_config(
                pairs=self.pairs, rows=rows,
                font=self.current_font, is_darkmode=self.is_darkmode,
            )
        except Exception as e:
            print(f"warning: could not save {CONFIG_FILE}: {e}")

    def on_row_added(self, _box):
        """Callback when a symbol is chosen from the dropdown for a box."""
        _box.input.hide()
        _box.dropdown.hide()
        self.reorder_boxes()
        self.update_add_buttons()

    def reorder_boxes(self):
        """
        Keep current relative order of active rows; move empty rows below them.
        """
        active = [b for b in self.boxes if b.symbol.text().strip() != ""]
        empty = [b for b in self.boxes if b.symbol.text().strip() == ""]
        ordered = active + empty

        # Remove & re-add into layout in the new order
        for b in self.boxes:
            try:
                self.rows_layout.removeWidget(b)
            except Exception:
                pass
        for i, b in enumerate(ordered):
            self.rows_layout.insertWidget(i, b)
            b.show()
        self.boxes = ordered

    def apply_theme(self):
        if self.is_darkmode:
            self.setStyleSheet("background-color: black;")
            self.header_frame.setStyleSheet("background-color: #111;")
            header_color = "gold"
            
        else :
            self.setStyleSheet("background-color: white;")
            self.header_frame.setStyleSheet("background-color: white;")
            header_color = "black"
            
        for lbl in self.header_frame.findChildren(QLabel):
            lbl.setStyleSheet(f"color: {header_color}; font-weight: bold; font-size: 18pt")

            
            
            
    def update_background(self, row_index):
        if self.parent_widget and self.parent_widget.is_darkmode:
            bg_color = "#2f3338" if row_index % 2 == 1 else "#22272b"
        else:
            bg_color = "#f5f4e9" if row_index % 2 == 1 else "#fcf6dc"
        self.setStyleSheet(f"QFrame {{ background-color: {bg_color}; border-radius: 5px;}}")
    
    def toggle_mode(self):
        self.is_darkmode = not self.is_darkmode
        self.apply_theme()
        for i, box in enumerate(self.boxes):
            box.update_background(i)
            box.apply_theme()
    
    
    def update_add_buttons(self):
        """Refresh per-row button visibility (hide ✖ on empty rows). The ➕
        flow was removed; new pairs are added via Ctrl+Shift+S."""
        for b in self.boxes:
            b.update_buttons(show_add=False)

    # --- Smooth visual row swap on arrow click ---
    def request_move(self, box, direction):
        """
        Smoothly swap 'box' with its neighbor using ghost overlays inside
        the scroll viewport, then reinsert widgets in the new order.
        """
        try:
            idx = self.boxes.index(box)
        except ValueError:
            return

        new_idx = idx + direction
        if not (0 <= new_idx < len(self.boxes)):
            return

        other = self.boxes[new_idx]

        viewport = self.scroll.viewport()

        # map widget top-left to viewport coordinates for correct animation when scrolled
        p1 = box.mapTo(viewport, QPoint(0, 0))
        p2 = other.mapTo(viewport, QPoint(0, 0))
        r1 = QRect(p1, box.size())
        r2 = QRect(p2, other.size())

        # Ghost overlays (screenshots) so we can animate freely
        pm1 = box.grab()
        pm2 = other.grab()

        ghost1 = QLabel(viewport)
        ghost1.setPixmap(pm1)
        ghost1.setGeometry(r1)
        ghost1.setAttribute(Qt.WA_TransparentForMouseEvents)
        ghost1.raise_()
        ghost1.show()

        ghost2 = QLabel(viewport)
        ghost2.setPixmap(pm2)
        ghost2.setGeometry(r2)
        ghost2.setAttribute(Qt.WA_TransparentForMouseEvents)
        ghost2.raise_()
        ghost2.show()

        # Hide the real widgets while animating
        box.hide()
        other.hide()

        # Animations (slide)
        a1 = QPropertyAnimation(ghost1, b"pos", self)
        a1.setDuration(220)
        a1.setStartValue(r1.topLeft())
        a1.setEndValue(r2.topLeft())
        a1.setEasingCurve(QEasingCurve.OutCubic)

        a2 = QPropertyAnimation(ghost2, b"pos", self)
        a2.setDuration(220)
        a2.setStartValue(r2.topLeft())
        a2.setEndValue(r1.topLeft())
        a2.setEasingCurve(QEasingCurve.OutCubic)

        group = QParallelAnimationGroup(self)
        group.addAnimation(a1)
        group.addAnimation(a2)

        def finalize():
            # Swap in the list
            self.boxes[idx], self.boxes[new_idx] = self.boxes[new_idx], self.boxes[idx]

            # Re-add all boxes in new order to preserve positions
            for b in self.boxes:
                try:
                    self.rows_layout.removeWidget(b)
                except Exception:
                    pass
            for i, b in enumerate(self.boxes):
                self.rows_layout.insertWidget(i, b)
                b.show()

            ghost1.deleteLater()
            ghost2.deleteLater()

            # Keep + placement correct and ensure moved row is visible
            self.update_add_buttons()
            self.scroll.ensureWidgetVisible(box)

        group.finished.connect(finalize)
        self._anim_group = group  # keep reference alive
        group.start()

    # --- Core refresh logic ---
    def refresh_once(self):
        try:
            rows = self.source.read_rows()
        except Exception as e:
            print("Read error:", e)
            traceback.print_exc()
            rows = []

        # update last rows dict for search & updates
        self.last_rows_dict = {sym: (bid, ask, low, high) for sym, bid, ask, low, high in rows}

        # initial fill: set symbols sequentially once
        if not self.initial_fill_done:
            for i, box in enumerate(self.boxes):
                if i < len(rows):
                    sym, bid, ask, low, high = rows[i]
                    box.symbol.setText(str(sym))
                    box.update_prices(bid, ask, low, high)
                else:
                    box.symbol.setText("")
                    box.update_prices("", "", "", "")
            self.initial_fill_done = True
            self.update_add_buttons()
            return

        # after initial fill: only update boxes that have a symbol
        for box in self.boxes:
            sym = box.symbol.text().strip()
            if sym and sym in self.last_rows_dict:
                bid, ask, low, high = self.last_rows_dict[sym]
                box.update_prices(bid, ask, low, high)

        self.update_add_buttons()

    def closeEvent(self, event):
        try: self.timer.stop()
        except Exception: pass
        try: self.source.close()
        except Exception: pass
        super().closeEvent(event)

    def toggle_fullscreen(self):
        if not self.is_fullscreen:
            self.showFullScreen()
            self.is_fullscreen = True
        else:
            self.showNormal()
            self.is_fullscreen = False

    # Close dropdown/input when clicking outside
    def eventFilter(self, obj, event):
        if event.type() == QEvent.MouseButtonPress:
            gp = event.globalPos()
            for box in self.boxes:
                if box.dropdown.isVisible() or box.input.isVisible():
                    inside_input = False
                    if box.input.isVisible():
                        local_pt = box.input.mapFromGlobal(gp)
                        if box.input.rect().contains(local_pt):
                            inside_input = True

                    inside_dropdown = False
                    if box.dropdown.isVisible():
                        if box.dropdown.geometry().contains(gp):
                            inside_dropdown = True

                    if not inside_input and not inside_dropdown:
                        box.dropdown.hide()
                        box.input.hide()
                        self.update_add_buttons()
        return super().eventFilter(obj, event)
    
    

    def open_font_changer(self):
        self.font_window = FontChanger(self)
        self.font_window.show()
        
    def apply_font(self):
        font_name = self.font_dropdown.currentText()
        font = QFont(font_name, 10)
        QApplication.setFont(font)
        self.close()
        
    def apply_font_to_widgets(self):
        for box in self.boxes:
            for lbl in [box.symbol, box.bid, box.ask, box.high, box.low]:
                lbl.setFont(self.current_font)
        # update header font 
        if hasattr(self, 'header_frame'):
            for lbl in self.header_frame.findChildren(QLabel):
                lbl.setFont(self.current_font)
                
    
    
    def closeEvent(self, event):
        try: self.timer.stop()
        except Exception: pass
        try: self.source.close()
        except Exception: pass

        # Save current config
        try:
            rows = [b.symbol.text().strip() for b in self.boxes]
            save_config(
                pairs=self.pairs,
                rows=rows,
                font=self.current_font,
                is_darkmode=self.is_darkmode,
            )
            print(f"saved {CONFIG_FILE}: {len(self.pairs)} symbol(s), {len([r for r in rows if r])} active row(s)")
        except Exception as e:
            print(f"warning: could not save {CONFIG_FILE} on close: {e}")

        super().closeEvent(event)



# -------------------------------
# Entry Point
# -------------------------------
if __name__ == "__main__":
    app = QApplication(sys.argv)
    config_data = load_config()

    if config_data and config_data.get("PAIRS"):
        pairs = config_data["PAIRS"]
        saved_rows = config_data.get("ROWS", [])
        is_darkmode = config_data.get("IS_DARKMODE", True)
        current_font = config_data.get("FONT", QFont("Arial", 10))
    else:
        dlg = SymbolsConfigDialog()
        if dlg.exec_() != QDialog.Accepted:
            sys.exit()
        pairs = dlg.pairs
        saved_rows = []
        is_darkmode = True
        current_font = QFont("Arial", 10)

    # Persist immediately so the symbol list survives a later DDE failure.
    # closeEvent saves again on graceful shutdown (rows order, theme, font).
    try:
        save_config(pairs=pairs, rows=saved_rows, font=current_font, is_darkmode=is_darkmode)
    except Exception as _e:
        print(f"warning: could not write {CONFIG_FILE}: {_e}")

    # Initialize main window
    window = MainWindow(pairs)
    window.is_darkmode = is_darkmode
    window.current_font = current_font
    window.apply_theme()
    window.apply_font_to_widgets()

    # Restore saved rows if any
    for i, box in enumerate(window.boxes):
        if i < len(saved_rows):
            box.symbol.setText(saved_rows[i])
    window.update_add_buttons()

    window.showMaximized()
    sys.exit(app.exec_())

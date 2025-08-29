import sys
import os
import traceback
import xlwings as xw
import pandas as pd
from PyQt5.QtWidgets import (
    QApplication, QWidget, QLabel, QLineEdit, QPushButton,
    QFileDialog, QVBoxLayout, QHBoxLayout, QMessageBox, QComboBox,
    QGridLayout, QGraphicsDropShadowEffect, QShortcut, QFrame,
    QListWidget, QListWidgetItem, QScrollArea, QStyledItemDelegate, QStyle, QShortcut
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
EXCLUDED = {""}
REFRESH_INTERVAL_MS = 100
MAX_BOXES = 12  # initial number of rows to create (list can grow)
# -------------------------------
# Config file handling
# -------------------------------
def save_config(file_path, sheet_name, rows=None, font=None, is_darkmode=True):
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        f.write(f"FILE_PATH={file_path}\n")
        f.write(f"SHEET_NAME={sheet_name}\n")
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
    if "ROWS" in config:
        config["ROWS"] = config["ROWS"].split(",")
    else:
        config["ROWS"] = []
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

        # High
        self.high = QLabel("")
        self.high.setStyleSheet("color: white; font-size: 22pt;")
        layout.addWidget(self.high,1)

        # Low
        self.low = QLabel("")
        self.low.setStyleSheet("color: white; font-size: 22pt;")
        layout.addWidget(self.low,1)

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
        """Only show + on first empty row; show ✖ only when symbol exists."""
        empty = (self.symbol.text().strip() == "")
        self.remove_btn.setVisible(not empty)
        self.add_btn.setVisible(empty and show_add)

        # 👇 NEW: Hide arrow buttons when row has no symbol
        self.up_btn.setVisible(not empty)
        self.down_btn.setVisible(not empty)
        
        
    def remove_self(self):
        """Clear this row but keep it in place."""
        self.symbol.setText("")
        self.update_prices("", "", "", "")
        if self.remove_callback:
            self.remove_callback(self)

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
# Excel Live Source
# -------------------------------
class ExcelLiveSource:
    def __init__(self, path, sheet_name):
        self.app = None
        self.wb = None
        self.sheet = None
        self.path = path
        self.sheet_name = sheet_name
        self._open()

    def _open(self):
        self.app = xw.App(visible=True)
        self.wb = self.app.books.open(self.path)
        self.sheet = self.wb.sheets[self.sheet_name]

    def read_rows(self):
        # expected range: B2:F500 -> [Symbol, Bid, Ask, Low, High]
        values = self.sheet.range("B2:F500").value
        rows = []
        if not values:
            return rows
        for row in values:
            if not row:
                continue
            symbol = row[0]
            if not symbol or (isinstance(symbol, str) and symbol.strip().upper() in EXCLUDED):
                continue
            bid = row[1] if len(row) > 1 else ""
            ask = row[2] if len(row) > 2 else ""
            low = row[3] if len(row) > 3 else ""
            high = row[4] if len(row) > 4 else ""
            
            rows.append((str(symbol), _fmt(bid), _fmt(ask), _fmt(low), _fmt(high)))

        return rows

    def close(self):
        try:
            if self.wb: self.wb.close()
        finally:
            if self.app: self.app.quit()
            
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
    def __init__(self, file_path, sheet_name):
        super().__init__()
        self.setWindowTitle("Live Prices")
        self.setStyleSheet("background-color: black;")
        
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
            self.source = ExcelLiveSource(file_path, sheet_name)
        except Exception as e:
            QMessageBox.critical(self, "Excel Error", f"Failed to open Excel file/sheet.\n\n{e}")
            raise

        # Boxes + state
        self.boxes = []
        for i in range(MAX_BOXES):
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
        
        
        # put it last , 
        self.apply_theme()   
         
        

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
                

    def on_row_cleared(self, _box):
        """Callback when a PriceBox clears itself (user clicked ✖)."""
        _box.symbol.setText("")
        _box.update_prices("", "", "", "")
        _box.input.hide()
        _box.dropdown.hide()
        self.reorder_boxes()
        self.update_add_buttons()

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
        """
        Show ➕ only on the first empty row; ensure there's always ONE empty row at
        the bottom when there are still unused symbols in Excel.
        """
        # calculate remaining symbols
        used = {b.symbol.text().strip() for b in self.boxes if b.symbol.text().strip()}
        all_syms = set(self.get_available_symbols_from_excel())
        remaining = [s for s in all_syms if s not in used]

        empty_boxes = [b for b in self.boxes if not b.symbol.text().strip()]
        if remaining and not empty_boxes:
            # create a new empty row at the bottom
            b = PriceBox(
                symbol="",
                row_index=len(self.boxes),
                remove_callback=self.on_row_cleared,
                add_callback=self.on_row_added,
                parent_widget=self
            )
            b.symbol.setFixedWidth(int(self.width() * 0.3))
            self.boxes.append(b)
            self.rows_layout.addWidget(b)
            empty_boxes = [b]

        # show ➕ only on the first empty row
        first_empty = empty_boxes[0] if empty_boxes else None
        for b in self.boxes:
            is_empty = (b.symbol.text().strip() == "")
            b.update_buttons(show_add=(b is first_empty))

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
        rows = [b.symbol.text().strip() for b in self.boxes]
        save_config(
            self.source.path,
            self.source.sheet_name,
            rows=rows,
            font=self.current_font,
            is_darkmode=self.is_darkmode
        )

        super().closeEvent(event)



# -------------------------------
# Entry Point
# -------------------------------
if __name__ == "__main__":
    app = QApplication(sys.argv)
    config_data = load_config()

    if config_data:
        file_path = config_data.get("FILE_PATH", "")
        sheet_name = config_data.get("SHEET_NAME", "")
        saved_rows = config_data.get("ROWS", [])
        is_darkmode = config_data.get("IS_DARKMODE", True)
        current_font = config_data.get("FONT", QFont("Arial", 10))
    else:
        # Ask user to select Excel file & sheet
        from PyQt5.QtWidgets import QDialog, QDialogButtonBox, QFormLayout, QComboBox

        class ExcelConfigDialog(QDialog):
            def __init__(self):
                super().__init__()
                self.setWindowTitle("Select Excel File & Sheet")
                self.resize(400, 150)
                self.selected_config = None

                layout = QFormLayout(self)

                # File input + browse
                self.file_input = QLineEdit()
                self.file_btn = QPushButton("Browse")
                hb_file = QHBoxLayout()
                hb_file.addWidget(self.file_input)
                hb_file.addWidget(self.file_btn)
                layout.addRow("Excel File:", hb_file)

                # Sheet input + dropdown
                self.sheet_input = QLineEdit()
                self.sheet_dropdown = QComboBox()
                hb_sheet = QHBoxLayout()
                hb_sheet.addWidget(self.sheet_input)
                hb_sheet.addWidget(self.sheet_dropdown)
                layout.addRow("Sheet name:", hb_sheet)

                # OK/Cancel buttons
                bb = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
                bb.accepted.connect(self.accept)
                bb.rejected.connect(self.reject)
                layout.addRow(bb)

                self.file_btn.clicked.connect(self.browse_file)
                self.sheet_dropdown.currentTextChanged.connect(self.update_sheet_input)

            def browse_file(self):
                file_path, _ = QFileDialog.getOpenFileName(
                    self, "Select Excel File", "", "Excel Files (*.xlsx *.xls)"
                )
                if file_path:
                    self.file_input.setText(file_path)
                    try:
                        xl = pd.ExcelFile(file_path)
                        self.sheet_dropdown.clear()
                        self.sheet_dropdown.addItems(xl.sheet_names)
                        if xl.sheet_names:
                            self.sheet_input.setText(xl.sheet_names[0])
                    except Exception as e:
                        QMessageBox.critical(self, "Error", f"Failed to read Excel: {e}")

            def update_sheet_input(self, sheet_name):
                self.sheet_input.setText(sheet_name)

            def accept(self):
                file_path = self.file_input.text().strip()
                sheet_name = self.sheet_input.text().strip()
                if not os.path.isfile(file_path):
                    QMessageBox.warning(self, "Error", "Invalid file path.")
                    return
                if not sheet_name:
                    QMessageBox.warning(self, "Error", "Sheet name cannot be empty.")
                    return
                save_config(file_path, sheet_name)
                self.selected_config = {"FILE_PATH": file_path, "SHEET_NAME": sheet_name}
                super().accept()

        dlg = ExcelConfigDialog()
        if dlg.exec_() != QDialog.Accepted:
            sys.exit()
        config_data = dlg.selected_config
        file_path = config_data["FILE_PATH"]
        sheet_name = config_data["SHEET_NAME"]
        saved_rows = []
        is_darkmode = True
        current_font = QFont("Arial", 10)

    # Initialize main window
    window = MainWindow(file_path, sheet_name)
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

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
MAX_BOXES = 12


def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and PyInstaller """
    if hasattr(sys, "_MEIPASS"):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)


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
# Price Box
# -------------------------------
class PriceBox(QFrame):
    def __init__(self, symbol="", row_index=0, remove_callback=None, add_callback=None, parent_widget=None, header_symbol_lbl=None, header_frame=None):
        super().__init__()
        self.last_bid = 0.0
        self.last_ask = 0.0
        self.bid_color = "white"  # Track current bid color state
        self.ask_color = "white"  # Track current ask color state
        self.remove_callback = remove_callback
        self.add_callback = add_callback
        self.parent_widget = parent_widget
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
        self.symbol.setStyleSheet("color: white; font-size: 20pt;")
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

        # arrow connections
        self.up_btn.clicked.connect(lambda: self.parent_widget.request_move(self, -1) if self.parent_widget else None)
        self.down_btn.clicked.connect(lambda: self.parent_widget.request_move(self, +1) if self.parent_widget else None)

        self.update_buttons(show_add=False)

    def update_buttons(self, show_add):
        """Only show + on first empty row; show ✖ only when symbol exists."""
        empty = (self.symbol.text().strip() == "")
        self.remove_btn.setVisible(not empty)
        self.add_btn.setVisible(empty and show_add)
        self.up_btn.setVisible(not empty)
        self.down_btn.setVisible(not empty)
        
    def remove_self(self):
        """Clear this row but keep it in place."""
        self.symbol.setText("")
        self.update_prices("", "", "", "")
        self.last_bid = 0.0
        self.last_ask = 0.0
        self.bid_color = "white"
        self.ask_color = "white"
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
            pos = self.input.mapToGlobal(self.input.rect().bottomLeft())
            self.dropdown.move(pos)
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
        """Update prices with persistent color logic."""
        # Get dynamic font size from parent window
        if self.parent_widget:
            width = self.parent_widget.width()
            base_width = 1920
            scale_factor = max(min(width / base_width, 1.5), 0.4)
            price_size = int(22 * scale_factor)
        else:
            price_size = 22
        
        # Update BID with persistent color
        try:
            new_bid = float(bid)
            
            # Determine new color based on price change
            if new_bid > self.last_bid:
                self.bid_color = "lime"  # Price went up, stay green
            elif new_bid < self.last_bid:
                self.bid_color = "red"   # Price went down, stay red
            # If equal, keep current color
            
            # Apply the color
            self.bid.setStyleSheet(f"color: {self.bid_color}; font-size: {price_size}pt; font-weight:bold;")
            self.bid.setText(_fmt(new_bid))
            self.last_bid = new_bid
        except:
            # Reset to default on error
            default_color = "white" if (self.parent_widget and self.parent_widget.is_darkmode) else "black"
            self.bid.setStyleSheet(f"color: {default_color}; font-size: {price_size}pt;")
            self.bid.setText(str(bid))
            self.bid_color = default_color

        # Update ASK with persistent color
        try:
            new_ask = float(ask)
            
            # Determine new color based on price change
            if new_ask > self.last_ask:
                self.ask_color = "lime"  # Price went up, stay green
            elif new_ask < self.last_ask:
                self.ask_color = "red"   # Price went down, stay red
            # If equal, keep current color
            
            # Apply the color
            self.ask.setStyleSheet(f"color: {self.ask_color}; font-size: {price_size}pt; font-weight:bold;")
            self.ask.setText(_fmt(new_ask))
            self.last_ask = new_ask
        except:
            # Reset to default on error
            default_color = "white" if (self.parent_widget and self.parent_widget.is_darkmode) else "black"
            self.ask.setStyleSheet(f"color: {default_color}; font-size: {price_size}pt;")
            self.ask.setText(str(ask))
            self.ask_color = default_color

        # Update High and Low (no color change logic needed)
        self.high.setText(_fmt("" if high == "" else high))
        self.low.setText(_fmt(low))

    def update_background(self, row_index):
        if self.parent_widget and self.parent_widget.is_darkmode:
            bg_color = "#2a2d30" if row_index % 2 == 1 else "#1B1E21"
        else:
            bg_color = "#f5f4e9" if row_index % 2 == 1 else "#f7f4e9"
        self.setStyleSheet(f"QFrame {{ background-color: {bg_color}; border-radius: 5px;}}")
        
    def apply_theme(self):
        """Update text colors for labels based on current theme."""
        if self.parent_widget:
            width = self.parent_widget.width()
            base_width = 1920
            scale_factor = max(min(width / base_width, 1.5), 0.4)
            symbol_size = int(20 * scale_factor)
            price_size = int(22 * scale_factor)
            button_size = int(18 * scale_factor)
        else:
            symbol_size = 20
            price_size = 22
            button_size = 18
            
        if self.parent_widget and self.parent_widget.is_darkmode:
            self.symbol.setStyleSheet(f"color: white; font-size: {symbol_size}pt;")
            self.high.setStyleSheet(f"color: white; font-size: {price_size}pt;")
            self.low.setStyleSheet(f"color: white; font-size: {price_size}pt;")
            self.up_btn.setStyleSheet(f"color: gray; font-size: {button_size}pt; background: transparent; border: none;")
            self.down_btn.setStyleSheet(f"color: gray; font-size: {button_size}pt; background: transparent; border: none;")
            
            # Update bid/ask but preserve their color state
            self.bid.setStyleSheet(f"color: {self.bid_color}; font-size: {price_size}pt; font-weight:bold;")
            self.ask.setStyleSheet(f"color: {self.ask_color}; font-size: {price_size}pt; font-weight:bold;")
        else:
            self.symbol.setStyleSheet(f"color: black; font-size: {symbol_size}pt;")
            self.high.setStyleSheet(f"color: black; font-size: {price_size}pt;")
            self.low.setStyleSheet(f"color: black; font-size: {price_size}pt;")
            self.up_btn.setStyleSheet(f"color: lightgray; font-size: {button_size}pt; background: transparent; border: none;")
            self.down_btn.setStyleSheet(f"color: lightgray; font-size: {button_size}pt; background: transparent; border: none;")
            
            # Update bid/ask but preserve their color state
            self.bid.setStyleSheet(f"color: {self.bid_color}; font-size: {price_size}pt; font-weight:bold;")
            self.ask.setStyleSheet(f"color: {self.ask_color}; font-size: {price_size}pt; font-weight:bold;")


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
        
        self.is_darkmode = True
        self.current_font = QFont("Arial", 10)

        main = QVBoxLayout(self)
        main.setContentsMargins(5,5,5,5)
        main.setSpacing(5)
        
        # -------------------------------
        # Logo at top
        # -------------------------------
        self.logo_label = QLabel()
        self.logo_label.setAlignment(Qt.AlignCenter)
        self.logo_dark = QPixmap(resource_path("./LANDMARK-MARKETS-WHITE-LOGO-2-scaled.png"))
        self.logo_light = QPixmap(resource_path("./LANDMARK-MARKETS-BLACK-LOGO-2.png"))
        
        # helper
        def update_logo_pixmap():
            current = self.logo_dark if self.is_darkmode else self.logo_light
            if current.isNull():
                self.logo_label.clear()
                return
            target_width = max(200, self.width() -40)
            max_logo_height = 150
            scaled = current.scaled(target_width, max_logo_height, Qt.KeepAspectRatio, Qt.SmoothTransformation)
            self.logo_label.setPixmap(scaled)
            
        self.update_logo_pixmap = update_logo_pixmap
        
        self.update_logo_pixmap()
        main.addWidget(self.logo_label)

        # Fixed header
        self.header_frame = QFrame()
        self.header_frame.setStyleSheet("background-color:#111;")
        hl = QHBoxLayout(self.header_frame)
        hl.setContentsMargins(10,8,10,8)
        hl.setSpacing(12)
        headers = ["Symbol","Bid","Ask","Low","High"]
        for i, h in enumerate(headers):
            lbl = QLabel(h)
            lbl.setStyleSheet("color: gold; font-size: 16pt; font-weight:bold;")
            lbl.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
            
            if h.lower() == "symbol":
                self.header_symbol_lbl = lbl
            
            hl.addWidget(lbl, 1)
        spacer = QFrame()
        spacer.setFixedWidth(5)
        hl.addWidget(spacer)
        
        # Dark mode button
        self.mode_btn = QPushButton("🌓")
        self.mode_btn.setFixedSize(35, 35)
        self.mode_btn.clicked.connect(self.toggle_mode)
        self.mode_btn.setStyleSheet("color: white; font-size: 18pt; border: 1px solid white;")
        hl.addWidget(self.mode_btn)
        
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

        self._anim_group = None

        self.initial_fill_done = False
        self.last_rows_dict = {}

        self.refresh_once()
        self.timer = QTimer()
        self.timer.timeout.connect(self.refresh_once)
        self.timer.start(REFRESH_INTERVAL_MS)

        self.is_fullscreen = False
        shortcut = QShortcut(QKeySequence("Ctrl+Shift+F1"), self)
        shortcut.activated.connect(self.toggle_fullscreen)

        self.installEventFilter(self)
        
        # Shortcut to open Font Changer
        self.font_shortcut = QShortcut(QKeySequence("Ctrl+Shift+F"), self)
        self.font_shortcut.activated.connect(self.open_font_changer)
        
        self.apply_theme()

    def get_available_symbols_from_excel(self):
        """Return list of symbols present in Excel (from last read)."""
        return list(self.last_rows_dict.keys())

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self.apply_dynamic_scaling()
        self.fill_extra_rows_if_space()
        
        # Scale logo to window width
        if hasattr(self, "update_logo_pixmap"):
            self.update_logo_pixmap()
    
    def apply_dynamic_scaling(self):
        width = self.width()
        height = self.height()
        
        base_width = 1920
        base_height = 1080
        
        scale_w = width / base_width
        scale_h = height / base_height
        scale_factor = min(scale_w, scale_h)
        
        scale_factor = max(scale_factor, 0.4)

        # Boost vertical size a bit in portrait
        if height > width:
            scale_factor *= 1.2  # increase 20% in portrait

        symbol_size = int(20 * scale_factor)
        price_size = int(22 * scale_factor)
        header_size = int(18 * scale_factor)
        button_size = int(18 * scale_factor)
        
        symbol_width = int(width * 0.3)
        
        if self.header_symbol_lbl:
            self.header_symbol_lbl.setFixedWidth(symbol_width)
        
        for lbl in self.header_frame.findChildren(QLabel):
            current_style = lbl.styleSheet()
            color = "gold" if "gold" in current_style else "black"
            lbl.setStyleSheet(f"color: {color}; font-weight: bold; font-size: {header_size}pt")
        
        for box in self.boxes:
            box.symbol.setFixedWidth(symbol_width)
            
            symbol_color = "white" if "white" in box.symbol.styleSheet() else "black"
            box.symbol.setStyleSheet(f"color: {symbol_color}; font-size: {symbol_size}pt;")
            
            # Bid/Ask preserve color
            box.bid.setStyleSheet(f"color: {box.bid_color}; font-size: {price_size}pt; font-weight: bold;")
            box.ask.setStyleSheet(f"color: {box.ask_color}; font-size: {price_size}pt; font-weight: bold;")
            
            for lbl in [box.high, box.low]:
                price_color = "white" if "white" in lbl.styleSheet() else "black"
                lbl.setStyleSheet(f"color: {price_color}; font-size: {price_size}pt;")
            
            button_pixel_size = int(28 * scale_factor)
            for btn in [box.up_btn, box.down_btn]:
                btn.setFixedSize(button_pixel_size, button_pixel_size)
                btn_color = "gray" if "gray" in btn.styleSheet() else "lightgray"
                btn.setStyleSheet(f"color: {btn_color}; font-size: {button_size}pt; background: transparent; border: none;")
            
            box.remove_btn.setStyleSheet(f"color: red; font-size: {button_size}pt; background: transparent; border: none;")
            box.add_btn.setStyleSheet(f"color: lime; font-size: {button_size}pt; background: transparent; border: none;")
            box.input.setStyleSheet(f"font-size: {int(18 * scale_factor)}pt;")
        
        mode_btn_size = int(35 * scale_factor)
        self.mode_btn.setFixedSize(mode_btn_size, mode_btn_size)
        self.mode_btn.setStyleSheet(f"color: white; font-size: {button_size}pt; border: 1px solid white;")


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
        """Keep current relative order of active rows; move empty rows below them."""
        active = [b for b in self.boxes if b.symbol.text().strip() != ""]
        empty = [b for b in self.boxes if b.symbol.text().strip() == ""]
        ordered = active + empty

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
        else:
            self.setStyleSheet("background-color: white;")
            self.header_frame.setStyleSheet("background-color: white;")
            header_color = "black"
            
        for lbl in self.header_frame.findChildren(QLabel):
            lbl.setStyleSheet(f"color: {header_color}; font-weight: bold; font-size: 18pt")
    
    def toggle_mode(self):
        self.is_darkmode = not self.is_darkmode
        self.apply_theme()

        # update boxes
        for i, box in enumerate(self.boxes):
            box.update_background(i)
            box.apply_theme()
        self.apply_dynamic_scaling()

        # ✅ update logo based on mode
        self.update_logo_pixmap()

    
    def update_add_buttons(self):
        """Show ➕ only on the first empty row."""
        used = {b.symbol.text().strip() for b in self.boxes if b.symbol.text().strip()}
        all_syms = set(self.get_available_symbols_from_excel())
        remaining = [s for s in all_syms if s not in used]

        empty_boxes = [b for b in self.boxes if not b.symbol.text().strip()]
        if remaining and not empty_boxes:
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

        first_empty = empty_boxes[0] if empty_boxes else None
        for b in self.boxes:
            is_empty = (b.symbol.text().strip() == "")
            b.update_buttons(show_add=(b is first_empty))

    def request_move(self, box, direction):
        """Smoothly swap 'box' with its neighbor using animations."""
        try:
            idx = self.boxes.index(box)
        except ValueError:
            return

        new_idx = idx + direction
        if not (0 <= new_idx < len(self.boxes)):
            return

        other = self.boxes[new_idx]
        viewport = self.scroll.viewport()

        p1 = box.mapTo(viewport, QPoint(0, 0))
        p2 = other.mapTo(viewport, QPoint(0, 0))
        r1 = QRect(p1, box.size())
        r2 = QRect(p2, other.size())

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

        box.hide()
        other.hide()

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
            self.boxes[idx], self.boxes[new_idx] = self.boxes[new_idx], self.boxes[idx]

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

            self.update_add_buttons()
            self.scroll.ensureWidgetVisible(box)

        group.finished.connect(finalize)
        self._anim_group = group
        group.start()

    def refresh_once(self):
        try:
            rows = self.source.read_rows()
        except Exception as e:
            print("Read error:", e)
            traceback.print_exc()
            rows = []

        self.last_rows_dict = {sym: (bid, ask, low, high) for sym, bid, ask, low, high in rows}

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

        for box in self.boxes:
            sym = box.symbol.text().strip()
            if sym and sym in self.last_rows_dict:
                bid, ask, low, high = self.last_rows_dict[sym]
                box.update_prices(bid, ask, low, high)

        self.update_add_buttons()

    def toggle_fullscreen(self):
        if not self.is_fullscreen:
            self.showFullScreen()
            self.is_fullscreen = True
        else:
            self.showNormal()
            self.is_fullscreen = False

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
        
    def apply_font_to_widgets(self):
        for box in self.boxes:
            for lbl in [box.symbol, box.bid, box.ask, box.high, box.low]:
                lbl.setFont(self.current_font)
        if hasattr(self, 'header_frame'):
            for lbl in self.header_frame.findChildren(QLabel):
                lbl.setFont(self.current_font)
    
    def closeEvent(self, event):
        try: 
            self.timer.stop()
        except Exception: 
            pass
        try: 
            self.source.close()
        except Exception: 
            pass

        rows = [b.symbol.text().strip() for b in self.boxes]
        save_config(
            self.source.path,
            self.source.sheet_name,
            rows=rows,
            font=self.current_font,
            is_darkmode=self.is_darkmode
        )

        super().closeEvent(event)
        
        
    def fill_extra_rows_if_space(self):
        if not hasattr(self, 'boxes') or not self.boxes:
            return

        viewport_height = self.scroll.viewport().height()
        if viewport_height <= 0:
            return

        example_box = self.boxes[0]
        max_row_height = max(b.sizeHint().height() for b in self.boxes)
        current_rows_height = sum(b.sizeHint().height() + self.rows_layout.spacing() for b in self.boxes)
        remaining_height = viewport_height - current_rows_height
        if remaining_height < max_row_height * 0.5:
            return

        extra_rows_count = int(remaining_height // max_row_height)
        if extra_rows_count <= 0:
            return

        used_symbols = {b.symbol.text().strip() for b in self.boxes if b.symbol.text().strip()}
        available_symbols = [s for s in self.get_available_symbols_from_excel() if s not in used_symbols]

        font = self.current_font if hasattr(self, 'current_font') else QFont("Arial", 10)
        symbol_width = int(self.width() * 0.3)  # match header width

        for _ in range(extra_rows_count):
            sym = available_symbols.pop(0) if available_symbols else ""
            new_box = PriceBox(
                symbol=sym,
                row_index=len(self.boxes),
                remove_callback=self.on_row_cleared,
                add_callback=self.on_row_added,
                parent_widget=self
            )

            # Apply font and fixed widths for all labels
            for lbl in [new_box.symbol, new_box.bid, new_box.ask, new_box.high, new_box.low]:
                lbl.setFont(font)
                if lbl == new_box.symbol:
                    lbl.setFixedWidth(symbol_width)
                lbl.setSizePolicy(new_box.symbol.sizePolicy())  # enforce same policy

            new_box.layout().setSpacing(12)
            new_box.layout().setContentsMargins(10,5,10,5)

            self.rows_layout.addWidget(new_box)
            self.boxes.append(new_box)

        # Ensure one blank last row
        if self.boxes[-1].symbol.text().strip() != "":
            blank_box = PriceBox(
                symbol="",
                row_index=len(self.boxes),
                remove_callback=self.on_row_cleared,
                add_callback=self.on_row_added,
                parent_widget=self
            )
            for lbl in [blank_box.symbol, blank_box.bid, blank_box.ask, blank_box.high, blank_box.low]:
                lbl.setFont(font)
                if lbl == blank_box.symbol:
                    lbl.setFixedWidth(symbol_width)
                lbl.setSizePolicy(blank_box.symbol.sizePolicy())
            self.rows_layout.addWidget(blank_box)
            self.boxes.append(blank_box)

        self.reorder_boxes()
        self.update_add_buttons()
        self.apply_dynamic_scaling()



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
        from PyQt5.QtWidgets import QDialog, QDialogButtonBox, QFormLayout

        class ExcelConfigDialog(QDialog):
            def __init__(self):
                super().__init__()
                self.setWindowTitle("Select Excel File & Sheet")
                self.resize(400, 150)
                self.selected_config = None

                layout = QFormLayout(self)

                self.file_input = QLineEdit()
                self.file_btn = QPushButton("Browse")
                hb_file = QHBoxLayout()
                hb_file.addWidget(self.file_input)
                hb_file.addWidget(self.file_btn)
                layout.addRow("Excel File:", hb_file)

                self.sheet_input = QLineEdit()
                self.sheet_dropdown = QComboBox()
                hb_sheet = QHBoxLayout()
                hb_sheet.addWidget(self.sheet_input)
                hb_sheet.addWidget(self.sheet_dropdown)
                layout.addRow("Sheet name:", hb_sheet)

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

    window = MainWindow(file_path, sheet_name)
    window.is_darkmode = is_darkmode
    window.current_font = current_font
    window.apply_theme()
    window.apply_font_to_widgets()

    for i, box in enumerate(window.boxes):
        if i < len(saved_rows):
            box.symbol.setText(saved_rows[i])
    window.update_add_buttons()

    window.showMaximized()
    sys.exit(app.exec_())

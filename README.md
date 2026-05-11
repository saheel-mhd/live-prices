# livepriceapp

A Windows desktop app that streams live bid / ask / high / low prices from MetaTrader 4 over DDE and renders them in a sortable, themable price grid.

Built with PyQt5. Uses MT4's built-in DDE server in **advise (hot-link) mode** — the same mechanism Excel uses — so values are pushed by MT4 as ticks arrive instead of polled.

---

## Features

- Live BID / ASK / HIGH / LOW for any symbol your broker quotes
- Bid/ask flash green on up-tick, red on down-tick
- Custom display names — show `GOLD` in the UI but subscribe to `XAUUSD` over DDE
- Drag-free row reordering with smooth animation (▲ / ▼ buttons)
- Connection-status dot in the header — green = ticks flowing, amber = connected but no recent ticks, red = no DDE conversation (hover for detail)
- Dark / light theme toggle
- Custom font picker
- Up to 20 watched symbols; rows shrink to fit when count exceeds 12
- Persistent config — symbols, row order, font, and theme survive restarts
- Hot-editable symbol list — `Ctrl+Shift+S` to add / rename / remove rows on the fly without restarting

---

## Requirements

- **Windows** (DDE is a Win32-only IPC)
- **MetaTrader 4** running and logged in to a broker
- **Python 3.10+** (3.13 tested)
- **MT4-compatible broker** — MT5 is not supported (it dropped DDE)

---

## Setup

### 1. Install Python dependencies

```powershell
pip install -r requirements.txt   # just PyQt5
```

That's it — the DDE client is implemented directly against the Win32 DDEML APIs via `ctypes`, no `pywin32` or third-party DDE packages required.

### 2. Enable MT4's DDE server

In MT4: **Tools → Options → Server tab → tick "Enable DDE server" → OK**.

If you toggle this while MT4 is already running, restart MT4 — DDE registration happens at startup.

### 3. Make sure your symbols are in Market Watch

MT4 only publishes quotes over DDE for symbols visible in the **Market Watch** panel (`Ctrl+M`). If a symbol isn't there:

- Right-click in Market Watch → **Show All**, or
- Right-click → **Symbols…** → find the symbol → **Show**

Many brokers add suffixes (e.g. `EURUSD.m`, `XAUUSDpro`, `EURUSD#`). Use the **exact** name as it appears in Market Watch when configuring the app.

---

## Running

```powershell
python liveprices.py
```

On first launch the symbols editor pops up. Fill in the table:

| Name (UI label) | MT4 Symbol (ticker) |
|---|---|
| GOLD | XAUUSD |
| SILVER | XAGUSD |
| EUR | EURUSD |

- **Name** is what shows in the UI's Symbol column.
- **MT4 Symbol** is the exact ticker MT4 uses — that's what we subscribe to.
- Multiple names can map to the same MT4 symbol (e.g. `GOLD` and `GLD` both → `XAUUSD`).
- Click **OK** — the main window opens with prices populating within a tick or two.

---

## Keyboard shortcuts

| Shortcut | Action |
|---|---|
| `Ctrl+Shift+S` | Open symbols editor (add / rename / remove rows live) |
| `Ctrl+Shift+F` | Change font |
| `Ctrl+Shift+F1` | Toggle fullscreen |
| `🌓` button | Toggle dark / light theme |
| `▲` / `▼` per row | Reorder |
| `✖` per row | Remove that pair (rows expand back if total > 12) |

---

## Tests

A small smoke test covers price formatting and config save/load round-trips (no extra deps — stdlib `unittest`):

```powershell
python -m unittest test_liveprices -v
```

---

## Building a standalone `.exe`

```powershell
pip install pyinstaller
python -m PyInstaller --clean --noconfirm --onefile --windowed `
    --name livepriceapp --icon icon.ico `
    --add-data "icon.ico;." liveprices.py
```

Output: `dist\livepriceapp.exe` — a single file, no Python installation required on the target machine.

---

## Configuration

State is stored next to the executable in `config.txt`:

```
SYMBOLS=GOLD:XAUUSD,SILVER:XAGUSD,EUR:EURUSD
FONT=Arial,10
IS_DARKMODE=True
ROWS=GOLD,SILVER,EUR
```

| Key | Purpose |
|---|---|
| `SYMBOLS` | Comma-separated `name:ticker` pairs |
| `FONT` | `family,size` |
| `IS_DARKMODE` | `True` / `False` |
| `ROWS` | Display order (subset of names from SYMBOLS) |

Legacy single-value entries (`SYMBOLS=EURUSD`) still load — name defaults to the ticker.

---

## Troubleshooting

### All values show blank / `N\A`

- MT4 isn't running, or DDE server is disabled. Check **Tools → Options → Server**.
- Symbol isn't in Market Watch. Open `Ctrl+M` — is the row there with ticking prices?
- Wrong ticker. Use the exact name MT4 shows (suffixes matter).

### `DdeConnect failed for topic 'BID' (DDE err 0x400a)`

`0x400a` = `DMLERR_NO_CONV_ESTABLISHED`. MT4 isn't accepting connections right now. Causes:

- MT4 was closed or restarted since the last successful run — relaunch the app.
- Broker connection dropped (bottom-right of MT4 doesn't show green bars). Reconnect.
- DDE server got toggled off. Re-enable in MT4 options and restart MT4.

The app also retries `DdeConnect` 5 times over 2 seconds at startup to absorb the transient race that happens right after a previous instance exits.

### `ModuleNotFoundError: No module named 'pyqt5'`

You're running with the wrong Python. Common gotcha: `WindowsApps\python3.13.exe` (the Microsoft Store stub) doesn't share `pip install` with `Programs\Python\Python313\python.exe`. Run with the explicit path:

```powershell
& "C:\Users\<you>\AppData\Local\Programs\Python\Python313\python.exe" liveprices.py
```

In VS Code: **Ctrl+Shift+P → Python: Select Interpreter** and pick the right one.

### Taskbar shows the Python icon, not `icon.ico`

The packaged `.exe` sets a Windows AppUserModelID (`livepriceapp`) so the running window is recognized as its own app. If you're running directly via `python liveprices.py` (not as the packaged exe), the taskbar may still show Python's icon — that's expected outside the bundle.

### High CPU / sluggish UI

Default refresh is 100ms. To slow it, edit `REFRESH_INTERVAL_MS` near the top of `liveprices.py`. Note that with advise-mode DDE the timer only redraws cached values — actual DDE traffic is push-based, so the timer interval mostly affects redraw smoothness, not network/IPC load.

---

## Architecture

A single file, intentionally — the whole app fits in one screen-scrollable script:

| Component | Responsibility |
|---|---|
| `MT4DDESource` | Win32 DDEML client. Subscribes to BID/ASK/HIGH/LOW topics for every unique MT4 symbol. Stores latest pushes in `self.cache`. |
| `SymbolsConfigDialog` | Modal table editor for `(name, mt4_symbol)` pairs. Always renders in light theme. |
| `PriceBox` | One row in the grid. Symbol + bid/ask/low/high + reorder arrows + ✖. |
| `FontChanger` | Font picker with per-family preview. |
| `MainWindow` | Header + scrollable rows + theme toggle. Refreshes from the source on a `QTimer`. |

DDE callback (`_on_dde_event`) runs on the main thread when Qt processes Windows messages, so no locking is needed around the cache.

---

## Why not just use Excel?

The original version of this app pulled prices from a live Excel sheet via `xlwings`. That worked but:

- 100 ms COM round-trips through Excel pegged CPU on busy workbooks
- Required Excel to stay open
- Excel itself was just doing the DDE work — one less moving part to subscribe directly

Switching to native DDE eliminated the Excel dependency and the polling overhead. MT4 pushes ticks straight into our cache.

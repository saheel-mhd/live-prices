"""
End-to-end DDE test using the same advise/hot-link approach as liveprices.py.

What it tests:
  1. DdeInitialize succeeds
  2. DdeConnect to MT4|BID, |ASK, |HIGH, |LOW
  3. XTYP_ADVSTART subscriptions for a few symbols
  4. Pumps Windows messages for ~6s and prints any XTYP_ADVDATA pushes
  5. Cleanup, then re-init and re-connect (catches stale-state bugs)

Run:
  & "C:\\Users\\techsupport\\AppData\\Local\\Programs\\Python\\Python313\\python.exe" dde_test.py
"""
import ctypes
import time
from ctypes import wintypes, byref, WINFUNCTYPE, POINTER, c_void_p, c_uint, c_ulong, c_int, c_size_t

user32 = ctypes.WinDLL("user32", use_last_error=True)

APPCMD_CLIENTONLY = 0x00000010
CP_WINUNICODE = 1200
CF_TEXT = 1

XCLASS_BOOL = 0x1000
XCLASS_FLAGS = 0x4000
XTYP_ADVDATA = 0x0010 | XCLASS_FLAGS
XTYP_ADVSTART = 0x0030 | XCLASS_BOOL
DDE_FACK = 0x8000

DDECALLBACK = WINFUNCTYPE(
    c_void_p, c_uint, c_uint, c_void_p, c_void_p, c_void_p, c_void_p, c_size_t, c_size_t
)
user32.DdeInitializeW.argtypes = [POINTER(c_ulong), DDECALLBACK, c_ulong, c_ulong]
user32.DdeInitializeW.restype = c_uint
user32.DdeUninitialize.argtypes = [c_ulong]
user32.DdeUninitialize.restype = wintypes.BOOL
user32.DdeCreateStringHandleW.argtypes = [c_ulong, ctypes.c_wchar_p, c_int]
user32.DdeCreateStringHandleW.restype = c_void_p
user32.DdeQueryStringW.argtypes = [c_ulong, c_void_p, ctypes.c_wchar_p, c_ulong, c_int]
user32.DdeQueryStringW.restype = c_ulong
user32.DdeConnect.argtypes = [c_ulong, c_void_p, c_void_p, c_void_p]
user32.DdeConnect.restype = c_void_p
user32.DdeClientTransaction.argtypes = [
    c_void_p, c_ulong, c_void_p, c_void_p, c_uint, c_uint, c_ulong, POINTER(c_ulong)
]
user32.DdeClientTransaction.restype = c_void_p
user32.DdeAccessData.argtypes = [c_void_p, POINTER(c_ulong)]
user32.DdeAccessData.restype = c_void_p
user32.DdeUnaccessData.argtypes = [c_void_p]
user32.DdeUnaccessData.restype = wintypes.BOOL
user32.DdeGetLastError.argtypes = [c_ulong]
user32.DdeGetLastError.restype = c_uint

DDE_ERRORS = {
    0x4000: "ADVACKTIMEOUT", 0x4001: "BUSY", 0x4002: "DATAACKTIMEOUT",
    0x4003: "DLL_NOT_INITIALIZED", 0x4004: "DLL_USAGE", 0x4005: "EXECACKTIMEOUT",
    0x4006: "INVALIDPARAMETER", 0x4007: "LOW_MEMORY", 0x4008: "MEMORY_ERROR",
    0x4009: "NOTPROCESSED", 0x400a: "NO_CONV_ESTABLISHED", 0x400b: "POKEACKTIMEOUT",
    0x400c: "POSTMSG_FAILED", 0x400d: "REENTRANCY", 0x400e: "SERVER_DIED",
    0x400f: "SYS_ERROR", 0x4010: "UNADVACKTIMEOUT", 0x4011: "UNFOUND_QUEUE_ID",
}

received = {}  # (topic, item) -> latest text

# Pump messages so DDE callbacks get dispatched.
PM_REMOVE = 0x0001
class MSG(ctypes.Structure):
    _fields_ = [("hwnd", c_void_p), ("message", c_uint), ("wParam", c_size_t),
                ("lParam", c_size_t), ("time", c_ulong),
                ("pt_x", c_int), ("pt_y", c_int)]
user32.PeekMessageW.argtypes = [POINTER(MSG), c_void_p, c_uint, c_uint, c_uint]
user32.PeekMessageW.restype = wintypes.BOOL
user32.TranslateMessage.argtypes = [POINTER(MSG)]
user32.DispatchMessageW.argtypes = [POINTER(MSG)]


def pump(seconds):
    """Pump the Windows message queue for `seconds` so DDE callbacks fire."""
    end = time.time() + seconds
    msg = MSG()
    while time.time() < end:
        if user32.PeekMessageW(byref(msg), None, 0, 0, PM_REMOVE):
            user32.TranslateMessage(byref(msg))
            user32.DispatchMessageW(byref(msg))
        else:
            time.sleep(0.01)


def err_name(code):
    return DDE_ERRORS.get(code, f"unknown 0x{code:04x}")


def run_round(label, idInst, callback, symbols):
    print(f"\n=== {label} ===")
    rc = user32.DdeInitializeW(byref(idInst), callback, APPCMD_CLIENTONLY, 0)
    print(f"  DdeInitialize -> {rc} ({'OK' if rc == 0 else 'FAIL'})")
    if rc != 0:
        return False

    service_hsz = user32.DdeCreateStringHandleW(idInst.value, "MT4", CP_WINUNICODE)
    print(f"  service HSZ for 'MT4' -> 0x{service_hsz or 0:x}")

    topics = ["BID", "ASK", "HIGH", "LOW"]
    topic_hsz = {t: user32.DdeCreateStringHandleW(idInst.value, t, CP_WINUNICODE) for t in topics}
    item_hsz = {s: user32.DdeCreateStringHandleW(idInst.value, s, CP_WINUNICODE) for s in symbols}

    convs = {}
    for t in topics:
        hConv = None
        for attempt in range(5):
            hConv = user32.DdeConnect(idInst, service_hsz, topic_hsz[t], None)
            if hConv:
                break
            err = user32.DdeGetLastError(idInst.value)
            print(f"    DdeConnect '{t}' attempt {attempt+1} failed: {err_name(err)}")
            time.sleep(0.4)
        if hConv:
            print(f"  DdeConnect '{t}' -> hConv 0x{hConv:x}")
            convs[t] = hConv
        else:
            print(f"  DdeConnect '{t}' -> FAILED after 5 attempts")

    if not convs:
        user32.DdeUninitialize(idInst.value)
        return False

    print(f"  Subscribing to {len(symbols)} symbol(s) on {len(convs)} topic(s)...")
    for t, hConv in convs.items():
        for s in symbols:
            pdw = c_ulong(0)
            user32.DdeClientTransaction(
                None, 0, hConv, item_hsz[s], CF_TEXT, XTYP_ADVSTART, 1500, byref(pdw)
            )

    print("  Pumping messages for 6s; expect XTYP_ADVDATA callbacks...")
    pump(6.0)

    print(f"  Received {len(received)} (topic,item) updates:")
    for (t, s) in sorted(received):
        print(f"    {t:6} {s:10} = {received[(t,s)]!r}")

    user32.DdeUninitialize(idInst.value)
    print("  DdeUninitialize done")
    return len(received) > 0


def main():
    symbols = ["EURUSD", "GBPUSD", "XAUUSD"]
    idInst = c_ulong(0)

    # Bind callback once (must outlive both rounds).
    def cb(uType, uFmt, hConv, hsz1, hsz2, hData, dwData1, dwData2):
        if uType == XTYP_ADVDATA:
            buf1 = ctypes.create_unicode_buffer(64)
            buf2 = ctypes.create_unicode_buffer(64)
            user32.DdeQueryStringW(idInst.value, hsz1, buf1, 64, CP_WINUNICODE)
            user32.DdeQueryStringW(idInst.value, hsz2, buf2, 64, CP_WINUNICODE)
            length = c_ulong(0)
            ptr = user32.DdeAccessData(hData, byref(length))
            if ptr:
                try:
                    raw = ctypes.string_at(ptr, length.value)
                    text = raw.decode("ascii", errors="ignore").rstrip("\x00").strip()
                    received[(buf1.value, buf2.value)] = text
                finally:
                    user32.DdeUnaccessData(hData)
            return DDE_FACK
        return 0

    callback = DDECALLBACK(cb)

    ok1 = run_round("ROUND 1", idInst, callback, symbols)
    received.clear()
    time.sleep(1.5)
    idInst2 = c_ulong(0)
    ok2 = run_round("ROUND 2 (re-init after cleanup)", idInst2, callback, symbols)

    print("\n=== summary ===")
    print(f"  Round 1: {'PASS — got data' if ok1 else 'FAIL — no data'}")
    print(f"  Round 2: {'PASS — reconnect works' if ok2 else 'FAIL — reconnect blocked'}")


if __name__ == "__main__":
    main()

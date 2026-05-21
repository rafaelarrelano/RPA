"""
test_sap_tcode.py
Test navigasi T-code di SAP via pyautogui
Shortcut: Ctrl+/ untuk fokus command bar
Jalankan dengan SAP Easy Access sudah terbuka
"""
import time
import pyautogui
import win32gui
import win32con

pyautogui.FAILSAFE = True
pyautogui.PAUSE    = 0.5

TCODE = "ZPGD_SAPSTK"   # ganti sesuai T-code yang mau ditest

# Class name window SAP GUI — ini yang membedakan SAP dari browser/app lain
SAP_WINDOW_CLASSES = [
    "SAP_FRONTEND_SESSION",   # SAP GUI 7.x
    "SAPFrontend",
    "SAPGUI",
]

# Keyword judul yang PASTI milik SAP GUI (bukan browser)
SAP_TITLE_EXACT = [
    "SAP Easy Access",
    "SAP R/3",
    "SAP NetWeaver",
    "ZPGD",
    "MIGO",
    "Program transfer",
]

# Keyword yang harus TIDAK ada di judul (browser, editor, dll)
SKIP_IF_CONTAINS = [
    "Firefox", "Chrome", "Edge", "Chromium",
    "Visual Studio", "Code", "Notepad",
    "Claude", "Thunderbird", "Explorer",
]


def get_sap_hwnd():
    """
    Cari window SAP GUI yang benar.
    Prioritas:
    1. Class name window = SAP_FRONTEND_SESSION / SAPFrontend
    2. Judul mengandung keyword SAP spesifik DAN tidak mengandung keyword browser
    """
    priority = []   # class name cocok → pasti SAP GUI
    fallback = []   # judul cocok tapi class tidak dikenal

    def cb(hwnd, _):
        if not win32gui.IsWindowVisible(hwnd):
            return
        title      = win32gui.GetWindowText(hwnd)
        class_name = win32gui.GetClassName(hwnd)

        if not title:
            return

        # Skip kalau judulnya mengandung browser/editor
        if any(s in title for s in SKIP_IF_CONTAINS):
            return

        # Prioritas 1: class name SAP
        if any(c in class_name for c in SAP_WINDOW_CLASSES):
            priority.append((hwnd, title, class_name))
            return

        # Prioritas 2: judul mengandung keyword SAP spesifik
        if any(kw in title for kw in SAP_TITLE_EXACT):
            fallback.append((hwnd, title, class_name))

    win32gui.EnumWindows(cb, None)
    return priority if priority else fallback


def focus_sap_window():
    wins = get_sap_hwnd()
    if not wins:
        print("SAP window tidak ditemukan!")
        print("Pastikan SAP Easy Access sudah terbuka dan login.")
        return False

    hwnd, title, class_name = wins[0]
    print(f"SAP window: {title!r}")
    print(f"Class name: {class_name!r}")
    win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)
    win32gui.SetForegroundWindow(hwnd)
    time.sleep(1)
    return True


# ── DEBUG: lihat semua window yang terbuka ────────────────
print("=== Semua window visible ===")
def _dbg(hwnd, _):
    if win32gui.IsWindowVisible(hwnd):
        title = win32gui.GetWindowText(hwnd)
        cls   = win32gui.GetClassName(hwnd)
        if title:
            print(f"  class={cls!r:35} title={title!r}")
win32gui.EnumWindows(_dbg, None)
print()
# ─────────────────────────────────────────────────────────

print("Mulai dalam 3 detik — jangan sentuh keyboard/mouse!")
time.sleep(3)

# 1. Fokus SAP
print("\n[1] Fokus SAP window...")
if not focus_sap_window():
    exit()

# 2. Ctrl+/ untuk fokus command bar, lalu ketik tcode
print(f"\n[2] Ketik T-code {TCODE!r} via Ctrl+/ ...")
pyautogui.keyDown("ctrl")
time.sleep(0.1)
pyautogui.press("/")
pyautogui.keyUp("ctrl")
time.sleep(0.4)

pyautogui.hotkey("ctrl", "a")
time.sleep(0.1)
pyautogui.typewrite(f"/{TCODE}", interval=0.1)
time.sleep(0.3)

print(f"\nSelesai! Cek SAP — harusnya ada '/{TCODE}' di command bar.")
print("Tekan Enter di SAP jika ingin masuk ke T-code tersebut.")
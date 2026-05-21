"""
test_sap_pyautogui.py
Test koneksi ke SAP via pyautogui - simulasi keyboard/mouse
READ ONLY - tidak ada perubahan di SAP
Jalankan: python test_sap_pyautogui.py
"""
import pyautogui
import time
import subprocess

# Matikan failsafe (geser mouse ke pojok = stop)
pyautogui.FAILSAFE = True
pyautogui.PAUSE = 0.5  # jeda 0.5 detik antar aksi

def find_sap_window():
    """Cari window SAP yang sedang terbuka."""
    import win32gui
    import win32con

    sap_windows = []

    def callback(hwnd, windows):
        if win32gui.IsWindowVisible(hwnd):
            title = win32gui.GetWindowText(hwnd)
            if "SAP" in title:
                windows.append((hwnd, title))

    win32gui.EnumWindows(callback, sap_windows)
    return sap_windows

def test_pyautogui_sap():
    print("=" * 50)
    print("Test SAP via pyautogui")
    print("=" * 50)

    # 1. Cari window SAP
    print("\n[1] Mencari window SAP...")
    windows = find_sap_window()

    if not windows:
        print("SAP window tidak ditemukan!")
        print("Pastikan SAP sudah terbuka dan sudah login.")
        return False

    print(f"SAP window ditemukan: {len(windows)} window")
    for hwnd, title in windows:
        print(f"  - {title}")

    # 2. Fokus ke window SAP pertama
    import win32gui, win32con
    hwnd = windows[0][0]
    win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
    win32gui.SetForegroundWindow(hwnd)
    time.sleep(1)
    print(f"\n[2] Fokus ke window: {windows[0][1]}")

    # 3. Test ketik di command bar SAP
    print("\n[3] Test ketik T-code MB52...")
    pyautogui.hotkey('alt', 'F6')  # fokus ke command bar
    time.sleep(0.5)
    pyautogui.hotkey('ctrl', 'a')  # select all
    pyautogui.typewrite('/nMB52', interval=0.1)
    time.sleep(0.5)

    print("\nHasil: SAP window ditemukan dan keyboard input berhasil!")
    print("CATATAN: Cek SAP kamu - harusnya ada '/nMB52' di command bar")
    print("\nJangan tekan Enter dulu - ini hanya test input keyboard.")
    return True

if __name__ == "__main__":
    test_pyautogui_sap()
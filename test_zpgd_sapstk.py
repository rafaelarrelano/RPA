"""
test_zpgd_sapstk.py
Test export ZPGD_SAPSTK untuk satu plant via pyautogui
"""
import pyautogui
import win32gui, win32con
import time, glob, os
from main import tab_to, type_field, focus_sap, sap_tcode, get_sap_hwnd

SAP_DOWNLOAD_DIR = r"C:\MATRIX\DOWNLOAD"
PLANT            = "4502"  # ganti sesuai plant yang mau di-test

def run_zpgd_sapstk(plant: str) -> str:
    print(f'Export ZPGD_SAPSTK untuk plant {plant}...')

    # Catat waktu sebelum export
    before = time.time()

    # Navigasi ke T-code
    sap_tcode("nZPGD_SAPSTK")
    time.sleep(2)

    # Fokus window SAP
    wins = []
    win32gui.EnumWindows(
        lambda h,_: wins.append(h) if win32gui.IsWindowVisible(h)
        and 'Program transfer' in win32gui.GetWindowText(h) else None, None
    )
    if not wins:
        raise Exception("Window ZPGD_SAPSTK tidak ditemukan!")

    win32gui.ShowWindow(wins[0], win32con.SW_MAXIMIZE)
    win32gui.SetForegroundWindow(wins[0])
    time.sleep(0.8)

# Reset ke awal
    pyautogui.hotkey("ctrl", "Home")
    time.sleep(0.4)

# Tab 0: Pilihan (1=ALL)
    type_field("1")

# Tab 1: Plant Code
    tab_to(1)
    type_field(plant)

# Execute F8
    print('Execute F8...')
    pyautogui.press("f8")
    time.sleep(3)

# Handle popup SAP GUI Security → klik Allow (kiri dari Deny)
    print('Handle popup Allow...')
    pyautogui.press("left")   # pindah dari Deny ke Allow
    time.sleep(0.3)
    pyautogui.press("enter")  # klik Allow
    time.sleep(5)             # tunggu file selesai dibuat

    # Cari file terbaru yang dibuat setelah timestamp sebelumnya
    pattern = os.path.join(SAP_DOWNLOAD_DIR, f"{plant}_*_SAPSTK_*.TXT")
    files   = [f for f in glob.glob(pattern) if os.path.getmtime(f) > before]

    if not files:
        # Fallback: ambil file terbaru
        all_files = glob.glob(pattern)
        if all_files:
            files = [max(all_files, key=os.path.getmtime)]

    if not files:
        raise FileNotFoundError(f"File SAPSTK tidak muncul di {SAP_DOWNLOAD_DIR}")

    latest = max(files, key=os.path.getmtime)
    print(f'File berhasil dibuat: {os.path.basename(latest)}')
    print(f'Ukuran: {os.path.getsize(latest):,} bytes')

    # Preview isi file
    with open(latest, 'r', encoding='utf-8', errors='ignore') as f:
        lines = [l.strip() for l in f if l.strip().startswith('FSTKGD')]
    print(f'Total baris FSTKGD: {len(lines)}')
    print('Contoh 3 baris:')
    for l in lines[:3]:
        print(' ', l)

    return latest


if __name__ == '__main__':
    print('Mulai dalam 3 detik - pastikan SAP sudah terbuka!')
    time.sleep(3)
    filepath = run_zpgd_sapstk(PLANT)
    print(f'\nSelesai! File: {filepath}')
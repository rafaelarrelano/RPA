"""
RPA Stock Reconciliation - Python + pyautogui
Alur: Export SAP .txt → Compare dengan Matrix Portal → Input MIGO_GI → Kirim laporan
SAP GUI R/3 | Python 3.8+
"""

import os
import re
import sys
import time
import math
import logging
import openpyxl
import pyautogui
import win32gui
import win32con
from datetime import datetime
from dataclasses import dataclass
from typing import Optional
from playwright.sync_api import sync_playwright

from config import Config
from logger import setup_logger

log = setup_logger()

pyautogui.FAILSAFE = True
pyautogui.PAUSE    = 0.6


# ─────────────────────────────────────────────
# MODEL DATA
# ─────────────────────────────────────────────

@dataclass
class StockDiff:
    param:        str
    plant:        str
    sloc:         str
    posting_date: str
    material:     str
    qty_matrix:   float
    qty_sap:      float
    diff:         float
    status:       int
    mvt_type:     str   = ""
    qty_adjust:   float = 0.0


# ─────────────────────────────────────────────
# MAPPING PLANT → COST CENTER
# ─────────────────────────────────────────────

def load_plant_mapping() -> dict:
    """Baca mapping Plant → Cost Center dari Excel."""
    try:
        wb = openpyxl.load_workbook(Config.PLANT_MAPPING_FILE, data_only=True)
        ws = wb["Plant_CostCenter"]
        mapping = {}
        for row in ws.iter_rows(min_row=5, values_only=True):
            if row[1] and row[3]:
                mapping[str(row[1]).strip()] = str(row[3]).strip()
        log.info(f"[CONFIG] {len(mapping)} plant mapping berhasil dibaca")
        return mapping
    except Exception as e:
        log.error(f"[CONFIG] Gagal baca plant mapping: {e}")
        raise


# ─────────────────────────────────────────────
# PARSING NOTIF
# ─────────────────────────────────────────────

def parse_decimal(raw: str) -> float:
    """Konversi desimal koma ke titik. '6,375' -> 6.375"""
    cleaned = raw.strip().replace(",", ".")
    return float(cleaned)


def convert_date(yyyymmdd: str) -> str:
    """Konversi tanggal SAP. '20260401' -> '01.04.2026'"""
    if len(yyyymmdd) != 8:
        raise ValueError(f"Format tanggal tidak valid: {yyyymmdd}")
    return f"{yyyymmdd[6:8]}.{yyyymmdd[4:6]}.{yyyymmdd[0:4]}"


def set_movement_type(item: StockDiff) -> bool:
    """Tentukan movement type. Return False jika selisih = 0."""
    if item.diff < 0:
        item.mvt_type   = "917"
        item.qty_adjust = abs(item.diff)
        return True
    elif item.diff > 0:
        item.mvt_type   = "918"
        item.qty_adjust = item.diff
        return True
    else:
        log.info(f"[SKIP] Selisih 0 untuk material {item.material}")
        return False


def parse_line(raw_line: str) -> Optional[StockDiff]:
    """Parse satu baris notif pipe-delimited dari portal."""
    if not raw_line.strip():
        return None
    fields = raw_line.strip().split("|")
    if len(fields) != 9:
        log.warning(f"[SKIP] Baris tidak valid ({len(fields)} field): {raw_line}")
        return None
    try:
        item = StockDiff(
            param        = fields[0].strip(),
            plant        = fields[1].strip(),
            sloc         = fields[2].strip(),
            posting_date = convert_date(fields[3].strip()),
            material     = fields[4].strip(),
            qty_matrix   = parse_decimal(fields[5]),
            qty_sap      = parse_decimal(fields[6]),
            diff         = parse_decimal(fields[7]),
            status       = int(fields[8].strip()),
        )
        return item
    except Exception as e:
        log.error(f"[ERROR] Gagal parse baris: {raw_line} | {e}")
        return None


def get_valid_items(raw_text: str) -> list:
    """Parse seluruh teks notif, filter, kembalikan list siap diproses."""
    valid = []
    for line in raw_text.strip().splitlines():
        item = parse_line(line)
        if item is None:
            continue
        if item.status != 0:
            log.info(f"[SKIP] Status bukan 0, material: {item.material}")
            continue
        if not set_movement_type(item):
            continue
        valid.append(item)
    log.info(f"[INFO] Total SKU valid: {len(valid)}")
    return valid


def format_qty_for_sap(qty: float) -> str:
    """Format qty untuk input SAP. 0.001 -> '0,001' | 1.0 -> '1'"""
    if qty == math.floor(qty):
        return str(int(qty))
    return f"{qty:.3f}".rstrip("0").rstrip(".").replace(".", ",")


# ─────────────────────────────────────────────
# HELPER - SAP WINDOW
# ─────────────────────────────────────────────

def get_sap_hwnd(title_keyword: str = "SAP") -> Optional[int]:
    """
    Cari window SAP GUI yang aktif (bukan SAP Logon launcher).

    Prioritas:
    1. Window dengan class name SAP GUI (SAP_FRONTEND_SESSION, SAPFrontend)
    2. Window dengan judul spesifik T-code SAP
    3. Fallback: window lain yang mengandung keyword

    Selalu skip: SAP Logon (launcher), browser, editor.
    """
    # Class name window SAP GUI yang sudah login
    SAP_CLASSES = [
        "SAP_FRONTEND_SESSION",
        "SAPFrontend",
        "SAPGUI",
    ]

    # Keyword judul yang pasti window SAP GUI aktif (bukan launcher)
    SAP_SPECIFIC = [
        "SAP Easy Access",
        "SAP R/3",
        "SAP NetWeaver",
        "Program transfer",      # ZPGD_SAPSTK
        "Report Flow Transfer",  # ZPGD_SAPSTK varian lain
        "MIGO",
        "ZPGD",
        "Display Material",
        "Goods Issue",
        "Import",                # Dialog import file
        "Transfer Branch",       # T-code flow transfer
    ]

    # Window yang WAJIB di-skip (launcher, browser, editor)
    SKIP_TITLES = [
        "SAP Logon",        # ← launcher, bukan session aktif
        "Chrome", "Firefox", "Edge", "Chromium",
        "Visual Studio", "Code", "Notepad",
        "Claude", "Thunderbird", "Explorer",
    ]

    wins_class    = []    # class name cocok → paling reliable
    wins_specific = []    # judul spesifik T-code
    wins_fallback = []    # fallback umum

    def cb(hwnd, _):
        if not win32gui.IsWindowVisible(hwnd):
            return
        title = win32gui.GetWindowText(hwnd)
        if not title:
            return

        # Skip launcher dan aplikasi non-SAP
        if any(s in title for s in SKIP_TITLES):
            return

        cls = win32gui.GetClassName(hwnd)

        # Prioritas 1: class name SAP GUI
        if any(c in cls for c in SAP_CLASSES):
            wins_class.append((hwnd, title))
            return

        # Prioritas 2: judul spesifik T-code
        if any(kw in title for kw in SAP_SPECIFIC):
            wins_specific.append((hwnd, title))
            return

        # Prioritas 3: fallback keyword umum
        if title_keyword in title:
            wins_fallback.append((hwnd, title))

    win32gui.EnumWindows(cb, None)

    # Ambil hwnd dari list dengan prioritas
    for lst in [wins_class, wins_specific, wins_fallback]:
        if lst:
            return lst[0][0]
    return None


def focus_sap(title_keyword: str = "SAP"):
    """Fokuskan window SAP ke depan."""
    hwnd = get_sap_hwnd(title_keyword)
    if hwnd:
        win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)
        win32gui.SetForegroundWindow(hwnd)
        time.sleep(0.8)
    else:
        raise Exception(f"SAP window '{title_keyword}' tidak ditemukan!")


def sap_tcode(tcode: str):
    """Navigasi ke T-code di SAP via Ctrl+/ untuk fokus command bar."""
    focus_sap()
    time.sleep(0.3)
    pyautogui.keyDown("ctrl")
    pyautogui.press("/")
    pyautogui.keyUp("ctrl")
    time.sleep(0.4)
    pyautogui.hotkey("ctrl", "a")
    time.sleep(0.1)
    pyautogui.typewrite(f"/{tcode}", interval=0.08)
    time.sleep(0.3)
    pyautogui.press("enter")
    time.sleep(2.5)


# ─────────────────────────────────────────────
# HELPER - TAB NAVIGATION
# ─────────────────────────────────────────────

def tab_to(n: int, delay: float = 0.15):
    """Tekan Tab sebanyak n kali dengan jeda."""
    for _ in range(n):
        pyautogui.press("tab")
        time.sleep(delay)


def type_field(value: str, interval: float = 0.07):
    """Clear field lalu ketik nilai."""
    pyautogui.hotkey("ctrl", "a")
    time.sleep(0.1)
    pyautogui.typewrite(str(value), interval=interval)


# ─────────────────────────────────────────────
# SCREENSHOT
# ─────────────────────────────────────────────

def take_screenshot(folder: str, item: StockDiff = None, doc_number: str = "") -> str:
    """
    Ambil screenshot layar SAP saat ini.
    Return: path file screenshot yang disimpan.
    """
    os.makedirs(folder, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

    if item:
        filename = f"MIGO_{item.plant}_{item.material}_{timestamp}.png"
    elif doc_number:
        filename = f"MIGO_{doc_number}_{timestamp}.png"
    else:
        filename = f"SAP_{timestamp}.png"

    filepath = os.path.join(folder, filename)
    pyautogui.screenshot(filepath)
    log.info(f"[SCREENSHOT] Disimpan: {filename}")
    return filepath


# ─────────────────────────────────────────────
# FASE 1 - EXPORT STOK SAP VIA T-CODE
# ─────────────────────────────────────────────

def run_tcode_export(output_path: str):
    """Jalankan T-code custom export stok ke file .txt via pyautogui."""
    try:
        log.info(f"[FASE1] Navigasi ke T-code {Config.SAP_TCODE_EXPORT}")
        sap_tcode(Config.SAP_TCODE_EXPORT)

        pyautogui.press("f8")
        time.sleep(3)

        pyautogui.hotkey("alt", "l")
        time.sleep(0.5)
        pyautogui.press("down", presses=3)
        pyautogui.press("enter")
        time.sleep(0.5)
        pyautogui.press("down")
        pyautogui.press("enter")
        time.sleep(1)

        pyautogui.press("enter")
        time.sleep(0.5)

        pyautogui.hotkey("ctrl", "a")
        pyautogui.typewrite(output_path, interval=0.05)
        pyautogui.press("enter")
        time.sleep(2)

        if not os.path.exists(output_path) or os.path.getsize(output_path) == 0:
            raise FileNotFoundError(f"File export kosong: {output_path}")

        log.info(f"[FASE1] Export berhasil: {output_path}")

    except Exception as e:
        log.error(f"[FASE1] Export gagal: {e}")
        raise


# ─────────────────────────────────────────────
# FASE 2 - UPLOAD KE PORTAL & AMBIL NOTIF
# ─────────────────────────────────────────────

def get_notif_from_portal(sap_txt_path: str, portal_url: str) -> str:
    """Buka portal, upload .txt SAP, ambil teks notif selisih."""
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        page    = browser.new_page()
        try:
            log.info(f"[FASE2] Membuka portal: {portal_url}")
            page.goto(portal_url, timeout=30000)
            page.wait_for_load_state("networkidle")

            page.set_input_files("input[type='file']", sap_txt_path)
            log.info("[FASE2] File .txt SAP berhasil diupload")

            page.click("button[type='submit']")
            page.wait_for_load_state("networkidle")
            time.sleep(3)

            notif_text = page.inner_text("#notif-result")
            log.info(f"[FASE2] Notif diambil: {len(notif_text.splitlines())} baris")

            backup = os.path.join(Config.FOLDER_OUTPUT,
                                  f"notif_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt")
            with open(backup, "w") as f:
                f.write(notif_text)

            return notif_text

        except Exception as e:
            log.error(f"[FASE2] Gagal ambil notif: {e}")
            raise
        finally:
            browser.close()


# ─────────────────────────────────────────────
# FASE 4 - INPUT MIGO_GI VIA PYAUTOGUI
# ─────────────────────────────────────────────

def _wait_migo_ready(timeout: int = 30):
    """Tunggu sampai MIGO screen siap (judul window berubah)."""
    start = time.time()
    while time.time() - start < timeout:
        hwnd = get_sap_hwnd("Goods Issue")
        if hwnd:
            return True
        hwnd = get_sap_hwnd("MIGO")
        if hwnd:
            return True
        time.sleep(0.5)
    raise TimeoutError("MIGO screen tidak muncul dalam batas waktu!")


def input_migo_single(item: StockDiff, plant_cc: dict, is_first: bool = True):
    """
    Input satu item ke MIGO_GI.
    - is_first=True  : buka MIGO baru, isi header
    - is_first=False : tambah baris baru di dokumen yang sama
    """
    cc = plant_cc.get(item.plant, "")
    if not cc:
        raise ValueError(f"Cost center untuk plant {item.plant} tidak ditemukan!")

    qty_str = format_qty_for_sap(item.qty_adjust)

    if is_first:
        # ── Buka MIGO_GI ─────────────────────────────────────
        log.info(f"[MIGO] Buka MIGO_GI | Plant={item.plant} | Mvt={item.mvt_type}")
        sap_tcode("MIGO_GI")
        time.sleep(2)

        focus_sap("MIGO")

        # Pilih movement type (field pertama)
        pyautogui.hotkey("ctrl", "Home")
        time.sleep(0.3)

        # Field: A07 Good Issue → pilih movement type
        # Tab 0: Action = A07 (Goods Issue)
        type_field("A07")
        tab_to(1)

        # Tab 1: Reference = Other (R10)
        type_field("R10")
        tab_to(1)
        time.sleep(0.5)

        # Isi movement type di field header
        # Navigasi ke tab "General" — field Movement Type
        pyautogui.hotkey("ctrl", "Home")
        time.sleep(0.3)

        # Klik tab General atau langsung isi movement type
        # Movement type ada di bagian atas form
        tab_to(2)
        type_field(item.mvt_type)
        pyautogui.press("enter")
        time.sleep(1)

        # Isi Posting Date
        tab_to(1)
        type_field(item.posting_date)
        pyautogui.press("tab")
        time.sleep(0.3)

    else:
        # ── Tambah baris baru ─────────────────────────────────
        log.info(f"[MIGO] Tambah baris | Material={item.material} | Qty={qty_str}")
        # Klik tombol "+" atau Insert Row
        pyautogui.hotkey("ctrl", "Insert")
        time.sleep(0.5)

    # ── Isi detail item ───────────────────────────────────────
    # Navigasi ke area item (baris tabel bawah)
    # Field: Material
    pyautogui.hotkey("ctrl", "F")  # cari field material di grid
    time.sleep(0.3)
    pyautogui.press("escape")
    time.sleep(0.2)

    # Tab ke kolom Material di baris item
    tab_to(1)
    type_field(item.material)
    tab_to(1)

    # Qty
    type_field(qty_str)
    tab_to(1)

    # Unit (biarkan default)
    tab_to(1)

    # Plant
    type_field(item.plant)
    tab_to(1)

    # SLoc
    type_field(item.sloc)
    tab_to(1)

    # Cost Center
    type_field(cc)
    time.sleep(0.3)

    log.info(f"[MIGO] Item terisi: {item.material} | {qty_str} | {item.mvt_type}")


def input_migo_batch(items: list, plant_cc: dict) -> str:
    """
    Input semua item selisih ke MIGO_GI dalam satu dokumen.
    Semua item dalam satu plant + movement type yang sama digabung.
    Return: nomor dokumen MIGO (string), atau "" jika gagal baca doc number.

    Alur:
    1. Buka MIGO_GI
    2. Isi header (movement type, posting date, plant)
    3. Isi setiap item di baris tabel
    4. Post dokumen (Ctrl+S atau tombol Post)
    5. Baca nomor dokumen dari status bar
    """
    if not items:
        log.warning("[MIGO] Tidak ada item untuk di-input!")
        return ""

    focus_sap()
    log.info(f"[MIGO] Mulai batch input | {len(items)} item")

    # Buka MIGO_GI
    sap_tcode("MIGO_GI")
    time.sleep(3)

    focus_sap("MIGO")
    pyautogui.hotkey("ctrl", "Home")
    time.sleep(0.5)

    first_item = items[0]

    # ── HEADER ───────────────────────────────────────────────
    # Field Movement Type — ketik langsung
    type_field(first_item.mvt_type)
    time.sleep(0.3)
    pyautogui.press("enter")
    time.sleep(1.5)

    # Posting Date
    # Cari field Posting Date dengan Tab
    tab_to(1)
    type_field(first_item.posting_date)
    tab_to(1)
    time.sleep(0.3)

    log.info(f"[MIGO] Header: Mvt={first_item.mvt_type} | Date={first_item.posting_date}")

    # ── ITEM LINES ────────────────────────────────────────────
    for idx, item in enumerate(items):
        log.info(
            f"[MIGO] Item {idx+1}/{len(items)} | "
            f"{item.material} | SLoc={item.sloc} | "
            f"Qty={format_qty_for_sap(item.qty_adjust)} | Mvt={item.mvt_type}"
        )

        cc = plant_cc.get(item.plant, "")
        if not cc:
            log.warning(f"[MIGO] Cost center plant {item.plant} tidak ada — skip item ini")
            continue

        qty_str = format_qty_for_sap(item.qty_adjust)

        if idx == 0:
            # Baris pertama sudah ada — langsung isi
            # Navigasi ke area tabel item (Ctrl+Home lalu Tab ke grid)
            pyautogui.hotkey("ctrl", "Home")
            time.sleep(0.3)
            # Tab ke field Material di baris pertama tabel item
            # Jumlah tab ini bisa berbeda tergantung layout SAP kamu
            # Sesuaikan angka tab_to() jika perlu
            tab_to(8, delay=0.1)
        else:
            # Tambah baris baru
            pyautogui.hotkey("ctrl", "Insert")
            time.sleep(0.5)

        # Isi Material
        type_field(item.material)
        tab_to(1)
        time.sleep(0.2)

        # Isi Qty
        type_field(qty_str)
        tab_to(1)
        time.sleep(0.2)

        # Skip Unit (biarkan default — biasanya auto-fill)
        tab_to(1)
        time.sleep(0.2)

        # Isi Plant
        type_field(item.plant)
        tab_to(1)
        time.sleep(0.2)

        # Isi Storage Location
        type_field(item.sloc)
        tab_to(1)
        time.sleep(0.2)

        # Isi Cost Center
        type_field(cc)
        tab_to(1)
        time.sleep(0.3)

        pyautogui.press("enter")
        time.sleep(0.5)

    # ── POST DOKUMEN ──────────────────────────────────────────
    log.info("[MIGO] Posting dokumen...")
    time.sleep(1)

    # Tombol Post = Ctrl+S di MIGO
    pyautogui.hotkey("ctrl", "s")
    time.sleep(4)

    # ── BACA NOMOR DOKUMEN ────────────────────────────────────
    doc_number = _read_migo_doc_number()
    if doc_number:
        log.info(f"[MIGO] ✓ Berhasil! Doc number: {doc_number}")
    else:
        log.warning("[MIGO] Dokumen mungkin berhasil dipost tapi nomor tidak terbaca")

    return doc_number


def _read_migo_doc_number() -> str:
    """
    Baca nomor dokumen MIGO dari status bar SAP setelah posting.
    SAP biasanya menampilkan: 'Material document 5000012345 posted'
    Return: nomor dokumen (string), atau "" jika tidak terbaca.
    """
    try:
        # Ambil screenshot status bar area (pojok kiri bawah SAP)
        # Cara lebih reliable: baca via pyautogui + OCR
        # Untuk sementara: coba baca via clipboard
        # Klik status bar untuk select text
        import pyperclip

        # Coba grab teks dari status bar SAP via keyboard shortcut
        # Di SAP, Ctrl+Shift+S kadang bisa copy status message
        # Cara paling simpel: screenshot + return kosong, user lihat sendiri
        time.sleep(1)
        pyautogui.screenshot(
            os.path.join(Config.FOLDER_SCREENSHOTS, 
                        f"migo_result_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png")
        )

        # Coba baca dari status bar dengan win32gui
        hwnd = get_sap_hwnd("MIGO")
        if not hwnd:
            hwnd = get_sap_hwnd("SAP")

        if hwnd:
            # Ambil semua child window text untuk cari doc number
            texts = []
            def cb(h, _):
                t = win32gui.GetWindowText(h)
                if t:
                    texts.append(t)
            win32gui.EnumChildWindows(hwnd, cb, None)

            for t in texts:
                # Pattern: angka 10 digit (nomor dokumen material SAP)
                match = re.search(r'\b(5\d{9}|4\d{9})\b', t)
                if match:
                    return match.group(1)

    except Exception as e:
        log.warning(f"[MIGO] Gagal baca doc number: {e}")

    return ""


# ─────────────────────────────────────────────
# MAIN - ORCHESTRATOR
# ─────────────────────────────────────────────

def main():
    log.info("=" * 60)
    log.info(f"RPA Stock Recon mulai: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    log.info("=" * 60)

    try:
        # Fase 1, 2, 3 — export SAP, ambil Matrix, compare
        from rpa_phase1_2 import run_phase1_2_3
        items_per_plant = run_phase1_2_3(Config.PLANTS)

        if not items_per_plant:
            log.info("[INFO] Tidak ada selisih hari ini. Robot selesai.")
            return

        total_item = sum(len(v) for v in items_per_plant.values())
        log.info(
            f"[INFO] Compare selesai | {len(items_per_plant)} plant | "
            f"{total_item} item selisih"
        )

        # Fase 4 — Kirim laporan Excel ke accounting
        log.info("[FASE4] Buat laporan dan kirim email ke accounting...")
        from send_email_report import send_stock_diff_report
        excel_path = send_stock_diff_report(items_per_plant)
        log.info(f"[FASE4] Email terkirim | File: {excel_path}")

    except Exception as e:
        log.critical(f"[FATAL] Robot berhenti: {e}")

    finally:
        log.info("=" * 60)
        log.info(f"RPA selesai: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        log.info("=" * 60)


if __name__ == "__main__":
    main()
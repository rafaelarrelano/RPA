"""
rpa_phase1_2.py
Fase 1: Export stok SAP via ZPGD_SAPSTK per plant (pyautogui)
Fase 2: Ambil data Matrix dari portal (Playwright CDP)
Fase 3: Compare Matrix vs SAP, kelompokkan per plant
"""

import os
import glob
import time
import pyautogui
import win32gui
import win32con
from datetime import datetime
from collections import defaultdict
from playwright.sync_api import sync_playwright

from config import Config
from logger import setup_logger
from main import (
    StockDiff, parse_decimal,
    tab_to, type_field, focus_sap, sap_tcode, get_sap_hwnd
)

log = setup_logger()

pyautogui.FAILSAFE = True
pyautogui.PAUSE    = 0.6


# ─────────────────────────────────────────────
# FASE 1 - EXPORT ZPGD_SAPSTK PER PLANT
# ─────────────────────────────────────────────

def run_zpgd_sapstk(plant: str, send_log=None) -> str:
    """
    Jalankan T-code SAPSTK untuk satu plant via pyautogui.
    T-code diambil dari Config.ACTIVE_TCODE_SAPSTK yang di-set GUI saat user
    memilih portal EOD:
      - Portal PGDMTX → /NZPGD_SAPSTK
      - Portal CMIS   → /NZCNS_SAPSTK
    Download file SAPSTK terbaru ke SAP_DOWNLOAD_DIR.
    Return: path file .txt yang baru didownload.
    """
    def _log(msg, level="INFO"):
        if send_log:
            send_log(msg, level)
        else:
            log.info(f"[FASE1] {msg}")

    # Baca T-code aktif dari Config (sudah di-set GUI berdasarkan pilihan portal)
    tcode = getattr(Config, "ACTIVE_TCODE_SAPSTK", "ZPGD_SAPSTK")

    def _wait_sap_window(keywords: list, timeout: int = 60) -> int:
        """
        Tunggu sampai window SAP dengan salah satu keyword judul muncul.
        Poll tiap 0.5 detik, max `timeout` detik.
        Return hwnd jika ketemu, raise Exception jika timeout.
        """
        deadline = time.time() + timeout
        while time.time() < deadline:
            found = []
            def _cb(h, _):
                if win32gui.IsWindowVisible(h):
                    t = win32gui.GetWindowText(h)
                    if any(kw in t for kw in keywords):
                        found.append(h)
            win32gui.EnumWindows(_cb, None)
            if found:
                return found[0]
            time.sleep(0.5)
        raise Exception(
            f"Window SAP dengan keyword {keywords} tidak muncul dalam {timeout} detik!"
        )

    def _wait_sapstk_file(pattern: str, after_ts: float,
                          timeout: int = 120) -> str:
        """
        Poll folder download sampai file SAPSTK baru muncul.
        `after_ts` = timestamp sebelum F8 ditekan.
        Timeout default 120 detik (file besar bisa lama).
        Return path file terbaru.
        """
        deadline = time.time() + timeout
        while time.time() < deadline:
            files = [f for f in glob.glob(pattern)
                     if os.path.getmtime(f) > after_ts]
            if files:
                latest = max(files, key=os.path.getmtime)
                # Tunggu file selesai ditulis (ukuran stabil 2x check)
                prev_size = -1
                stable    = 0
                for _ in range(20):          # max 10 detik stabilitas
                    size = os.path.getsize(latest)
                    if size > 0 and size == prev_size:
                        stable += 1
                        if stable >= 2:
                            return latest
                    else:
                        stable = 0
                    prev_size = size
                    time.sleep(0.5)
                return latest               # kembalikan meskipun belum stabil
            time.sleep(1.0)
        raise FileNotFoundError(
            f"File SAPSTK tidak muncul dalam {timeout} detik di {Config.SAP_DOWNLOAD_DIR}"
        )

    try:
        _log(f"Export /{tcode} plant {plant}...")
        before = time.time()

        # Navigasi ke T-code — tunggu window muncul (adaptif, max 60 detik)
        sap_tcode(f"n{tcode}")
        _log(f"Menunggu window /{tcode} muncul...")
        hwnd = _wait_sap_window(
            ["Program transfer", "Report Flow Transfer", tcode],
            timeout=60
        )

        win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)
        win32gui.SetForegroundWindow(hwnd)
        time.sleep(0.8)

        pyautogui.hotkey("ctrl", "Home")
        time.sleep(0.4)

        # Tab 0: Pilihan = 1 (ALL)
        type_field("1")

        # Tab 1: Plant Code
        tab_to(1)
        type_field(plant)

        # Execute F8
        _log(f"Execute F8 plant {plant}...")
        pyautogui.press("f8")

        # Tunggu popup SAP GUI Security dengan polling (max 30 detik)
        _log("Menunggu popup Allow SAP GUI Security...")
        popup_deadline = time.time() + 30
        popup_found    = False
        while time.time() < popup_deadline:
            popups = []
            def _pcb(h, _):
                if win32gui.IsWindowVisible(h):
                    t = win32gui.GetWindowText(h)
                    if "Security" in t or "Allow" in t:
                        popups.append(h)
            win32gui.EnumWindows(_pcb, None)
            if popups:
                win32gui.SetForegroundWindow(popups[0])
                time.sleep(0.3)
                pyautogui.press("left")   # pilih Allow
                time.sleep(0.2)
                pyautogui.press("enter")
                popup_found = True
                _log("Popup Allow diklik", "OK")
                break
            time.sleep(0.5)

        if not popup_found:
            _log("Popup Allow tidak terdeteksi — lanjut polling file", "WARN")

        # Tunggu file SAPSTK muncul di folder download (polling, max 120 detik)
        _log(f"Menunggu file SAPSTK plant {plant} di {Config.SAP_DOWNLOAD_DIR}...")
        pattern = os.path.join(Config.SAP_DOWNLOAD_DIR, f"{plant}_*_SAPSTK_*.TXT")
        latest  = _wait_sapstk_file(pattern, before, timeout=120)

        _log(f"[/{tcode}] File berhasil: {os.path.basename(latest)}", "OK")
        return latest

    except Exception as e:
        _log(f"[/{tcode}] Export gagal plant {plant}: {e}", "ERROR")
        raise


def parse_sapstk_file(filepath: str) -> dict:
    """
    Baca file .txt SAPSTK.
    Return: { (material, sloc, param): qty_sap }
    Key 3 elemen supaya konsisten dengan compare() di test_compare.py.
    Baca FSTKGD dan FSTKVN, skip material kepala 2 dan 7.
    """
    VALID_PARAMS = {"FSTKGD", "FSTKVN"}
    stok_sap = {}
    try:
        with open(filepath, "r", encoding="utf-8", errors="ignore") as f:
            for line in f:
                line = line.strip()
                if not line:
                    continue
                prefix = line.split("|")[0].strip() if "|" in line else ""
                if prefix not in VALID_PARAMS:
                    continue
                fields = line.split("|")
                if len(fields) < 6:
                    continue
                sloc     = fields[2].strip()
                material = fields[4].strip()
                qty_str  = fields[5].strip()

                if not qty_str:
                    continue
                if material.startswith("7") or material.startswith("2"):
                    continue

                try:
                    qty_sap = parse_decimal(qty_str)
                except Exception:
                    continue

                key = (material, sloc, prefix)
                stok_sap[key] = stok_sap.get(key, 0.0) + qty_sap

        fstkgd_c = sum(1 for k in stok_sap if k[2] == "FSTKGD")
        fstkvn_c = sum(1 for k in stok_sap if k[2] == "FSTKVN")
        log.info(
            f"[FASE1] {len(stok_sap)} material+sloc dari {os.path.basename(filepath)} "
            f"(FSTKGD={fstkgd_c}, FSTKVN={fstkvn_c})"
        )
    except Exception as e:
        log.error(f"[FASE1] Gagal baca file SAPSTK: {e}")
        raise
    return stok_sap


# ─────────────────────────────────────────────
# FASE 2 - AMBIL DATA MATRIX DARI PORTAL
# ─────────────────────────────────────────────

def get_matrix_from_portal_cdp(plant: str) -> dict:
    """
    Ambil data Matrix dari portal via Chrome CDP.
    Chrome harus sudah terbuka dengan --remote-debugging-port=9222
    dan halaman ViewDetail sudah terbuka di tab INPUT.
    Return: { (material, sloc, param): {'qty', 'sloc', 'tgl', 'param'} }
    """
    with sync_playwright() as p:
        browser = p.chromium.connect_over_cdp("http://localhost:9222")

        page = None
        for ctx in browser.contexts:
            for pg in ctx.pages:
                if "ViewDetail" in pg.url:
                    page = pg
                    break

        if not page:
            raise Exception("Halaman ViewDetail tidak ditemukan! "
                            "Pastikan portal sudah terbuka di Chrome.")

        log.info(f"[FASE2] Ambil data dari: {page.url}")

        result = page.evaluate("""() => {
            const els = document.querySelectorAll('textarea, pre, .tab-pane.active');
            const texts = Array.from(els)
                .map(e => e.value || e.innerText || e.textContent)
                .filter(t => t && t.length > 10);
            return texts.join('\\n');
        }""")

        browser.close()

    matrix    = {}
    skipped_7 = 0
    for line in result.splitlines():
        line = line.strip()
        if not line.startswith("FSTKGD"):
            continue
        fields = line.split("|")
        if len(fields) < 6:
            continue
        if fields[1].strip() != plant:
            continue

        sloc     = fields[2].strip()
        tgl      = fields[3].strip()
        material = fields[4].strip()
        qty_str  = fields[5].strip()
        prefix   = fields[0].strip()   # FSTKGD atau FSTKVN

        if not qty_str:
            continue

        # Skip material kepala 2 dan 7
        if material.startswith("7") or material.startswith("2"):
            skipped_7 += 1
            continue

        try:
            qty_matrix = parse_decimal(qty_str)
        except Exception:
            continue

        key = (material, sloc, prefix)
        matrix[key] = {
            "qty":  qty_matrix,
            "sloc": sloc,
            "tgl":  tgl,
            "param": prefix,
        }

    log.info(f"[FASE2] Matrix plant {plant}: {len(matrix)} material | skip kepala 2&7: {skipped_7}")
    return matrix


# ─────────────────────────────────────────────
# FASE 3 - COMPARE & HITUNG SELISIH PER PLANT
# ─────────────────────────────────────────────

def compare_plant(
    plant: str,
    matrix: dict,
    stok_sap: dict,
    posting_date: str,
    tolerance: float = 0.0
) -> list:
    """
    Bandingkan Matrix vs SAP per material+sloc.
    posting_date per item diambil dari field 'tgl' di data Matrix portal
    (format YYYYMMDD → dikonversi ke dd.mm.yyyy).
    Fallback ke parameter posting_date jika tgl kosong/tidak valid.
    Return: list StockDiff yang perlu adjustment.
    """
    from main import convert_date

    items = []
    for (material, sloc, param), data in matrix.items():
        qty_matrix = data["qty"]
        qty_sap    = stok_sap.get((material, sloc, param), 0.0)
        diff       = round(qty_matrix - qty_sap, 6)

        if abs(diff) <= tolerance:
            continue

        # Ambil tanggal dari data Matrix portal (field tgl = "YYYYMMDD")
        tgl_raw = data.get("tgl", "").strip()
        try:
            item_posting_date = convert_date(tgl_raw)  # "20260504" → "04.05.2026"
        except Exception:
            item_posting_date = posting_date  # fallback ke datetime.now()

        item = StockDiff(
            param        = param,
            plant        = plant,
            sloc         = sloc,
            posting_date = item_posting_date,
            material     = material,
            qty_matrix   = qty_matrix,
            qty_sap      = qty_sap,
            diff         = diff,
            status       = 0,
        )

        if diff < 0:
            item.mvt_type   = "917"
            item.qty_adjust = abs(diff)
        else:
            item.mvt_type   = "918"
            item.qty_adjust = diff

        items.append(item)

    log.info(f"[FASE3] Plant {plant}: {len(items)} item selisih")
    return items


# ─────────────────────────────────────────────
# MAIN FASE 1-2-3
# ─────────────────────────────────────────────

def run_phase1_2_3(plants: list) -> dict:
    """
    Jalankan Fase 1, 2, 3 untuk semua plant.
    Return: { plant: [StockDiff] } siap untuk Fase 4 (MIGO)
    """
    from limit_adjustment import load_limit_adjustment, filter_by_limit

    posting_date     = datetime.now().strftime("%d.%m.%Y")
    items_per_plant  = {}
    limits           = load_limit_adjustment()

    for plant in plants:
        log.info(f"[RUN] Proses plant {plant}")
        try:
            # Fase 1: Export ZPGD_SAPSTK
            filepath = run_zpgd_sapstk(plant)
            stok_sap = parse_sapstk_file(filepath)

            # Fase 2: Ambil Matrix dari portal
            matrix = get_matrix_from_portal_cdp(plant)

            if not matrix:
                log.info(f"[RUN] Plant {plant}: tidak ada data Matrix")
                continue

            # Fase 3: Compare
            items = compare_plant(plant, matrix, stok_sap, posting_date,
                                  Config.DIFF_THRESHOLD)

            if not items:
                log.info(f"[RUN] Plant {plant}: tidak ada selisih")
                continue

            # Apply limit adjustment
            items_ok, items_exceeded = filter_by_limit(items, limits)

            if items_exceeded:
                log.warning(f"[RUN] Plant {plant}: {len(items_exceeded)} item lewat batas limit")

            if items_ok:
                items_per_plant[plant] = items_ok
                log.info(f"[RUN] Plant {plant}: {len(items_ok)} item siap ke MIGO")

        except Exception as e:
            log.error(f"[RUN] Plant {plant} gagal: {e}")
            continue

    log.info(f"[RUN] Selesai | {len(items_per_plant)} plant ada selisih")
    return items_per_plant


if __name__ == "__main__":
    # Test satu plant
    plant  = "4502"
    result = run_phase1_2_3([plant])
    for p, items in result.items():
        print(f"\nPlant {p}: {len(items)} item")
        for item in items[:5]:
            print(f"  {item.material} | {item.sloc} | diff={item.diff} | {item.mvt_type} | adj={item.qty_adjust}")
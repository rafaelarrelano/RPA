"""
rpa_phase1_2.py
Fase 1: Export stok SAP via ZPGD_SAPSTK per plant (pyautogui)
Fase 2: Ambil data Matrix dari portal (Playwright CDP)
Fase 3: Compare Matrix vs SAP, kelompokkan per plant
"""

import os
import glob
import re
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


def _interruptible_sleep(duration: float):
    """Sleep yang bisa di-interrupt dengan check stop_event tiap 0.1s."""
    if _is_stopped():
        return
    end_time = time.time() + duration
    while time.time() < end_time:
        if _is_stopped():
            return
        remaining = end_time - time.time()
        sleep_chunk = min(0.1, remaining)
        if sleep_chunk > 0:
            time.sleep(sleep_chunk)

log = setup_logger()

pyautogui.FAILSAFE = True
pyautogui.PAUSE    = 0.6


# ─────────────────────────────────────────────
# CHECK STOP EVENT
# ─────────────────────────────────────────────

def _is_stopped() -> bool:
    """Cek apakah user klik Stop di GUI."""
    try:
        from rpa_gui import stop_event
        return stop_event.is_set()
    except Exception:
        return False


def _copy_text_to_clipboard(text: str):
    """Salin teks ke clipboard untuk paste ke SAP."""
    try:
        import pyperclip
        pyperclip.copy(text)
    except Exception:
        import tkinter as tk
        root = tk.Tk()
        root.withdraw()
        root.clipboard_clear()
        root.clipboard_append(text)
        root.update()
        root.destroy()


def _wait_sap_window(keywords: list, timeout: int = 90) -> int:
    """
    Tunggu sampai window SAP dengan salah satu keyword judul muncul.
    Poll tiap 0.5 detik, max `timeout` detik.
    Return hwnd jika ketemu, raise Exception jika timeout.
    """
    deadline = time.time() + timeout
    while time.time() < deadline:
        if _is_stopped():
            raise Exception("Robot dihentikan oleh user")

        found = []
        def _cb(h, _):
            if win32gui.IsWindowVisible(h):
                t = win32gui.GetWindowText(h)
                if any(kw in t for kw in keywords):
                    found.append(h)
        win32gui.EnumWindows(_cb, None)
        if found:
            return found[0]
        _interruptible_sleep(0.5)
    raise Exception(
        f"Window SAP dengan keyword {keywords} tidak muncul dalam {timeout} detik!"
    )


def _wait_sapstk_file(pattern: str, after_ts: float,
                      timeout: int = 300, log_fn=None) -> str:
    """
    Poll folder download sampai file SAPSTK baru muncul.
    `after_ts` = timestamp sebelum F8 ditekan.
    Timeout default 300 detik — tunggu sampai file benar-benar terdownload.
    Tiap 30 detik cetak log progress agar tidak terlihat hang.
    Return path file terbaru.
    """
    deadline      = time.time() + timeout
    last_log_time = time.time()
    while time.time() < deadline:
        if _is_stopped():
            raise Exception("Robot dihentikan oleh user")

        if time.time() - last_log_time >= 30:
            elapsed = int(time.time() - after_ts)
            if log_fn:
                log_fn(f"Masih menunggu file SAPSTK... ({elapsed}s)", "INFO")
            else:
                log.info(f"[FASE1] Masih menunggu file SAPSTK... ({elapsed}s)")
            last_log_time = time.time()

        files = [f for f in glob.glob(pattern) if os.path.getmtime(f) > after_ts]
        if files:
            latest = max(files, key=os.path.getmtime)
            prev_size = -1
            stable    = 0
            for _ in range(20):
                if _is_stopped():
                    raise Exception("Robot dihentikan oleh user")

                size = os.path.getsize(latest)
                if size > 0 and size == prev_size:
                    stable += 1
                    if stable >= 2:
                        return latest
                else:
                    stable = 0
                prev_size = size
                _interruptible_sleep(0.5)
            return latest
        _interruptible_sleep(1.0)
    raise FileNotFoundError(
        f"File SAPSTK tidak muncul dalam {timeout} detik di {Config.SAP_DOWNLOAD_DIR}"
    )


# ─────────────────────────────────────────────
# FASE 1 - EXPORT ZPGD_SAPSTK PER PLANT
# ─────────────────────────────────────────────

def run_zpgd_sapstk(plant: str | list[str], send_log=None) -> str:
    """
    Jalankan T-code SAPSTK via pyautogui.
    T-code diambil dari Config.ACTIVE_TCODE_SAPSTK yang di-set GUI saat user
    memilih portal EOD:
      - Portal PGDMTX → /NZPGD_SAPSTK
      - Portal CMIS   → /NZCNS_SAPSTK
    `plant` dapat berupa string untuk satu plant atau list/string newline
    untuk beberapa plant.
    Paste plant ke multiple selection dengan urutan vertikal (satu plant per baris).
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

    try:
        _log(f"Export /{tcode} plant {plant}...")
        before = time.time()

        # Navigasi ke T-code — tunggu window muncul (max 90 detik)
        sap_tcode(f"n{tcode}")
        _log(f"Menunggu window /{tcode} muncul...")
        hwnd = _wait_sap_window(
            ["Program transfer", "Report Flow Transfer", tcode],
            timeout=90
        )

        win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)
        win32gui.SetForegroundWindow(hwnd)
        _interruptible_sleep(0.8)

        pyautogui.hotkey("ctrl", "Home")
        _interruptible_sleep(0.4)

        # Field pertama: isi dengan 2
        type_field("2")

        # Tab 3 kali lalu Enter ke tombol multiple selection
        tab_to(3)
        pyautogui.press("enter")

        # Hapus data lama di multiple selection
        pyautogui.hotkey("shift", "f4")
        _interruptible_sleep(0.4)

        # Paste plant list ke multiple selection, satu plant per baris
        if isinstance(plant, str):
            plant_list = [p.strip() for p in re.split(r"[\r\n,]+", plant) if p.strip()]
        else:
            plant_list = [str(p).strip() for p in plant if str(p).strip()]
        if not plant_list:
            raise ValueError("Daftar plant tidak boleh kosong")

        plant_rows = "\r\n".join(plant_list)
        _copy_text_to_clipboard(plant_rows)
        debug_path = os.path.join(Config.SAP_DOWNLOAD_DIR, "sapstk_multi_plant_debug.txt")
        try:
            with open(debug_path, "w", encoding="utf-8") as debug_file:
                debug_file.write(
                    f"RAW PLANT TYPE: {type(plant).__name__}\n"
                    f"RAW PLANT REPR: {repr(plant)}\n"
                    f"PLANT LIST ({len(plant_list)}): {plant_list!r}\n"
                    f"PLANT ROWS REPR: {plant_rows!r}\n"
                    "-----\n"
                    f"{plant_rows}\n"
                )
            _log(f"Debug plant list ditulis ke: {debug_path}")
        except Exception:
            _log("Gagal tulis debug plant list ke file.", "WARN")

        _log(f"Plant raw input type={type(plant).__name__} repr={repr(plant)}")
        _log(f"Plant list split ({len(plant_list)}): {plant_list}")
        _log(f"Plant list untuk multiple selection ({len(plant_list)}):\n{plant_rows}")
        pyautogui.hotkey("shift", "f12")
        _interruptible_sleep(0.4)

        # Execute multiple selection
        _log(f"Execute F8 multiple selection plant {plant}...")
        pyautogui.press("f8")
        _interruptible_sleep(0.8)

        # Execute program sekali lagi
        _log(f"Execute F8 program plant {plant}...")
        pyautogui.press("f8")

        # Tunggu file SAPSTK muncul di folder download (polling, max 300 detik)
        _log(f"Menunggu file SAPSTK plant {plant} di {Config.SAP_DOWNLOAD_DIR}...")
        # Pertama coba baca pesan SAP (Ctrl+C) untuk dapatkan nama file persis.
        # Jika berhasil dan file ada di folder download → pakai itu.
        try:
            win32gui.SetForegroundWindow(hwnd)
            _interruptible_sleep(0.4)
            import pyperclip
            clipboard_text = ""
            for _ in range(3):
                try:
                    pyautogui.hotkey("ctrl", "c")
                    _interruptible_sleep(0.25)
                    clipboard_text = pyperclip.paste() or ""
                    if clipboard_text and "SAPSTK" in clipboard_text.upper():
                        break
                except Exception:
                    _interruptible_sleep(0.2)

            if clipboard_text and "SAPSTK" in clipboard_text.upper():
                _log("Mencoba ekstrak nama file SAP dari clipboard...")
                txt = clipboard_text.strip()
                up = txt.upper()
                idx = up.find("SAPSTK")
                if idx != -1:
                    # cari batas kiri (backslash atau space) dan batas kanan .TXT
                    left = txt.rfind("\\", 0, idx)
                    if left == -1:
                        left = txt.rfind(" ", 0, idx)
                    right = up.find('.TXT', idx)
                    if right != -1:
                        candidate_name = txt[left+1:right+4] if left != -1 else txt[:right+4]
                        candidate_name = candidate_name.strip('"').strip()
                        # jika clipboard berisi full path gunakan itu, else combine with download dir
                        if os.path.isabs(candidate_name) and os.path.exists(candidate_name):
                            _log(f"Ditemukan file SAP dari clipboard: {candidate_name}")
                            return candidate_name
                        else:
                            candidate_path = os.path.join(Config.SAP_DOWNLOAD_DIR, os.path.basename(candidate_name))
                            if os.path.exists(candidate_path):
                                _log(f"Ditemukan file SAP di folder download: {candidate_path}")
                                return candidate_path
            else:
                _log("Clipboard tidak berisi nama file SAP atau gagal ambil clipboard", "WARN")
        except Exception as _e:
            _log(f"Ekstraksi nama file dari SAP clipboard gagal: {_e}", "WARN")

        # Gunakan pola yang lebih longgar untuk mencakup variasi penamaan file SAP.
        # Sebelumnya pola khusus single-plant mengasumsikan plant ada di awal nama
        # file; beberapa sistem SAP menamai file berbeda sehingga file tidak terdeteksi.
        pattern = os.path.join(Config.SAP_DOWNLOAD_DIR, "*_*_SAPSTK_*.TXT")

        # Tunggu file baru muncul, lalu verifikasi bahwa file tersebut
        # memang berisi data untuk plant yang diminta. Jika tidak, lanjut
        # tunggu file berikutnya sampai timeout total tercapai.
        total_deadline = time.time() + 300
        current_before = before
        latest = None
        while time.time() < total_deadline:
            # hitungan timeout untuk _wait_sapstk_file: sisa waktu hingga total_deadline
            remaining = int(total_deadline - time.time())
            try:
                candidate = _wait_sapstk_file(pattern, current_before, timeout=remaining, log_fn=_log)
            except FileNotFoundError:
                break

            # Verifikasi isi file — ambil plant yang ada di file
            try:
                stok_candidate = parse_sapstk_file(candidate)
                plants_in_file = {k[0] for k in stok_candidate.keys()}
            except Exception:
                plants_in_file = set()

            # Jika ada minimal satu plant yang diminta ada di file → terima
            if any(p in plants_in_file for p in plant_list):
                latest = candidate
                break

            # Jika tidak cocok, catat dan tunggu file berikutnya (update before)
            _log(f"File {os.path.basename(candidate)} tidak berisi plant yang diminta — tunggu file berikutnya", "WARN")
            current_before = os.path.getmtime(candidate)

        if not latest:
            raise FileNotFoundError(
                f"File SAPSTK yang berisi plant {plant_list} tidak ditemukan dalam {Config.SAP_DOWNLOAD_DIR}"
            )

        _log(f"[/{tcode}] File berhasil: {os.path.basename(latest)}", "OK")
        return latest

    except Exception as e:
        _log(f"[/{tcode}] Export gagal plant {plant}: {e}", "ERROR")
        raise


def parse_sapstk_file(filepath: str) -> dict:
    """
    Baca file .txt SAPSTK.
    Return: { (plant, material, sloc, param): qty_sap }
    Key 4 elemen supaya bisa membandingkan satu file SAPSTK multi-plant.
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
                plant     = fields[1].strip()
                sloc      = fields[2].strip()
                material  = fields[4].strip()
                qty_str   = fields[5].strip()

                if not qty_str:
                    continue
                if material.startswith("7") or material.startswith("2"):
                    continue

                try:
                    qty_sap = parse_decimal(qty_str)
                except Exception:
                    continue

                key = (plant, material, sloc, prefix)
                stok_sap[key] = stok_sap.get(key, 0.0) + qty_sap

        fstkgd_c = sum(1 for k in stok_sap if k[3] == "FSTKGD")
        fstkvn_c = sum(1 for k in stok_sap if k[3] == "FSTKVN")
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
        qty_sap    = stok_sap.get((plant, material, sloc, param), 0.0)
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
    posting_date    = datetime.now().strftime("%d.%m.%Y")
    items_per_plant = {}

    log.info(f"[RUN] Eksekusi SAPSTK multi-plant: {plants}")
    filepath = run_zpgd_sapstk(plants)
    stok_sap = parse_sapstk_file(filepath)

    for plant in plants:
        log.info(f"[RUN] Proses plant {plant}")
        try:
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

            # Semua item langsung masuk — tidak ada filter limit adjustment
            items_per_plant[plant] = items
            log.info(f"[RUN] Plant {plant}: {len(items)} item siap ke MIGO")

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
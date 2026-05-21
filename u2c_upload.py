"""
u2c_upload.py
Step baru setelah compare:
  1. Buat file .txt format U2C dari hasil selisih
  2. Upload ke SAP via T-code nZPGD_U2C (input lokasi file)

Format baris U2C:
  FSTKGD|{plant}|{sloc}|{tanggal_YYYYMMDD}|{material}||{qty_matrix_koma}

Contoh:
  FSTKGD|4507|WH01|20260507|378025||0,500

Catatan:
- Kolom ke-6 (index 5) selalu kosong
- qty_matrix pakai format koma (contoh: 0,500 bukan 0.500)
- Tanggal format YYYYMMDD (bukan dd.mm.yyyy)
- Satu file per plant, disimpan di U2C_OUTPUT_DIR
"""

import os
import time
import pyautogui
import win32gui
import win32con
from datetime import datetime

from config import Config
from logger import setup_logger

log = setup_logger()

pyautogui.FAILSAFE = True
pyautogui.PAUSE    = 0.5

# ─────────────────────────────────────────────
# KONFIGURASI & PERSISTENT CONFIG
# ─────────────────────────────────────────────

# Default path — bisa di-override via GUI atau load_u2c_config()
U2C_OUTPUT_DIR  = r"C:\Users\User\Documents\PGD\EOD"
U2C_FIXED_NAME  = "U2C.txt"   # nama file tetap, di-overwrite setiap run

# File config path (simpan pilihan user)
_U2C_CONFIG_FILE = os.path.join(
    os.path.dirname(os.path.abspath(__file__)),
    "config", "u2c_config.json"
)


def load_u2c_config() -> dict:
    """
    Load konfigurasi U2C dari file JSON.
    Return: { 'u2c_filepath': '...path lengkap ke U2C.txt...' }
    """
    import json
    try:
        with open(_U2C_CONFIG_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {"u2c_filepath": os.path.join(U2C_OUTPUT_DIR, U2C_FIXED_NAME)}


def save_u2c_config(u2c_filepath: str):
    """
    Simpan path file U2C ke config JSON.
    Dipanggil dari GUI saat user mengubah path.
    """
    import json
    os.makedirs(os.path.dirname(_U2C_CONFIG_FILE), exist_ok=True)
    with open(_U2C_CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump({"u2c_filepath": u2c_filepath}, f, indent=2)
    log.info(f"[U2C] Config disimpan: {u2c_filepath}")


def get_u2c_filepath() -> str:
    """
    Ambil path file U2C yang akan dipakai (dari config atau default).
    """
    cfg = load_u2c_config()
    return cfg.get("u2c_filepath", os.path.join(U2C_OUTPUT_DIR, U2C_FIXED_NAME))


# ─────────────────────────────────────────────
# FORMAT HELPER
# ─────────────────────────────────────────────

def _date_to_sap(posting_date: str) -> str:
    """
    Konversi tanggal dd.mm.yyyy → YYYYMMDD untuk format U2C.
    Contoh: '07.05.2026' → '20260507'
    """
    try:
        d = datetime.strptime(posting_date, "%d.%m.%Y")
        return d.strftime("%Y%m%d")
    except Exception:
        # Kalau sudah format YYYYMMDD, kembalikan langsung
        if len(posting_date) == 8 and posting_date.isdigit():
            return posting_date
        raise ValueError(f"Format tanggal tidak dikenal: {posting_date!r}")


def _qty_to_u2c(qty: float) -> str:
    """
    Format qty untuk U2C: pakai koma sebagai desimal.
    Contoh: 0.5 → '0,500' | 65.0 → '65,000' | 6.375 → '6,375'
    """
    return f"{qty:.3f}".replace(".", ",")


# ─────────────────────────────────────────────
# BUAT FILE U2C
# ─────────────────────────────────────────────

def build_u2c_file(plant: str, items: list, output_dir: str = None) -> str:
    """
    Buat file .txt U2C dari list StockDiff untuk satu plant.
    File di-overwrite setiap run dengan nama tetap.
    Return: path file .txt yang dibuat.
    """
    if output_dir is None:
        output_dir = U2C_OUTPUT_DIR

    os.makedirs(output_dir, exist_ok=True)

    filename = f"U2C_{plant}.txt"
    filepath = os.path.join(output_dir, filename)

    lines = []
    for item in items:
        tgl_sap = _date_to_sap(item.posting_date)
        qty_str = _qty_to_u2c(item.qty_matrix)
        line    = f"FSTKGD|{item.plant}|{item.sloc}|{tgl_sap}|{item.material}||{qty_str}"
        lines.append(line)

    with open(filepath, "w", encoding="utf-8", newline="\n") as f:
        f.write("\n".join(lines))

    log.info(f"[U2C] File di-overwrite: {filename} | {len(lines)} baris")
    return filepath


def build_u2c_all_plants(items_per_plant: dict,
                         filepath: str = None) -> dict:
    """
    Gabungkan semua plant ke SATU file U2C tetap (di-overwrite tiap run).
    filepath: path file U2C (dari config GUI). Jika None → pakai get_u2c_filepath().
    Return: { plant: filepath } — semua plant mengarah ke file yang sama.
    """
    if filepath is None:
        filepath = get_u2c_filepath()

    os.makedirs(os.path.dirname(filepath), exist_ok=True)

    lines = []
    for plant, items in sorted(items_per_plant.items()):
        if not items:
            continue
        for item in items:
            tgl_sap = _date_to_sap(item.posting_date)
            qty_str = _qty_to_u2c(item.qty_matrix)
            line    = f"FSTKGD|{item.plant}|{item.sloc}|{tgl_sap}|{item.material}||{qty_str}"
            lines.append(line)

    with open(filepath, "w", encoding="utf-8", newline="\n") as f:
        f.write("\n".join(lines))

    total = len(lines)
    log.info(f"[U2C] File di-overwrite: {os.path.basename(filepath)} | {total} baris total")
    return {plant: filepath for plant in items_per_plant}




def build_u2c_from_matrix(matrix_per_plant: dict, filepath: str = None) -> int:
    """
    Buat file U2C langsung dari raw Matrix portal (dict dari get_fstkgd_from_view_detail).
    Ini memastikan SEMUA item di tab INPUT tertulis ke U2C,
    bukan hanya yang selisih.

    matrix_per_plant: { plant: { (material, sloc, param): {qty, sloc, tgl, param} } }
    filepath        : path file U2C tetap. Jika None → pakai get_u2c_filepath()
    Return          : jumlah baris yang ditulis
    """
    if filepath is None:
        filepath = get_u2c_filepath()

    os.makedirs(os.path.dirname(filepath), exist_ok=True)

    lines      = []
    fstkgd_cnt = 0
    fstkvn_cnt = 0

    for plant, matrix in sorted(matrix_per_plant.items()):
        for (material, sloc, param), data in matrix.items():
            tgl_raw = data.get("tgl", "").strip()
            # Konversi tgl: bisa YYYYMMDD atau dd.mm.yyyy
            if len(tgl_raw) == 8 and tgl_raw.isdigit():
                tgl_sap = tgl_raw
            else:
                try:
                    tgl_sap = _date_to_sap(tgl_raw)
                except Exception:
                    tgl_sap = datetime.now().strftime("%Y%m%d")

            qty_str = _qty_to_u2c(data["qty"])
            # Gunakan param asli dari portal (FSTKGD atau FSTKVN)
            line = f"{param}|{plant}|{sloc}|{tgl_sap}|{material}||{qty_str}"
            lines.append(line)

            if param == "FSTKVN":
                fstkvn_cnt += 1
            else:
                fstkgd_cnt += 1

    with open(filepath, "w", encoding="utf-8", newline="\n") as f:
        f.write("\n".join(lines))

    total = len(lines)
    log.info(
        f"[U2C] File di-overwrite: {os.path.basename(filepath)} | "
        f"{total} baris (FSTKGD={fstkgd_cnt}, FSTKVN={fstkvn_cnt})"
    )
    return total

# ─────────────────────────────────────────────
# FOKUS SAP WINDOW
# ─────────────────────────────────────────────

SAP_WINDOW_CLASSES = ["SAP_FRONTEND_SESSION", "SAPFrontend", "SAPGUI"]
SAP_TITLE_KEYWORDS = [
    "SAP Easy Access", "SAP R/3", "SAP NetWeaver",
    "ZPGD", "MIGO", "Program transfer",
    "Report Flow Transfer", "Transfer Branch",
    "Import", "Goods Issue",
]
SKIP_TITLES = [
    "SAP Logon",    # launcher — bukan session aktif
    "Firefox", "Chrome", "Edge", "Chromium",
    "Visual Studio", "Code", "Notepad",
    "Claude", "Thunderbird", "Explorer",
]


def focus_sap():
    """Fokuskan window SAP ke depan."""
    priority = []
    fallback = []

    def cb(hwnd, _):
        if not win32gui.IsWindowVisible(hwnd):
            return
        title = win32gui.GetWindowText(hwnd)
        cls   = win32gui.GetClassName(hwnd)
        if not title or any(s in title for s in SKIP_TITLES):
            return
        if any(c in cls for c in SAP_WINDOW_CLASSES):
            priority.append(hwnd)
        elif any(kw in title for kw in SAP_TITLE_KEYWORDS):
            fallback.append(hwnd)

    win32gui.EnumWindows(cb, None)
    hwnd = priority[0] if priority else (fallback[0] if fallback else None)
    if not hwnd:
        raise Exception("SAP window tidak ditemukan!")

    win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)
    win32gui.SetForegroundWindow(hwnd)
    time.sleep(0.8)


def sap_tcode(tcode: str, wait: float = 2.5):
    """Navigasi ke T-code via Ctrl+/."""
    focus_sap()
    time.sleep(0.3)
    pyautogui.keyDown("ctrl")
    time.sleep(0.1)
    pyautogui.press("/")
    pyautogui.keyUp("ctrl")
    time.sleep(0.4)
    pyautogui.hotkey("ctrl", "a")
    time.sleep(0.1)
    pyautogui.typewrite(f"/{tcode}", interval=0.1)
    time.sleep(0.3)
    pyautogui.press("enter")
    time.sleep(wait)


# ─────────────────────────────────────────────
# UPLOAD KE SAP VIA ZPGD_U2C
# ─────────────────────────────────────────────

def upload_u2c_to_sap(filepath: str, plant: str,
                      send_log=None) -> bool:
    """
    Upload file U2C ke SAP.
    T-code diambil dari Config.ACTIVE_TCODE_U2C yang di-set GUI saat user
    memilih portal EOD:
      - Portal PGDMTX → /NZPGD_U2C
      - Portal CMIS   → /NZCNS_U2C

    Alur:
    1. Navigasi ke T-code U2C yang sesuai portal
    2. Isi field lokasi file dengan filepath
    3. Execute (Enter / Transfer)
    4. Handle popup Allow jika muncul
    5. Tunggu selesai

    Return: True jika berhasil, False jika gagal
    """
    from config import Config

    def _log(msg, level="INFO"):
        if send_log:
            send_log(msg, level)
        log.info(f"[U2C] {msg}")

    # Baca T-code aktif dari Config (sudah di-set GUI berdasarkan pilihan portal)
    tcode = getattr(Config, "ACTIVE_TCODE_U2C", "ZPGD_U2C")

    try:
        _log(f"Upload U2C via /{tcode}: {os.path.basename(filepath)}")

        # 1. Navigasi ke T-code U2C
        # sap_tcode() sudah fokus ke SAP, lalu Enter → dialog "Import from a Local File" muncul
        sap_tcode(f"n{tcode}", wait=3.0)
        _log(f"Navigasi ke /{tcode}...")

        # 2. Dialog "Import from a Local File" sudah aktif & field File name sudah fokus
        # JANGAN focus_sap() di sini — akan pindahkan fokus dari dialog ke parent window!
        # Langsung select all & paste path file
        time.sleep(1.0)   # tunggu dialog render sempurna
        pyautogui.hotkey("ctrl", "a")
        time.sleep(0.2)
        import pyperclip
        pyperclip.copy(filepath)
        time.sleep(0.3)
        pyautogui.hotkey("ctrl", "v")
        time.sleep(0.5)
        _log(f"Path diisi: {filepath}")

        # 3. Klik Transfer (Enter)
        pyautogui.press("enter")
        time.sleep(1.0)

        # 4. Handle popup SAP GUI Security → Allow
        _handle_allow_popup()
        time.sleep(5)

        _log(f"Upload selesai [/{tcode}] plant {plant}", "OK")
        return True

    except Exception as e:
        _log(f"Upload gagal [/{tcode}] plant {plant}: {e}", "ERROR")
        return False


def _type_path(filepath: str):
    """
    Ketik path file ke field SAP yang aktif.
    Pakai clipboard (pyperclip) supaya path dengan backslash tidak error.
    """
    try:
        import pyperclip
        pyperclip.copy(filepath)
        pyautogui.hotkey("ctrl", "v")
    except ImportError:
        # Fallback: typewrite langsung (backslash mungkin bermasalah)
        # Ganti backslash dengan double backslash untuk pyautogui
        safe_path = filepath.replace("\\", "\\\\")
        pyautogui.typewrite(safe_path, interval=0.05)


def _handle_allow_popup():
    """
    Handle popup 'SAP GUI Security' yang muncul saat akses file lokal.
    Klik Allow (tombol kiri dari Deny).
    """
    try:
        time.sleep(1)
        # Cek apakah ada popup
        popup_hwnd = None
        def cb(hwnd, _):
            nonlocal popup_hwnd
            if win32gui.IsWindowVisible(hwnd):
                title = win32gui.GetWindowText(hwnd)
                if "Security" in title or "Allow" in title or "SAP GUI" in title:
                    popup_hwnd = hwnd
        win32gui.EnumWindows(cb, None)

        if popup_hwnd:
            win32gui.SetForegroundWindow(popup_hwnd)
            time.sleep(0.3)
            pyautogui.press("left")   # pindah ke tombol Allow
            time.sleep(0.2)
            pyautogui.press("enter")  # klik Allow
            time.sleep(1)
            log.info("[U2C] Popup Allow diklik")
        else:
            # Coba tekan Enter/Space saja kalau tidak ada popup terdeteksi
            log.info("[U2C] Tidak ada popup terdeteksi — lanjut")
    except Exception as e:
        log.warning(f"[U2C] Handle popup gagal: {e}")


# ─────────────────────────────────────────────
# PIPELINE LENGKAP U2C
# ─────────────────────────────────────────────

def run_u2c_pipeline(items_per_plant: dict,
                     output_dir: str = None,
                     send_log=None) -> dict:
    """
    Pipeline lengkap U2C untuk semua plant:
    1. Buat file .txt U2C per plant
    2. Upload tiap file ke SAP via nZPGD_U2C

    Return: { plant: {'file': filepath, 'uploaded': bool} }
    """
    def _log(msg, level="INFO"):
        if send_log:
            send_log(msg, level)
        log.info(f"[U2C] {msg}")

    if not items_per_plant:
        _log("Tidak ada data selisih — U2C dilewati")
        return {}

    if output_dir is None:
        output_dir = U2C_OUTPUT_DIR

    result = {}

    _log(f"Mulai pipeline U2C | {len(items_per_plant)} plant")
    _log(f"Output dir: {output_dir}")

    for plant, items in sorted(items_per_plant.items()):
        if not items:
            continue

        _log(f"{'='*40}")
        _log(f"Plant {plant} | {len(items)} item")

        # Step 1: Buat file U2C
        try:
            filepath = build_u2c_file(plant, items, output_dir)
            _log(f"File U2C dibuat: {os.path.basename(filepath)}", "OK")
        except Exception as e:
            _log(f"Gagal buat file U2C plant {plant}: {e}", "ERROR")
            result[plant] = {"file": None, "uploaded": False}
            continue

        # Step 2: Upload ke SAP
        _log(f"Upload ke SAP via nZPGD_U2C...")
        uploaded = upload_u2c_to_sap(filepath, plant, send_log)

        result[plant] = {"file": filepath, "uploaded": uploaded}

        # Kembali ke SAP Easy Access untuk plant berikutnya
        if len(items_per_plant) > 1:
            try:
                sap_tcode("SESSION_MANAGER", wait=1.5)
            except Exception:
                pass

    # Ringkasan
    ok    = sum(1 for v in result.values() if v["uploaded"])
    fail  = len(result) - ok
    _log(f"{'='*40}")
    _log(f"U2C selesai | {ok} berhasil | {fail} gagal",
         "OK" if fail == 0 else "WARN")

    return result


# ─────────────────────────────────────────────
# TEST MANDIRI
# ─────────────────────────────────────────────

if __name__ == "__main__":
    """
    U2C Pipeline — langsung dari data Matrix portal.
    U2C selalu dibuat selama ada data di tab INPUT portal,
    tidak peduli hasil compare (compare hanya untuk email laporan).

    Alur:
      1. Baca SEMUA data FSTKGD dari tab INPUT ViewDetail di Chrome
      2. Konversi langsung ke StockDiff (qty = qty_matrix)
      3. Buat file U2C
      4. Upload ke SAP via nZPGD_U2C

    SYARAT:
      1. Chrome sudah terbuka dengan --remote-debugging-port=9222
      2. Tab ViewDetail plant TARGET sudah terbuka, tab INPUT sudah diklik
      3. SAP sudah terbuka dan login (untuk step upload)
    """
    import sys
    from playwright.sync_api import sync_playwright
    from main import StockDiff, convert_date

    TEST_PLANT = "4507"   # ganti sesuai plant
    CDP_URL    = "http://127.0.0.1:9222"

    print("=" * 60)
    print(f"U2C Pipeline — Plant {TEST_PLANT}")
    print("=" * 60)
    print("SYARAT: Chrome terbuka, tab ViewDetail+INPUT sudah aktif")

    # ── Step 1: Baca Matrix dari tab Chrome ──────────────────
    print(f"\n[1] Baca data Matrix dari portal plant {TEST_PLANT}...")
    matrix = {}
    try:
        from test_compare import get_fstkgd_from_view_detail

        with sync_playwright() as pw:
            browser = pw.chromium.connect_over_cdp(CDP_URL, timeout=5000)
            ctx     = browser.contexts[0] if browser.contexts else None
            if not ctx:
                print("    GAGAL: Tidak ada context di Chrome!")
                sys.exit(1)

            target_page = None
            for pg in ctx.pages:
                if "ViewDetail" in pg.url:
                    target_page = pg
                    print(f"    Tab: {pg.url}")
                    break
            if not target_page:
                for pg in ctx.pages:
                    if "portal.mayora" in pg.url:
                        target_page = pg
                        print(f"    Tab portal: {pg.url}")
                        break
            if not target_page:
                print("    GAGAL: Tidak ada tab ViewDetail di Chrome!")
                print("    Tab yang terbuka:")
                for pg in ctx.pages:
                    print(f"      {pg.url}")
                sys.exit(1)

            matrix = get_fstkgd_from_view_detail(
                target_page, target_page.url, TEST_PLANT
            )

        if not matrix:
            print("    WARNING: 0 material terbaca dari tab INPUT!")
            print("    Pastikan tab INPUT sudah diklik di ViewDetail.")
            sys.exit(0)

        print(f"    OK: {len(matrix)} material terbaca")
        print("    Contoh 5 data:")
        for (mat, sloc, *_), data in list(matrix.items())[:5]:
            print(f"      {mat} | SLoc={sloc} | qty={data['qty']} | tgl={data['tgl']}")

    except Exception as e:
        print(f"    GAGAL: {e}")
        sys.exit(1)

    # ── Step 2: Konversi Matrix → StockDiff untuk U2C ────────
    # qty_matrix langsung dari portal, tidak perlu compare
    print(f"\n[2] Konversi {len(matrix)} item ke format U2C...")
    posting_date_default = datetime.now().strftime("%d.%m.%Y")
    items = []

    for (material, sloc, *_), data in matrix.items():
        tgl_raw = data.get("tgl", "").strip()
        try:
            posting_date = convert_date(tgl_raw)   # "20260507" -> "07.05.2026"
        except Exception:
            posting_date = posting_date_default

        item = StockDiff(
            param        = "FSTKGD",
            plant        = TEST_PLANT,
            sloc         = sloc,
            posting_date = posting_date,
            material     = material,
            qty_matrix   = data["qty"],
            qty_sap      = 0.0,   # tidak relevan untuk U2C
            diff         = 0.0,
            status       = 0,
            mvt_type     = "",
            qty_adjust   = 0.0,
        )
        items.append(item)

    print(f"    OK: {len(items)} item siap dibuat U2C")

    # ── Step 3: Buat file U2C ─────────────────────────────────
    print(f"\n[3] Buat file U2C...")
    items_per_plant = {TEST_PLANT: items}
    try:
        files = build_u2c_all_plants(items_per_plant, U2C_OUTPUT_DIR)
        for plant, fp in files.items():
            print(f"    File: {fp}")
            print(f"    Isi ({len(items)} baris):")
            with open(fp, encoding="utf-8") as f:
                for line in f:
                    print(f"      {line.rstrip()}")
    except Exception as e:
        print(f"    GAGAL buat file U2C: {e}")
        sys.exit(1)

    # ── Step 4: Upload ke SAP ─────────────────────────────────
    print(f"\n[4] Upload U2C ke SAP via nZPGD_U2C")
    print("    Pastikan SAP sudah terbuka dan login!")
    answer = input("    Ketik 'y' untuk upload, Enter untuk skip: ").strip().lower()

    if answer == "y":
        print("\n    Mulai dalam 3 detik -- jangan sentuh keyboard/mouse!")
        time.sleep(3)
        try:
            result = run_u2c_pipeline(items_per_plant)
            print()
            for plant, r in result.items():
                status = "BERHASIL" if r["uploaded"] else "GAGAL"
                print(f"    Plant {plant}: {status} | {r.get('file', '-')}")
        except Exception as e:
            print(f"    GAGAL upload: {e}")
    else:
        print("    Upload dilewati -- file U2C sudah tersimpan, bisa upload manual.")
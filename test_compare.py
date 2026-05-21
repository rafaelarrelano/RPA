"""
test_compare.py
Compare data Matrix (portal) vs SAP (file SAPSTK) per material+sloc

ALUR BARU (otomatis penuh):
  1. Buka Chrome (atau konek ke Chrome yang sudah buka) — TIDAK perlu buka manual
  2. Login ke portal jika belum login
  3. Buka halaman List Upload EOD
  4. Scan semua baris: jika Result "Not Completed" / warna merah → buka ViewDetail
  5. Di ViewDetail baca tab INPUT → ambil baris FSTKGD
  6. Compare dengan file SAPSTK terbaru
  7. Kirim email jika ada selisih

Jika Chrome sudah terbuka dengan portal, robot konek tanpa buka browser baru.
"""

import glob
import os
import subprocess
import time
from datetime import datetime
from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout

from main import parse_decimal, StockDiff, convert_date

# ─────────────────────────────────────────────
# KONFIGURASI
# ─────────────────────────────────────────────
SAP_DOWNLOAD_DIR  = r"C:\MATRIX\DOWNLOAD"
TOLERANCE         = 0.0
CDP_PORT          = 9222

# URL portal — sesuaikan jika berbeda
PORTAL_BASE       = "https://portal.mayora.co.id"
PORTAL_LOGIN_URL  = f"{PORTAL_BASE}/Account/Login"

# Daftar URL ListEod yang tersedia — bisa ditambah sesuai kebutuhan
PORTAL_EOD_URLS = {
    "PGDMTX": f"{PORTAL_BASE}/PGDMTX/Upload/ListEod",
    "CMIS":   f"{PORTAL_BASE}/CMIS/Upload/ListEod",
}

# URL aktif — diubah dari GUI sebelum robot jalan
PORTAL_LIST_EOD = PORTAL_EOD_URLS["PGDMTX"]


def set_portal_url(url: str):
    """Set URL portal ListEod aktif — dipanggil dari GUI."""
    global PORTAL_LIST_EOD
    PORTAL_LIST_EOD = url

# Path Chrome executable — robot pakai ini jika CDP belum aktif
CHROME_PATHS = [
    r"C:\Program Files\Google\Chrome\Application\chrome.exe",
    r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
    r"C:\Users\{}\AppData\Local\Google\Chrome\Application\chrome.exe".format(
        os.environ.get("USERNAME", "User")
    ),
]

# ─────────────────────────────────────────────
# HELPER: KONEK ATAU LAUNCH CHROME
# ─────────────────────────────────────────────

def _find_chrome() -> str:
    for path in CHROME_PATHS:
        if os.path.exists(path):
            return path
    raise FileNotFoundError(
        "Chrome tidak ditemukan! Isi CHROME_PATHS di test_compare.py "
        "dengan path chrome.exe yang benar."
    )


def _is_cdp_active() -> bool:
    """Cek apakah Chrome sudah berjalan dengan CDP aktif di port yang ditentukan."""
    import urllib.request
    for host in ["127.0.0.1", "localhost"]:
        try:
            urllib.request.urlopen(
                f"http://{host}:{CDP_PORT}/json/version", timeout=2
            )
            return True
        except Exception:
            continue
    return False


def _kill_chrome():
    """Tutup semua proses Chrome yang sedang berjalan."""
    try:
        subprocess.run(
            ["taskkill", "/F", "/IM", "chrome.exe"],
            capture_output=True, timeout=8
        )
        time.sleep(2)
    except Exception:
        pass


def _launch_chrome_with_cdp(send_log=None):
    """
    Buka Chrome dengan --remote-debugging-port.
    Jika Chrome sudah berjalan tanpa CDP → tutup dulu lalu buka ulang.
    """
    chrome = _find_chrome()

    # Cek apakah Chrome sedang berjalan
    chrome_running = False
    try:
        result = subprocess.run(
            ["tasklist", "/FI", "IMAGENAME eq chrome.exe", "/NH"],
            capture_output=True, text=True, timeout=5
        )
        chrome_running = "chrome.exe" in result.stdout
    except Exception:
        pass

    if chrome_running:
        if send_log:
            send_log("Chrome sedang berjalan — restart dengan CDP flag...", "WARN")
        _kill_chrome()

    cmd = [
        chrome,
        f"--remote-debugging-port={CDP_PORT}",
        "--remote-debugging-address=0.0.0.0",
        "--user-data-dir=C:\\ChromeRPA",        # profile terpisah — wajib agar CDP aktif
        "--no-first-run",
        "--no-default-browser-check",
        PORTAL_LIST_EOD,
    ]
    subprocess.Popen(cmd)

    if send_log:
        send_log(f"Chrome dibuka dengan CDP port {CDP_PORT}...", "INFO")

    # Tunggu Chrome siap — cek tiap 0.5 detik, max 20 detik
    for _ in range(40):
        time.sleep(0.5)
        if _is_cdp_active():
            if send_log:
                send_log("Chrome siap!", "OK")
            return
    raise TimeoutError("Chrome tidak siap dalam 20 detik!")


def connect_browser(playwright, send_log=None):
    """
    Konek ke Chrome via CDP.
    Jika belum aktif → buka Chrome dengan profile RPA → buka halaman login portal.
    Robot TIDAK auto-login — user input sendiri username & password.
    Setelah user selesai login, robot lanjut otomatis.
    Return: (browser, page, is_new)
    """
    cdp_url = f"http://127.0.0.1:{CDP_PORT}"

    if _is_cdp_active():
        _log(send_log, "Konek ke Chrome yang sudah terbuka...", "INFO")
        # Retry connect sampai 3x — kadang Chrome butuh waktu setelah CDP aktif
        browser = None
        for attempt in range(3):
            try:
                browser = playwright.chromium.connect_over_cdp(cdp_url, timeout=30000)
                break
            except Exception as e:
                if attempt < 2:
                    _log(send_log, f"Retry konek CDP ({attempt+1}/3)...", "WARN")
                    time.sleep(2)
                else:
                    raise

        ctx  = browser.contexts[0] if browser.contexts else browser.new_context()
        # Cari tab ListEod atau portal
        page = None
        for pg in ctx.pages:
            if "portal.mayora" in pg.url:
                page = pg
                break
        if not page:
            page = ctx.new_page()
        return browser, page, False

    # CDP belum aktif → buka Chrome baru
    _log(send_log, "Membuka Chrome dengan profil RPA...", "INFO")

    chrome = _find_chrome()
    chrome_running = False
    try:
        result = subprocess.run(
            ["tasklist", "/FI", "IMAGENAME eq chrome.exe", "/NH"],
            capture_output=True, text=True, timeout=5
        )
        chrome_running = "chrome.exe" in result.stdout
    except Exception:
        pass

    if chrome_running:
        _log(send_log,
             "Chrome terbuka tanpa CDP — menutup dan buka ulang dengan CDP...", "WARN")
        _log(send_log,
             "CATATAN: Tutup tab portal yang penting sebelum ini!", "WARN")
        time.sleep(2)
        _kill_chrome()

    # Buka Chrome ke halaman LOGIN dulu
    cmd = [
        chrome,
        f"--remote-debugging-port={CDP_PORT}",
        "--remote-debugging-address=0.0.0.0",
        "--user-data-dir=C:\\ChromeRPA",
        "--no-first-run",
        "--no-default-browser-check",
        "https://portal.mayora.co.id/v2login",   # halaman login portal
    ]
    subprocess.Popen(cmd)

    # Tunggu Chrome siap — cek tiap 0.5 detik, max 30 detik
    for _ in range(60):
        time.sleep(0.5)
        if _is_cdp_active():
            break
    else:
        raise TimeoutError("Chrome tidak siap dalam 30 detik!")

    # Tambah delay agar Chrome benar-benar siap terima koneksi
    time.sleep(2)
    _log(send_log, "Chrome siap — silakan LOGIN di browser yang terbuka", "OK")

    browser = playwright.chromium.connect_over_cdp(cdp_url, timeout=30000)
    ctx     = browser.contexts[0] if browser.contexts else browser.new_context()

    # Tunggu halaman aktif
    time.sleep(2)
    page = ctx.pages[0] if ctx.pages else ctx.new_page()

    return browser, page, True


# ─────────────────────────────────────────────
# HELPER: LOGIN PORTAL JIKA BELUM LOGIN
# ─────────────────────────────────────────────

def ensure_logged_in(page, send_log=None):
    """
    Cek apakah sudah login. Jika belum, login otomatis.
    Kredensial diambil dari Config (env var / config.py).
    """
    from config import Config

    # Jika URL mengandung 'Login' → belum login
    if "login" in page.url.lower() or "account" in page.url.lower():
        _log(send_log, "Login ke portal...")
        try:
            page.fill("input[name='username'], input[name='UserName'], #username", Config.PORTAL_USER)
            page.fill("input[name='password'], input[name='Password'], #password", Config.PORTAL_PASSWORD)
            page.click("button[type='submit'], input[type='submit']")
            page.wait_for_load_state("networkidle", timeout=20000)
            _log(send_log, f"Login berhasil | user: {Config.PORTAL_USER}")
        except Exception as e:
            _log(send_log, f"Login gagal: {e}", "ERROR")
            raise
    else:
        _log(send_log, "Sudah login ke portal")


def _log(send_log, msg, level="INFO"):
    if send_log:
        send_log(msg, level)
    else:
        print(f"[{level}] {msg}")


# ─────────────────────────────────────────────
# HELPER: ISI TANGGAL & KLIK SEARCH DI LISTEOD
# ─────────────────────────────────────────────

def search_by_date(page, posting_date: str, send_log=None):
    """
    Isi field tanggal di halaman ListEod dan klik Search.
    posting_date format: "dd.mm.yyyy" (dari GUI)
    Portal Mayora pakai format: "dd-Mon-yy" (contoh: 02-May-26)

    Selector sudah dikonfirmasi dari inspect portal:
    - Field tanggal : id="dateFilter"
    - Tombol Search : id="btnSearch"
    """
    try:
        from datetime import datetime as dt
        d           = dt.strptime(posting_date, "%d.%m.%Y")
        portal_date = d.strftime("%d-%b-%y")   # contoh: 02-May-26
        _log(send_log, f"Filter tanggal: {posting_date} → {portal_date}")

        # ── Isi field dateFilter ──────────────────────────────
        page.wait_for_selector("#dateFilter", timeout=10000)
        date_field = page.query_selector("#dateFilter")

        if not date_field:
            _log(send_log, "Field #dateFilter tidak ditemukan!", "WARN")
            return

        # Clear dan isi via JS supaya trigger event Kendo/jQuery
        page.evaluate("""(val) => {
            const el = document.getElementById('dateFilter');
            el.value = val;
            el.dispatchEvent(new Event('input',  { bubbles: true }));
            el.dispatchEvent(new Event('change', { bubbles: true }));
        }""", portal_date)

        time.sleep(0.3)
        _log(send_log, f"dateFilter diisi: {portal_date}", "OK")

        # ── Klik tombol btnSearch ─────────────────────────────
        page.wait_for_selector("#btnSearch", timeout=5000)
        page.click("#btnSearch")
        page.wait_for_load_state("networkidle", timeout=20000)
        _log(send_log, "Search diklik — tabel diperbarui", "OK")

    except Exception as e:
        _log(send_log, f"Filter tanggal gagal: {e}", "WARN")


# ─────────────────────────────────────────────
# STEP 1: BACA LIST EOD — CARI BARIS NOT COMPLETED
# ─────────────────────────────────────────────

def get_not_completed_rows(page, posting_date: str = None, send_log=None) -> list:
    """
    Scan halaman ListEod yang sudah terbuka, kumpulkan semua baris
    yang kolom Difference endstock (cell[9]) ada nilainya (bukan '--').
    Jika posting_date diberikan dan halaman belum di-filter, isi filter dulu.
    """
    # Navigasi ke ListEod jika belum di sana
    if PORTAL_LIST_EOD not in page.url:
        _log(send_log, "Buka halaman List Upload EOD...")
        page.goto(PORTAL_LIST_EOD, wait_until="networkidle", timeout=30000)
        ensure_logged_in(page, send_log)
        if posting_date:
            search_by_date(page, posting_date, send_log)
    else:
        _log(send_log, "Sudah di halaman List Upload EOD")

    rows_to_check = []
    page_num = 1

    while True:
        _log(send_log, f"Scan halaman {page_num} List EOD...", "INFO")

        # Cek total halaman dari Kendo pager
        total_pages = page.evaluate("""() => {
            const pager = document.querySelector('.k-pager-info');
            return pager ? pager.innerText.trim() : '';
        }""")
        if total_pages:
            _log(send_log, f"  {total_pages}", "INFO")

        try:
            page.wait_for_selector("table tbody tr", timeout=10000)
        except PWTimeout:
            _log(send_log, "Tabel tidak ditemukan", "WARN")
            break

        # Tunggu tabel k-selectable muncul
        try:
            page.wait_for_selector("table.k-selectable tbody tr", timeout=10000)
        except PWTimeout:
            _log(send_log, "Tabel tidak ditemukan", "WARN")
            break

        rows_data = page.evaluate("""() => {
            const tbl = document.querySelector('table.k-selectable');
            if (!tbl) return [];
            const rows = tbl.querySelectorAll('tbody tr');
            const result = [];
            for (const row of rows) {
                const cells = Array.from(row.querySelectorAll('td'));
                if (cells.length < 10) continue;
                const plant   = cells[1] ? cells[1].innerText.trim() : '';
                const diffEnd = cells[9] ? cells[9].innerText.trim() : '--';
                const link    = cells[0].querySelector('a') || row.querySelector('a');
                const href    = link ? link.href : '';
                const docNo   = cells[0].innerText.trim().split('\\n')[0].trim();
                const hasDiff = diffEnd !== '--' && diffEnd !== '' && diffEnd !== null;
                if (plant) result.push({ docNo, plant, href, diffEnd, hasDiff });
            }
            return result;
        }""")

        _log(send_log, f"Scan {len(rows_data)} baris di tabel...")
        for r in rows_data:
            _log(send_log, f"  Plant={r['plant']} diff={r['diffEnd']!r} hasDiff={r['hasDiff']} href={r['href'][:50] if r['href'] else 'NONE'}")
            if r["hasDiff"]:
                rows_to_check.append({
                    "doc_no": r["docNo"],
                    "plant":  r["plant"].strip(),
                    "url":    r["href"],
                    "diff":   r["diffEnd"],
                })
                _log(send_log,
                     f"✓ Plant {r['plant'].strip()} — Difference: {r['diffEnd']}", "WARN")

        # Next page — Kendo Grid pagination
        has_next = page.evaluate("""() => {
            // Kendo Grid: tombol next (.k-i-arrow-e atau .k-next)
            const selectors = [
                'a.k-link[aria-label="Next Page"]:not(.k-state-disabled)',
                '.k-pager-nav.k-next:not(.k-state-disabled)',
                'a[title="Next Page"]:not(.k-state-disabled)',
                '.k-i-arrow-e:not(.k-state-disabled)',
            ];
            for (const sel of selectors) {
                const el = document.querySelector(sel);
                if (el && el.offsetParent !== null) {
                    el.click();
                    return true;
                }
            }
            return false;
        }""")
        if has_next:
            page.wait_for_load_state("networkidle", timeout=10000)
            page_num += 1
        else:
            break

    _log(send_log,
         f"Ditemukan {len(rows_to_check)} baris Not Completed untuk dicek",
         "OK" if rows_to_check else "INFO")
    return rows_to_check


# ─────────────────────────────────────────────
# STEP 2: BACA TAB INPUT DARI VIEWDETAIL
# ─────────────────────────────────────────────

def get_fstkgd_from_view_detail(page, detail_url: str,
                                plant: str, send_log=None) -> dict:
    """
    Buka halaman ViewDetail, klik tab INPUT, ambil semua baris FSTKGD.
    Return: { (material, sloc): {'qty', 'sloc', 'tgl'} }
    """
    _log(send_log, f"Buka ViewDetail plant {plant}: {detail_url}")
    page.goto(detail_url, wait_until="networkidle", timeout=30000)

    # Klik tab INPUT jika belum aktif
    try:
        input_tab = page.query_selector(
            "a[href*='INPUT'], li a:has-text('INPUT'), .nav-tabs a:has-text('INPUT')"
        )
        if input_tab:
            input_tab.click()
            page.wait_for_load_state("networkidle", timeout=5000)
            time.sleep(0.5)
    except Exception:
        pass  # Tab mungkin sudah aktif

    # Ambil semua teks dari area konten tab INPUT
    raw_text = page.evaluate("""() => {
        // Prioritas: textarea, pre, tab aktif
        const selectors = [
            '.tab-pane.active textarea',
            '.tab-pane.active pre',
            '.tab-pane.active',
            '#INPUT textarea',
            '#INPUT pre',
            'textarea',
            'pre',
        ];
        for (const sel of selectors) {
            const el = document.querySelector(sel);
            if (el) {
                const txt = el.value || el.innerText || el.textContent || '';
                if (txt.includes('FSTKGD') || txt.includes('FSTKVN')) return txt;
            }
        }
        // Fallback: body text
        return document.body.innerText;
    }""")

    VALID_PARAMS = {"FSTKGD", "FSTKVN"}
    matrix    = {}
    skipped_7 = 0
    for line in raw_text.splitlines():
        line = line.strip()
        if not line:
            continue
        # Ambil prefix (FSTKGD atau FSTKVN)
        prefix = line.split("|")[0].strip() if "|" in line else ""
        if prefix not in VALID_PARAMS:
            continue
        fields = line.split("|")
        if len(fields) < 6:
            continue
        # Filter hanya baris yang plant-nya cocok (field index 1)
        if fields[1].strip() != plant:
            continue

        sloc     = fields[2].strip()
        tgl      = fields[3].strip()
        material = fields[4].strip()
        qty_str  = fields[5].strip()

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

        # Key unik per material + sloc + param (FSTKGD vs FSTKVN beda baris)
        key = (material, sloc, prefix)
        matrix[key] = {"qty": qty_matrix, "sloc": sloc, "tgl": tgl, "param": prefix}

    fstkgd_count = sum(1 for k in matrix if k[2] == "FSTKGD")
    fstkvn_count = sum(1 for k in matrix if k[2] == "FSTKVN")
    _log(send_log,
         f"Plant {plant}: {len(matrix)} material Matrix "
         f"(FSTKGD={fstkgd_count}, FSTKVN={fstkvn_count}) | skip kepala-2&7: {skipped_7}")
    return matrix


# ─────────────────────────────────────────────
# STEP 3: BACA FILE SAPSTK
# ─────────────────────────────────────────────

def get_sap_from_file(plant: str, send_log=None) -> dict:
    pattern = os.path.join(SAP_DOWNLOAD_DIR, f"{plant}_*_SAPSTK_*.TXT")
    files   = glob.glob(pattern)
    if not files:
        raise FileNotFoundError(f"File SAPSTK tidak ditemukan untuk plant {plant}")
    latest = max(files, key=os.path.getmtime)
    _log(send_log, f"Baca file SAP: {os.path.basename(latest)}")

    VALID_PARAMS = {"FSTKGD", "FSTKVN"}
    stok_sap = {}
    with open(latest, "r", encoding="utf-8", errors="ignore") as f:
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
    _log(send_log,
         f"SAP plant {plant}: {len(stok_sap)} material+sloc "
         f"(FSTKGD={fstkgd_c}, FSTKVN={fstkvn_c})")
    return stok_sap


# ─────────────────────────────────────────────
# STEP 4: COMPARE
# ─────────────────────────────────────────────

def compare(plant: str, matrix: dict, stok_sap: dict,
            posting_date: str, send_log=None) -> list:
    items = []
    for (material, sloc, param), data in matrix.items():
        qty_matrix = data["qty"]
        qty_sap    = stok_sap.get((material, sloc, param), 0.0)
        diff       = round(qty_matrix - qty_sap, 6)
        if abs(diff) <= TOLERANCE:
            continue

        tgl_raw = data.get("tgl", "").strip()
        try:
            item_posting_date = convert_date(tgl_raw)
        except Exception:
            item_posting_date = posting_date

        item = StockDiff(
            param=param, plant=plant, sloc=sloc,
            posting_date=item_posting_date, material=material,
            qty_matrix=qty_matrix, qty_sap=qty_sap,
            diff=diff, status=0,
        )
        if diff < 0:
            item.mvt_type   = "917"
            item.qty_adjust = abs(diff)
        else:
            item.mvt_type   = "918"
            item.qty_adjust = diff
        items.append(item)

    fstkgd_c = sum(1 for i in items if i.param == "FSTKGD")
    fstkvn_c = sum(1 for i in items if i.param == "FSTKVN")
    _log(send_log,
         f"Selisih plant {plant}: {len(items)} material "
         f"(FSTKGD={fstkgd_c}, FSTKVN={fstkvn_c})")
    return items


# ─────────────────────────────────────────────
# FUNGSI UTAMA — dipanggil dari rpa_gui.py
# ─────────────────────────────────────────────

def get_matrix_from_portal(plant: str, send_log=None) -> dict:
    """
    Kompatibel dengan rpa_gui.py:
    Konek ke Chrome (atau buka baru), navigasi ke ViewDetail plant tsb,
    ambil data Matrix dari tab INPUT.

    Jika Chrome sudah terbuka dengan ViewDetail plant yang benar → pakai langsung.
    Jika belum → cari dari ListEod.
    """
    with sync_playwright() as p:
        browser, work_page, is_new = connect_browser(p, send_log)
        try:
            ctx = browser.contexts[0] if browser.contexts else browser.new_context()

            # Cek apakah sudah ada tab ViewDetail untuk plant ini
            target_page = None
            for pg in ctx.pages:
                if f"ViewDetail" in pg.url and f"/{plant}" in pg.url:
                    target_page = pg
                    break

            if target_page:
                _log(send_log, f"Pakai tab yang sudah terbuka: {target_page.url}")
                return _read_input_tab(target_page, plant, send_log)

            # Belum ada tab yang tepat → cari dari ListEod
            list_page = work_page
            ensure_logged_in(list_page, send_log)
            list_page.goto(PORTAL_LIST_EOD, wait_until="networkidle", timeout=30000)
            ensure_logged_in(list_page, send_log)

            detail_url = _find_detail_url_for_plant(list_page, plant, send_log)
            if not detail_url:
                raise Exception(
                    f"Tidak ditemukan ViewDetail untuk plant {plant} di ListEod"
                )

            return get_fstkgd_from_view_detail(list_page, detail_url, plant, send_log)

        finally:
            if is_new:
                browser.close()


def _read_input_tab(page, plant: str, send_log=None) -> dict:
    """Baca tab INPUT dari halaman ViewDetail yang sudah terbuka."""
    return get_fstkgd_from_view_detail(page, page.url, plant, send_log)


def _goto_first_page(page):
    """
    Reset ke halaman 1.
    Selector confirmed: a[data-page="1"] dan a[aria-label="Go to the first page"]
    """
    try:
        went = page.evaluate("""() => {
            // Cek apakah sudah di halaman 1
            const cur = document.querySelector('span.k-state-selected, li.k-current-page span');
            if (cur && (cur.innerText || cur.textContent).trim() === '1') return 'skip';

            // Klik link data-page="1" (tombol First atau angka 1)
            const first = document.querySelector(
                'a[aria-label="Go to the first page"], a[data-page="1"]'
            );
            if (first && !first.classList.contains('k-state-disabled')) {
                first.click();
                return 'first-clicked';
            }
            return 'skip';
        }""")
        if went == 'first-clicked':
            _wait_pager_page(page, target_page=1, timeout=15.0)
    except Exception:
        pass


def _wait_grid_data_loaded(page, timeout: float = 20.0):
    """
    Tunggu sampai Kendo Grid selesai render setelah pindah halaman.
    Kendo load data via AJAX — networkidle saja tidak cukup.

    Strategi yang benar:
    - JANGAN cek isi sel — baris bisa memang kosong (Pending, belum diproses)
    - Tunggu loading mask hilang  ← ini sinyal Kendo selesai AJAX
    - Tunggu row count stabil 2x poll berturutan ← render selesai
    """
    import time as _t

    # Jeda awal — beri waktu Kendo trigger AJAX request dulu
    _t.sleep(0.6)

    deadline   = _t.time() + timeout
    prev_count = -1

    while _t.time() < deadline:
        result = page.evaluate("""() => {
            // Cek loading mask / overlay Kendo masih tampil
            const masks = document.querySelectorAll(
                '.k-loading-mask, .k-loading-image, .k-loading-color'
            );
            for (const m of masks) {
                const style = window.getComputedStyle(m);
                if (style.display !== 'none' && style.visibility !== 'hidden'
                    && parseFloat(style.opacity) > 0) {
                    return { loading: true, rows: -1 };
                }
            }
            // Hitung baris tbody yang ada (bisa 0 jika tidak ada data)
            const tbl  = document.querySelector('table.k-selectable');
            const rows = tbl ? tbl.querySelectorAll('tbody tr').length : 0;
            return { loading: false, rows: rows };
        }""")

        if result['loading']:
            # Masih loading — reset stabilitas, tunggu
            prev_count = -1
            _t.sleep(0.3)
            continue

        cur_count = result['rows']
        if cur_count == prev_count:
            # Row count stabil 2x poll berturutan → render selesai
            return True

        prev_count = cur_count
        _t.sleep(0.3)

    return True  # timeout — lanjut saja, data mungkin memang kosong


def _wait_pager_page(page, target_page: int = None, timeout: float = 15.0):
    """
    Tunggu sampai pager Kendo menampilkan nomor halaman target.
    Strategi:
    1. Kalau target_page diketahui → tunggu angka itu aktif di pager
    2. Kalau tidak diketahui → tunggu loading mask hilang + 0.8s
    """
    import time as _t

    deadline = _t.time() + timeout

    if target_page is not None:
        # Poll sampai halaman aktif di pager = target_page
        while _t.time() < deadline:
            current = page.evaluate("""() => {
                // Cara 1: Kendo API
                try {
                    const grid = $('.k-grid').data('kendoGrid');
                    if (grid && grid.dataSource) return grid.dataSource.page();
                } catch(e) {}
                // Selector confirmed: span.k-state-selected atau li.k-current-page span
                const active = document.querySelector(
                    'span.k-state-selected, li.k-current-page span'
                );
                if (active) return parseInt(active.innerText || active.textContent, 10);
                return null;
            }""")
            if current == target_page:
                # Halaman sudah benar, tunggu sebentar untuk render selesai
                _t.sleep(0.5)
                return
            _t.sleep(0.2)
        # Timeout — beri extra wait
        _t.sleep(1.0)
    else:
        # Tidak tahu target — tunggu loading mask hilang
        _t.sleep(0.3)
        deadline2 = _t.time() + timeout
        while _t.time() < deadline2:
            loading = page.evaluate("""() => {
                const masks = document.querySelectorAll(
                    '.k-loading-mask, .k-loading-image, .k-loading-color'
                );
                for (const m of masks) {
                    const s = window.getComputedStyle(m);
                    if (s.display !== 'none' && s.visibility !== 'hidden'
                        && parseFloat(s.opacity || '1') > 0) return true;
                }
                return false;
            }""")
            if not loading:
                _t.sleep(0.5)
                return
            _t.sleep(0.2)


def _click_next_page(page) -> bool:
    """
    Klik Next Page Kendo Grid dan tunggu data benar-benar ter-load.
    Coba Kendo dataSource API dulu, lalu fallback berbagai selector DOM.
    Return True jika berhasil pindah, False jika sudah halaman terakhir.
    """
    try:
        result = page.evaluate("""() => {
            // Selector confirmed dari debug: a[aria-label="Go to the next page"]
            // Disabled check: elemen <a> langsung punya class k-state-disabled
            const nextBtn = document.querySelector(
                'a[aria-label="Go to the next page"]'
            );
            if (!nextBtn) return 'not-found';
            if (nextBtn.classList.contains('k-state-disabled')) return 'last-page';

            // Baca halaman sekarang untuk hitung target
            const curEl = document.querySelector(
                'span.k-state-selected, li.k-current-page span'
            );
            const curPage = curEl ? parseInt(
                curEl.innerText || curEl.textContent, 10
            ) : null;

            nextBtn.click();
            return 'clicked:' + ((curPage || 0) + 1);
        }""")

        if result and result.startswith('clicked:'):
            target = None
            try:
                target = int(result.split(':')[1])
            except Exception:
                pass
            _wait_pager_page(page, target_page=target, timeout=15.0)
            return True

        return False  # last-page atau not-found
    except Exception:
        return False


def _find_detail_url_for_plant(page, plant: str, send_log=None,
                               posting_date: str = None) -> str:
    """
    Scan halaman ListEod dari halaman 1, cari link ViewDetail untuk plant tertentu.
    Selalu mulai dari halaman 1.
    """
    # Selalu kembali ke halaman 1 dulu
    _goto_first_page(page)

    page_num = 1

    while True:
        try:
            page.wait_for_selector("table.k-selectable tbody tr", timeout=10000)
        except PWTimeout:
            _log(send_log, f"  Tabel tidak muncul di halaman {page_num}", "WARN")
            break

        # Scroll tabel dari atas ke bawah supaya semua baris ter-render
        _scroll_table_to_render(page)

        # Info pager untuk log
        pager_info = page.evaluate("""() => {
            const p = document.querySelector('.k-pager-info');
            return p ? p.innerText.trim() : '';
        }""")

        # Debug: ambil sample isi sel kolom Plant (kolom index 1) dari semua baris
        sample = page.evaluate("""() => {
            const tbl = document.querySelector('table.k-selectable');
            if (!tbl) return [];
            return Array.from(tbl.querySelectorAll('tbody tr'))
                .slice(0, 5)
                .map(row => {
                    const cells = row.querySelectorAll('td');
                    return {
                        plant: cells[1] ? cells[1].innerText.trim() : '?',
                        doc:   cells[0] ? cells[0].innerText.trim().substring(0,30) : '?',
                    };
                });
        }""")
        sample_txt = ', '.join(f"{r['plant']}" for r in sample if r['plant'])
        _log(send_log,
             f"  Scan hal {page_num}"
             + (f"  [{pager_info}]" if pager_info else "")
             + f"  → cari {plant}  |  sample plants: [{sample_txt}]")

        url = page.evaluate(f"""() => {{
            const tbl = document.querySelector('table.k-selectable');
            if (!tbl) return '';
            const rows = tbl.querySelectorAll('tbody tr');
            for (const row of rows) {{
                const cells = Array.from(row.querySelectorAll('td'));
                if (cells.length < 2) continue;
                const cellPlant = cells[1] ? cells[1].innerText.trim() : '';
                if (cellPlant !== '{plant}') continue;
                const link = cells[0].querySelector('a') || row.querySelector('a');
                if (link) return link.href;
            }}
            return '';
        }}""")

        if url:
            _log(send_log, f"  ✓ Plant {plant} ditemukan di halaman {page_num}", "OK")
            return url

        # Coba pindah ke halaman berikutnya
        if _click_next_page(page):
            page_num += 1
        else:
            _log(send_log,
                 f"  Plant {plant} tidak ditemukan setelah scan {page_num} halaman",
                 "WARN")
            break

    return ""


def _scroll_table_to_render(page):
    """
    Scroll tabel dari atas ke bawah perlahan supaya semua baris ter-render.
    Portal ini pakai virtual/lazy render — baris baru muncul setelah di-scroll.
    """
    try:
        page.evaluate("""() => {
            const tbl = document.querySelector('table.k-selectable');
            if (!tbl) return;
            // Scroll container tabel (bisa .k-grid-content atau parent scroll)
            const scrollable = tbl.closest('.k-grid-content')
                            || tbl.closest('[style*="overflow"]')
                            || tbl.parentElement;
            if (!scrollable) return;
            // Scroll ke bawah bertahap
            const h = scrollable.scrollHeight;
            scrollable.scrollTop = Math.floor(h * 0.33);
        }""")
        time.sleep(0.2)
        page.evaluate("""() => {
            const tbl = document.querySelector('table.k-selectable');
            if (!tbl) return;
            const scrollable = tbl.closest('.k-grid-content')
                            || tbl.closest('[style*="overflow"]')
                            || tbl.parentElement;
            if (!scrollable) return;
            scrollable.scrollTop = scrollable.scrollHeight;
        }""")
        time.sleep(0.2)
        # Scroll kembali ke atas sebelum scan DOM
        page.evaluate("""() => {
            const tbl = document.querySelector('table.k-selectable');
            if (!tbl) return;
            const scrollable = tbl.closest('.k-grid-content')
                            || tbl.closest('[style*="overflow"]')
                            || tbl.parentElement;
            if (!scrollable) return;
            scrollable.scrollTop = 0;
        }""")
        # Scroll window juga (untuk lazy load berbasis viewport)
        page.evaluate("""() => {
            window.scrollTo(0, document.body.scrollHeight);
        }""")
        time.sleep(0.2)
        page.evaluate("() => { window.scrollTo(0, 0); }")
        time.sleep(0.1)
    except Exception:
        pass


# ─────────────────────────────────────────────
# PIPELINE LENGKAP: dipanggil dari rpa_gui.py
# run_robot() di rpa_gui.py sudah memanggil:
#   get_matrix_from_portal(plant)
#   get_sap_from_file(plant)
#   compare(plant, matrix, stok_sap, posting_date)
# Fungsi di atas sudah kompatibel.
#
# ATAU pakai fungsi all-in-one di bawah ini:
# ─────────────────────────────────────────────

def _is_stopped() -> bool:
    """Cek apakah user klik Stop di GUI."""
    try:
        from rpa_gui import stop_event
        return stop_event.is_set()
    except Exception:
        return False


def run_full_pipeline(plants: list = None, posting_date: str = None,
                      email_to: str = None, email_cc: str = None,
                      send_log=None):
    """
    Pipeline lengkap:
    1. Konek/buka Chrome
    2. Login portal jika belum
    3. Scan ListEod → cari Not Completed / semua plant yang ada
    4. Compare Matrix vs SAP per plant
    5. Kirim email jika ada selisih

    Jika `plants` diisi → hanya proses plant di list tsb.
    Jika `plants` kosong → proses semua yang Not Completed di ListEod.
    """
    from limit_adjustment import filter_by_limit
    from send_email_report import send_stock_diff_report
    from config import Config

    if posting_date is None:
        posting_date = datetime.now().strftime("%d.%m.%Y")

    items_per_plant  = {}
    matrix_per_plant = {}   # { plant: matrix dict } raw dari portal → untuk U2C
    sap_data = {}  # { plant: stok_sap dict } hasil download semua plant sekaligus

    with sync_playwright() as p:
        browser, work_page, is_new = connect_browser(p, send_log)
        try:
            ctx = browser.contexts[0] if browser.contexts else browser.new_context()
            if not work_page or work_page.is_closed():
                work_page = ctx.pages[0] if ctx.pages else ctx.new_page()

            # Navigasi ke ListEod & pastikan sudah login
            work_page.goto(PORTAL_LIST_EOD, wait_until="networkidle", timeout=30000)
            ensure_logged_in(work_page, send_log)
            if work_page.url != PORTAL_LIST_EOD:
                work_page.goto(PORTAL_LIST_EOD, wait_until="networkidle", timeout=30000)

            search_by_date(work_page, posting_date, send_log)

            # ── STEP 1: Kumpulkan semua plant yang ada selisih ───
            if plants:
                # Mode list: cari detail_url untuk tiap plant
                # Setiap pencarian selalu mulai dari halaman 1 (_goto_first_page
                # dipanggil di dalam _find_detail_url_for_plant).
                plant_rows = []
                _log(send_log,
                     f"Mencari {len(plants)} plant di ListEod "
                     f"(scan semua halaman per plant)...", "INFO")
                for plant in plants:
                    url = _find_detail_url_for_plant(
                        work_page, plant, send_log, posting_date
                    )
                    if url:
                        plant_rows.append({"plant": plant, "url": url})
                    else:
                        _log(send_log,
                             f"Plant {plant} tidak ditemukan di ListEod "
                             f"(mungkin belum ada data untuk tanggal ini)",
                             "WARN")
            else:
                # Mode auto-scan: ambil semua yang ada difference
                plant_rows = get_not_completed_rows(work_page, posting_date, send_log)

            if not plant_rows:
                _log(send_log, "Tidak ada plant dengan selisih.", "OK")
                return {}

            plants_with_diff = [r["plant"] for r in plant_rows]
            _log(send_log,
                 f"Plant dengan selisih: {', '.join(plants_with_diff)}", "OK")

            # ── STEP 2: Download SAPSTK semua plant sekaligus di SAP ──
            _log(send_log, "=" * 40, "INFO")
            _log(send_log,
                 f"Download SAPSTK {len(plants_with_diff)} plant dari SAP...", "INFO")

            # Cek apakah SAP sudah terbuka
            try:
                import win32gui
                sap_open = False
                def check_sap(hwnd, _):
                    nonlocal sap_open
                    if win32gui.IsWindowVisible(hwnd):
                        title = win32gui.GetWindowText(hwnd)
                        if "SAP" in title:
                            sap_open = True
                win32gui.EnumWindows(check_sap, None)
            except Exception:
                sap_open = True  # skip check jika error

            if not sap_open:
                _log(send_log, "SAP belum terbuka!", "WARN")
                # Kirim signal ke GUI via send_log dengan level SAP_WAIT
                # GUI akan set sap_ready_event setelah user klik OK
                if hasattr(send_log, '__self__') or callable(send_log):
                    # Buat event lokal untuk tunggu konfirmasi
                    import threading as _threading
                    _wait_event = _threading.Event()

                    # Wrap send_log untuk intercept SAP_WAIT_DONE
                    _original_log = send_log
                    def _patched_log(msg, level="INFO"):
                        if level == "SAP_WAIT_DONE":
                            _wait_event.set()
                            return
                        _original_log(msg, level)

                    # Kirim signal SAP_WAIT ke GUI
                    send_log("SAP belum terbuka — popup akan muncul", "SAP_WAIT")
                    # Tunggu konfirmasi dari GUI (blocking)
                    _wait_event.wait(timeout=300)  # max 5 menit
                else:
                    # Mode CLI: tanya langsung
                    input("SAP belum terbuka! Buka SAP lalu tekan Enter...")

                _log(send_log, "SAP dikonfirmasi terbuka — mulai download...", "OK")

            _log(send_log,
                 "SAP akan dikontrol otomatis — jangan sentuh keyboard/mouse!", "WARN")

            try:
                from rpa_phase1_2 import run_zpgd_sapstk, parse_sapstk_file
                for plant in plants_with_diff:
                    if _is_stopped():
                        _log(send_log, "Robot dihentikan saat download SAP.", "WARN")
                        break
                    try:
                        _log(send_log, f"Download SAPSTK plant {plant}...")
                        filepath = run_zpgd_sapstk(plant, send_log)
                        sap_data[plant] = parse_sapstk_file(filepath)
                        _log(send_log,
                             f"Plant {plant}: {len(sap_data[plant])} material+sloc", "OK")
                    except Exception as e:
                        _log(send_log, f"Plant {plant} gagal download: {e}", "ERROR")
            except ImportError:
                _log(send_log, "rpa_phase1_2 tidak tersedia — pakai file SAPSTK lokal", "WARN")

            _log(send_log, "Download SAP selesai — proses portal dimulai...", "OK")
            _log(send_log, "=" * 40, "INFO")

            # ── STEP 3: Proses tiap plant dari portal ───────────
            for row in plant_rows:
                plant      = row["plant"]
                detail_url = row["url"]

                # Ambil stok SAP dari hasil download (atau fallback file lokal)
                if plant in sap_data:
                    stok_sap = sap_data[plant]
                else:
                    try:
                        stok_sap = get_sap_from_file(plant, send_log)
                    except FileNotFoundError as e:
                        _log(send_log, str(e), "ERROR")
                        continue

                # Cek stop sebelum proses tiap plant
                if _is_stopped():
                    _log(send_log, "Robot dihentikan oleh user.", "WARN")
                    break

                _process_plant_with_sap(
                    work_page, plant, detail_url, posting_date,
                    stok_sap, items_per_plant, matrix_per_plant, send_log
                )

                # Kembali ke ListEod untuk plant berikutnya
                # Goto + filter ulang + reset ke halaman 1
                if len(plant_rows) > 1:
                    work_page.goto(PORTAL_LIST_EOD, wait_until="networkidle", timeout=20000)
                    search_by_date(work_page, posting_date, send_log)
                    _goto_first_page(work_page)   # pastikan mulai dari hal 1

        finally:
            if is_new:
                browser.close()

    # ── STEP 4: U2C — buat file dari Matrix portal & upload ke SAP ──
    # U2C dibuat dari SEMUA data tab INPUT portal (matrix_per_plant),
    # bukan hanya yang selisih — supaya semua qty Matrix terupdate ke SAP.
    u2c_source = matrix_per_plant if matrix_per_plant else {}
    if u2c_source:
        _log(send_log, "=" * 40, "INFO")
        _log(send_log, "Buat file U2C dan upload ke SAP...", "INFO")
        try:
            from u2c_upload import build_u2c_from_matrix, upload_u2c_to_sap, get_u2c_filepath
            u2c_filepath = get_u2c_filepath()
            n_lines      = build_u2c_from_matrix(u2c_source, filepath=u2c_filepath)
            _log(send_log, f"File U2C: {u2c_filepath} | {n_lines} baris", "OK")
            uploaded = upload_u2c_to_sap(u2c_filepath, send_log)
            status   = "OK" if uploaded else "GAGAL"
            _log(send_log, f"Upload U2C → {status}", "OK" if uploaded else "ERROR")
        except Exception as e:
            _log(send_log, f"U2C gagal: {e}", "ERROR")

    # ── STEP 5: Kirim email laporan jika ada selisih ────────────
    if items_per_plant:
        total_item = sum(len(v) for v in items_per_plant.values())
        _log(send_log,
             f"Buat laporan & kirim email | "
             f"{len(items_per_plant)} plant | {total_item} item")
        try:
            excel_path = send_stock_diff_report(
                items_per_plant,
                override_to=email_to,
                override_cc=email_cc,
            )
            _log(send_log, f"Email terkirim | File: {os.path.basename(excel_path)}", "OK")
        except Exception as e:
            _log(send_log, f"Gagal kirim email: {e}", "ERROR")
    else:
        _log(send_log, "Tidak ada selisih — email tidak dikirim.", "OK")

    return items_per_plant


def _process_plant_with_sap(page, plant: str, detail_url: str, posting_date: str,
                            stok_sap: dict, items_per_plant: dict,
                            matrix_per_plant: dict = None, send_log=None):
    """
    Proses satu plant: baca Matrix dari portal → compare → simpan hasil.
    matrix_per_plant: dict untuk simpan raw Matrix (dipakai untuk U2C).
    Hanya diisi jika plant punya item yang LOLOS filter limit.
    """
    try:
        _log(send_log, f"{'='*40}")
        _log(send_log, f"Compare plant {plant}")

        # Baca Matrix dari portal
        matrix = get_fstkgd_from_view_detail(page, detail_url, plant, send_log)
        if not matrix:
            _log(send_log, f"Plant {plant}: tidak ada data Matrix FSTKGD", "WARN")
            return

        # Simpan Matrix mentah untuk U2C — SELALU, tidak peduli hasil compare
        # U2C harus dibuat selama ada data di tab INPUT portal
        if matrix_per_plant is not None:
            matrix_per_plant[plant] = matrix
            _log(send_log,
                 f"Plant {plant}: {len(matrix)} item Matrix disimpan untuk U2C", "OK")

        # Compare
        items = compare(plant, matrix, stok_sap, posting_date, send_log)
        if not items:
            _log(send_log, f"Plant {plant}: tidak ada selisih — U2C tetap dibuat", "OK")
            return

        # Load limit adjustment dari file Excel
        try:
            from limit_adjustment import load_limit_adjustment, filter_by_limit
            limits = load_limit_adjustment()
        except Exception as e:
            _log(send_log, f"Gagal baca file limit adjustment: {e} — semua item lolos", "WARN")
            from limit_adjustment import filter_by_limit
            limits = {}

        items_ok, items_skip = filter_by_limit(items, limits)

        # ── Log detail item yang LOLOS ────────────────────
        if items_ok:
            _log(send_log,
                 f"Plant {plant}: {len(items_ok)} item lolos limit → masuk laporan", "OK")

        # ── Log detail item yang LEWAT BATAS ──────────────
        if items_skip:
            _log(send_log,
                 f"Plant {plant}: {len(items_skip)} item LEWAT LIMIT ADJUSTMENT ↓", "WARN")
            for item in items_skip:
                lim   = limits.get(item.material, {})
                lim_p = lim.get('limit_plus',  'N/A')
                lim_m = lim.get('limit_minus', 'N/A')
                if item.diff > 0:
                    batas = f"limit+ = {lim_p}"
                    jarak = f"selisih {item.diff:+.3f} > {lim_p}"
                else:
                    batas = f"limit- = {lim_m}"
                    jarak = f"selisih {item.diff:+.3f} < {lim_m}"
                _log(send_log,
                     f"  \u26a0 {item.material} | SLoc={item.sloc} | "
                     f"diff={item.diff:+.6f} | {batas} | {jarak}", "WARN")
            _log(send_log,
                 f"Plant {plant}: item lewat limit \u2192 SKIP dari U2C & laporan", "WARN")

            # Kirim ke panel Limit Alert di GUI
            import json
            items_data  = [{"material": i.material, "sloc": i.sloc,
                             "diff": i.diff} for i in items_skip]
            limits_data = {k: {"limit_plus": v["limit_plus"],
                                "limit_minus": v["limit_minus"]}
                           for k, v in limits.items()}
            _log(send_log,
                 f"{plant}|||{json.dumps(items_data)}|||{json.dumps(limits_data)}",
                 "LIMIT_ALERT")

        # ── Simpan hasil ke laporan ───────────────────────
        if items_ok:
            items_per_plant[plant] = items_ok
        else:
            # Semua item lewat limit → tidak masuk email laporan
            # tapi U2C tetap sudah disimpan di matrix_per_plant (di atas)
            _log(send_log,
                 f"Plant {plant}: semua item lewat limit → skip laporan email",
                 "WARN")

    except Exception as e:
        _log(send_log, f"Plant {plant} GAGAL: {e}", "ERROR")


def _process_plant(page, plant: str, detail_url: str, posting_date: str,
                   items_per_plant: dict, send_log=None):
    """Proses satu plant dengan download SAPSTK dulu (legacy — dipakai jika perlu)."""
    try:
        stok_sap = get_sap_from_file(plant, send_log)
    except FileNotFoundError as e:
        _log(send_log, str(e), "ERROR")
        return
    _process_plant_with_sap(
        page, plant, detail_url, posting_date,
        stok_sap, items_per_plant, send_log
    )


# ─────────────────────────────────────────────
# ENTRY POINT (test manual)
# ─────────────────────────────────────────────

if __name__ == "__main__":
    from config import Config

    print("=" * 60)
    print("Test pipeline otomatis — scan ListEod Not Completed")
    print("=" * 60)

    # Jalankan pipeline: auto-scan semua Not Completed
    results = run_full_pipeline(
        plants       = Config.PLANTS,   # atau kosongkan [] untuk auto-scan
        posting_date = datetime.now().strftime("%d.%m.%Y"),
    )

    print(f"\nSelesai | {len(results)} plant ada selisih")
    for plant, items in results.items():
        print(f"  Plant {plant}: {len(items)} item")
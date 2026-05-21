"""
test_login_filter.py
Test: login portal → isi tanggal → klik Search → lihat hasil tabel
"""
from playwright.sync_api import sync_playwright
import time

PORTAL_LIST_EOD = ""
POSTING_DATE    = "02.05.2026"   # ganti sesuai kebutuhan

with sync_playwright() as p:
    browser = p.chromium.connect_over_cdp('http://127.0.0.1:9222')
    ctx     = browser.contexts[0]

    # Cari atau buat halaman
    page = None
    for pg in ctx.pages:
        if 'portal.mayora' in pg.url:
            page = pg
            break
    if not page:
        page = ctx.new_page()

    # Navigasi ke ListEod
    print(f"Navigasi ke {PORTAL_LIST_EOD}...")
    page.goto(PORTAL_LIST_EOD, wait_until="networkidle", timeout=30000)
    print(f"URL sekarang: {page.url}")

    # Cek apakah perlu login
    if "login" in page.url.lower() or "account" in page.url.lower():
        print("Perlu login — inspect form login...")
        inputs = page.evaluate("""() => {
            return Array.from(document.querySelectorAll('input')).map(i => ({
                id: i.id, name: i.name, type: i.type,
                placeholder: i.placeholder, visible: i.offsetWidth > 0
            }));
        }""")
        for inp in inputs:
            print(f"  {inp}")
    else:
        print("Sudah login!")

        # Isi tanggal
        from datetime import datetime
        d           = datetime.strptime(POSTING_DATE, "%d.%m.%Y")
        portal_date = d.strftime("%d-%b-%y")
        print(f"\nIsi dateFilter: {portal_date}")

        page.wait_for_selector("#dateFilter", timeout=5000)

        # Isi via JS
        page.evaluate("""(val) => {
            const el = document.getElementById('dateFilter');
            el.value = val;
            el.dispatchEvent(new Event('input',  { bubbles: true }));
            el.dispatchEvent(new Event('change', { bubbles: true }));
        }""", portal_date)
        time.sleep(0.5)

        # Verifikasi nilai
        val = page.eval_on_selector("#dateFilter", "el => el.value")
        print(f"Nilai dateFilter sekarang: {val}")

        # Klik Search
        print("Klik #btnSearch...")
        page.click("#btnSearch")
        page.wait_for_load_state("networkidle", timeout=15000)
        print("Search selesai!")

        # Lihat berapa baris di tabel
        row_count = page.evaluate("""() => {
            return document.querySelectorAll('table tbody tr').length;
        }""")
        print(f"Jumlah baris di tabel: {row_count}")

        # Ambil teks baris pertama
        if row_count > 0:
            first_row = page.evaluate("""() => {
                const row = document.querySelector('table tbody tr');
                return row ? row.innerText : '';
            }""")
            print(f"Baris pertama: {first_row[:200]}")

    browser.close()
    print("\nSelesai!")
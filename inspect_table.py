"""
inspect_table.py
Inspect struktur kolom tabel ListEod portal Mayora
"""
from playwright.sync_api import sync_playwright

with sync_playwright() as p:
    browser = p.chromium.connect_over_cdp('http://127.0.0.1:9222')
    ctx     = browser.contexts[0]

    page = None
    for pg in ctx.pages:
        if 'ListEod' in pg.url:
            page = pg
            break

    if not page:
        print("Halaman ListEod tidak ditemukan!")
        browser.close()
        exit()

    print(f"URL: {page.url}\n")

    # Ambil header kolom
    headers = page.evaluate("""() => {
        const ths = document.querySelectorAll('table thead th, table tr:first-child th, table tr:first-child td');
        return Array.from(ths).map((th, i) => ({ index: i, text: th.innerText.trim() }));
    }""")

    print("=== HEADER KOLOM ===")
    for h in headers:
        print(f"  [{h['index']}] {h['text']!r}")

    # Ambil semua sel baris pertama
    print("\n=== BARIS PERTAMA (semua sel) ===")
    first_row = page.evaluate("""() => {
        const row = document.querySelector('table tbody tr');
        if (!row) return [];
        const cells = row.querySelectorAll('td');
        return Array.from(cells).map((td, i) => ({
            index: i,
            text:  td.innerText.trim().substring(0, 60),
        }));
    }""")

    for cell in first_row:
        print(f"  [{cell['index']}] {cell['text']!r}")

    # Ambil semua baris — cek kolom mana yang ada angka selisih
    print("\n=== SEMUA BARIS (cari kolom Difference) ===")
    all_rows = page.evaluate("""() => {
        const rows = document.querySelectorAll('table tbody tr');
        return Array.from(rows).map(row => {
            const cells = Array.from(row.querySelectorAll('td'));
            const link  = row.querySelector('a[href*="ViewDetail"]');
            return {
                plant: cells[1] ? cells[1].innerText.trim() : '',
                href:  link ? link.href : '',
                cells: cells.map((td, i) => ({
                    i,
                    v: td.innerText.trim()
                }))
            };
        });
    }""")

    for row in all_rows[:3]:  # tampilkan 3 baris saja
        print(f"\n  Plant {row['plant']}:")
        for cell in row['cells']:
            print(f"    [{cell['i']}] {cell['v']!r}")

    browser.close()
    print("\nSelesai!")

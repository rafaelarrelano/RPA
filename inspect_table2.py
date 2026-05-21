"""
inspect_table2.py - Inspect HTML tabel ListEod secara detail
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

    # Ambil semua tr di tbody
    result = page.evaluate("""() => {
        const tables = document.querySelectorAll('table');
        const out = [];
        tables.forEach((tbl, ti) => {
            const rows = tbl.querySelectorAll('tbody tr');
            if (rows.length > 0) {
                out.push({
                    tableIndex: ti,
                    tableId: tbl.id,
                    tableClass: tbl.className,
                    rowCount: rows.length,
                    firstRowCellCount: rows[0].querySelectorAll('td').length,
                    firstRowText: rows[0].innerText.substring(0, 200),
                });
            }
        });
        return out;
    }""")

    print(f"=== SEMUA TABLE DENGAN TBODY TR ===")
    for t in result:
        print(f"\nTable [{t['tableIndex']}] id={t['tableId']!r} class={t['tableClass']!r}")
        print(f"  Rows: {t['rowCount']} | CellsFirstRow: {t['firstRowCellCount']}")
        print(f"  FirstRow: {t['firstRowText']!r}")

    # Cari tabel yang berisi data plant (4502, 4503, dll)
    print("\n=== TABEL YANG BERISI DATA PLANT ===")
    plant_table = page.evaluate("""() => {
        const tables = document.querySelectorAll('table');
        for (let ti = 0; ti < tables.length; ti++) {
            const tbl = tables[ti];
            if (tbl.innerText.includes('4502') && tbl.innerText.includes('ViewDetail')) {
                const rows = tbl.querySelectorAll('tbody tr');
                return {
                    tableIndex: ti,
                    id: tbl.id,
                    className: tbl.className,
                    rowCount: rows.length,
                    rows: Array.from(rows).slice(0, 3).map((row, ri) => ({
                        rowIndex: ri,
                        cellCount: row.querySelectorAll('td').length,
                        hasLink: !!row.querySelector('a[href*="ViewDetail"]'),
                        linkHref: row.querySelector('a[href*="ViewDetail"]') ?
                                  row.querySelector('a[href*="ViewDetail"]').href : '',
                        cells: Array.from(row.querySelectorAll('td')).map((td, ci) => ({
                            i: ci,
                            v: td.innerText.trim().substring(0, 60)
                        }))
                    }))
                };
            }
        }
        return null;
    }""")

    if plant_table:
        print(f"Table [{plant_table['tableIndex']}] id={plant_table['id']!r}")
        print(f"Rows: {plant_table['rowCount']}")
        for row in plant_table['rows']:
            print(f"\n  Row [{row['rowIndex']}] cells={row['cellCount']} hasLink={row['hasLink']}")
            print(f"  LinkHref: {row['linkHref']!r}")
            for cell in row['cells']:
                print(f"    [{cell['i']}] {cell['v']!r}")
    else:
        print("Tabel dengan data plant tidak ditemukan!")

    browser.close()
    print("\nSelesai!")

"""
debug_rows.py - Cek isi baris tabel k-selectable secara detail
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

    result = page.evaluate("""() => {
        const tbl = document.querySelector('table.k-selectable');
        if (!tbl) return { error: 'table.k-selectable tidak ditemukan' };

        const rows = tbl.querySelectorAll('tbody tr');
        const out  = [];

        for (let i = 0; i < Math.min(rows.length, 3); i++) {
            const row   = rows[i];
            const cells = Array.from(row.querySelectorAll('td'));

            // Cari semua link di row ini
            const links = Array.from(row.querySelectorAll('a')).map(a => ({
                text: a.innerText.trim(),
                href: a.href,
            }));

            // Ambil nilai sel kunci
            out.push({
                rowIndex:  i,
                cellCount: cells.length,
                cell0:     cells[0] ? cells[0].innerText.trim().substring(0, 80) : '',
                cell1:     cells[1] ? cells[1].innerText.trim() : '',   // Plant
                cell9:     cells[9] ? cells[9].innerText.trim() : '',   // Difference endstock
                links:     links,
            });
        }
        return { rows: out, totalRows: rows.length };
    }""")

    if 'error' in result:
        print(f"ERROR: {result['error']}")
    else:
        print(f"Total rows: {result['totalRows']}")
        for row in result['rows']:
            print(f"\nRow [{row['rowIndex']}] cells={row['cellCount']}")
            print(f"  cell[0] (DocNo) : {row['cell0']!r}")
            print(f"  cell[1] (Plant) : {row['cell1']!r}")
            print(f"  cell[9] (Diff)  : {row['cell9']!r}")
            print(f"  Links ({len(row['links'])}):")
            for lnk in row['links']:
                print(f"    text={lnk['text']!r}  href={lnk['href']!r}")

    browser.close()
    print("\nSelesai!")

from playwright.sync_api import sync_playwright

with sync_playwright() as p:
    browser = p.chromium.connect_over_cdp('http://localhost:9222')

    page = None
    for ctx in browser.contexts:
        for pg in ctx.pages:
            if 'ViewDetail' in pg.url:
                page = pg
                break

    if not page:
        print('Halaman ViewDetail tidak ditemukan!')
    else:
        print('Page:', page.url)

        # Ambil semua teks dari elemen aktif
        result = page.evaluate("""() => {
            const els = document.querySelectorAll('textarea, pre, .tab-pane.active');
            const texts = Array.from(els)
                .map(e => e.value || e.innerText || e.textContent)
                .filter(t => t && t.length > 10);
            return texts.join('\\n');
        }""")

        # Filter hanya baris FSTKGD
        lines = result.splitlines()
        fstkgd_lines = [l.strip() for l in lines if l.strip().startswith('FSTKGD')]

        print(f'Total baris FSTKGD: {len(fstkgd_lines)}')
        print('Contoh 3 baris pertama:')
        for l in fstkgd_lines[:3]:
            print(' ', l)

    browser.close()
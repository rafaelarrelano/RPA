"""
debug_pager.py
Jalankan saat Chrome sudah terbuka dengan halaman ListEod aktif.
Dump HTML pager dan coba semua cara klik Next Page, laporkan mana yang berhasil.
"""
from playwright.sync_api import sync_playwright
import time

with sync_playwright() as p:
    browser = p.chromium.connect_over_cdp('http://127.0.0.1:9222')
    ctx     = browser.contexts[0]

    page = None
    for pg in ctx.pages:
        if 'ListEod' in pg.url or 'Upload' in pg.url:
            page = pg
            break

    if not page:
        print("Halaman ListEod tidak ditemukan!")
        print("Tab yang terbuka:")
        for pg in ctx.pages:
            print(f"  {pg.url}")
        browser.close()
        exit()

    print(f"URL: {page.url}")
    print("=" * 70)

    # 1. Dump HTML pager lengkap
    pager_html = page.evaluate("""() => {
        const selectors = [
            '.k-pager-wrap', '.k-grid-pager', '[data-role=pager]',
            '.k-pager', 'nav[role=navigation]', '.pagination',
        ];
        for (const sel of selectors) {
            const el = document.querySelector(sel);
            if (el) return { sel, html: el.outerHTML };
        }
        // Fallback: cari elemen yang ada angka halaman
        const all = document.querySelectorAll('*');
        for (const el of all) {
            if (el.children.length > 2 && el.innerText && el.innerText.includes('169')) {
                return { sel: el.tagName + '.' + el.className, html: el.outerHTML.substring(0, 800) };
            }
        }
        return { sel: 'NOT FOUND', html: '' };
    }""")
    print(f"\n[PAGER] Selector: {pager_html['sel']}")
    print(f"[PAGER] HTML:\n{pager_html['html']}\n")

    # 2. Baca halaman aktif saat ini
    current_info = page.evaluate("""() => {
        // Kendo API
        try {
            const grids = document.querySelectorAll('[data-role=grid]');
            for (const g of grids) {
                const kGrid = $(g).data('kendoGrid');
                if (kGrid) return {
                    method: 'kendo-api',
                    page: kGrid.dataSource.page(),
                    total: kGrid.dataSource.totalPages(),
                };
            }
        } catch(e) {}

        // Pager info text
        const info = document.querySelector('.k-pager-info');
        return { method: 'pager-info', text: info ? info.innerText : 'not found' };
    }""")
    print(f"[CURRENT PAGE] {current_info}")

    # 3. Coba semua selector Next Page dan laporkan mana yang ADA di DOM
    print("\n[NEXT PAGE SELECTORS] Cek mana yang ada di DOM:")
    selectors = [
        'a[aria-label="Next Page"]',
        'a[title="Next Page"]',
        '.k-pager-next a',
        '.k-pager-next',
        'a.k-next',
        'li.k-next a',
        'li.k-next',
        'span.k-i-arrow-e',
        '.k-i-arrow-e',
        'a[data-page]',
        'button[data-page]',
        '.page-next',
        '.next-page',
        'a:contains(">")',
    ]
    for sel in selectors:
        found = page.evaluate(f"""() => {{
            try {{
                const el = document.querySelector('{sel}');
                if (!el) return null;
                return {{
                    tag: el.tagName,
                    text: el.innerText || el.textContent,
                    classes: el.className,
                    disabled: el.hasAttribute('disabled') || el.classList.contains('k-state-disabled'),
                    href: el.getAttribute('href') || '',
                    onclick: el.getAttribute('onclick') || '',
                }};
            }} catch(e) {{ return {{ error: e.message }}; }}
        }}""")
        if found:
            print(f"  ✓ '{sel}': {found}")
        else:
            print(f"  ✗ '{sel}': tidak ada")

    # 4. Coba klik Next pakai querySelector dengan teks ">"
    print("\n[NEXT PAGE] Cari semua <a> dan <button> yang mungkin Next:")
    candidates = page.evaluate("""() => {
        const result = [];
        // Semua link dan button
        const els = document.querySelectorAll('a, button, span[onclick], li[onclick]');
        for (const el of els) {
            const txt = (el.innerText || el.textContent || '').trim();
            const cls = el.className || '';
            const aria = el.getAttribute('aria-label') || '';
            const title = el.getAttribute('title') || '';
            if (txt === '>' || txt === '›' || txt === '»' ||
                aria.toLowerCase().includes('next') ||
                title.toLowerCase().includes('next') ||
                cls.includes('next') || cls.includes('k-next') ||
                cls.includes('arrow-e') || cls.includes('forward')) {
                result.push({
                    tag: el.tagName,
                    text: txt.substring(0, 20),
                    classes: cls.substring(0, 60),
                    aria,
                    title,
                    disabled: el.hasAttribute('disabled') ||
                              el.classList.contains('k-state-disabled') ||
                              el.classList.contains('disabled'),
                    id: el.id || '',
                });
            }
        }
        return result;
    }""")
    for c in candidates:
        print(f"  → {c}")

    browser.close()
    print("\nSelesai!")

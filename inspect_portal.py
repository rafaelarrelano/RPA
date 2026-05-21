from playwright.sync_api import sync_playwright

with sync_playwright() as p:
    browser = p.chromium.connect_over_cdp('http://localhost:9222')
    ctx = browser.contexts[0]

    page = None
    for pg in ctx.pages:
        if 'ListEod' in pg.url:
            page = pg
            break

    if not page:
        print("Halaman ListEod tidak ditemukan!")
        print("URL yang terbuka:")
        for pg in ctx.pages:
            print(f"  {pg.url}")
        browser.close()
        exit()

    print(f"URL: {page.url}")
    print("=" * 60)

    # Inspect semua input fields
    inputs = page.evaluate("""() => {
        const inputs = document.querySelectorAll('input');
        return Array.from(inputs).map(inp => ({
            id:          inp.id,
            name:        inp.name,
            type:        inp.type,
            placeholder: inp.placeholder,
            value:       inp.value,
            className:   inp.className,
            visible:     inp.offsetWidth > 0 && inp.offsetHeight > 0,
        }));
    }""")

    print(f"\nSemua INPUT ({len(inputs)} total):")
    for inp in inputs:
        print(f"  id={inp['id']!r:25} name={inp['name']!r:20} "
              f"type={inp['type']!r:10} value={inp['value']!r:15} "
              f"placeholder={inp['placeholder']!r:20} visible={inp['visible']}")

    # Inspect semua button / submit
    buttons = page.evaluate("""() => {
        const btns = document.querySelectorAll('button, input[type=submit], input[type=button], a.btn');
        return Array.from(btns).map(b => ({
            tag:       b.tagName,
            id:        b.id,
            text:      b.innerText || b.value || '',
            className: b.className,
            visible:   b.offsetWidth > 0 && b.offsetHeight > 0,
        }));
    }""")

    print(f"\nSemua BUTTON ({len(buttons)} total):")
    for btn in buttons:
        print(f"  tag={btn['tag']:8} id={btn['id']!r:20} "
              f"text={btn['text']!r:20} visible={btn['visible']}")

    # HTML area filter/form
    filter_html = page.evaluate("""() => {
        const sel = [
            'form', '.filter', '#filter', '.row:first-child',
            '.panel', '.card', '.search-area'
        ];
        for (const s of sel) {
            const el = document.querySelector(s);
            if (el) return el.outerHTML.substring(0, 2000);
        }
        return document.body.innerHTML.substring(0, 2000);
    }""")

    print("\nHTML Filter Area (2000 chars):")
    print(filter_html)

    browser.close()

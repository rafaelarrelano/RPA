"""
debug_u2c_param.py
Jalankan saat Chrome sudah terbuka dengan tab ViewDetail aktif.
Tujuan: cek apakah tab INPUT portal benar-benar punya baris FSTKVN atau tidak.
"""
from playwright.sync_api import sync_playwright

TARGET_PLANT = "B356"   # ganti sesuai plant yang mau dicek

with sync_playwright() as p:
    browser = p.chromium.connect_over_cdp("http://127.0.0.1:9222")
    ctx     = browser.contexts[0]

    page = None
    for pg in ctx.pages:
        if "ViewDetail" in pg.url:
            page = pg
            break

    if not page:
        print("Tab ViewDetail tidak ditemukan!")
        print("Tab yang terbuka:")
        for pg in ctx.pages:
            print(f"  {pg.url}")
        browser.close()
        exit()

    print(f"URL: {page.url}")
    print("=" * 60)

    raw_text = page.evaluate("""() => {
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
        return document.body.innerText;
    }""")

    lines = raw_text.splitlines()

    fstkgd = [l for l in lines if l.strip().startswith("FSTKGD")]
    fstkvn = [l for l in lines if l.strip().startswith("FSTKVN")]

    print(f"Total baris FSTKGD : {len(fstkgd)}")
    print(f"Total baris FSTKVN : {len(fstkvn)}")
    print()

    if fstkvn:
        print("Contoh baris FSTKVN (5 pertama):")
        for l in fstkvn[:5]:
            print(f"  {l.strip()}")
    else:
        print(">> TIDAK ADA baris FSTKVN di tab INPUT ini!")
        print(">> Kemungkinan portal hanya kirim FSTKGD untuk plant ini.")

    print()
    print("Contoh 3 baris FSTKGD:")
    for l in fstkgd[:3]:
        print(f"  {l.strip()}")

    # Cek dari matrix build
    print()
    print("=" * 60)
    print("Simulasi build_u2c_from_matrix — cek key param:")
    from main import parse_decimal
    VALID_PARAMS = {"FSTKGD", "FSTKVN"}
    matrix = {}
    for line in lines:
        line = line.strip()
        if not line:
            continue
        prefix = line.split("|")[0].strip() if "|" in line else ""
        if prefix not in VALID_PARAMS:
            continue
        fields = line.split("|")
        if len(fields) < 6:
            continue
        if fields[1].strip() != TARGET_PLANT:
            continue
        sloc     = fields[2].strip()
        tgl      = fields[3].strip()
        material = fields[4].strip()
        qty_str  = fields[5].strip()
        if not qty_str:
            continue
        if material.startswith("7") or material.startswith("2"):
            continue
        try:
            qty = parse_decimal(qty_str)
        except Exception:
            continue
        key = (material, sloc, prefix)
        matrix[key] = {"qty": qty, "sloc": sloc, "tgl": tgl, "param": prefix}

    fstkgd_keys = [k for k in matrix if k[2] == "FSTKGD"]
    fstkvn_keys = [k for k in matrix if k[2] == "FSTKVN"]
    print(f"Matrix keys FSTKGD : {len(fstkgd_keys)}")
    print(f"Matrix keys FSTKVN : {len(fstkvn_keys)}")

    if fstkvn_keys:
        print("Contoh FSTKVN di matrix:")
        for k in fstkvn_keys[:3]:
            d = matrix[k]
            print(f"  key={k}  data={d}")
    else:
        print(">> Tidak ada FSTKVN di matrix — data portal memang tidak punya FSTKVN")

    browser.close()
    print("\nSelesai!")

"""
limit_adjustment.py
Baca file Excel Daftar Limit Adjustment Decimal dari SAP
dan apply ke logika compare sebelum input MIGO
"""
import openpyxl
from config import Config
from logger import setup_logger

log = setup_logger()


def load_limit_adjustment(filepath: str = None) -> dict:
    """
    Baca file Excel Daftar Limit Adjustment Decimal.
    Return: { material_code: {'limit_plus': float, 'limit_minus': float} }

    Struktur Excel:
    - Header di baris 4 (kolom: None, Kode, Nama Barang, Limit Plus, Limit Minus)
    - Data mulai baris 5
    """
    if filepath is None:
        filepath = Config.LIMIT_ADJ_FILE

    try:
        wb = openpyxl.load_workbook(filepath, data_only=True)
        ws = wb["Limit adjustment SKU"]

        limits = {}
        data_started = False

        for row in ws.iter_rows(values_only=True):
            # Cari baris header (kolom ke-2 = 'Kode')
            if row[1] == 'Kode':
                data_started = True
                continue

            if not data_started:
                continue

            # Skip baris kosong
            if not row[1]:
                continue

            # Ambil data
            material    = str(int(row[1])).strip()
            limit_plus  = float(row[3]) if row[3] is not None else 0.0
            limit_minus = float(row[4]) if row[4] is not None else 0.0

            limits[material] = {
                'limit_plus':  limit_plus,
                'limit_minus': limit_minus,
            }

        log.info(f"[LIMIT] {len(limits)} material limit adjustment dibaca")
        return limits

    except Exception as e:
        log.error(f"[LIMIT] Gagal baca file limit: {e}")
        raise


def is_within_limit(material: str, diff: float, limits: dict) -> bool:
    """
    Cek apakah selisih masih dalam batas limit adjustment.
    Return True = BOLEH di-adjust | False = LEWAT BATAS, skip

    Logika:
    - diff positif: cek apakah diff <= limit_plus
    - diff negatif: cek apakah diff >= limit_minus (limit_minus bernilai negatif)
    - Material tidak ada di mapping: gunakan Config.DIFF_THRESHOLD sebagai default
    """
    if material not in limits:
        # Material tidak ada di daftar limit → LEWAT BATAS, masuk laporan email
        # (tidak di-skip, tapi tidak ada acuan limit → anggap selisih signifikan)
        return True

    limit = limits[material]

    if diff > 0:
        # Selisih positif → cek limit plus
        within = diff <= limit['limit_plus']
        if not within:
            log.warning(
                f"[LIMIT] LEWAT BATAS | {material} | diff={diff} > limit_plus={limit['limit_plus']}"
            )
        return within

    elif diff < 0:
        # Selisih negatif → cek limit minus
        within = diff >= limit['limit_minus']
        if not within:
            log.warning(
                f"[LIMIT] LEWAT BATAS | {material} | diff={diff} < limit_minus={limit['limit_minus']}"
            )
        return within

    else:
        # Selisih = 0, tidak perlu adjustment
        return False


def filter_by_limit(items: list, limits: dict) -> tuple:
    """
    Filter list StockDiff berdasarkan limit adjustment.
    Return: (items_ok, items_exceeded)
    - items_ok       : boleh di-adjust ke MIGO
    - items_exceeded : lewat batas, perlu cek manual
    """
    items_ok       = []
    items_exceeded = []

    for item in items:
        if is_within_limit(item.material, item.diff, limits):
            items_ok.append(item)
        else:
            items_exceeded.append(item)
            log.warning(
                f"[LIMIT] Skip MIGO | {item.material} | SLoc={item.sloc} | "
                f"diff={item.diff} | mvt={item.mvt_type}"
            )

    log.info(f"[LIMIT] Lolos filter : {len(items_ok)} item")
    log.info(f"[LIMIT] Lewat batas  : {len(items_exceeded)} item → perlu cek manual")
    return items_ok, items_exceeded


if __name__ == '__main__':
    # Test baca file limit
    import os

    # Pakai file yang diupload
    filepath = r"C:\Users\User\Downloads\RPA_StockRecon_Python\rpa_stock_recon\config\List Limit Adj. Material SAP.xlsx"

    print('Baca file limit adjustment...')
    limits = load_limit_adjustment(filepath)

    print(f'\nTotal material: {len(limits)}')
    print('\nContoh 5 material:')
    for mat, lim in list(limits.items())[:5]:
        print(f'  {mat} | Limit+: {lim["limit_plus"]} | Limit-: {lim["limit_minus"]}')

    # Test is_within_limit
    print('\nTest filter:')
    test_cases = [
        ('377001', 0.030),   # diff=0.030, limit_plus=0.042 → BOLEH
        ('377001', 0.050),   # diff=0.050, limit_plus=0.042 → LEWAT BATAS
        ('377003', -0.015),  # diff=-0.015, limit_minus=-0.021 → BOLEH
        ('377003', -0.025),  # diff=-0.025, limit_minus=-0.021 → LEWAT BATAS
        ('999999', 0.100),   # material tidak ada → pakai default threshold
    ]
    for mat, diff in test_cases:
        result = is_within_limit(mat, diff, limits)
        status = 'BOLEH' if result else 'LEWAT BATAS'
        print(f'  {mat} | diff={diff:+.3f} | → {status}')
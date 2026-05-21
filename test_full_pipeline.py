"""
test_full_pipeline.py
Full pipeline: Compare Matrix vs SAP → Filter limit → Input MIGO_GI
"""
import time
from datetime import datetime
from test_compare import get_matrix_from_portal, get_sap_from_file, compare
from limit_adjustment import load_limit_adjustment, filter_by_limit
from main import input_migo_batch, load_plant_mapping, take_screenshot
from config import Config

PLANT        = '4502'   # ganti sesuai plant
posting_date = datetime.now().strftime('%d.%m.%Y')

print('=' * 55)
print(f'Full pipeline plant {PLANT}')
print('=' * 55)

# Step 1: Load mapping plant → cost center
print('\n[1] Load plant mapping...')
PLANT_CC = load_plant_mapping()
cc = PLANT_CC.get(PLANT, "")
if not cc:
    print(f'ERROR: Cost center untuk plant {PLANT} tidak ada di mapping!')
    exit()
print(f'Cost center {PLANT}: {cc}')

# Step 2: Ambil data Matrix dari portal
print('\n[2] Ambil data Matrix dari portal...')
matrix = get_matrix_from_portal(PLANT)

# Step 3: Ambil data SAP dari file terbaru
print('\n[3] Ambil data SAP dari file SAPSTK...')
stok_sap = get_sap_from_file(PLANT)

# Step 4: Compare
print('\n[4] Compare Matrix vs SAP...')
items = compare(PLANT, matrix, stok_sap, posting_date)

if not items:
    print('Tidak ada selisih! Robot selesai.')
    exit()

# Step 5: Apply limit adjustment
print('\n[5] Apply limit adjustment...')
limits   = load_limit_adjustment()
items_ok, items_exceeded = filter_by_limit(items, limits)

print(f'\nHasil filter:')
print(f'  Boleh di-adjust   : {len(items_ok)} item')
print(f'  Lewat batas limit : {len(items_exceeded)} item (skip)')

if items_exceeded:
    print(f'\nItem yang di-skip:')
    for item in items_exceeded:
        lim = limits.get(item.material, {})
        print(f'  {item.material} | diff={item.diff:+.6f} | limit+={lim.get("limit_plus","N/A")} | limit-={lim.get("limit_minus","N/A")}')

if not items_ok:
    print('\nSemua item lewat batas limit. Robot selesai.')
    exit()

# Step 6: Input MIGO
print(f'\n[6] Input MIGO_GI untuk {len(items_ok)} item...')
print('Mulai dalam 5 detik - pastikan SAP full screen!')
time.sleep(5)

try:
    doc_number = input_migo_batch(items_ok, PLANT_CC)
    take_screenshot(Config.FOLDER_SCREENSHOTS, items_ok[0], doc_number)
    print(f'\nBerhasil! Doc number: {doc_number}')
    print(f'Total item diproses: {len(items_ok)}')
except Exception as e:
    print(f'\nGagal input MIGO: {e}')
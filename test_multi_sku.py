from main import StockDiff, input_migo_batch
from config import Config
import time

test_items = [
    StockDiff('FSTKGD','4503','WH01','09.04.2026','378001',100,100.001,-0.001,0,'917',0.001),
    StockDiff('FSTKGD','4503','WH01','09.04.2026','378002',100,100.002,-0.002,0,'917',0.002),
    StockDiff('FSTKGD','4503','WH01','09.04.2026','378003',100,100.003,-0.003,0,'917',0.003),
]

plant_cc = {'4503': 'SPLG450500'}

print(f'Mulai batch {len(test_items)} item dalam 3 detik...')
time.sleep(3)

doc = input_migo_batch(test_items, plant_cc)
print(f'Selesai! Doc number: {doc}')
#!/usr/bin/env python3
import re

filepath = "c:/Users/user/source/repos/RPA/rpa_phase1_2.py"

with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

# Cari posisi akhir imports
lines = content.split('\n')
import_end = 0
for i, line in enumerate(lines):
    if line.startswith('import ') or line.startswith('from '):
        import_end = i
    elif import_end > 0 and line and not line.startswith(('import ', 'from ', '#', ' ', '\t')):
        break

# Posisi untuk insert function
insert_pos = len('\n'.join(lines[:import_end+1])) + 1

# Check apakah _is_stopped sudah ada
if '_is_stopped' not in content:
    print("ERROR: _is_stopped tidak ditemukan. Pastikan sudah ada di file.")
    exit(1)

# Check apakah _interruptible_sleep sudah ada
if '_interruptible_sleep' in content:
    print("INFO: _interruptible_sleep sudah ada di file. Skip inject.")
else:
    func_code = '''

def _interruptible_sleep(duration: float):
    """Sleep yang bisa di-interrupt dengan check stop_event tiap 0.1s."""
    if _is_stopped():
        return
    end_time = time.time() + duration
    while time.time() < end_time:
        if _is_stopped():
            return
        remaining = end_time - time.time()
        sleep_chunk = min(0.1, remaining)
        if sleep_chunk > 0:
            time.sleep(sleep_chunk)
'''
    
    content = content[:insert_pos] + func_code + content[insert_pos:]
    
    with open(filepath, 'w', encoding='utf-8') as f:
        f.write(content)
    
    print(f"✓ Injected _interruptible_sleep() ke rpa_phase1_2.py")

print("Done!")

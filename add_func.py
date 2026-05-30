import sys

with open('test_compare.py', 'r', encoding='utf-8') as f:
    lines = f.readlines()

# Find line index after 'from main import...'
insert_idx = None
for i, line in enumerate(lines):
    if 'from main import parse_decimal' in line:
        insert_idx = i + 1
        break

if insert_idx:
    # Insert the function
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
    lines.insert(insert_idx, func_code)
    
    with open('test_compare.py', 'w', encoding='utf-8') as f:
        f.writelines(lines)
    
    print('Function added successfully')
else:
    print('Could not find import line')

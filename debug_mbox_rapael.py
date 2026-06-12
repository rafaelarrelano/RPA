# Paste di Python console / run as script
import os, re

MBOX_PATH = r"D:\Email Rafael\Drafts"

with open(MBOX_PATH, "rb") as f:
    raw = f.read()

messages = re.split(rb'\nFrom ', raw)
print(f"Total pesan di mbox: {len(messages)}")

# Cek 2 pesan terakhir - hanya header penting
for i, msg in enumerate(messages[-2:], start=len(messages)-1):
    print(f"\n=== PESAN #{i} ===")
    try:
        text = msg.decode('utf-8', errors='replace')
    except:
        text = msg.decode('latin-1', errors='replace')
    
    # Header saja
    for j, line in enumerate(text.split('\n')[:15]):
        print(f"  {j:02d}: {repr(line)}")
    
    # Cek line endings
    print(f"  CRLF count: {text.count(chr(13)+chr(10))}")
    print(f"  LF-only: {text.count(chr(10)) - text.count(chr(13)+chr(10))}")
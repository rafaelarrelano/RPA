"""
debug_mbox_format.py
Lihat format pesan terakhir di mbox Drafts untuk bandingkan
dengan yang ditulis RPA.
"""
import os, re

MBOX_PATH = r"D:\Email Rafael\Drafts"

with open(MBOX_PATH, "rb") as f:
    raw = f.read()

# Split pesan berdasarkan "From " di awal baris
messages = re.split(rb'\nFrom ', raw)
print(f"Total pesan di mbox: {len(messages)}")
print()

# Tampilkan header 30 baris dari 3 pesan terakhir
for i, msg in enumerate(messages[-3:], start=len(messages)-2):
    print(f"{'='*60}")
    print(f"PESAN #{i} (dari {len(messages)} total)")
    print(f"{'='*60}")
    # Decode dengan error handling
    header_bytes = msg[:2000]
    try:
        header_text = header_bytes.decode('utf-8', errors='replace')
    except:
        header_text = header_bytes.decode('latin-1', errors='replace')
    
    # Tampilkan hanya header (sampai baris kosong pertama)
    lines = header_text.split('\n')
    for j, line in enumerate(lines[:40]):
        print(f"  {j:02d}: {repr(line)}")
    print()

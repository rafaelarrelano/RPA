"""
Baca 2000 bytes terakhir dari file mbox Drafts untuk verifikasi.
"""
import os

mbox = r"D:\Email Rafael\Drafts"

with open(mbox, "rb") as f:
    f.seek(0, 2)  # ke akhir file
    size = f.tell()
    # Baca 2000 bytes terakhir
    f.seek(max(0, size - 2000))
    tail = f.read()

# Coba decode
for enc in ("utf-8", "latin-1", "cp1252"):
    try:
        text = tail.decode(enc)
        print(f"Encoding: {enc}")
        print("-" * 40)
        print(text)
        print("-" * 40)
        # Cek apakah RPA TEST ada
        print(f"[RPA TEST] ditemukan: {'[RPA TEST]' in text}")
        break
    except Exception as e:
        print(f"{enc}: {e}")

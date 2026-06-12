# Jalankan sebagai script: python patch_imap.py

import re

filepath = r"C:\Users\User\Documents\PGD\rpa_stock_recon\send_email_report.py"

with open(filepath, "r", encoding="utf-8") as f:
    content = f.read()

# Cek apakah sudah ada kode IMAP
if "imaplib" in content:
    print("File sudah pakai IMAP — tidak perlu patch")
else:
    print(f"File belum diupdate. Perlu copy manual.")
    print(f"Copy file dari: C:\\RPA_StockRecon\\... atau dari download Claude")
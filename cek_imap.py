"""
cek_imap.py
Cek konfigurasi IMAP dari Thunderbird dan test koneksi.
"""
import os
import re
import socket

appdata  = os.environ.get("APPDATA", "")
tb_base  = os.path.join(appdata, "Thunderbird")

# Baca prefs.js dari profil aktif (8krq6v0x.default-release)
prefs_file = os.path.join(
    tb_base, "Profiles", "8krq6v0x.default-release", "prefs.js"
)

print(f"Baca prefs.js: {prefs_file}")
print(f"Exists: {os.path.exists(prefs_file)}\n")

if os.path.exists(prefs_file):
    with open(prefs_file, "r", encoding="utf-8", errors="ignore") as f:
        content = f.read()

    # Cari semua setting yang berkaitan dengan mail server
    keywords = [
        "mail.server",
        "hostname",
        "port",
        "userName",
        "type",
        "socketType",
    ]
    print("=== Setting mail server dari prefs.js ===")
    for line in content.splitlines():
        if any(kw in line for kw in keywords):
            if "mail.server" in line or "mail.account" in line:
                print(f"  {line.strip()}")

    # Cari khusus hostname dan port IMAP
    print("\n=== Hostname & Port ===")
    servers = {}
    for line in content.splitlines():
        m = re.search(r'user_pref\("mail\.server\.(server\d+)\.(\w+)",\s*(.+)\);', line)
        if m:
            srv_id  = m.group(1)
            key     = m.group(2)
            val     = m.group(3).strip('"')
            if srv_id not in servers:
                servers[srv_id] = {}
            servers[srv_id][key] = val

    for srv_id, data in servers.items():
        if data.get("type") in ("imap", "pop3", None):
            print(f"\n  [{srv_id}]")
            for k, v in data.items():
                print(f"    {k} = {v}")

# Test koneksi ke port IMAP umum
print("\n=== Test koneksi IMAP ke mail.mayora.co.id ===")
host = "mail.mayora.co.id"
for port in [993, 143, 587, 465]:
    try:
        s = socket.create_connection((host, port), timeout=3)
        s.close()
        print(f"  Port {port}: TERBUKA ✓")
    except Exception as e:
        print(f"  Port {port}: TIDAK BISA ({e})")

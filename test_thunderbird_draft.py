"""
test_thunderbird_draft.py
Test simpan draft langsung ke mbox Thunderbird (POP3 lokal).
Jalankan: python test_thunderbird_draft.py
"""
import os, re, sys
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText


def find_profile_dir() -> str:
    appdata = os.environ.get("APPDATA", "")
    tb_base = os.path.join(appdata, "Thunderbird")
    profiles_ini = os.path.join(tb_base, "profiles.ini")
    if not os.path.exists(profiles_ini):
        return ""
    with open(profiles_ini, "r", encoding="utf-8", errors="ignore") as f:
        content = f.read()

    # Install block — hanya satu baris
    m = re.search(r'^\[Install[^\]]+\][^\[]*?^Default=([^\n\r]+)',
                  content, re.MULTILINE)
    if m:
        raw  = m.group(1).strip()
        path = os.path.join(tb_base, raw.replace("/", os.sep)) \
               if not os.path.isabs(raw) else raw
        if os.path.isdir(path):
            return path

    # Profil default-release
    blocks = re.split(r'\[Profile\d+\]', content)
    all_paths = []
    for block in blocks:
        pm = re.search(r'^Path=([^\n\r]+)', block, re.MULTILINE)
        rm = re.search(r'^IsRelative=1',    block, re.MULTILINE)
        if not pm:
            continue
        raw  = pm.group(1).strip()
        path = os.path.join(tb_base, raw.replace("/", os.sep)) if rm else raw
        if os.path.isdir(path):
            all_paths.append((path, raw))
    for path, raw in all_paths:
        if "default-release" in raw:
            return path
    return all_paths[0][0] if all_paths else ""


def find_drafts_mbox(profile_dir: str, email_from: str = "") -> str:
    domain   = email_from.split("@")[-1].lower() if "@" in email_from else ""
    username = email_from.split("@")[0].lower()  if "@" in email_from else ""
    candidates = []

    # Baca direktori custom dari prefs.js
    prefs_file = os.path.join(profile_dir, "prefs.js")
    if os.path.exists(prefs_file):
        with open(prefs_file, "r", encoding="utf-8", errors="ignore") as f:
            prefs = f.read()
        for raw in re.findall(
            r'user_pref\("mail\.server\.server\d+\.directory",\s*"([^"]+)"\)',
            prefs
        ):
            decoded = raw.replace("\\\\", "\\").replace("/", os.sep)
            if not os.path.isdir(decoded):
                continue
            drafts = os.path.join(decoded, "Drafts")
            score  = 2 if domain in decoded.lower() else (1 if username in decoded.lower() else 0)
            candidates.append((score, drafts, decoded))

    # Subfolder Mail/ImapMail di dalam profil
    for sub in ("Mail", "ImapMail"):
        base = os.path.join(profile_dir, sub)
        if not os.path.isdir(base):
            continue
        for acct in os.listdir(base):
            ap = os.path.join(base, acct)
            if not os.path.isdir(ap):
                continue
            drafts = os.path.join(ap, "Drafts")
            score  = 2 if domain in acct.lower() else (1 if username in acct.lower() else 0)
            candidates.append((score, drafts, ap))

    candidates.sort(key=lambda x: -x[0])

    # Kembalikan yang sudah ada dulu, lalu yang perlu dibuat
    for _, drafts, _ in candidates:
        if os.path.isfile(drafts):
            return drafts
    # Belum ada → kembalikan kandidat terbaik untuk dibuat
    if candidates:
        return candidates[0][1]
    return ""


def write_draft(mbox_path: str, email_from: str, email_to: str) -> bool:
    msg = MIMEMultipart("mixed")
    msg["Subject"] = f"[RPA TEST] Draft test — {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}"
    msg["From"]    = email_from
    msg["To"]      = email_to
    msg["Date"]    = datetime.now().strftime("%a, %d %b %Y %H:%M:%S +0700")
    msg["X-Mozilla-Draft-Info"] = "internal/draft; vcard=0; receipt=0; DSN=0; uuencode=0"
    msg["X-Mailer"] = "RPA Stock Recon Test"
    msg.attach(MIMEText(
        f"<p>Draft test dari RPA — {datetime.now().strftime('%H:%M:%S')}</p>",
        "html", "utf-8"
    ))

    os.makedirs(os.path.dirname(mbox_path), exist_ok=True)
    ts      = datetime.now().strftime("%a %b %e %H:%M:%S %Y")
    escaped = "\n".join(
        (">" + line if line.startswith("From ") else line)
        for line in msg.as_string().split("\n")
    )
    with open(mbox_path, "a", encoding="utf-8", errors="replace") as f:
        f.write(f"From {email_from} {ts}\n")
        f.write("X-Mozilla-Status: 0008\n")
        f.write("X-Mozilla-Status2: 00000000\n")
        f.write(escaped)
        f.write("\n\n")
    return True


def _find_thunderbird_exe() -> str:
    """Cari path thunderbird.exe."""
    import shutil
    candidates = [
        shutil.which("thunderbird"),
        r"C:\Program Files\Mozilla Thunderbird\thunderbird.exe",
        r"C:\Program Files (x86)\Mozilla Thunderbird\thunderbird.exe",
    ]
    for c in candidates:
        if c and os.path.exists(c):
            return c
    return ""


def refresh_thunderbird(mbox_path: str):
    """
    Hapus .msf index agar Thunderbird rebuild draft list.
    Jika Thunderbird sedang buka dan mengunci .msf:
      1. Tutup Thunderbird
      2. Hapus .msf
      3. Buka kembali Thunderbird
    """
    import subprocess, time

    msf    = mbox_path + ".msf"
    tb_exe = _find_thunderbird_exe()

    if not os.path.exists(msf):
        print(f"  Tidak ada .msf — Thunderbird akan buat index baru otomatis")
        return

    # Coba hapus langsung dulu
    try:
        os.remove(msf)
        print(f"  Index .msf dihapus ✓")
        return
    except PermissionError:
        pass

    # .msf sedang dikunci Thunderbird → tutup dulu
    print(f"  .msf sedang dipakai Thunderbird — tutup Thunderbird...")
    try:
        subprocess.run(
            ["taskkill", "/F", "/IM", "thunderbird.exe"],
            capture_output=True, timeout=8
        )
        # Tunggu proses benar-benar mati
        for _ in range(20):
            time.sleep(0.5)
            try:
                os.remove(msf)
                print(f"  Index .msf dihapus ✓")
                break
            except PermissionError:
                continue
        else:
            print(f"  WARN: gagal hapus .msf setelah tutup Thunderbird")
            return
    except Exception as e:
        print(f"  WARN: {e}")
        return

    # Buka kembali Thunderbird
    if tb_exe:
        print(f"  Membuka kembali Thunderbird...")
        subprocess.Popen(
            [tb_exe],
            creationflags=0x00000008  # DETACHED_PROCESS
        )
        time.sleep(2)
        print(f"  Thunderbird dibuka ✓ — klik folder Drafts untuk lihat email baru")
    else:
        print(f"  Thunderbird.exe tidak ditemukan — buka manual lalu klik folder Drafts")


# ─────────────────────────────────────────────

EMAIL_FROM = "rafael.arrelano@mayora.co.id"
EMAIL_TO   = "rafael.arrelano@mayora.co.id"

print("=" * 55)
print("TEST: Simpan Draft ke Thunderbird (POP3 lokal)")
print("=" * 55)

print("\n[1] Cari profil aktif...")
profile_dir = find_profile_dir()
if not profile_dir:
    print("  GAGAL: profil tidak ditemukan")
    sys.exit(1)
print(f"  OK: {profile_dir}")

print("\n[2] Cari file mbox Drafts...")
mbox_path = find_drafts_mbox(profile_dir, EMAIL_FROM)
if not mbox_path:
    print("  GAGAL: tidak ada direktori mail yang valid")
    sys.exit(1)
exists = os.path.isfile(mbox_path)
size_before = os.path.getsize(mbox_path) if exists else 0
print(f"  OK: {mbox_path}")
print(f"  File exists: {exists} | Ukuran: {size_before:,} bytes")

print("\n[3] Tulis draft test...")
write_draft(mbox_path, EMAIL_FROM, EMAIL_TO)
size_after = os.path.getsize(mbox_path)
print(f"  OK! Ukuran sesudah: {size_after:,} bytes (+{size_after - size_before:,} bytes)")

print("\n[4] Refresh Thunderbird (hapus .msf index)...")
refresh_thunderbird(mbox_path)

print()
print("=" * 55)
print("SELESAI!")
print()
print("Langkah selanjutnya:")
print("  1. Buka (atau switch ke) Thunderbird")
print("  2. Klik folder 'Drafts' di sidebar")
print("     → draft baru harus langsung muncul")
print("  3. Jika belum muncul: klik kanan Drafts → Refresh")
print("=" * 55)
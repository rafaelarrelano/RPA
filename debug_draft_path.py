"""
debug_draft_path.py
Jalankan untuk diagnosa kenapa draft tidak masuk Thunderbird.
"""
import os, re

appdata  = os.environ.get("APPDATA", "")
tb_base  = os.path.join(appdata, "Thunderbird")

print("=" * 60)
print("STEP 1: Cari profil aktif dari profiles.ini")
print("=" * 60)

profiles_ini = os.path.join(tb_base, "profiles.ini")
print(f"profiles.ini: {profiles_ini}")
print(f"Exists: {os.path.exists(profiles_ini)}")

profile_dir = ""
if os.path.exists(profiles_ini):
    with open(profiles_ini, "r", encoding="utf-8", errors="ignore") as f:
        content = f.read()

    # Cari Install block
    m = re.search(r'\[Install[^\]]+\].*?^Default=(.+)$', content, re.MULTILINE | re.DOTALL)
    if m:
        raw  = m.group(1).strip()
        path = os.path.join(tb_base, raw.replace("/", os.sep)) if not os.path.isabs(raw) else raw
        print(f"\nInstall block Default: {raw!r}")
        print(f"  → Path: {path}")
        print(f"  → Exists: {os.path.isdir(path)}")
        if os.path.isdir(path):
            profile_dir = path

    if not profile_dir:
        blocks = re.split(r'\[Profile\d+\]', content)
        for block in blocks:
            pm = re.search(r'^Path=(.+)$', block, re.MULTILINE)
            dm = re.search(r'^IsDefault=1', block, re.MULTILINE)
            rm = re.search(r'^IsRelative=1', block, re.MULTILINE)
            if not pm:
                continue
            raw  = pm.group(1).strip()
            path = os.path.join(tb_base, raw.replace("/", os.sep)) if rm else raw
            print(f"\nProfile block: {raw!r} → {path}")
            print(f"  Exists: {os.path.isdir(path)}, IsDefault: {bool(dm)}")
            if os.path.isdir(path):
                if not profile_dir:
                    profile_dir = path

print(f"\nProfil aktif: {profile_dir!r}")

print()
print("=" * 60)
print("STEP 2: Baca semua 'directory' dari prefs.js")
print("=" * 60)

if profile_dir:
    prefs_file = os.path.join(profile_dir, "prefs.js")
    print(f"prefs.js: {prefs_file}")
    print(f"Exists: {os.path.exists(prefs_file)}")

    if os.path.exists(prefs_file):
        with open(prefs_file, "r", encoding="utf-8", errors="ignore") as f:
            prefs_content = f.read()

        matches = re.findall(
            r'user_pref\("mail\.server\.(server\d+)\.directory",\s*"([^"]+)"\)',
            prefs_content
        )
        print(f"\nDitemukan {len(matches)} direktori di prefs.js:")
        for srv, raw in matches:
            decoded = raw.replace("\\\\", "\\").replace("/", os.sep)
            exists  = os.path.isdir(decoded)
            drafts  = os.path.join(decoded, "Drafts")
            drafts_exists = os.path.isfile(drafts)
            print(f"\n  [{srv}]")
            print(f"    raw value : {raw!r}")
            print(f"    decoded   : {decoded!r}")
            print(f"    dir exists: {exists}")
            if exists:
                print(f"    Drafts    : {drafts}")
                print(f"    Drafts exists: {drafts_exists}")
                if not drafts_exists:
                    print(f"    !! File Drafts belum ada — akan dibuat otomatis")
                    files_in_dir = os.listdir(decoded)
                    print(f"    Files di direktori: {files_in_dir[:10]}")
else:
    print("Profil tidak ditemukan!")

print()
print("=" * 60)
print("STEP 3: Cek subfolder Mail/ di profil")
print("=" * 60)

if profile_dir:
    for sub in ("Mail", "ImapMail"):
        base = os.path.join(profile_dir, sub)
        if os.path.isdir(base):
            print(f"\n[{sub}] {base}")
            for acct in os.listdir(base):
                ap = os.path.join(base, acct)
                drafts = os.path.join(ap, "Drafts")
                print(f"  {acct}: Drafts={os.path.isfile(drafts)}")
        else:
            print(f"[{sub}] tidak ada")

print()
print("=" * 60)
print("SELESAI — salin output ini dan kirim ke developer")
print("=" * 60)

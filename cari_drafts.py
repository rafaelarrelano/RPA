"""
cari_drafts.py
Scan SEMUA profil Thunderbird, cari file Drafts di mana saja.
"""
import os

appdata  = os.environ.get("APPDATA", "")
tb_base  = os.path.join(appdata, "Thunderbird")
profiles = os.path.join(tb_base, "Profiles")

print(f"Scan semua profil di: {profiles}\n")

for profile_name in os.listdir(profiles):
    profile_path = os.path.join(profiles, profile_name)
    if not os.path.isdir(profile_path):
        continue
    print(f"=== PROFIL: {profile_name} ===")
    for subfolder in ("Mail", "ImapMail"):
        base = os.path.join(profile_path, subfolder)
        if not os.path.isdir(base):
            print(f"  [{subfolder}] tidak ada")
            continue
        print(f"  [{subfolder}]")
        for acc in os.listdir(base):
            ap = os.path.join(base, acc)
            if not os.path.isdir(ap):
                continue
            print(f"    [{acc}]")
            for f in sorted(os.listdir(ap)):
                full = os.path.join(ap, f)
                size = os.path.getsize(full) if os.path.isfile(full) else "DIR"
                mark = " ← DRAFTS MBOX!" if f == "Drafts" else ""
                print(f"      {f}  ({size} bytes){mark}")
    print()

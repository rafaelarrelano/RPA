import os, re

appdata = os.environ.get("APPDATA", "")
tb_base = os.path.join(appdata, "Thunderbird")
profiles_ini = os.path.join(tb_base, "profiles.ini")

print(f"profiles.ini: {profiles_ini}")
print(f"Exists: {os.path.exists(profiles_ini)}")

if os.path.exists(profiles_ini):
    with open(profiles_ini) as f:
        print(f.read())

# Scan semua file di ImapMail
for server in os.listdir(os.path.join(tb_base, "Profiles")):
    profile_path = os.path.join(tb_base, "Profiles", server)
    imap = os.path.join(profile_path, "ImapMail")
    if os.path.isdir(imap):
        for s in os.listdir(imap):
            sp = os.path.join(imap, s)
            print(f"\n[{s}]")
            for f in os.listdir(sp):
                print(f"  {f}")
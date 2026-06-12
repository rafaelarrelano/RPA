import os

appdata = os.environ.get("APPDATA", "")
tb_base = os.path.join(appdata, "Thunderbird")

# Profile0 = default-release (yang dipakai sehari-hari)
profile_path = os.path.join(tb_base, "Profiles", "8krq6v0x.default-release")

print(f"Profile path: {profile_path}")
print(f"Exists: {os.path.isdir(profile_path)}")

# Scan ImapMail
imap_base = os.path.join(profile_path, "ImapMail")
print(f"\nImapMail path: {imap_base}")
print(f"Exists: {os.path.isdir(imap_base)}")

if os.path.isdir(imap_base):
    for server in os.listdir(imap_base):
        sp = os.path.join(imap_base, server)
        if os.path.isdir(sp):
            print(f"\n  [{server}]")
            for f in sorted(os.listdir(sp)):
                full = os.path.join(sp, f)
                size = os.path.getsize(full) if os.path.isfile(full) else "DIR"
                print(f"    {f}  ({size} bytes)")

# Scan Mail juga
mail_base = os.path.join(profile_path, "Mail")
print(f"\nMail path: {mail_base}")
if os.path.isdir(mail_base):
    for acc in os.listdir(mail_base):
        ap = os.path.join(mail_base, acc)
        if os.path.isdir(ap):
            print(f"\n  [{acc}]")
            for f in sorted(os.listdir(ap)):
                full = os.path.join(ap, f)
                size = os.path.getsize(full) if os.path.isfile(full) else "DIR"
                print(f"    {f}  ({size} bytes)")
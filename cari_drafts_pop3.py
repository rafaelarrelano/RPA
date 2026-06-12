import os

email_dir = r"D:\Email Rafael"

print(f"Exists: {os.path.isdir(email_dir)}")
print("\nIsi folder:")
for root, dirs, files in os.walk(email_dir):
    level = root.replace(email_dir, "").count(os.sep)
    indent = "  " * level
    print(f"{indent}{os.path.basename(root)}/")
    for f in sorted(files):
        full = os.path.join(root, f)
        size = os.path.getsize(full)
        mark = " ← INI DIA!" if f == "Drafts" else ""
        print(f"{indent}  {f}  ({size:,} bytes){mark}")

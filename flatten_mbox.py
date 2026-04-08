"""
Flatten E:\Mail_mbox\ into E:\TB_Mail\ for Thunderbird Local Folders.
- New-style (my converter): E:\Mail_mbox\<name>\Inbox  → E:\TB_Mail\<name>
- Old-style (readpst raw):  E:\Mail_mbox\<name>\...\mbox files → merged into E:\TB_Mail\<name>
"""
import os
import shutil
from pathlib import Path

SRC = Path(r"E:\Mail_mbox")
DEST = Path(r"E:\TB_Mail")
DEST.mkdir(exist_ok=True)

def find_mbox_files(folder):
    """Recursively find all files named 'mbox' or 'Inbox' in a directory."""
    results = []
    for root, dirs, files in os.walk(folder):
        for f in files:
            if f in ("mbox", "Inbox"):
                results.append(Path(root) / f)
    return results

converted = 0
failed = 0

for entry in sorted(SRC.iterdir()):
    if not entry.is_dir():
        continue

    dest_file = DEST / entry.name
    if dest_file.exists():
        print(f"SKIP (exists): {entry.name}")
        continue

    mbox_files = find_mbox_files(entry)
    if not mbox_files:
        print(f"SKIP (no mbox): {entry.name}")
        continue

    try:
        if len(mbox_files) == 1:
            # Simple copy
            shutil.copy2(mbox_files[0], dest_file)
        else:
            # Merge multiple mbox files (concatenate)
            with open(dest_file, "wb") as out:
                for mf in sorted(mbox_files):
                    with open(mf, "rb") as inp:
                        shutil.copyfileobj(inp, out)
                    out.write(b"\n")

        size_mb = dest_file.stat().st_size / (1024*1024)
        print(f"OK: {entry.name} ({len(mbox_files)} mbox, {size_mb:.1f} MB)")
        converted += 1
    except Exception as e:
        print(f"FAIL: {entry.name} — {e}")
        failed += 1

print()
print(f"Done: {converted} OK, {failed} failed")
print(f"Thunderbird folder: {DEST}")
print("Point Thunderbird Local Folders to this directory.")

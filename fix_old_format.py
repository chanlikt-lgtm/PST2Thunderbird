"""Fix 2001.pst and 2003.pst by copying from E:\Mail_mbox\ raw readpst output."""
import shutil
from pathlib import Path

OUT_BASE = Path(r"E:\TB_Mail_v2")

def sanitize(name):
    for c in r'<>:"/\|?*':
        name = name.replace(c, "_")
    return name.strip() or "Unknown"

def copy_raw_to_tb(mbox_src_root, tb_name):
    """
    Copy readpst mbox structure to Thunderbird .sbd format.
    mbox_src_root: E:\Mail_mbox\<period>\<pst_root>\
    tb_name: name to use in TB_Mail_v2
    """
    container = OUT_BASE / tb_name
    sbd_dir   = OUT_BASE / (tb_name + ".sbd")
    container.touch()
    sbd_dir.mkdir(exist_ok=True)

    # Recursively find all mbox files and map to .sbd structure
    def copy_folder(src_folder, dest_dir):
        mbox_file = src_folder / "mbox"
        total = 0
        if mbox_file.exists() and mbox_file.stat().st_size > 0:
            dest_file = dest_dir / sanitize(src_folder.name)
            dest_file.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(mbox_file, dest_file)
            size = mbox_file.stat().st_size
            print(f"  {src_folder.name}: {size//1024} KB", flush=True)
            total += size
        # Recurse into subdirs
        for sub in sorted(src_folder.iterdir()):
            if sub.is_dir() and not sub.name.startswith("."):
                sub_dest = dest_dir / (sanitize(src_folder.name) + ".sbd")
                total += copy_folder(sub, sub_dest)
        return total

    total = 0
    for sub in sorted(mbox_src_root.iterdir()):
        if sub.is_dir() and not sub.name.startswith("."):
            total += copy_folder(sub, sbd_dir)

    print(f"  -> {total//1024//1024} MB total\n", flush=True)

# 2001.pst — root dir is E:\Mail_mbox\2001\2001 cl tan\
print("Fixing: 2001.pst", flush=True)
root_2001 = Path(r"E:\Mail_mbox\2001")
subdirs = [d for d in root_2001.iterdir() if d.is_dir()]
if subdirs:
    copy_raw_to_tb(subdirs[0], "2001")
else:
    print("  No subdirs found\n", flush=True)

# 2003.pst
print("Fixing: 2003 PART1  1Si cltan.pst", flush=True)
root_2003 = Path(r"E:\Mail_mbox\2003 PART1  1Si cltan")
subdirs = [d for d in root_2003.iterdir() if d.is_dir()]
if subdirs:
    copy_raw_to_tb(subdirs[0], "2003 PART1  1Si cltan")
else:
    print("  No subdirs found\n", flush=True)

# INFINEON Sep 2007 — no mbox files, skip
print("INFINEON  Sep 2007.pst: no recoverable data (0 mbox files in raw backup)", flush=True)

print("Done.", flush=True)

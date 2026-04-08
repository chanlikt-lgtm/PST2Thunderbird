"""
Convert F:\Mail_mbox\raw\ (readpst directory format) to Thunderbird .sbd structure.
Each 'mbox' file inside subdirectories becomes a separate folder in Thunderbird.

readpst structure:
  Mail_mbox\raw\2001\
    2001 cl tan\          <- root folder (skip this level, use as container name)
      Inbox\
        mbox              <- actual MBOX file
      Sent Items\
        mbox
      ...

Thunderbird output:
  TB_Mail_v2\2001         <- empty container file
  TB_Mail_v2\2001.sbd\
    Inbox                 <- MBOX file
    Sent Items            <- MBOX file
    ...
"""
import os
import shutil
from pathlib import Path

SRC      = Path(r"E:\Mail_mbox")
OUT_BASE = Path(r"E:\TB_Mail_v2")
OUT_BASE.mkdir(exist_ok=True)

def sanitize(name):
    for c in r'<>:"/\|?*':
        name = name.replace(c, "_")
    return name.strip() or "Unknown"

def find_mbox_entries(folder):
    """
    Recursively find all (relative_path_parts, mbox_file_path) pairs.
    Skips the first directory level (the PST root folder name).
    """
    results = []
    for item in sorted(folder.iterdir()):
        if item.name.startswith("."):
            continue
        if item.is_dir():
            # Recurse
            sub = find_mbox_entries_with_path(item, [])
            results.extend(sub)
        elif item.name == "mbox":
            results.append(([], item))
    return results

def find_mbox_entries_with_path(folder, path_parts):
    results = []
    mbox_file = folder / "mbox"
    if mbox_file.exists() and mbox_file.stat().st_size > 100:
        results.append((path_parts, mbox_file))
    for item in sorted(folder.iterdir()):
        if item.name.startswith(".") or item.name == "mbox":
            continue
        if item.is_dir():
            results.extend(find_mbox_entries_with_path(item, path_parts + [item.name]))
    return results

def tb_mbox_path(out_base, period_name, folder_path):
    """
    Build Thunderbird path for a folder.
    folder_path = ['Inbox'] -> out_base/period_name.sbd/Inbox
    folder_path = ['Top', 'Sent Items'] -> out_base/period_name.sbd/Top.sbd/Sent Items
    """
    sbd = out_base / (period_name + ".sbd")
    if not folder_path:
        return sbd / "Inbox"
    path = sbd
    for part in folder_path[:-1]:
        path = path / (sanitize(part) + ".sbd")
    path = path / sanitize(folder_path[-1])
    return path

converted = 0
skipped = 0

for entry in sorted(SRC.iterdir()):
    if not entry.is_dir():
        continue

    period_name = sanitize(entry.name)
    container   = OUT_BASE / period_name
    sbd_dir     = OUT_BASE / (period_name + ".sbd")

    if container.exists() or sbd_dir.exists():
        print(f"SKIP (exists): {period_name}")
        skipped += 1
        continue

    print(f"Processing: {period_name} ...", flush=True)

    # Find the root subfolder (one level down, e.g. "2001 cl tan")
    subdirs = [d for d in entry.iterdir() if d.is_dir() and not d.name.startswith(".")]
    if not subdirs:
        print(f"  SKIP: no subdirs")
        continue

    # Collect all mbox files with their relative paths
    all_mboxes = []
    for root_subdir in subdirs:
        # Skip the first level (PST root folder name), go one deeper
        entries = find_mbox_entries_with_path(root_subdir, [])
        all_mboxes.extend(entries)

    if not all_mboxes:
        print(f"  SKIP: no mbox files found")
        continue

    # Create container file and .sbd directory
    container.touch()
    sbd_dir.mkdir()

    total_size = 0
    for path_parts, mbox_file in all_mboxes:
        dest = tb_mbox_path(OUT_BASE, period_name, path_parts)
        dest.parent.mkdir(parents=True, exist_ok=True)
        try:
            shutil.copy2(mbox_file, dest)
            size = mbox_file.stat().st_size
            total_size += size
            label = "/".join(path_parts) if path_parts else "Inbox"
            print(f"  {label}: {size//1024//1024} MB", flush=True)
        except Exception as e:
            print(f"  ERROR copying {path_parts}: {e}", flush=True)

    print(f"  -> Total: {total_size//1024//1024} MB", flush=True)
    converted += 1

print()
print(f"Done: {converted} processed, {skipped} skipped")
print(f"Output: {OUT_BASE}")

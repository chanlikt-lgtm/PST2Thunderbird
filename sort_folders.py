"""
sort_folders.py — Prefix Thunderbird Local Folder names with YYYY-MM
so they sort chronologically in Thunderbird.

  "April 2014"  →  "2014-04 April 2014"
  "Aug 2012"    →  "2012-08 Aug 2012"
  "2007 June"   →  "2007-06 2007 June"
  "2022"        →  "2022-00 2022"

Renames: container file, .sbd directory, .msf file.
Folders with no detectable date get prefix "zz " and sort to the bottom.

Usage:
    python sort_folders.py --dry-run    (preview, no changes)
    python sort_folders.py              (apply)
"""

import sys
import io
import re
from pathlib import Path

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

OUT_BASE = Path(r"E:\TB_Mail_v2")

# Month name → zero-padded number  (Mac = March in Malay)
MONTHS = {
    'jan': '01', 'january': '01',
    'feb': '02', 'february': '02',
    'mac': '03', 'mar': '03', 'march': '03',
    'apr': '04', 'april': '04',
    'may': '05',
    'jun': '06', 'june': '06',
    'jul': '07', 'july': '07',
    'aug': '08', 'august': '08',
    'sep': '09', 'sept': '09', 'september': '09',
    'oct': '10', 'october': '10',
    'nov': '11', 'november': '11',
    'dec': '12', 'december': '12',
}

# These are Thunderbird system entries — never rename
SKIP = {
    'Trash', 'Unsent Messages', 'msgFilterRules.dat',
    'Inbox', 'Sent', 'Drafts', 'Templates', 'Archives',
}


def parse_date(name: str):
    """
    Returns (year_str, month_str) e.g. ('2014', '04').
    Falls back to ('9999', '99') if no date found.
    """
    # 4-digit year (no adjacent digit check avoids underscore word-boundary issue)
    year_match = re.search(r'(?<!\d)((?:19|20)\d{2})(?!\d)', name)
    year = year_match.group(1) if year_match else None

    # 2-digit year fallback (e.g. "TanChanNov05", "Nov05")
    if year is None:
        two = re.search(r'(?<!\d)(\d{2})(?!\d)', name)
        if two:
            n = int(two.group(1))
            year = f"20{n:02d}" if n <= 30 else f"19{n:02d}"

    # Month name — check whole words first, then substrings (handles camelCase like TanChanNov05)
    month = '00'
    name_lower = name.lower()
    for word in re.findall(r'[A-Za-z]+', name):
        m = MONTHS.get(word.lower())
        if m:
            month = m
            break
    if month == '00':
        # substring search for camelCase names
        for mon_name, mon_num in sorted(MONTHS.items(), key=lambda x: -len(x[0])):
            if mon_name in name_lower:
                month = mon_num
                break

    if year:
        return year, month
    return '9999', '99'


def make_prefix(name: str) -> str:
    year, month = parse_date(name)
    if year == '9999':
        return 'zz'          # undated → bottom
    return f"{year}-{month}"


def already_prefixed(name: str) -> bool:
    """True if name already starts with YYYY-MM or 'zz '."""
    return bool(re.match(r'^(20|19)\d{2}-\d{2} |^zz ', name))


def main():
    dry_run = '--dry-run' in sys.argv

    # Collect all container files (no extension, not .sbd/.msf)
    entries = []
    for p in sorted(OUT_BASE.iterdir()):
        if p.suffix in ('.sbd', '.msf', '.dat'):
            continue
        if p.name in SKIP:
            continue
        if already_prefixed(p.name):
            continue
        entries.append(p)

    if not entries:
        print("Nothing to rename — all folders already prefixed.")
        return

    print(f"{'DRY RUN — ' if dry_run else ''}Renaming {len(entries)} folder(s) in {OUT_BASE}\n")
    print(f"  {'OLD NAME':<45}  ->  NEW NAME")
    print(f"  {'-'*45}     {'-'*45}")

    renames = []
    for p in entries:
        prefix   = make_prefix(p.name)
        new_name = f"{prefix} {p.name}"
        print(f"  {p.name:<45}  ->  {new_name}")
        renames.append((p.name, new_name))

    print()

    if dry_run:
        print("DRY RUN — no changes made.")
        print("Re-run without --dry-run to apply.")
        return

    # Apply renames
    errors = 0
    for old_name, new_name in renames:
        for suffix in ('', '.sbd', '.msf'):
            old_path = OUT_BASE / (old_name + suffix)
            new_path = OUT_BASE / (new_name + suffix)
            if old_path.exists():
                try:
                    old_path.rename(new_path)
                except Exception as e:
                    print(f"  ERROR renaming {old_path.name}: {e}")
                    errors += 1

    print(f"Done. {len(renames)} folders renamed, {errors} error(s).")
    print("Restart Thunderbird to see sorted folders.")


if __name__ == '__main__':
    main()

"""
gap_check.py — Check for monthly gaps in PST coverage.

Parses all PST filenames, maps them to YYYY-MM, then finds missing
months between the earliest and latest date found (or user-defined range).

Usage:
    py -3.11 gap_check.py                    # auto range from PST dates
    py -3.11 gap_check.py --from 1999-01 --to 2024-01
"""

import sys
import io
import re
from pathlib import Path
from datetime import date

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

PST_DIR = Path(r"E:\PST")

MONTHS = {
    'jan': 1, 'january': 1,
    'feb': 2, 'february': 2,
    'mac': 3, 'mar': 3, 'march': 3,
    'apr': 4, 'april': 4,
    'may': 5,
    'jun': 6, 'june': 6,
    'jul': 7, 'july': 7,
    'aug': 8, 'august': 8,
    'sep': 9, 'sept': 9, 'september': 9,
    'oct': 10, 'october': 10,
    'nov': 11, 'november': 11,
    'dec': 12, 'december': 12,
}

# PSTs too small to contain real mail — skip from coverage map
SKIP_SMALL = {
    'archive', 'mailbox', 'outlook', 'tanchangov05',
    'infineon 25 sep 2007',
}


def parse_pst_date(stem: str):
    """
    Returns (year, month_or_None) from a PST stem.
    month=None means we only know the year (covers full year).
    Returns None if no date found.
    """
    name = stem.strip()

    # 4-digit year
    year_m = re.search(r'(?<!\d)((?:19|20)\d{2})(?!\d)', name)
    year = int(year_m.group(1)) if year_m else None

    # 2-digit year fallback (e.g. TanChanNov05)
    if year is None:
        two = re.search(r'(?<!\d)(\d{2})(?!\d)', name)
        if two:
            n = int(two.group(1))
            year = 2000 + n if n <= 30 else 1900 + n

    if year is None:
        return None

    # Month
    month = None
    for word in re.findall(r'[A-Za-z]+', name):
        m = MONTHS.get(word.lower())
        if m:
            month = m
            break
    if month is None:
        name_lower = name.lower()
        for mon, num in sorted(MONTHS.items(), key=lambda x: -len(x[0])):
            if mon in name_lower:
                month = num
                break

    return (year, month)


def main():
    # Parse --from / --to args
    arg_from = arg_to = None
    args = sys.argv[1:]
    for i, a in enumerate(args):
        if a == "--from" and i + 1 < len(args):
            try:
                y, m = args[i+1].split("-")
                arg_from = (int(y), int(m))
            except: pass
        if a == "--to" and i + 1 < len(args):
            try:
                y, m = args[i+1].split("-")
                arg_to = (int(y), int(m))
            except: pass

    pst_files = sorted(PST_DIR.glob("*.pst"))
    print(f"Analysing {len(pst_files)} PST files...\n", flush=True)

    # Map each PST to covered months
    covered = set()      # set of (year, month) tuples
    year_only = set()    # years with no month info (whole year assumed covered)
    parsed = []

    for pst in pst_files:
        stem = pst.stem
        size_mb = pst.stat().st_size / 1024 / 1024

        if size_mb < 1:  # skip tiny/empty PSTs
            continue
        if stem.lower() in SKIP_SMALL:
            continue

        result = parse_pst_date(stem)
        if result is None:
            continue

        year, month = result
        if month:
            covered.add((year, month))
            parsed.append((stem, year, month, size_mb))
        else:
            year_only.add(year)
            # Mark all 12 months for year-only PSTs
            for m in range(1, 13):
                covered.add((year, m))
            parsed.append((stem, year, None, size_mb))

    if not covered:
        print("No dates found in PST filenames.", flush=True)
        return

    # Determine range
    all_years  = [y for y, m in covered]
    start_year = arg_from[0] if arg_from else min(all_years)
    start_mon  = arg_from[1] if arg_from else 1
    end_year   = arg_to[0]   if arg_to   else max(all_years)
    end_mon    = arg_to[1]   if arg_to   else 12

    print(f"Coverage range: {start_year}-{start_mon:02d}  to  {end_year}-{end_mon:02d}\n", flush=True)

    # Find gaps
    gaps = []
    y, m = start_year, start_mon
    while (y, m) <= (end_year, end_mon):
        if (y, m) not in covered:
            gaps.append((y, m))
        m += 1
        if m > 12:
            m = 1
            y += 1

    # Print calendar view
    MONTH_NAMES = ['', 'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun',
                       'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']

    print("Calendar (O=covered  .=missing  *=year-only):\n", flush=True)
    print(f"  {'Year':>4}  " + "  ".join(MONTH_NAMES[1:]), flush=True)
    print(f"  {'----':>4}  " + "  ".join(["---"]*12), flush=True)

    for yr in range(start_year, end_year + 1):
        row = f"  {yr:>4}  "
        for mo in range(1, 13):
            if yr == start_year and mo < start_mon:
                row += "    "
            elif yr == end_year and mo > end_mon:
                row += "    "
            elif (yr, mo) in covered:
                if yr in year_only:
                    row += " *  "
                else:
                    row += " O  "
            else:
                row += " .  "
        print(row, flush=True)

    # Gap list
    print(f"\n{'─'*50}", flush=True)
    if gaps:
        print(f"GAPS ({len(gaps)} missing months):\n", flush=True)
        # Group consecutive months
        groups = []
        start = gaps[0]
        prev  = gaps[0]
        for g in gaps[1:]:
            py, pm = prev
            gy, gm = g
            next_m = pm + 1 if pm < 12 else 1
            next_y = py if pm < 12 else py + 1
            if (gy, gm) == (next_y, next_m):
                prev = g
            else:
                groups.append((start, prev))
                start = prev = g
        groups.append((start, prev))

        for (sy, sm), (ey, em) in groups:
            if (sy, sm) == (ey, em):
                print(f"  {sy}-{sm:02d}  ({MONTH_NAMES[sm]} {sy})", flush=True)
            else:
                print(f"  {sy}-{sm:02d} → {ey}-{em:02d}  "
                      f"({MONTH_NAMES[sm]} {sy} – {MONTH_NAMES[em]} {ey})", flush=True)
    else:
        print("No gaps found — full coverage!", flush=True)

    print(f"\nTotal covered: {len(covered)} months", flush=True)
    print(f"Total missing: {len(gaps)} months", flush=True)


if __name__ == "__main__":
    main()

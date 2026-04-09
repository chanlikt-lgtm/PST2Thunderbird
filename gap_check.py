"""
gap_check.py — Check for monthly gaps in email coverage.

Scans ACTUAL email dates inside every mbox file in E:\\TB_Mail_v2\\
to build a true coverage map. PST filenames are unreliable — a file
named "2001.pst" may contain emails from 1999 onwards.

Usage:
    py -3.11 gap_check.py                    # auto range from actual dates
    py -3.11 gap_check.py --from 1999-01 --to 2024-01
"""

import sys
import io
import re
import mailbox
import email.utils
from pathlib import Path
from datetime import datetime

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

OUT_BASE = Path(r"E:\TB_Mail_v2")

MONTH_NAMES = ['', 'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun',
                   'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']


def parse_date(msg) -> tuple:
    """Extract (year, month) from email Date header. Returns None if unparseable."""
    date_str = msg.get("Date", "")
    if not date_str:
        return None
    try:
        t = email.utils.parsedate_tz(date_str)
        if t:
            dt = datetime(*t[:6])
            if 1990 <= dt.year <= 2030:
                return (dt.year, dt.month)
    except Exception:
        pass
    return None


def main():
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

    folders = sorted(OUT_BASE.glob("*.sbd"))
    print(f"Scanning actual email dates in {len(folders)} folders...", flush=True)
    print(f"(This reads every email header — may take a few minutes)\n", flush=True)

    covered    = set()   # (year, month) tuples with at least 1 email
    total_read = 0
    total_skip = 0

    for sbd in folders:
        mbox_files = [f for f in sbd.rglob("*")
                      if f.is_file() and f.suffix != ".msf"]
        for mf in mbox_files:
            try:
                mb = mailbox.mbox(str(mf))
                for msg in mb:
                    total_read += 1
                    ym = parse_date(msg)
                    if ym:
                        covered.add(ym)
                    else:
                        total_skip += 1
            except Exception:
                pass
        sys.stdout.write(f"\r  Processed: {total_read:,} emails, {len(covered)} months covered...   ")
        sys.stdout.flush()

    print(f"\n\nDone. Read {total_read:,} emails ({total_skip:,} had unparseable dates).\n", flush=True)

    if not covered:
        print("No dates found.", flush=True)
        return

    all_years  = [y for y, m in covered]
    start_year = arg_from[0] if arg_from else min(all_years)
    start_mon  = arg_from[1] if arg_from else 1
    end_year   = arg_to[0]   if arg_to   else max(all_years)
    end_mon    = arg_to[1]   if arg_to   else 12

    print(f"Actual email date range: {min(all_years)}-{min(m for y,m in covered if y==min(all_years)):02d}"
          f"  to  {max(all_years)}-{max(m for y,m in covered if y==max(all_years)):02d}", flush=True)
    print(f"Checking range         : {start_year}-{start_mon:02d}  to  {end_year}-{end_mon:02d}\n", flush=True)

    # Find gaps
    gaps = []
    y, m = start_year, start_mon
    while (y < end_year) or (y == end_year and m <= end_mon):
        if (y, m) not in covered:
            gaps.append((y, m))
        m += 1
        if m > 12:
            m = 1
            y += 1

    # Calendar view
    print("Calendar (O=has email  .=no email):\n", flush=True)
    print(f"  {'Year':>4}  " + " ".join(f"{n:>3}" for n in MONTH_NAMES[1:]), flush=True)
    print(f"  {'----':>4}  " + " ".join(["---"]*12), flush=True)

    for yr in range(start_year, end_year + 1):
        row = f"  {yr:>4}  "
        for mo in range(1, 13):
            if yr == start_year and mo < start_mon:
                row += "    "
            elif yr == end_year and mo > end_mon:
                row += "    "
            elif (yr, mo) in covered:
                row += "  O "
            else:
                row += "  . "
        print(row, flush=True)

    # Gap summary
    print(f"\n{'─'*50}", flush=True)
    if gaps:
        print(f"GAPS ({len(gaps)} months with zero emails):\n", flush=True)
        groups = []
        start_g = prev = gaps[0]
        for g in gaps[1:]:
            py, pm = prev
            next_ym = (py, pm + 1) if pm < 12 else (py + 1, 1)
            if g == next_ym:
                prev = g
            else:
                groups.append((start_g, prev))
                start_g = prev = g
        groups.append((start_g, prev))

        for (sy, sm), (ey, em) in groups:
            if (sy, sm) == (ey, em):
                print(f"  {sy}-{sm:02d}  ({MONTH_NAMES[sm]} {sy})", flush=True)
            else:
                print(f"  {sy}-{sm:02d} -> {ey}-{em:02d}  "
                      f"({MONTH_NAMES[sm]} {sy} - {MONTH_NAMES[em]} {ey})", flush=True)
    else:
        print("No gaps — full coverage!", flush=True)

    print(f"\nTotal covered : {len(covered)} months", flush=True)
    print(f"Total missing : {len(gaps)} months", flush=True)


if __name__ == "__main__":
    main()

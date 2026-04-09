"""
verify_sizes.py — Check that every converted PST folder is a reasonable
size compared to its source PST file.

Rule: output mbox total >= 5% of PST size. Below that = suspect.
(A 2 GB PST converting to 20 MB mbox is clearly wrong.)

The script matches PST stems to TB folders by stripping any YYYY-MM prefix
so that "2019-12 December 2019.sbd" matches "December 2019.pst".

Usage:
    py -3.11 verify_sizes.py
    py -3.11 verify_sizes.py --fix    (re-run reconvert.py on suspect ones)
"""

import sys
import io
import re
import subprocess
from pathlib import Path

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

PST_DIR  = Path(r"E:\PST")
OUT_BASE = Path(r"E:\TB_Mail_v2")
MIN_RATIO = 0.05   # 5%

STRIP_PREFIX = re.compile(r'^(?:19|20)\d{2}-\d{2} ')


def sbd_content_bytes(sbd_dir: Path) -> int:
    """Total mbox bytes inside an .sbd folder (excludes .msf and .lock)."""
    return sum(f.stat().st_size for f in sbd_dir.rglob("*")
               if f.is_file() and f.suffix not in (".msf", ".lock"))


def main():
    fix_mode = "--fix" in sys.argv

    # Index all .sbd dirs by their stem with prefix stripped
    sbd_map = {}   # stripped_stem_lower -> sbd Path
    for sbd in OUT_BASE.glob("*.sbd"):
        raw = sbd.stem                          # e.g. "2019-12 December 2019"
        stripped = STRIP_PREFIX.sub("", raw)   # e.g. "December 2019"
        sbd_map[stripped.lower()] = sbd

    pst_files = sorted(PST_DIR.glob("*.pst"))
    print(f"PSTs in {PST_DIR}  : {len(pst_files)}")
    print(f"Converted folders : {len(sbd_map)}")
    print()

    fmt = "  {flag}  {pst:<35}  PST={pst_mb:>6} MB   out={out_mb:>6} MB   {ratio:>6}   {note}"
    header = fmt.format(flag=" ", pst="PST file", pst_mb=0, out_mb=0, ratio="ratio", note="status")
    header = "     {:<35}  {:>10}   {:>10}   {:>8}   {}".format(
        "PST file", "PST size", "out size", "ratio", "status")
    print(header)
    print("  " + "-" * 80)

    ok_count      = 0
    warn_count    = 0
    missing_count = 0
    suspects      = []

    for pst in pst_files:
        stem  = pst.stem
        pst_bytes = pst.stat().st_size
        pst_mb    = pst_bytes // (1024 * 1024)

        sbd = sbd_map.get(stem.lower())
        if sbd is None:
            print(f"  ??  {stem:<35}  PST={pst_mb:>6} MB   NO OUTPUT FOLDER", flush=True)
            missing_count += 1
            continue

        out_bytes = sbd_content_bytes(sbd)
        out_mb    = out_bytes // (1024 * 1024)
        ratio     = out_bytes / pst_bytes if pst_bytes else 0

        if ratio < MIN_RATIO:
            flag  = "!!"
            note  = f"SUSPECT  ({ratio*100:.1f}%)"
            warn_count += 1
            suspects.append(pst)
        else:
            flag  = "  "
            note  = f"OK       ({ratio*100:.0f}%)"
            ok_count += 1

        print(f"  {flag}  {stem:<35}  PST={pst_mb:>6} MB   out={out_mb:>6} MB   {note}",
              flush=True)

    print()
    print(f"{'─'*55}", flush=True)
    print(f"  OK      : {ok_count}", flush=True)
    print(f"  SUSPECT : {warn_count}", flush=True)
    print(f"  MISSING : {missing_count}", flush=True)

    if suspects:
        print()
        print("SUSPECT conversions (output < 5% of PST size):", flush=True)
        for p in suspects:
            print(f"  {p.name}", flush=True)

        if fix_mode:
            print()
            print("-- FIX MODE: re-running reconvert.py on suspects --", flush=True)
            stems = [p.stem for p in suspects]
            cmd = [sys.executable, str(Path(__file__).parent / "reconvert.py")] + stems
            subprocess.run(cmd)
        else:
            print()
            print("Run with --fix to reconvert suspect PSTs:", flush=True)
            print("  py -3.11 verify_sizes.py --fix", flush=True)


if __name__ == "__main__":
    main()

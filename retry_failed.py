"""
retry_failed.py — Re-run only failed PSTs from the most recent reconvert batch.

Sources (tried in order):
  1. failed_psts.log  — written by hardened reconvert.py (future runs)
  2. overnight.log    — legacy batch output
  3. Any *.output file passed as argument

Usage:
    py -3.11 retry_failed.py                   # auto-detect source
    py -3.11 retry_failed.py path\to\run.output  # explicit output file
"""

import sys
import re
import subprocess
from pathlib import Path

PST_DIR    = Path(r"E:\PST")
OUT_BASE   = Path(r"E:\TB_Mail_v2")
FAILED_LOG = Path(r"E:\claude\Pst2Thunder\failed_psts.log")
OVERNIGHT  = Path(r"E:\claude\Pst2Thunder\overnight.log")
RECONVERT  = Path(r"E:\claude\Pst2Thunder\reconvert.py")


def stems_from_failed_log(path: Path) -> list:
    """Parse failed_psts.log: '<full_path> | <error>'"""
    stems = []
    with path.open(encoding="utf-8", errors="replace") as f:
        for line in f:
            line = line.strip()
            if not line:
                continue
            pst = line.split("|")[0].strip()
            if pst:
                stems.append(Path(pst).stem)
    return stems


def stems_from_output_log(path: Path) -> list:
    """
    Parse a batch output log.
    Looks for pairs:
        ── SomePST.pst
        ...
          FAILED: ...
    """
    data = path.read_bytes().decode("utf-8", errors="replace")
    lines = data.replace("\r", "\n").split("\n")
    stems = []
    current_pst = None
    for line in lines:
        # Header line: ── filename.pst
        m = re.match(r"[\u2500\-]{2}\s+(.+\.pst)", line.strip())
        if m:
            current_pst = Path(m.group(1)).stem
            continue
        if current_pst and re.search(r"\bFAILED\b", line):
            stems.append(current_pst)
            current_pst = None   # only capture first FAILED per PST
    return stems


def collect_failed_stems() -> list:
    # Explicit output file passed as arg
    if len(sys.argv) > 1:
        src = Path(sys.argv[1])
        if src.exists():
            print(f"Reading failures from: {src}")
            return stems_from_output_log(src)
        else:
            print(f"File not found: {src}")
            sys.exit(1)

    # Prefer failed_psts.log (hardened runs)
    if FAILED_LOG.exists() and FAILED_LOG.stat().st_size > 0:
        print(f"Reading failures from: {FAILED_LOG}")
        return stems_from_failed_log(FAILED_LOG)

    # Fall back to overnight.log (legacy / current batch)
    if OVERNIGHT.exists():
        print(f"failed_psts.log not found — parsing: {OVERNIGHT}")
        return stems_from_output_log(OVERNIGHT)

    print("No failure source found. Checked:")
    print(f"  {FAILED_LOG}")
    print(f"  {OVERNIGHT}")
    sys.exit(1)


def main():
    stems = collect_failed_stems()
    unique = sorted(set(s for s in stems if s))

    if not unique:
        print("No failed PSTs found.")
        return

    # Verify each PST file exists before retrying
    to_retry = []
    missing  = []
    for stem in unique:
        pst = PST_DIR / (stem + ".pst")
        if pst.exists():
            to_retry.append(stem)
        else:
            missing.append(stem)

    if missing:
        print(f"\nSkipping (PST file not found): {len(missing)}")
        for s in missing[:20]:
            print(f"  {s}")
        if len(missing) > 20:
            print(f"  ... +{len(missing) - 20} more")

    if not to_retry:
        print("Nothing to retry.")
        return

    # Filter out PSTs that already have non-empty output (from a previous success)
    skipped   = []
    pending   = []
    for stem in to_retry:
        out_dir = OUT_BASE / (stem + ".sbd")
        if out_dir.exists() and any(out_dir.iterdir()):
            skipped.append(stem)
        else:
            pending.append(stem)

    if skipped:
        print(f"\nSkipping (already has output): {len(skipped)}")
        for s in skipped[:20]:
            print(f"  {s}")
        if len(skipped) > 20:
            print(f"  ... +{len(skipped) - 20} more")

    if not pending:
        print("Nothing left to retry after skipping already-converted PSTs.")
        return

    print(f"\nRetrying {len(pending)} PST(s):\n")
    succeeded = 0
    failed    = 0
    for name in pending:
        print(f"── {name}.pst")
        result = subprocess.run(
            ["py", "-3.11", str(RECONVERT), name],
            cwd=str(RECONVERT.parent),
        )
        if result.returncode != 0:
            print(f"  !! Returned exit code {result.returncode}")
            try:
                with FAILED_LOG.open("a", encoding="utf-8") as lf:
                    lf.write(f"{PST_DIR / (name + '.pst')} | retry exit {result.returncode}\n")
            except Exception:
                pass
            failed += 1
        else:
            succeeded += 1
        print()

    print(f"Retried : {len(pending)} PSTs")
    print(f"  OK    : {succeeded}")
    print(f"  Failed: {failed}")
    if skipped:
        print(f"  Skipped (had output): {len(skipped)}")
    print("Restart Thunderbird to see updated content.")


if __name__ == "__main__":
    main()

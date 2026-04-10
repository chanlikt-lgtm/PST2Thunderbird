"""
check_pst.py — Verify if a PST file is valid before copying.
Usage: py check_pst.py "path\to\file.pst"
       py check_pst.py          (checks all .pst in current folder)
"""
import os, sys, glob

def check(path):
    size = os.path.getsize(path)
    print(f"\n{'='*50}")
    print(f"File : {path}")
    print(f"Size : {size/1024/1024/1024:.2f} GB ({size:,} bytes)")

    with open(path, "rb") as f:
        header = f.read(8)
        magic_ok = header[:4] == b'!BDN'
        print(f"Magic: {header.hex()}  =>  {'VALID !BDN' if magic_ok else 'BAD / NOT PST'}")

        # Find first non-zero byte
        f.seek(0)
        pos = 0
        first_data = None
        while pos < min(size, 10*1024*1024):
            buf = f.read(65536)
            if not buf: break
            idx = next((i for i,b in enumerate(buf) if b != 0), None)
            if idx is not None:
                first_data = pos + idx
                break
            pos += len(buf)

        if first_data == 0:
            print(f"Data : starts at byte 0  =>  GOOD")
        elif first_data:
            print(f"Data : first non-zero at {first_data:,} bytes ({first_data/1024:.0f} KB)  =>  HEADER CORRUPT")
        else:
            print(f"Data : all zeros in first 10 MB  =>  CORRUPT / EMPTY")

    verdict = "GOOD — safe to copy" if magic_ok and first_data == 0 else "BAD — do not use"
    print(f"=> {verdict}")

if len(sys.argv) > 1:
    for arg in sys.argv[1:]:
        matches = glob.glob(arg)
        if matches:
            for m in sorted(matches):
                check(m)
        elif os.path.exists(arg):
            check(arg)
        else:
            print(f"Not found: {arg}")
else:
    psts = glob.glob("*.pst") + glob.glob("*.PST")
    if not psts:
        print("No .pst files found in current folder.")
        print("Usage: py check_pst.py path\\to\\file.pst")
    else:
        for p in sorted(psts):
            check(p)

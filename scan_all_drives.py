"""
scan_all_drives.py — Find all PST files on all drives, compare with E:\PST\
"""
import os, sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

EPST = r'E:\PST'
SKIP = {'windows', 'program files', 'program files (x86)', 'programdata',
        'appdata', '$recycle.bin', 'system volume information', 'recovery'}

# Build index of known PSTs in E:\PST
known = {}
for f in os.listdir(EPST):
    if f.lower().endswith('.pst'):
        p = os.path.join(EPST, f)
        known[f.lower()] = os.path.getsize(p)

def check_magic(path):
    try:
        with open(path, 'rb') as f:
            return f.read(4) == b'!BDN'
    except:
        return False

results = {'new': [], 'larger': [], 'corrupt': [], 'dup': []}

DRIVES = ['C:\\', 'F:\\', 'G:\\']

for drive in DRIVES:
    print(f'\nScanning {drive} ...', flush=True)
    for root, dirs, files in os.walk(drive, topdown=True):
        dirs[:] = [d for d in dirs if d.lower() not in SKIP]
        for f in files:
            if not f.lower().endswith('.pst'):
                continue
            p = os.path.join(root, f)
            try:
                size = os.path.getsize(p)
                if size < 1024 * 100:  # skip under 100 KB
                    continue
                valid = check_magic(p)
                key = f.lower()
                in_known = key in known
                size_diff = size - known.get(key, 0) if in_known else None

                if not valid:
                    results['corrupt'].append((p, size))
                elif not in_known:
                    results['new'].append((p, size))
                elif size_diff > 10 * 1024 * 1024:
                    results['larger'].append((p, size, size_diff))
                else:
                    results['dup'].append((p, size))
            except:
                pass

print('\n' + '='*60)
print(f"*** NEW PSTs not in E:\\PST ({len(results['new'])}):")
for p, s in sorted(results['new']): print(f"  {s//1024//1024:>6} MB  {p}")

print(f"\n*** LARGER than E:\\PST version ({len(results['larger'])}):")
for p, s, d in sorted(results['larger']): print(f"  {s//1024//1024:>6} MB  (+{d//1024//1024} MB)  {p}")

print(f"\n*** CORRUPT ({len(results['corrupt'])}):")
for p, s in sorted(results['corrupt']): print(f"  {s//1024//1024:>6} MB  {p}")

print(f"\n  Duplicates: {len(results['dup'])} (skipped)")
print('\nDone.')

"""
reconvert_with_backup.py — Backup existing mbox folders then reconvert PSTs.
Backups go to E:\TB_Mail_v2\_backup\YYYY-MM-DD\
"""
import os, sys, io, shutil, datetime
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

PST_DIR  = r'E:\PST'
OUT_BASE = r'E:\TB_Mail_v2'
BACKUP   = os.path.join(OUT_BASE, '_backup', datetime.date.today().isoformat())

STEMS = [
    'PART2 TAN CHAN LIK 2004',
    'cltan_2003',
    'cltan_backup_2002',
    '2003 PART1  1Si cltan',
    'APRIL 2011',
    'Dec 2012',
    'July 2012',
    'July 2017',
    '2024 Jan',
    'April 2025',
    'April 2012',
]

os.makedirs(BACKUP, exist_ok=True)
print(f'Backup folder: {BACKUP}')
print()

# ── Backup existing mbox + .sbd for each stem ─────────────────────────────
backed_up = 0
for stem in STEMS:
    # Find matching folder in TB_Mail_v2 (may have YYYY-MM prefix)
    matches = [
        e for e in os.listdir(OUT_BASE)
        if stem.lower() in e.lower()
        and not e.startswith('_backup')
    ]
    if not matches:
        print(f'  (no existing folder for: {stem})')
        continue
    for name in matches:
        src = os.path.join(OUT_BASE, name)
        dst = os.path.join(BACKUP, name)
        try:
            if os.path.isdir(src):
                shutil.copytree(src, dst)
            else:
                shutil.copy2(src, dst)
            print(f'  BACKED UP: {name}')
            backed_up += 1
        except Exception as e:
            print(f'  BACKUP ERROR {name}: {e}')

print(f'\nBacked up {backed_up} items to {BACKUP}')
print('Starting reconversion...')
print()

# ── Reconvert ─────────────────────────────────────────────────────────────
import subprocess
args = [r'py', '-3.11', r'E:\claude\Pst2Thunder\reconvert.py'] + STEMS
result = subprocess.run(args)
sys.exit(result.returncode)

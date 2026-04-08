"""
auto_flow.py — Automatically process new PST files dropped into E:\PST\Incoming\

Run manually or via Windows Task Scheduler (recommended: every hour).

Flow:
    1. Scan E:\PST\Incoming\ for *.pst files
    2. Duplicate check: filename (case-insensitive) + MD5 hash vs E:\PST\
    3. Copy new PSTs to E:\PST\
    4. Convert to E:\TB_Mail_v2\ (full MIME + attachments)
    5. Prefix new folders with YYYY-MM for chronological sort
    6. Move processed file out of Incoming (to E:\PST\Incoming\Done\)
    7. Log everything to auto_flow.log

Usage:
    py -3.11 auto_flow.py              # process Incoming folder
    py -3.11 auto_flow.py --dry-run    # preview only, no changes
"""

import sys
import io
import hashlib
import shutil
import mailbox
import mimetypes
import email
import email.mime.multipart
import email.mime.text
import email.mime.base
import re
import logging
from email import encoders
from datetime import datetime
from pathlib import Path
from libratom.lib.pff import PffArchive

# ── Config ────────────────────────────────────────────────────────────────────
INCOMING  = Path(r"E:\PST\Incoming")
DONE_DIR  = Path(r"E:\PST\Incoming\Done")
PST_DIR   = Path(r"E:\PST")
OUT_BASE  = Path(r"E:\TB_Mail_v2")
LOG_FILE  = Path(r"E:\claude\Pst2Thunder\auto_flow.log")
SCRIPT_DIR = Path(r"E:\claude\Pst2Thunder")

# Force UTF-8 stdout
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

# ── Logging ───────────────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-7s  %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
    handlers=[
        logging.FileHandler(LOG_FILE, encoding="utf-8"),
        logging.StreamHandler(sys.stdout),
    ]
)
log = logging.getLogger("auto_flow")

# ── Skip list ─────────────────────────────────────────────────────────────────
SKIP = {
    "spam search folder 2", "search root", "ipm_views", "ipm_common_views",
    "freebusy data", "reminders", "to-do search", "itemprocsearch",
    "tracked mail processing", "calendar", "contacts", "journal",
    "notes", "tasks", "outbox", "drafts", "junk e-mail",
    "sync issues", "rss feeds", "conversation action settings",
    "social activity feeds", "quick step settings", "suggested contacts",
    "top of outlook data file", "top of personal folders",
}

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


# ── Helpers ───────────────────────────────────────────────────────────────────

def sanitize(name):
    for c in r'<>:"/\|?*':
        name = name.replace(c, "_")
    return name.strip() or "Unknown"


def md5(path: Path, chunk=1 << 20) -> str:
    h = hashlib.md5()
    with open(path, "rb") as f:
        while True:
            data = f.read(chunk)
            if not data:
                break
            h.update(data)
    return h.hexdigest()


def parse_date(name: str):
    year_match = re.search(r'(?<!\d)((?:19|20)\d{2})(?!\d)', name)
    year = year_match.group(1) if year_match else None
    if year is None:
        two = re.search(r'(?<!\d)(\d{2})(?!\d)', name)
        if two:
            n = int(two.group(1))
            year = f"20{n:02d}" if n <= 30 else f"19{n:02d}"
    month = '00'
    for word in re.findall(r'[A-Za-z]+', name):
        m = MONTHS.get(word.lower())
        if m:
            month = m
            break
    if month == '00':
        name_lower = name.lower()
        for mon_name, mon_num in sorted(MONTHS.items(), key=lambda x: -len(x[0])):
            if mon_name in name_lower:
                month = mon_num
                break
    if year:
        return year, month
    return '9999', '99'


def make_prefix(name: str) -> str:
    year, month = parse_date(name)
    return f"{year}-{month}" if year != '9999' else 'zz'


def already_prefixed(name: str) -> bool:
    return bool(re.match(r'^(20|19)\d{2}-\d{2} |^zz ', name))


# ── MIME message builder ──────────────────────────────────────────────────────

def build_mime_message(message) -> email.message.Message:
    header_data = {}
    raw_headers = message.transport_headers or ""
    if isinstance(raw_headers, bytes):
        raw_headers = raw_headers.decode("utf-8", errors="replace")
    if raw_headers:
        try:
            hdr = email.message_from_string(raw_headers + "\r\n\r\n")
            for key in ("From", "To", "CC", "Subject", "Date",
                        "Message-ID", "Reply-To", "In-Reply-To", "References"):
                val = hdr.get(key)
                if val:
                    header_data[key] = val
        except Exception:
            pass

    plain = ""
    html  = ""
    try:
        plain = message.plain_text_body or ""
        if isinstance(plain, bytes):
            plain = plain.decode("utf-8", errors="replace")
    except Exception:
        pass
    try:
        html = message.html_body or ""
        if isinstance(html, bytes):
            html = html.decode("utf-8", errors="replace")
    except Exception:
        pass

    att_parts = []
    try:
        for att in message.attachments:
            try:
                att_name = (att.name or "attachment").strip()
                att_size = att.size or 0
                if att_size <= 0:
                    continue
                att_data = att.read_buffer(att_size)
                if not att_data:
                    continue
                mime_type, _ = mimetypes.guess_type(att_name)
                if mime_type and "/" in mime_type:
                    main_type, sub_type = mime_type.split("/", 1)
                else:
                    main_type, sub_type = "application", "octet-stream"
                part = email.mime.base.MIMEBase(main_type, sub_type)
                part.set_payload(att_data)
                encoders.encode_base64(part)
                part.add_header("Content-Disposition", "attachment", filename=att_name)
                att_parts.append(part)
            except Exception:
                pass
    except Exception:
        pass

    if plain and html:
        body_part = email.mime.multipart.MIMEMultipart("alternative")
        body_part.attach(email.mime.text.MIMEText(plain, "plain", "utf-8"))
        body_part.attach(email.mime.text.MIMEText(html,  "html",  "utf-8"))
    elif html:
        body_part = email.mime.text.MIMEText(html,  "html",  "utf-8")
    elif plain:
        body_part = email.mime.text.MIMEText(plain, "plain", "utf-8")
    else:
        body_part = email.mime.text.MIMEText("", "plain", "utf-8")

    if att_parts:
        msg = email.mime.multipart.MIMEMultipart("mixed")
        msg.attach(body_part)
        for part in att_parts:
            msg.attach(part)
    else:
        msg = body_part

    for key, val in header_data.items():
        if key not in msg:
            msg[key] = val
    if "Subject" not in msg:
        try:
            msg["Subject"] = message.subject or "(no subject)"
        except Exception:
            msg["Subject"] = "(no subject)"

    return msg


def write_messages(messages, mbox_path: Path) -> int:
    mbox_path.parent.mkdir(parents=True, exist_ok=True)
    for lock in mbox_path.parent.glob(mbox_path.name + ".lock*"):
        try:
            lock.unlink()
        except Exception:
            pass
    count = 0
    try:
        mbox = mailbox.mbox(str(mbox_path))
        try:
            mbox.lock()
        except Exception:
            pass
        try:
            for message in messages:
                try:
                    msg = build_mime_message(message)
                    mbox.add(msg)
                    count += 1
                except Exception:
                    pass
        finally:
            mbox.flush()
            mbox.unlock()
    except Exception as e:
        log.error(f"mbox error on {mbox_path.name}: {e}")
    return count


# ── Convert one PST ───────────────────────────────────────────────────────────

def convert_pst(pst_path: Path, dry_run: bool) -> int:
    stem      = pst_path.stem
    container = OUT_BASE / stem
    sbd_dir   = OUT_BASE / (stem + ".sbd")

    if container.exists() or sbd_dir.exists():
        log.warning(f"Output already exists for '{stem}' — skipping convert")
        return 0

    if dry_run:
        log.info(f"  [DRY-RUN] Would convert: {pst_path.name}")
        return 0

    container.touch()
    sbd_dir.mkdir(parents=True, exist_ok=True)
    total = 0

    try:
        with PffArchive(str(pst_path)) as archive:
            try:
                root = archive.archive.get_root_folder()
                is_ansi = (root is None)
            except Exception:
                is_ansi = False

            if is_ansi:
                log.info(f"  ANSI format → Inbox")
                try:
                    all_msgs = list(archive.messages())
                    count = write_messages(all_msgs, sbd_dir / "Inbox")
                    log.info(f"  Inbox (all): {count} messages")
                    total = count
                except Exception as e:
                    log.error(f"  ANSI failed: {e}")
            else:
                for folder in archive.folders():
                    if folder is None:
                        continue
                    try:
                        name = folder.name or ""
                    except Exception:
                        continue
                    if not name or name.lower() in SKIP:
                        continue
                    try:
                        msgs = list(folder.sub_messages)
                    except Exception:
                        msgs = []
                    if not msgs:
                        continue
                    count = write_messages(msgs, sbd_dir / sanitize(name))
                    if count:
                        log.info(f"  {name}: {count} messages")
                        total += count

    except Exception as e:
        log.error(f"  FAILED: {e}")
        try:
            container.unlink(missing_ok=True)
            shutil.rmtree(sbd_dir, ignore_errors=True)
        except Exception:
            pass
        return 0

    log.info(f"  Total: {total} messages → {sbd_dir}")
    return total


# ── Sort new folder ───────────────────────────────────────────────────────────

def sort_new_folder(stem: str, dry_run: bool):
    """Add YYYY-MM prefix to the newly created TB folder."""
    if already_prefixed(stem):
        return
    prefix   = make_prefix(stem)
    new_name = f"{prefix} {stem}"
    for suffix in ('', '.sbd', '.msf'):
        old_path = OUT_BASE / (stem + suffix)
        new_path = OUT_BASE / (new_name + suffix)
        if old_path.exists():
            if dry_run:
                log.info(f"  [DRY-RUN] Would rename: {old_path.name} -> {new_path.name}")
            else:
                try:
                    old_path.rename(new_path)
                except Exception as e:
                    log.error(f"  Rename failed {old_path.name}: {e}")
    if not dry_run:
        log.info(f"  Sorted as: {new_name}")


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    dry_run = "--dry-run" in sys.argv

    log.info("=" * 60)
    log.info(f"auto_flow started  {'[DRY-RUN]' if dry_run else ''}")
    log.info(f"Watching: {INCOMING}")

    # Find PSTs in Incoming (exclude Done subfolder)
    incoming_psts = [
        p for p in INCOMING.glob("*.pst")
        if p.is_file() and p.parent == INCOMING
    ]

    if not incoming_psts:
        log.info("No new PSTs in Incoming folder. Nothing to do.")
        log.info("=" * 60)
        return

    log.info(f"Found {len(incoming_psts)} PST(s) in Incoming")

    # Build existing index
    existing_psts  = list(PST_DIR.rglob("*.pst"))
    existing_stems = {p.stem.strip().lower(): p for p in existing_psts}
    existing_hashes = None  # lazy

    stats = {"new": 0, "dup_name": 0, "dup_hash": 0, "failed": 0}

    DONE_DIR.mkdir(parents=True, exist_ok=True)
    OUT_BASE.mkdir(exist_ok=True)

    for src in incoming_psts:
        log.info(f"── {src.name}  ({src.stat().st_size // 1024 // 1024} MB)")
        stem_lower = src.stem.strip().lower()

        # 1. Filename duplicate check
        if stem_lower in existing_stems:
            log.warning(f"  DUPLICATE (name) — matches {existing_stems[stem_lower].name} — skipping")
            stats["dup_name"] += 1
            if not dry_run:
                shutil.move(str(src), str(DONE_DIR / src.name))
                log.info(f"  Moved to Done/")
            continue

        # 2. Hash duplicate check (lazy build)
        if existing_hashes is None:
            log.info(f"  Computing hashes of {len(existing_psts)} existing PSTs...")
            existing_hashes = {}
            for ex in existing_psts:
                try:
                    existing_hashes[md5(ex)] = ex
                except Exception as e:
                    log.warning(f"  Cannot hash {ex.name}: {e}")

        try:
            src_hash = md5(src)
        except Exception as e:
            log.error(f"  Cannot read file: {e} — skipping")
            stats["failed"] += 1
            continue

        if src_hash in existing_hashes:
            match = existing_hashes[src_hash]
            log.warning(f"  DUPLICATE (hash) — same content as {match.name} — skipping")
            stats["dup_hash"] += 1
            if not dry_run:
                shutil.move(str(src), str(DONE_DIR / src.name))
                log.info(f"  Moved to Done/")
            continue

        # 3. Genuinely new — copy to PST_DIR
        dest = PST_DIR / src.name
        log.info(f"  NEW — copying to {dest}")
        if not dry_run:
            try:
                shutil.copy2(src, dest)
            except Exception as e:
                log.error(f"  Copy failed: {e}")
                stats["failed"] += 1
                continue

        # 4. Convert
        log.info(f"  Converting...")
        total = convert_pst(dest if not dry_run else src, dry_run)

        # 5. Sort (add YYYY-MM prefix)
        sort_new_folder(src.stem, dry_run)

        # 6. Move to Done
        if not dry_run:
            shutil.move(str(src), str(DONE_DIR / src.name))
            log.info(f"  Moved to Done/")
            # Update index for subsequent files in same run
            existing_stems[stem_lower] = dest
            existing_hashes[src_hash]  = dest

        stats["new"] += 1
        log.info(f"  Done: {total} messages")

    log.info("── Summary ──────────────────────────────────────────")
    log.info(f"  New converted  : {stats['new']}")
    log.info(f"  Duplicate name : {stats['dup_name']}")
    log.info(f"  Duplicate hash : {stats['dup_hash']}")
    log.info(f"  Failed         : {stats['failed']}")
    log.info(f"  Restart Thunderbird to see new folders." if stats['new'] else "  No new mail added.")
    log.info("=" * 60)


if __name__ == "__main__":
    main()

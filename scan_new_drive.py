"""
scan_new_drive.py — Scan a drive/folder for PST files, find genuinely new ones,
and convert them to Thunderbird format.

Usage:
    python scan_new_drive.py G:\pst
    python scan_new_drive.py G:\pst --dry-run     (report only, no copy/convert)

Steps:
    1. Find all .pst in source folder
    2. For each: filename check vs E:\PST\ (case-insensitive)
    3. For non-matching names: MD5 hash check (catches renames)
    4. Report: DUPLICATE / RENAMED-DUPE / NEW
    5. Copy NEW ones to E:\PST\ and convert to E:\TB_Mail_v2\
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
from email import encoders
import re
from pathlib import Path
from libratom.lib.pff import PffArchive

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

def _make_sort_prefix(name: str) -> str:
    year_m = re.search(r'(?<!\d)((?:19|20)\d{2})(?!\d)', name)
    year = year_m.group(1) if year_m else None
    if not year:
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
        for mon, num in sorted(MONTHS.items(), key=lambda x: -len(x[0])):
            if mon in name.lower():
                month = num
                break
    return f"{year}-{month}" if year else 'zz'

def _sort_folder(stem: str):
    """Add YYYY-MM prefix to a TB folder if not already prefixed."""
    if re.match(r'^(20|19)\d{2}-\d{2} |^zz ', stem):
        return
    prefix   = _make_sort_prefix(stem)
    new_name = f"{prefix} {stem}"
    for suffix in ('', '.sbd', '.msf'):
        old = OUT_BASE / (stem + suffix)
        new = OUT_BASE / (new_name + suffix)
        if old.exists():
            try:
                old.rename(new)
            except Exception as e:
                print(f"  Sort rename failed {old.name}: {e}", flush=True)
    print(f"  Sorted as: {new_name}", flush=True)

# Force UTF-8 stdout
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

PST_DIR  = Path(r"E:\PST")
OUT_BASE = Path(r"E:\TB_Mail_v2")

SKIP = {
    "spam search folder 2", "search root", "ipm_views", "ipm_common_views",
    "freebusy data", "reminders", "to-do search", "itemprocsearch",
    "tracked mail processing", "calendar", "contacts", "journal",
    "notes", "tasks", "outbox", "drafts", "junk e-mail",
    "sync issues", "rss feeds", "conversation action settings",
    "social activity feeds", "quick step settings", "suggested contacts",
    "top of outlook data file", "top of personal folders",
}


# ── Helpers ──────────────────────────────────────────────────────────────────

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
        if isinstance(plain, bytes): plain = plain.decode("utf-8", errors="replace")
    except Exception: pass
    try:
        html = message.html_body or ""
        if isinstance(html, bytes): html = html.decode("utf-8", errors="replace")
    except Exception: pass
    att_parts = []
    try:
        for att in message.attachments:
            try:
                att_name = (att.name or "attachment").strip()
                att_size = att.size or 0
                if att_size <= 0: continue
                att_data = att.read_buffer(att_size)
                if not att_data: continue
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
            except Exception: pass
    except Exception: pass
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
        for part in att_parts: msg.attach(part)
    else:
        msg = body_part
    for key, val in header_data.items():
        if key not in msg: msg[key] = val
    if "Subject" not in msg:
        try: msg["Subject"] = message.subject or "(no subject)"
        except Exception: msg["Subject"] = "(no subject)"
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
        print(f"    mbox error: {e}", flush=True)
    return count


def convert_pst(pst_path: Path):
    stem      = pst_path.stem
    container = OUT_BASE / stem
    sbd_dir   = OUT_BASE / (stem + ".sbd")

    if container.exists() or sbd_dir.exists():
        print(f"  [SKIP] Output already exists: {stem}", flush=True)
        return

    container.touch()
    sbd_dir.mkdir(parents=True, exist_ok=True)
    total = 0

    try:
        with PffArchive(str(pst_path)) as archive:
            # Detect ANSI (old) format
            try:
                root = archive.archive.get_root_folder()
                is_ansi = (root is None)
            except Exception:
                is_ansi = False

            if is_ansi:
                print(f"  Detected: ANSI/old format — all messages → Inbox", flush=True)
                try:
                    all_msgs = list(archive.messages())
                    count = write_messages(all_msgs, sbd_dir / "Inbox")
                    print(f"  Inbox (all): {count} messages", flush=True)
                    total = count
                except Exception as e:
                    print(f"  FAILED (ANSI): {e}", flush=True)
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
                        print(f"  {name}: {count} messages", flush=True)
                        total += count

    except Exception as e:
        print(f"  FAILED: {e}", flush=True)
        try:
            container.unlink(missing_ok=True)
            shutil.rmtree(sbd_dir, ignore_errors=True)
        except Exception:
            pass
        return

    print(f"  → {total} messages  →  {sbd_dir}", flush=True)


# ── Main ─────────────────────────────────────────────────────────────────────

def main():
    dry_run = "--dry-run" in sys.argv
    args = [a for a in sys.argv[1:] if not a.startswith("--")]

    if not args:
        print("Usage: python scan_new_drive.py <source_folder> [--dry-run]", flush=True)
        sys.exit(1)

    src = Path(args[0])
    if not src.exists():
        print(f"ERROR: Source folder not found: {src}", flush=True)
        sys.exit(1)

    print(f"Source : {src}", flush=True)
    print(f"Archive: {PST_DIR}", flush=True)
    print(f"Output : {OUT_BASE}", flush=True)
    if dry_run:
        print(f"Mode   : DRY RUN (no changes)", flush=True)
    print()

    # Find all PSTs in source
    src_psts = sorted(src.rglob("*.pst"))
    print(f"Found {len(src_psts)} .pst file(s) in {src}", flush=True)
    print()

    if not src_psts:
        print("Nothing to do.", flush=True)
        return

    # Build lookup index of existing PSTs
    existing_psts = list(PST_DIR.rglob("*.pst"))
    existing_stems = {p.stem.strip().lower(): p for p in existing_psts}

    results = {"duplicate_name": [], "duplicate_hash": [], "new": [], "error": []}

    # Phase 1: classify each source PST
    print("── Scanning ─────────────────────────────────────────", flush=True)

    # Pre-compute hashes of existing PSTs once (only for stem-mismatches)
    existing_hashes = None  # lazy — computed only if needed

    new_psts = []

    for src_pst in src_psts:
        stem_lower = src_pst.stem.strip().lower()
        size_mb = src_pst.stat().st_size // 1024 // 1024

        if stem_lower in existing_stems:
            match = existing_stems[stem_lower]
            print(f"  [DUP-NAME ] {src_pst.name}  ({size_mb} MB)  →  matches {match.name}", flush=True)
            results["duplicate_name"].append(src_pst)
        else:
            # Name is new — do hash check
            if existing_hashes is None:
                print(f"  (Computing hashes of {len(existing_psts)} existing PSTs...)", flush=True)
                existing_hashes = {}
                for ex in existing_psts:
                    try:
                        existing_hashes[md5(ex)] = ex
                    except Exception as e:
                        print(f"  WARNING: Cannot hash {ex.name}: {e}", flush=True)

            try:
                src_hash = md5(src_pst)
            except Exception as e:
                print(f"  [ERROR    ] {src_pst.name}: cannot read — {e}", flush=True)
                results["error"].append(src_pst)
                continue

            if src_hash in existing_hashes:
                match = existing_hashes[src_hash]
                print(f"  [DUP-HASH ] {src_pst.name}  ({size_mb} MB)  →  same content as {match.name}", flush=True)
                results["duplicate_hash"].append(src_pst)
            else:
                print(f"  [NEW      ] {src_pst.name}  ({size_mb} MB)  MD5={src_hash[:8]}...", flush=True)
                results["new"].append(src_pst)
                new_psts.append(src_pst)

    print()
    print("── Summary ──────────────────────────────────────────", flush=True)
    print(f"  Duplicate (name) : {len(results['duplicate_name'])}", flush=True)
    print(f"  Duplicate (hash) : {len(results['duplicate_hash'])}", flush=True)
    print(f"  Read errors      : {len(results['error'])}", flush=True)
    print(f"  NEW              : {len(results['new'])}", flush=True)
    print()

    if not new_psts:
        print("No new PSTs found. Nothing to convert.", flush=True)
        return

    if dry_run:
        print("DRY RUN — skipping copy and convert.", flush=True)
        print("Re-run without --dry-run to process new files.", flush=True)
        return

    # Phase 2: copy + convert new PSTs
    print("── Copy & Convert ───────────────────────────────────", flush=True)
    OUT_BASE.mkdir(exist_ok=True)
    converted = 0
    failed = 0

    for src_pst in new_psts:
        dest = PST_DIR / src_pst.name
        print(f"\nProcessing: {src_pst.name}", flush=True)

        # Copy
        try:
            print(f"  Copying to {dest} ...", flush=True)
            shutil.copy2(src_pst, dest)
        except Exception as e:
            print(f"  COPY FAILED: {e}", flush=True)
            failed += 1
            continue

        # Convert
        convert_pst(dest)

        # Sort — add YYYY-MM prefix for chronological order in Thunderbird
        _sort_folder(src_pst.stem)

        converted += 1

    print()
    print(f"── Done ─────────────────────────────────────────────", flush=True)
    print(f"  Converted : {converted}", flush=True)
    print(f"  Failed    : {failed}", flush=True)
    if converted:
        print(f"  Restart Thunderbird (or right-click folder → Repair) to see new mail.", flush=True)


if __name__ == "__main__":
    main()

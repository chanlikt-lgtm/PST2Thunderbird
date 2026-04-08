"""
add_pst.py — Add a new PST file, check for duplicates, and convert to Thunderbird format.

Usage:
    python add_pst.py "path\to\new_file.pst"

Flow:
    1. Normalize + filename duplicate check against E:\PST\
    2. MD5 hash duplicate check (catches renames)
    3. Copy to E:\PST\
    4. Detect PST format (Unicode vs ANSI) and convert to E:\TB_Mail_v2\
"""

import sys
import io
import hashlib
import shutil
import mailbox
import email
from pathlib import Path
from libratom.lib.pff import PffArchive

# Force UTF-8 stdout (handles Chinese/non-ASCII folder names)
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
                    raw = message.transport_headers or ""
                    if isinstance(raw, bytes):
                        raw = raw.decode("utf-8", errors="replace")
                    body = ""
                    try:
                        body = message.plain_text_body or ""
                        if isinstance(body, bytes):
                            body = body.decode("utf-8", errors="replace")
                    except Exception:
                        pass
                    full = raw + "\r\n" + body if raw else body
                    try:
                        msg = email.message_from_string(full)
                    except Exception:
                        msg = email.message.Message()
                        msg["Subject"] = getattr(message, "subject", "") or "(no subject)"
                        msg.set_payload(body)
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


# ── Step 1 & 2: Duplicate checks ─────────────────────────────────────────────

def check_duplicates(new_pst: Path) -> bool:
    """Returns True if duplicate found (abort). False if safe to proceed."""
    existing_psts = list(PST_DIR.rglob("*.pst"))

    # 1. Filename check (case-insensitive, stem only)
    new_stem = new_pst.stem.strip().lower()
    for ex in existing_psts:
        if ex.stem.strip().lower() == new_stem:
            print(f"DUPLICATE (filename): '{new_pst.name}' matches '{ex}'", flush=True)
            print("Aborting — no changes made.", flush=True)
            return True

    # 2. Hash check (detects renames of the same file)
    print("Computing MD5 of new file...", flush=True)
    new_hash = md5(new_pst)
    print(f"  MD5: {new_hash}", flush=True)
    print(f"  Checking against {len(existing_psts)} existing PST(s)...", flush=True)
    for ex in existing_psts:
        try:
            ex_hash = md5(ex)
            if ex_hash == new_hash:
                print(f"DUPLICATE (content): '{new_pst.name}' is identical to '{ex}'", flush=True)
                print("Aborting — no changes made.", flush=True)
                return True
        except Exception as e:
            print(f"  Could not hash {ex.name}: {e}", flush=True)

    print("  No duplicates found. Proceeding.", flush=True)
    return False


# ── Step 3: Copy to PST_DIR ───────────────────────────────────────────────────

def copy_to_pst_dir(new_pst: Path) -> Path:
    dest = PST_DIR / new_pst.name
    if dest.exists():
        # Shouldn't happen after duplicate check, but guard anyway
        print(f"WARNING: {dest} already exists — skipping copy.", flush=True)
        return dest
    print(f"Copying to {dest} ...", flush=True)
    shutil.copy2(new_pst, dest)
    size_mb = dest.stat().st_size // 1024 // 1024
    print(f"  Copied: {size_mb} MB", flush=True)
    return dest


# ── Step 4 & 5: Convert ───────────────────────────────────────────────────────

def convert_pst(pst_path: Path):
    stem      = pst_path.stem
    container = OUT_BASE / stem
    sbd_dir   = OUT_BASE / (stem + ".sbd")

    if container.exists() or sbd_dir.exists():
        print(f"WARNING: Output already exists for '{stem}' — skipping conversion.", flush=True)
        return

    container.touch()
    sbd_dir.mkdir(parents=True, exist_ok=True)
    total = 0

    print(f"Converting: {pst_path.name}", flush=True)

    try:
        with PffArchive(str(pst_path)) as archive:

            # Detect ANSI (old) format: root_folder is None
            try:
                root = archive.archive.get_root_folder()
                is_ansi = (root is None)
            except Exception:
                is_ansi = False

            if is_ansi:
                # ── ANSI format: flat dump all messages into Inbox ──
                print("  Detected: ANSI/old format — all messages → Inbox", flush=True)
                try:
                    all_msgs = list(archive.messages())
                    mbox_path = sbd_dir / "Inbox"
                    count = write_messages(all_msgs, mbox_path)
                    print(f"  Inbox (all): {count} messages", flush=True)
                    total = count
                except Exception as e:
                    print(f"  FAILED (ANSI): {e}", flush=True)
            else:
                # ── Unicode format: per-folder conversion ──
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
                    mbox_path = sbd_dir / sanitize(name)
                    count = write_messages(msgs, mbox_path)
                    if count:
                        print(f"  {name}: {count} messages", flush=True)
                        total += count

    except Exception as e:
        print(f"  FAILED: {e}", flush=True)
        # Clean up empty output on failure
        try:
            container.unlink(missing_ok=True)
            shutil.rmtree(sbd_dir, ignore_errors=True)
        except Exception:
            pass
        return

    print(f"  → Total: {total} messages", flush=True)
    print(f"  → Output: {sbd_dir}", flush=True)


# ── Main ─────────────────────────────────────────────────────────────────────

def main():
    if len(sys.argv) < 2:
        print("Usage: python add_pst.py <path_to_new.pst>", flush=True)
        sys.exit(1)

    new_pst = Path(sys.argv[1])

    if not new_pst.exists():
        print(f"ERROR: File not found: {new_pst}", flush=True)
        sys.exit(1)

    if new_pst.suffix.lower() != ".pst":
        print(f"ERROR: Not a .pst file: {new_pst}", flush=True)
        sys.exit(1)

    print(f"New PST: {new_pst}", flush=True)
    print(f"Size:    {new_pst.stat().st_size // 1024 // 1024} MB", flush=True)
    print()

    # Step 1+2: Duplicate check
    print("── Duplicate Check ──────────────────────────────", flush=True)
    if check_duplicates(new_pst):
        sys.exit(0)
    print()

    # Step 3: Copy to E:\PST\
    print("── Copy to PST archive ──────────────────────────", flush=True)
    dest_pst = copy_to_pst_dir(new_pst)
    print()

    # Step 4+5: Convert
    print("── Convert to Thunderbird format ────────────────", flush=True)
    OUT_BASE.mkdir(exist_ok=True)
    convert_pst(dest_pst)
    print()

    print("Done. Restart Thunderbird (or right-click folder → Repair) to see new mail.", flush=True)


if __name__ == "__main__":
    main()

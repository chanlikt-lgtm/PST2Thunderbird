"""
Convert PST files in E:\PST\ to Thunderbird folder structure in E:\TB_Mail_v2\
Preserves Inbox / Sent Items / Deleted Items as separate MBOX files.

Output structure:
  E:\TB_Mail_v2\
    April 2014          <- empty container file
    April 2014.sbd\
      Inbox             <- MBOX file
      Sent Items        <- MBOX file
      Deleted Items     <- MBOX file
"""
import sys
import mailbox
import email
from pathlib import Path
from libratom.lib.pff import PffArchive

PST_DIR  = Path(r"E:\PST")
OUT_BASE = Path(r"E:\TB_Mail_v2")
OUT_BASE.mkdir(exist_ok=True)

# System/non-mail folders to skip
SKIP = {
    "spam search folder 2", "search root", "ipm_views", "ipm_common_views",
    "freebusy data", "reminders", "to-do search", "itemprocsearch",
    "tracked mail processing", "calendar", "contacts", "journal",
    "notes", "tasks", "outbox", "drafts", "junk e-mail",
    "sync issues", "rss feeds", "conversation action settings",
    "social activity feeds", "quick step settings", "suggested contacts",
    "top of outlook data file", "top of personal folders",
}

def sanitize(name):
    for c in r'<>:"/\|?*':
        name = name.replace(c, "_")
    return name.strip() or "Unknown"

def write_messages(messages, mbox_path):
    mbox_path.parent.mkdir(parents=True, exist_ok=True)
    count = 0
    # Remove stale lock files
    for lock in mbox_path.parent.glob(mbox_path.name + ".lock*"):
        try:
            lock.unlink()
        except Exception:
            pass
    try:
        mbox = mailbox.mbox(str(mbox_path))
        try:
            mbox.lock()
        except Exception:
            pass  # continue without lock on Windows if needed
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
        print(f"    Error: {e}", flush=True)
    return count


pst_files = sorted(PST_DIR.glob("*.pst"))
to_convert = [p for p in pst_files
              if not (OUT_BASE / p.stem).exists()
              and not (OUT_BASE / (p.stem + ".sbd")).exists()]

print(f"PST files: {len(pst_files)}  |  Already done: {len(pst_files)-len(to_convert)}  |  To convert: {len(to_convert)}")
print()

converted = 0
failed = 0

for pst_path in to_convert:
    print(f"Converting: {pst_path.name}", flush=True)
    container = OUT_BASE / pst_path.stem
    sbd_dir   = OUT_BASE / (pst_path.stem + ".sbd")

    try:
        with PffArchive(str(pst_path)) as archive:
            container.touch()
            sbd_dir.mkdir(exist_ok=True)
            pst_total = 0

            for folder in archive.folders():
                if folder is None:
                    continue
                try:
                    name = folder.name or ""
                except Exception:
                    continue

                if not name or name.lower() in SKIP:
                    continue

                # Check for messages
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
                    pst_total += count

        print(f"  -> Total: {pst_total} messages\n", flush=True)
        converted += 1
    except Exception as e:
        print(f"  FAILED: {e}\n", flush=True)
        failed += 1

print(f"Done: {converted} converted, {failed} failed")
print(f"Output: {OUT_BASE}")
sys.stdout.flush()

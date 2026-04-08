"""
Recovery script for PSTs that failed in pst_to_tb_v2.py:
- 2001.pst, 2003 PART1  1Si cltan.pst, INFINEON  Sep 2007.pst:
  Older PST format — use archive.messages() directly, put all in Inbox
- Aug 2013.pst:
  Charmap error on folder name — fix with UTF-8 stdout and skip bad names
"""
import sys
import io
import mailbox
import email
from pathlib import Path
from libratom.lib.pff import PffArchive

# Force UTF-8 stdout to avoid charmap errors on Chinese folder names
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

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

def sanitize(name):
    for c in r'<>:"/\|?*':
        name = name.replace(c, "_")
    return name.strip() or "Unknown"

def write_messages(messages, mbox_path):
    mbox_path.parent.mkdir(parents=True, exist_ok=True)
    for lock in mbox_path.parent.glob(mbox_path.name + ".lock*"):
        try: lock.unlink()
        except: pass
    count = 0
    try:
        mbox = mailbox.mbox(str(mbox_path))
        try: mbox.lock()
        except: pass
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
                    except: pass
                    full = raw + "\r\n" + body if raw else body
                    try:
                        msg = email.message_from_string(full)
                    except:
                        msg = email.message.Message()
                        msg.set_payload(body)
                    mbox.add(msg)
                    count += 1
                except: pass
        finally:
            mbox.flush()
            mbox.unlock()
    except Exception as e:
        print(f"    mbox error: {e}", flush=True)
    return count

# --- Aug 2013: re-run with UTF-8, using folder-by-folder approach ---
pst_path = Path(r"E:\PST\Aug 2013.pst")
print(f"Fixing: {pst_path.name}", flush=True)
container = OUT_BASE / pst_path.stem
sbd_dir   = OUT_BASE / (pst_path.stem + ".sbd")
container.touch()
sbd_dir.mkdir(exist_ok=True)
total = 0
try:
    with PffArchive(str(pst_path)) as archive:
        for folder in archive.folders():
            if folder is None: continue
            try: name = folder.name or ""
            except: continue
            if not name or name.lower() in SKIP: continue
            try: msgs = list(folder.sub_messages)
            except: msgs = []
            if not msgs: continue
            mbox_path = sbd_dir / sanitize(name)
            count = write_messages(msgs, mbox_path)
            print(f"  {name}: {count} messages", flush=True)
            total += count
    print(f"  -> Total: {total}\n", flush=True)
except Exception as e:
    print(f"  FAILED: {e}\n", flush=True)

# --- Older format PSTs: fall back to archive.messages() -> all in Inbox ---
OLD_PSTS = [
    r"E:\PST\2001.pst",
    r"E:\PST\2003 PART1  1Si cltan.pst",
    r"E:\PST\INFINEON  Sep 2007.pst",
]

for pst_str in OLD_PSTS:
    pst_path = Path(pst_str)
    print(f"Fixing (old format): {pst_path.name}", flush=True)
    container = OUT_BASE / pst_path.stem
    sbd_dir   = OUT_BASE / (pst_path.stem + ".sbd")
    container.touch()
    sbd_dir.mkdir(exist_ok=True)
    try:
        with PffArchive(str(pst_path)) as archive:
            # Use messages() which does flat iteration without folder tree
            all_msgs = list(archive.messages())
            mbox_path = sbd_dir / "Inbox"
            count = write_messages(all_msgs, mbox_path)
            print(f"  Inbox (all messages): {count}", flush=True)
            print(f"  -> Total: {count}\n", flush=True)
    except Exception as e:
        print(f"  FAILED: {e}\n", flush=True)

print("Recovery done.", flush=True)

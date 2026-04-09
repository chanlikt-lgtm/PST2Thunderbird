"""
reconvert.py — Reconvert specific PSTs with full MIME support:
  - Plain text + HTML body (whichever is available)
  - Attachments preserved (ppt, xlsx, dat, pdf, etc.)

Usage:
    python reconvert.py "Dec 2021" "Nov 2021"
    python reconvert.py --all        (reconvert everything in E:\PST\)
"""

import sys
import io
import shutil
import mailbox
import mimetypes
import email
import email.mime.multipart
import email.mime.text
import email.mime.base
from email import encoders
from pathlib import Path
from libratom.lib.pff import PffArchive

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


def sanitize(name):
    for c in r'<>:"/\|?*':
        name = name.replace(c, "_")
    return name.strip() or "Unknown"


def build_mime_message(message) -> email.message.Message:
    """Build a proper MIME message with body (plain/html) and attachments."""

    # ── Extract key headers ──────────────────────────────────────────────────
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

    # ── Extract bodies ───────────────────────────────────────────────────────
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

    # ── Extract attachments ──────────────────────────────────────────────────
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
                part.add_header("Content-Disposition", "attachment",
                                filename=att_name)
                att_parts.append(part)
            except Exception:
                pass
    except Exception:
        pass

    # ── Build body part ──────────────────────────────────────────────────────
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

    # ── Wrap in multipart/mixed if there are attachments ────────────────────
    if att_parts:
        msg = email.mime.multipart.MIMEMultipart("mixed")
        msg.attach(body_part)
        for part in att_parts:
            msg.attach(part)
    else:
        msg = body_part

    # ── Copy headers ─────────────────────────────────────────────────────────
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
        print(f"    mbox error: {e}", flush=True)
    return count


def reconvert(pst_path: Path):
    stem      = pst_path.stem
    container = OUT_BASE / stem
    sbd_dir   = OUT_BASE / (stem + ".sbd")

    # Remove old output
    print(f"  Removing old output...", flush=True)
    try:
        container.unlink(missing_ok=True)
    except Exception:
        pass
    if sbd_dir.exists():
        shutil.rmtree(sbd_dir, ignore_errors=True)

    container.touch()
    sbd_dir.mkdir(parents=True, exist_ok=True)
    total = 0

    try:
        with PffArchive(str(pst_path)) as archive:
            # Detect ANSI format
            try:
                root = archive.archive.get_root_folder()
                is_ansi = (root is None)
            except Exception:
                is_ansi = False

            if is_ansi:
                print(f"  ANSI format → Inbox", flush=True)
                try:
                    all_msgs = list(archive.messages())
                    count = write_messages(all_msgs, sbd_dir / "Inbox")
                    print(f"  Inbox (all): {count} messages", flush=True)
                    total = count
                except Exception as e:
                    print(f"  FAILED: {e}", flush=True)
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
        return 0

    print(f"  → {total} messages", flush=True)
    verify_size(pst_path, sbd_dir)
    return total


MIN_RATIO = 0.05   # output must be >= 5% of PST size; below this = suspect


def verify_size(pst_path: Path, sbd_dir: Path):
    """Warn if output mbox total is suspiciously small vs PST input."""
    try:
        pst_bytes = pst_path.stat().st_size
        out_bytes = sum(f.stat().st_size for f in sbd_dir.rglob("*")
                        if f.is_file() and f.suffix not in (".msf", ".lock"))
        if pst_bytes == 0:
            return
        ratio = out_bytes / pst_bytes
        pst_mb  = pst_bytes  // (1024 * 1024)
        out_mb  = out_bytes  // (1024 * 1024)
        if ratio < MIN_RATIO:
            print(f"  !! SIZE WARNING: PST={pst_mb} MB  output={out_mb} MB  "
                  f"({ratio*100:.1f}%) — likely incomplete conversion!", flush=True)
        else:
            print(f"  Size OK: PST={pst_mb} MB  output={out_mb} MB  "
                  f"({ratio*100:.0f}%)", flush=True)
    except Exception as e:
        print(f"  Size check failed: {e}", flush=True)


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    do_all = "--all" in sys.argv

    if do_all:
        targets = [p for p in sorted(PST_DIR.glob("*.pst"))]
    elif args:
        targets = []
        for stem in args:
            p = PST_DIR / (stem if stem.endswith(".pst") else stem + ".pst")
            if not p.exists():
                print(f"WARNING: {p} not found — skipping", flush=True)
            else:
                targets.append(p)
    else:
        print("Usage:", flush=True)
        print("  python reconvert.py \"Dec 2021\" \"Nov 2021\"", flush=True)
        print("  python reconvert.py --all", flush=True)
        sys.exit(1)

    print(f"Reconverting {len(targets)} PST(s) with full MIME + attachment support", flush=True)
    print(f"Output: {OUT_BASE}", flush=True)
    print()

    OUT_BASE.mkdir(exist_ok=True)
    total_msgs = 0
    for pst_path in targets:
        print(f"── {pst_path.name}", flush=True)
        total_msgs += reconvert(pst_path)
        print()

    print(f"Done. Total messages: {total_msgs}", flush=True)
    print("Restart Thunderbird to see updated content.", flush=True)


if __name__ == "__main__":
    main()

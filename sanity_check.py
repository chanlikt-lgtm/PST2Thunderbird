"""
sanity_check.py — Sanity check all converted Thunderbird folders:
  1. Emails with no body (title only)
  2. Attachments that are .dat (should be office format)
  3. Summary per folder

Usage:
    py -3.11 sanity_check.py
    py -3.11 sanity_check.py --fix-dat    (also attempt to decode winmail.dat via tnefparse)
"""

import sys
import io
import mailbox
import email
from pathlib import Path
from collections import defaultdict

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

OUT_BASE = Path(r"E:\TB_Mail_v2")
LOG_FILE = Path(r"E:\claude\Pst2Thunder\sanity_check.log")

# Known good office/document extensions
OFFICE_EXTS = {
    '.doc', '.docx', '.xls', '.xlsx', '.ppt', '.pptx',
    '.pdf', '.txt', '.csv', '.zip', '.rar', '.7z',
    '.jpg', '.jpeg', '.png', '.gif', '.bmp', '.tif', '.tiff',
    '.msg', '.eml', '.htm', '.html', '.xml',
    '.mp3', '.mp4', '.avi', '.mov',
    '.vsd', '.vsdx', '.one', '.pub',
}

def check_encrypted(msg) -> str:
    """Returns encryption type or empty string."""
    ct = msg.get_content_type() or ""
    if "pkcs7-mime" in ct or "pkcs7mime" in ct:
        return "S/MIME-ENCRYPTED"
    if "pkcs7-signature" in ct:
        return "S/MIME-SIGNED"
    if msg.is_multipart():
        for part in msg.walk():
            pct = part.get_content_type() or ""
            fname = (part.get_filename() or "").lower()
            if "pkcs7-mime" in pct:
                return "S/MIME-ENCRYPTED"
            if "pkcs7-signature" in pct or fname.endswith(".p7s"):
                return "S/MIME-SIGNED"
            if "pgp-encrypted" in pct or fname.endswith(".pgp") or fname.endswith(".gpg"):
                return "PGP-ENCRYPTED"
            if fname == "smime.p7m":
                return "S/MIME-ENCRYPTED"
    return ""


def check_body(msg) -> bool:
    """Returns True if message has any body content."""
    if msg.is_multipart():
        for part in msg.walk():
            ct = part.get_content_type()
            cd = part.get("Content-Disposition", "")
            if "attachment" in cd:
                continue
            if ct in ("text/plain", "text/html"):
                payload = part.get_payload(decode=True)
                if payload and payload.strip():
                    return True
    else:
        payload = msg.get_payload(decode=True)
        if payload and payload.strip():
            return True
        # Also check string payload
        payload_str = msg.get_payload()
        if isinstance(payload_str, str) and payload_str.strip():
            return True
    return False


def get_attachments(msg):
    """Returns list of (filename, size) for all attachments."""
    atts = []
    if msg.is_multipart():
        for part in msg.walk():
            cd = part.get("Content-Disposition", "")
            if "attachment" in cd:
                fname = part.get_filename() or ""
                payload = part.get_payload(decode=True) or b""
                atts.append((fname, len(payload)))
    return atts


def main():
    fix_dat = "--fix-dat" in sys.argv

    results = []  # (folder_stem, no_body, total, dat_atts, total_atts)

    folders = sorted(OUT_BASE.glob("*.sbd"))
    print(f"Checking {len(folders)} folders...\n", flush=True)

    for sbd in folders:
        stem = sbd.stem
        mbox_files = [f for f in sbd.rglob("*")
                      if f.is_file() and f.suffix != ".msf"]

        total_msgs   = 0
        no_body      = 0
        total_atts   = 0
        dat_atts     = 0
        encrypted    = 0
        signed       = 0
        bad_att_examples = []

        for mf in mbox_files:
            try:
                mb = mailbox.mbox(str(mf))
                for msg in mb:
                    total_msgs += 1
                    enc = check_encrypted(msg)
                    if "ENCRYPTED" in enc:
                        encrypted += 1
                    elif "SIGNED" in enc:
                        signed += 1
                    if not enc and not check_body(msg):
                        no_body += 1
                    for fname, size in get_attachments(msg):
                        total_atts += 1
                        ext = Path(fname).suffix.lower() if fname else ""
                        if ext == ".dat" or (fname and fname.lower() == "winmail.dat"):
                            dat_atts += 1
                            if len(bad_att_examples) < 3:
                                bad_att_examples.append(fname)
            except Exception as e:
                pass

        if total_msgs == 0:
            continue

        no_body_pct = no_body * 100 // total_msgs
        has_issue = no_body_pct > 10 or dat_atts > 0 or encrypted > 0

        results.append({
            "stem": stem,
            "total": total_msgs,
            "no_body": no_body,
            "no_body_pct": no_body_pct,
            "total_atts": total_atts,
            "dat_atts": dat_atts,
            "encrypted": encrypted,
            "signed": signed,
            "examples": bad_att_examples,
            "has_issue": has_issue,
        })

        flags = []
        if no_body_pct > 10: flags.append("NO-BODY")
        if dat_atts > 0:     flags.append("DAT-ATT")
        if encrypted > 0:    flags.append("ENCRYPTED")
        status = "+".join(flags) if flags else "OK"

        flag = "  " if not has_issue else "!!"
        print(f"{flag} [{status:<18}] {stem}", flush=True)
        if no_body > 0:
            print(f"           No body  : {no_body}/{total_msgs} ({no_body_pct}%)", flush=True)
        if encrypted > 0:
            print(f"           Encrypted: {encrypted} msgs (cannot read without key)", flush=True)
        if signed > 0:
            print(f"           Signed   : {signed} msgs (readable, has signature)", flush=True)
        if dat_atts > 0:
            print(f"           .dat atts: {dat_atts}/{total_atts}  e.g. {bad_att_examples}", flush=True)

    # Summary
    total_issues   = sum(1 for r in results if r["has_issue"])
    total_no_body  = sum(r["no_body"] for r in results)
    total_dat      = sum(r["dat_atts"] for r in results)
    total_encrypted= sum(r["encrypted"] for r in results)
    total_signed   = sum(r["signed"] for r in results)
    total_msgs_all = sum(r["total"] for r in results)
    total_atts_all = sum(r["total_atts"] for r in results)

    summary = f"""
══════════════════════════════════════════════
SANITY CHECK SUMMARY
══════════════════════════════════════════════
Folders checked    : {len(results)}
Folders with issues: {total_issues}
Total messages     : {total_msgs_all:,}
No-body emails     : {total_no_body:,}  ({total_no_body*100//max(total_msgs_all,1)}%)
Encrypted emails   : {total_encrypted:,}  (unreadable without private key)
Signed emails      : {total_signed:,}  (readable, S/MIME signature)
Total attachments  : {total_atts_all:,}
.dat attachments   : {total_dat:,}
══════════════════════════════════════════════
"""
    print(summary, flush=True)

    # Save log
    with open(LOG_FILE, "w", encoding="utf-8") as f:
        f.write(summary)
        for r in results:
            if r["has_issue"]:
                f.write(f"\n!! {r['stem']}\n")
                f.write(f"   No body  : {r['no_body']}/{r['total']} ({r['no_body_pct']}%)\n")
                f.write(f"   Encrypted: {r['encrypted']} msgs\n")
                f.write(f"   Signed   : {r['signed']} msgs\n")
                f.write(f"   .dat atts: {r['dat_atts']}/{r['total_atts']}\n")
                if r["examples"]:
                    f.write(f"   Examples : {r['examples']}\n")

    print(f"Log saved: {LOG_FILE}", flush=True)

    if total_dat > 0:
        print("\nNOTE: .dat attachments are Outlook TNEF/winmail.dat files.", flush=True)
        print("Run: py -3.11 fix_dat.py  to decode them into proper office files.", flush=True)


if __name__ == "__main__":
    main()

"""
make_v3_pdf.py — Generate PDF documenting V3 reconvert fixes and housekeeping.
Rev 2: From/To header fix + archive._data folder split fix + V3 batch.
"""
from pathlib import Path
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.lib import colors
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    HRFlowable, Preformatted
)
from reportlab.lib.enums import TA_LEFT, TA_CENTER

OUT = Path(r"E:\claude\Pst2Thunder\PST_V3_Fixes.pdf")

doc = SimpleDocTemplate(
    str(OUT),
    pagesize=A4,
    leftMargin=2*cm, rightMargin=2*cm,
    topMargin=2*cm, bottomMargin=2*cm,
    title="PST->Thunderbird: V3 Fixes — From Header + Folder Split",
    author="PST2Thunderbird Project",
)

styles = getSampleStyleSheet()
H1 = ParagraphStyle("H1", parent=styles["Heading1"], fontSize=16,
                    spaceAfter=6, textColor=colors.HexColor("#1a3a5c"))
H2 = ParagraphStyle("H2", parent=styles["Heading2"], fontSize=13,
                    spaceAfter=4, textColor=colors.HexColor("#1a3a5c"))
H3 = ParagraphStyle("H3", parent=styles["Heading3"], fontSize=11,
                    spaceAfter=3, textColor=colors.HexColor("#2d5986"))
BODY = ParagraphStyle("Body", parent=styles["Normal"], fontSize=10,
                      spaceAfter=6, leading=14)
MONO = ParagraphStyle("Mono", parent=styles["Code"], fontSize=8.5,
                      leading=13, backColor=colors.HexColor("#f4f4f4"),
                      borderPadding=(4, 6, 4, 6), spaceAfter=8)
BULLET = ParagraphStyle("Bullet", parent=BODY, leftIndent=16,
                         bulletIndent=6, spaceAfter=3)
NOTE = ParagraphStyle("Note", parent=BODY, fontSize=9,
                      textColor=colors.HexColor("#555555"),
                      backColor=colors.HexColor("#fffbe6"),
                      borderPadding=4)

def h1(t): return Paragraph(t, H1)
def h2(t): return Paragraph(t, H2)
def h3(t): return Paragraph(t, H3)
def p(t):  return Paragraph(t, BODY)
def b(t):  return Paragraph(f"* {t}", BULLET)
def code(t): return Preformatted(t, MONO)
def note(t): return Paragraph(f"<i>{t}</i>", NOTE)
def sp(n=1): return Spacer(1, n * 0.35 * cm)
def hr():  return HRFlowable(width="100%", thickness=0.5,
                              color=colors.HexColor("#cccccc"), spaceAfter=6)

def tbl(data, col_widths=None, header=True):
    t = Table(data, colWidths=col_widths, hAlign="LEFT")
    style = [
        ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
        ("FONTSIZE", (0,0), (-1,-1), 9),
        ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#dce6f1")),
        ("ROWBACKGROUNDS", (0,1), (-1,-1),
         [colors.white, colors.HexColor("#f5f5f5")]),
        ("GRID", (0,0), (-1,-1), 0.3, colors.HexColor("#aaaaaa")),
        ("TOPPADDING", (0,0), (-1,-1), 4),
        ("BOTTOMPADDING", (0,0), (-1,-1), 4),
        ("LEFTPADDING", (0,0), (-1,-1), 6),
    ]
    t.setStyle(TableStyle(style))
    return t

story = []

# ── Title ───────────────────────────────────────────────────────────────────
story += [
    sp(1),
    Paragraph("PST→Thunderbird: V3 Fixes", H1),
    Paragraph("From/To Header Restoration &amp; Folder Split Fix", H2),
    Paragraph("Rev 2 — April 2026", BODY),
    hr(), sp(),
]

# ── 1. Overview ─────────────────────────────────────────────────────────────
story += [
    h2("1. Overview"),
    p("The V2 batch (April 11 2026) successfully converted 444,480 emails from 120 PSTs "
      "with RTF/OLE inline image support. However two bugs were found after opening "
      "Thunderbird:"),
    b("From/To columns empty for all Infineon internal emails (2007–2022)"),
    b("All PSTs falling through to ANSI flat-folder fallback due to a broken attribute reference, "
      "causing all folders to be merged into a single 'Inbox (all)'"),
    p("V3 fixes both issues and reconverts all 117 PSTs to <b>E:\\TB_Mail_v3</b> "
      "(V2 preserved as backup)."),
    sp(),
]

# ── 2. Bug 1 — From/To Headers Missing ──────────────────────────────────────
story += [
    h2("2. Bug 1 — From/To Headers Missing"),
    h3("Root Cause"),
    p("Infineon internal emails (Exchange/MAPI format) do not include a standard "
      "<b>From:</b> or <b>To:</b> header in their <tt>transport_headers</tt> field. "
      "The transport_headers for these messages contains only Received:, Content-Type:, "
      "Date:, and Subject: — no sender or recipient."),
    p("The original reconvert.py only extracted headers present in transport_headers. "
      "When none were found, the From/To fields were simply omitted from the mbox output. "
      "Thunderbird then displayed an empty Correspondents/From column."),
    h3("Example — before fix"),
    code(
        "From MAILER-DAEMON Wed Jan 26 02:53:36 2011\n"
        "Content-Type: multipart/alternative; ...\n"
        "MIME-Version: 1.0\n"
        "Date: Wed, 26 Jan 2011 10:53:36 +0800\n"
        "Subject: We have unlocked and changed the password\n"
        "                          <-- no From: header"
    ),
    h3("Fix — PST sender property fallback"),
    p("Added fallback in <tt>build_mime_message()</tt>: if From is absent from "
      "transport_headers, read <tt>message.sender_name</tt> and "
      "<tt>message.sender_email_address</tt> directly from the PST MAPI properties."),
    code(
        "# Fallback From/To from PST sender/recipient properties\n"
        'if "From" not in header_data:\n'
        "    try:\n"
        '        name      = getattr(message, "sender_name", None) or ""\n'
        '        email_addr = getattr(message, "sender_email_address", None) or ""\n'
        "        if email_addr:\n"
        '            header_data["From"] = email.utils.formataddr((name, email_addr))\n'
        "        elif name:\n"
        '            header_data["From"] = name\n'
        "    except Exception:\n"
        "        pass\n"
        "\n"
        'if "To" not in header_data:\n'
        "    try:\n"
        "        recips = []\n"
        "        for r in (message.recipients or []):\n"
        '            rname = getattr(r, "display_name", "") or ""\n'
        '            raddr = getattr(r, "email_address", "") or ""\n'
        "            if raddr:\n"
        "                recips.append(email.utils.formataddr((rname, raddr)))\n"
        "            elif rname:\n"
        "                recips.append(rname)\n"
        "        if recips:\n"
        '            header_data["To"] = ", ".join(recips)\n'
        "    except Exception:\n"
        "        pass"
    ),
    h3("Example — after fix"),
    code(
        "From MAILER-DAEMON Wed Jan 26 02:53:36 2011\n"
        "Content-Type: multipart/alternative; ...\n"
        "MIME-Version: 1.0\n"
        "Date: Wed, 26 Jan 2011 10:53:36 +0800\n"
        "Subject: We have unlocked and changed the password\n"
        "From: DMZ Support (IFAG IT)\n"
        "To: Tan Chan Lik (IFKM OP FEP T PI IC 2)"
    ),
    h3("Validation"),
    p("Sample check on 4 PSTs (50 messages each) before full batch:"),
    tbl([
        ["PST", "Era", "From coverage", "Sample sender"],
        ["cl_tan-Silterra", "2001 Silterra", "50/50 (100%)", '"Tan Chan Lik" <tan_chan_lik@yahoo.com>'],
        ["MAy 2009", "2009 Infineon", "50/50 (100%)", "Kishore Kamal (IFKM OP FEP T PI IC 2)"],
        ["Jan 2018", "2018 Infineon", "50/50 (100%)", '"Oh Guan Kai (IFKM FE TK IC SMT)" <...>'],
        ["April 2025", "2025 Recent", "50/50 (100%)", "TechInsights Platform <platform@...>"],
    ], col_widths=[3.5*cm, 3*cm, 3.5*cm, 7.5*cm]),
    sp(),
]

# ── 3. Bug 2 — All PSTs Falling to ANSI Fallback ───────────────────────────
story += [
    h2("3. Bug 2 — All PSTs Falling to ANSI Flat-Folder Fallback"),
    h3("Root Cause"),
    p("The broken-root detection used <tt>archive.archive.get_root_folder()</tt> to check "
      "if a PST has a valid folder tree. However the installed version of libratom does not "
      "expose an <tt>archive</tt> attribute on the PffArchive object — the underlying pypff "
      "file is stored as <tt>archive._data</tt>."),
    p("This caused an AttributeError on every PST, which was silently caught and set "
      "<tt>is_bad_root = True</tt>, triggering the ANSI fallback for all 117 PSTs. "
      "The fallback dumps all messages from all folders into a single 'Inbox (all)' mbox, "
      "destroying the folder hierarchy."),
    h3("Symptom"),
    code(
        "-- MAy 2009.pst\n"
        "  Broken/ANSI PST -> fallback to archive.messages()\n"
        "  Inbox (all): 3802 messages   <-- all folders merged\n"
        "\n"
        "-- Jan 2018.pst\n"
        "  Broken/ANSI PST -> fallback to archive.messages()\n"
        "  Inbox (all): 5864 messages   <-- Jan 2018 is NOT ANSI"
    ),
    h3("Fix — one-line attribute correction"),
    code(
        "# Before (broken):\n"
        "root = archive.archive.get_root_folder()\n"
        "\n"
        "# After (fixed):\n"
        "root = archive._data.get_root_folder()"
    ),
    h3("After fix — proper folder split"),
    code(
        "-- MAy 2009.pst\n"
        "  Deleted Items: 552 messages\n"
        "  Inbox: 2143 messages\n"
        "  Sent Items: 811 messages\n"
        "  PENDING: 20 messages\n"
        "  IMPLANT: 48 messages\n"
        "  ... (10 folders total)\n"
        "  -> 3638 messages"
    ),
    sp(),
]

# ── 4. V3 Batch ─────────────────────────────────────────────────────────────
story += [
    h2("4. V3 Batch — Full Reconvert"),
    p("Both fixes applied. Full batch fired to <b>E:\\TB_Mail_v3</b>. "
      "V2 preserved at E:\\TB_Mail_v2 as backup until V3 validated."),
    tbl([
        ["Parameter", "Value"],
        ["Output directory", r"E:\TB_Mail_v3"],
        ["PSTs", "117 (same set as V2)"],
        ["Task ID", "b2e0o75x6"],
        ["Started", "2026-04-11 ~19:xx"],
        ["Expected messages", "~444,000"],
        ["V2 status", "Preserved (not deleted)"],
    ], col_widths=[5*cm, 12.5*cm]),
    sp(),
]

# ── 5. Housekeeping Done ─────────────────────────────────────────────────────
story += [
    h2("5. Housekeeping Done (V2 Session)"),
    p("The following cleanup was performed on E:\\TB_Mail_v2 before V3 was started:"),
    tbl([
        ["Action", "Detail"],
        ["Deleted old date-prefix dirs", "117 dirs (212.7 GB) — April 10 batch output superseded by V2"],
        ["Deleted orphan companion files", "168 files — .msf and companion mbox files with no matching .sbd"],
        ["Renamed all folders", "YYYY-MM prefix added to all 117 dirs for chronological sort in TB"],
        ["Deleted zz* orphan dirs", "5 dirs — zz cl_tan-Silterra, zz cl_tan-lsi, zz archive, etc."],
        ["Created missing companions", "March 2008 .sbd companion file created"],
        ["Disk freed", "~213 GB freed; E: drive: 33 GB -> 246 GB free"],
    ], col_widths=[5.5*cm, 12*cm]),
    sp(),
]

# ── 6. Unfixable PSTs ────────────────────────────────────────────────────────
story += [
    h2("6. Permanently Unfixable PSTs"),
    tbl([
        ["PST", "Size", "Reason"],
        ["Aug 2023.pst", "9.6 GB", "Invalid file signature (zero header) — corrupt"],
        ["Oct 2022.pst", "2.3 GB", "Invalid file signature (zero header) — corrupt"],
        ["mailbox.pst", "?", "Missing name-to-ID map — libpff cannot open"],
        ["INFINEON Sep 2007.pst", "0 bytes", "Zero-byte file — no data"],
        ["Oct 2013.pst", "0 bytes", "Zero-byte file — no data"],
        ["Dec 2023.pst", "0 bytes", "Zero-byte file — no data"],
        ["2001.pst", "743 MB", "ANSI broken root — needs WSL+readpst to recover"],
        ["2003 PART1 1Si cltan.pst", "264 MB", "ANSI broken root — needs WSL+readpst to recover"],
    ], col_widths=[5.5*cm, 2*cm, 10*cm]),
    sp(),
]

# ── 7. File Inventory ────────────────────────────────────────────────────────
story += [
    h2("7. File Inventory"),
    tbl([
        ["File", "Purpose"],
        [r"E:\claude\Pst2Thunder\reconvert.py", "Main conversion script — V3 with all fixes"],
        [r"E:\claude\Pst2Thunder\retry_failed.py", "Re-run failed PSTs from batch log"],
        [r"E:\claude\Pst2Thunder\gap_check.py", "Scan mbox output for coverage gaps"],
        [r"E:\TB_Mail_v2" + "\\", "V2 output (444,480 msgs) - preserved as backup"],
        [r"E:\TB_Mail_v3" + "\\", "V3 output - in progress, fixes From/To + folders"],
        [r"E:\PST" + "\\", "Source PSTs (122 files, 196 GB)"],
        ["PST_OLE_Image_Solution.pdf", "Rev 1 — RTF/OLE inline image solution"],
        ["PST_V3_Fixes.pdf", "Rev 2 — this document"],
    ], col_widths=[8*cm, 9.5*cm]),
    sp(),
]

doc.build(story)
print(f"Written: {OUT}")

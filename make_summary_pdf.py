"""
make_summary_pdf.py  —  PST-to-Thunderbird conversion project summary PDF
"""

from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    HRFlowable, KeepTogether,
)
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT
from datetime import date

OUT_PATH = r"E:\claude\Pst2Thunder\PST_Conversion_Summary.pdf"

# ── colour palette ─────────────────────────────────────────────────────────
NAVY   = colors.HexColor("#1B3A6B")
STEEL  = colors.HexColor("#3A7FC1")
LIGHT  = colors.HexColor("#EAF2FB")
GREEN  = colors.HexColor("#1D6A38")
LGRE   = colors.HexColor("#E6F4EA")
RED    = colors.HexColor("#A02020")
LRED   = colors.HexColor("#FDECEA")
AMBER  = colors.HexColor("#8B5E00")
LAMB   = colors.HexColor("#FFF8E1")
GREY   = colors.HexColor("#555555")
WHITE  = colors.white
BLACK  = colors.black

styles = getSampleStyleSheet()

def S(name, **kw):
    base = styles[name]
    return ParagraphStyle(name + "_custom", parent=base, **kw)

h1   = S("Heading1", fontSize=20, textColor=NAVY, spaceAfter=4)
h2   = S("Heading2", fontSize=13, textColor=NAVY, spaceBefore=14, spaceAfter=4)
h3   = S("Heading3", fontSize=11, textColor=STEEL, spaceBefore=8, spaceAfter=3)
body = S("Normal",   fontSize=9,  textColor=BLACK, leading=13)
mono = S("Code",     fontSize=8,  textColor=GREY,  leading=11)
bold = S("Normal",   fontSize=9,  textColor=BLACK, leading=13, fontName="Helvetica-Bold")
note = S("Normal",   fontSize=8,  textColor=GREY,  leading=11, leftIndent=10)
warn = S("Normal",   fontSize=9,  textColor=AMBER, leading=13)
err  = S("Normal",   fontSize=9,  textColor=RED,   leading=13)
ok   = S("Normal",   fontSize=9,  textColor=GREEN, leading=13)

def HR():
    return HRFlowable(width="100%", thickness=0.5, color=STEEL, spaceAfter=6)

def SP(h=6):
    return Spacer(1, h)

def tbl(data, col_widths, style_extra=None):
    base_style = [
        ("FONTNAME",    (0,0), (-1,0), "Helvetica-Bold"),
        ("FONTSIZE",    (0,0), (-1,-1), 9),
        ("BACKGROUND",  (0,0), (-1,0), NAVY),
        ("TEXTCOLOR",   (0,0), (-1,0), WHITE),
        ("ROWBACKGROUNDS", (0,1), (-1,-1), [WHITE, LIGHT]),
        ("GRID",        (0,0), (-1,-1), 0.3, colors.HexColor("#CCCCCC")),
        ("VALIGN",      (0,0), (-1,-1), "MIDDLE"),
        ("LEFTPADDING",  (0,0), (-1,-1), 5),
        ("RIGHTPADDING", (0,0), (-1,-1), 5),
        ("TOPPADDING",   (0,0), (-1,-1), 3),
        ("BOTTOMPADDING",(0,0), (-1,-1), 3),
    ]
    if style_extra:
        base_style += style_extra
    t = Table(data, colWidths=col_widths)
    t.setStyle(TableStyle(base_style))
    return t

# ── content builder ────────────────────────────────────────────────────────
story = []

# ── Title block ──────────────────────────────────────────────────────────
story += [
    SP(10),
    Paragraph("PST → Thunderbird Conversion", h1),
    Paragraph("Project Summary Report", S("Heading2", fontSize=14, textColor=STEEL, spaceBefore=0, spaceAfter=2)),
    Paragraph(f"Generated: {date.today().strftime('%d %B %Y')}  |  Owner: Tan Chan Lik", note),
    HR(),
    SP(4),
]

# ── 1. Project Overview ──────────────────────────────────────────────────
story += [Paragraph("1. Project Overview", h2), HR()]
overview = [
    ["Objective",    "Convert all personal Outlook PST archives to Thunderbird-readable mbox format"],
    ["Source",       "E:\\PST\\  (120 PST files, 172 GB total)"],
    ["Destination",  "E:\\TB_Mail_v2\\  (Thunderbird Local Folders)"],
    ["Scripts",      "E:\\claude\\Pst2Thunder\\  |  GitHub: chanlikt-lgtm/PST2Thunderbird"],
    ["Date Range",   "November 1999 → August 2024  (25 years of email)"],
    ["Status",       "COMPLETE — all available PSTs converted"],
]
ov_t = Table([[Paragraph(k, bold), Paragraph(v, body)] for k, v in overview],
             colWidths=[3.8*cm, 13.2*cm])
ov_t.setStyle(TableStyle([
    ("FONTSIZE",    (0,0), (-1,-1), 9),
    ("VALIGN",      (0,0), (-1,-1), "TOP"),
    ("ROWBACKGROUNDS", (0,0), (-1,-1), [WHITE, LIGHT]),
    ("GRID",        (0,0), (-1,-1), 0.3, colors.HexColor("#CCCCCC")),
    ("LEFTPADDING",  (0,0), (-1,-1), 5),
    ("RIGHTPADDING", (0,0), (-1,-1), 5),
    ("TOPPADDING",   (0,0), (-1,-1), 4),
    ("BOTTOMPADDING",(0,0), (-1,-1), 4),
]))
story += [ov_t, SP(10)]

# ── 2. Conversion Statistics ─────────────────────────────────────────────
story += [Paragraph("2. Conversion Statistics", h2), HR()]
stats_data = [
    ["Metric", "Value", "Notes"],
    ["PST files processed",   "120",          "of which 7 are zero-byte placeholders"],
    ["Total PST size",        "172 GB",        "source data on E:\\PST\\"],
    ["Output size",           "178 GB",        "E:\\TB_Mail_v2\\  (MIME encoding adds ~7%)"],
    ["Size ratio",            "~107%",         "output / source  (expected: >5%  →  all PASS)"],
    ["Emails converted",      "386,424",       "across 111 Thunderbird folders"],
    ["Total attachments",     "379,811",       "across all folders"],
    ["Thunderbird folders",   "111",           "top-level mail folders (YYYY-MM sorted)"],
    ["Conversion quality",    "111 OK / 3 Suspect / 6 Missing", "suspect = zero-byte source PSTs"],
]
story += [tbl(stats_data,
              [5.5*cm, 3.5*cm, 9.0*cm],
              style_extra=[
                  ("BACKGROUND", (0,3), (-1,3), LGRE),
                  ("BACKGROUND", (0,5), (-1,5), LGRE),
              ]),
          SP(10)]

# ── 3. Email Coverage Map ────────────────────────────────────────────────
story += [Paragraph("3. Email Coverage  (by year/month)", h2), HR(),
          Paragraph("Based on actual email Date headers.  O = covered  |  . = missing  |  <b>■ = full year</b>", note),
          SP(4)]

cal_header = ["Year", "Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]
CAL = [
    ["1999",".",".",".",".",".",".",".",".",".",".", "O","O"],
    ["2000","O","O","O","O","O","O","O","O",".",".",".","O"],
    ["2001","O","O","O","O","O","O","O","O","O",".","O","."],
    ["2002",".",".",  "O",".","O",".",".",".", "O","O","O","O"],
    ["2003","O","O","O","O",".","O","O","O","O",".",".","O"],
    ["2004","O","O","O",".",".",".",".",".",".",".",".","." ],
    ["2005",".",".",".",".",".",".",".",".",".",".",".","."],
    ["2006",".",".",".",".",".",".",".",".",".",".",".","."],
    ["2007",".",".",".",".",".",".",".","O",".",".",".","."],
    ["2008",".",".",".",".",".",".",".",".",".",".",".","."],
    ["2009",".",".",".",".",".",".",".",".",".","O","O","O"],
]
full_years = [str(y) for y in range(2010, 2023)]
for yr in full_years:
    CAL.append([yr, "■","■","■","■","■","■","■","■","■","■","■","■"])

CAL += [
    ["2022","■","■","■","■","■","■","■","■","■","■","■","■"],
    ["2023","~","~","~","~","~","~","~","~","~","~","~","~"],
    ["2024",".",".",".",".",".",".",".","O","-","-","-","-"],
]

cal_data = [cal_header] + CAL
col_w = [1.3*cm] + [1.0*cm]*12

cal_style = [
    ("FONTNAME",    (0,0), (-1,0), "Helvetica-Bold"),
    ("FONTSIZE",    (0,0), (-1,-1), 7.5),
    ("BACKGROUND",  (0,0), (-1,0), NAVY),
    ("TEXTCOLOR",   (0,0), (-1,0), WHITE),
    ("ALIGN",       (0,0), (-1,-1), "CENTER"),
    ("GRID",        (0,0), (-1,-1), 0.2, colors.HexColor("#BBBBBB")),
    ("TOPPADDING",  (0,0), (-1,-1), 2),
    ("BOTTOMPADDING",(0,0), (-1,-1), 2),
    ("ROWBACKGROUNDS", (0,1), (-1,-1), [WHITE, LIGHT]),
]
# Highlight fully-covered years
for i, row in enumerate(cal_data[1:], start=1):
    yr = row[0]
    if yr in full_years:
        cal_style.append(("BACKGROUND", (0,i), (-1,i), LGRE))
        cal_style.append(("TEXTCOLOR",  (1,i), (-1,i), GREEN))
        cal_style.append(("FONTNAME",   (1,i), (-1,i), "Helvetica-Bold"))
# Highlight 2023 gap row
for i, row in enumerate(cal_data[1:], start=1):
    if row[0] == "2023":
        cal_style.append(("BACKGROUND", (0,i), (-1,i), LRED))

ct = Table(cal_data, colWidths=col_w)
ct.setStyle(TableStyle(cal_style))
story += [ct, SP(6),
          Paragraph("■ 2010–2022: Fully covered (every month — 156 consecutive months)", ok),
          Paragraph("● 2005–2006, 2008: No emails found — PSTs existed but are corrupt/zero-byte", warn),
          Paragraph("~ 2023: 81 sent emails exist but have no Date header — show as undated in Thunderbird", warn),
          Paragraph("✗ 2004–2007, 2007–2009: Major gaps — no PSTs found for those months", err),
          SP(10)]

# ── 4. Gap Analysis ──────────────────────────────────────────────────────
story += [Paragraph("4. Gap Analysis  (confirmed)", h2), HR()]
gap_data = [
    ["Gap Period", "Missing Months", "Status / Notes"],
    ["Jan 1999 – Oct 1999",    "10 months", "Pre-archive — likely never backed up"],
    ["Sep 2000 – Nov 2000",    "3 months",  "Searching for backup"],
    ["Oct 2001",               "1 month",   "Single month gap"],
    ["Dec 2001 – Feb 2002",    "3 months",  "Searching for backup"],
    ["Apr 2002",               "1 month",   "Single month gap"],
    ["Jun 2002 – Aug 2002",    "3 months",  "Searching for backup"],
    ["May 2003",               "1 month",   "Single month gap"],
    ["Oct 2003 – Nov 2003",    "2 months",  "Searching for backup"],
    ["Apr 2004 – Jul 2007",    "40 months", "⚠ MAJOR GAP — 3+ years  (IFX Silterra era)"],
    ["Sep 2007 – Sep 2009",    "25 months", "⚠ MAJOR GAP — 2 years  (INFINEON Sep 2007 corrupt)"],
    ["2022",                   "0 months",  "✓ FULLY COVERED — all 12 months confirmed (prev. scan was wrong)"],
    ["2023",                   "—",         "81 sent emails exist but NO Date header — appear undated"],
    ["2024",                   "—",         "Aug 2024 only (Sent folder); rest not archived"],
]
gap_style_extra = [
    ("BACKGROUND", (0,10), (-1,10), LAMB),
    ("BACKGROUND", (0,11), (-1,11), LAMB),
    ("BACKGROUND", (0,12), (-1,12), LGRE),
    ("TEXTCOLOR",  (0,12), (-1,12), GREEN),
    ("BACKGROUND", (0,13), (-1,13), LAMB),
    ("BACKGROUND", (0,14), (-1,14), LIGHT),
]
story += [tbl(gap_data, [4.5*cm, 2.5*cm, 10.0*cm], style_extra=gap_style_extra), SP(10)]

# ── 5. Known Issues & Pending Items ─────────────────────────────────────
story += [Paragraph("5. Known Issues & Pending Items", h2), HR()]

issues = [
    ["#", "Issue", "Severity", "Action Required"],
    ["1", "2023 emails entirely missing — no PST on any drive",
         "HIGH", "Find backup  (BaiduNetdisk download in progress)"],
    ["2", "INFINEON 25 Sep 2007.pst — ANSI format, unreadable by libpff",
         "HIGH", "Install WSL + readpst to recover 1.8 GB PST"],
    ["3", "59 .dat attachments (TNEF/winmail.dat)",
         "MED", "Build fix_dat.py with tnefparse to decode to proper Office files"],
    ["4", "1,103 S/MIME encrypted emails — unreadable",
         "MED", "Import private key (.p12/.pfx) into Thunderbird"],
    ["5", "Oct 2013.pst — suspect (zero-byte source)",
         "LOW", "Verify source PST is truly empty; search for backup"],
    ["6", "Dec 2023.pst — suspect (zero-byte source)",
         "LOW", "Part of 2023 gap; find full 2023 backup"],
    ["7", "5 duplicate unsorted folders (2004/2007/2008 era)",
         "LOW", "Close Thunderbird, delete old prefixed duplicates, resort"],
]
iss_style = [
    ("BACKGROUND", (0,2), (-1,2), LRED),
    ("BACKGROUND", (0,3), (-1,3), LRED),
    ("BACKGROUND", (0,4), (-1,4), LAMB),
    ("BACKGROUND", (0,5), (-1,5), LAMB),
]
story += [tbl(issues, [0.5*cm, 6.5*cm, 1.5*cm, 8.5*cm], style_extra=iss_style), SP(10)]

# ── 6. Suspect PSTs ──────────────────────────────────────────────────────
story += [Paragraph("6. PST Quality Audit", h2), HR()]
suspect_data = [
    ["PST File", "PST Size", "Output Size", "Ratio", "Status"],
    ["All 111 converted PSTs",        "—",     "—",    ">5%",  "PASS"],
    ["archive.pst",                   "0 MB",  "0 MB", "—",    "zero-byte placeholder"],
    ["Outlook.pst",                   "0 MB",  "0 MB", "—",    "zero-byte placeholder"],
    ["mailbox.pst",                   "0 MB",  "0 MB", "—",    "zero-byte placeholder"],
    ["TanChanNov05.pst",              "0 MB",  "0 MB", "—",    "zero-byte placeholder"],
    ["Dec 2023.pst",                  "0 MB",  "0 MB", "—",    "SUSPECT — part of 2023 gap"],
    ["Oct 2013.pst",                  "0 MB",  "0 MB", "—",    "SUSPECT — find backup"],
    ["INFINEON 25 Sep 2007.pst",      "0.3 MB",  "0 MB", "—",  "near-empty (7 folders, structural)"],
    ["INFINEON  Sep 2007.pst",        "1,800 MB","0 MB","0%",  "CORRUPT — ANSI format, needs WSL+readpst"],
    ["2001.pst",                      "?",       "0 MB","—",   "CORRUPT — sub_folders missing"],
    ["2003 PART1 1Si cltan.pst",      "?",       "0 MB","—",   "CORRUPT — sub_folders missing"],
    ["MAc 2015.pst  (deleted)",       "10,900 MB","—", "—",    "CORRUPT — all 0xFF bytes, deleted"],
]
susp_style = [
    ("BACKGROUND", (0,1), (-1,1), LGRE),
    ("TEXTCOLOR",  (0,1), (-1,1), GREEN),
    ("BACKGROUND", (0,7), (-1,7), LAMB),
    ("BACKGROUND", (0,8), (-1,8), LAMB),
    ("BACKGROUND", (0,9), (-1,9), LRED),
    ("TEXTCOLOR",  (0,9), (-1,9), RED),
    ("BACKGROUND", (0,10), (-1,10), LRED),
    ("TEXTCOLOR",  (0,10), (-1,10), RED),
]
story += [tbl(suspect_data,
              [5.5*cm, 2.0*cm, 2.0*cm, 1.5*cm, 6.0*cm],
              style_extra=susp_style), SP(10)]

# ── 7. Attachment & Content Analysis ────────────────────────────────────
story += [Paragraph("7. Attachment & Content Analysis", h2), HR()]
att_data = [
    ["Category", "Count", "Notes"],
    ["Total emails",              "386,424",    "across 111 Thunderbird folders"],
    ["Total attachments",         "379,811",    "all formats"],
    ["Standard attachments",      "379,752",    "PDF, Office, images, ZIP, etc. — correct format"],
    [".dat (TNEF/winmail.dat)",   "59",         "Outlook RTF format — need fix_dat.py to decode"],
    ["S/MIME encrypted emails",   "1,103",      "Unreadable without private key (.p12/.pfx)"],
]
att_style = [
    ("BACKGROUND", (0,5), (-1,5), LAMB),
    ("BACKGROUND", (0,6), (-1,6), LAMB),
]
story += [tbl(att_data, [5.0*cm, 2.5*cm, 9.5*cm], style_extra=att_style),
          SP(6),
          Paragraph("Note: 379,752 attachments (99.98%) are in correct native format. "
                    "Only 59 .dat files require conversion via fix_dat.py (tnefparse library).", note),
          SP(10)]

# ── 8. Tools & Scripts ───────────────────────────────────────────────────
story += [Paragraph("8. Tools & Scripts", h2), HR()]
tools_data = [
    ["Script", "Purpose", "Status"],
    ["reconvert.py",      "Re-convert PST → mbox with correct MIME + From_ dates",  "DONE — ran reconvert --all"],
    ["auto_flow.py",      "Hourly Task Scheduler job to pick up new PSTs automatically", "ACTIVE — PST2Thunderbird_AutoFlow"],
    ["scan_new_drive.py", "Scan all drives for new PST files",                       "DONE — found 5 new on G:\\"],
    ["sort_folders.py",   "Prefix folder names with YYYY-MM for chronological sort", "DONE — 111 folders sorted"],
    ["verify_sizes.py",   "Audit PST vs output size ratios (flag if < 5%)",          "DONE — 111 OK, 3 suspect"],
    ["gap_check.py",      "Scan mbox Date headers to find coverage gaps",            "DONE — 111 missing months found"],
    ["sanity_check.py",   "Full audit: messages, attachments, encrypted, .dat",      "DONE — report logged"],
    ["fix_dat.py",        "Decode 59 TNEF .dat attachments to Office format",        "PENDING — not yet built"],
]
story += [tbl(tools_data, [3.8*cm, 8.0*cm, 5.2*cm]), SP(10)]

# ── 9. Next Steps ────────────────────────────────────────────────────────
story += [Paragraph("9. Next Steps (Priority Order)", h2), HR()]
steps = [
    ("HIGH",  "RED",  "Find 2023 email backup — check BaiduNetdisk download (E:\\BaiduNetdiskDownload)"),
    ("HIGH",  "RED",  "Install WSL + readpst to recover INFINEON 25 Sep 2007.pst (1.8 GB ANSI format)"),
    ("MED",   "AMBER","Build fix_dat.py to decode 59 TNEF attachments to proper Office files"),
    ("MED",   "AMBER","Import S/MIME private key (.p12/.pfx) into Thunderbird to read 1,103 encrypted emails"),
    ("LOW",   "GREY", "Delete 5 duplicate unsorted folders (close TB first, then re-run sort_folders.py)"),
    ("LOW",   "GREY", "Re-run gap_check.py after BaiduNetdisk + INFINEON recovery"),
    ("LOW",   "GREY", "Git push final state to chanlikt-lgtm/PST2Thunderbird"),
]
step_data = [["Pri", "Action"]]
for pri, col, text in steps:
    c = {"RED": RED, "AMBER": AMBER, "GREY": GREY}[col]
    step_data.append([
        Paragraph(f"<b>{pri}</b>", ParagraphStyle("x", fontSize=8, textColor=c)),
        Paragraph(text, body)
    ])
st = Table(step_data, colWidths=[1.2*cm, 15.8*cm])
st.setStyle(TableStyle([
    ("FONTSIZE",    (0,0), (-1,-1), 9),
    ("ROWBACKGROUNDS", (0,1), (-1,-1), [WHITE, LIGHT]),
    ("GRID",        (0,0), (-1,-1), 0.3, colors.HexColor("#CCCCCC")),
    ("VALIGN",      (0,0), (-1,-1), "TOP"),
    ("LEFTPADDING",  (0,0), (-1,-1), 5),
    ("RIGHTPADDING", (0,0), (-1,-1), 5),
    ("TOPPADDING",   (0,0), (-1,-1), 4),
    ("BOTTOMPADDING",(0,0), (-1,-1), 4),
    ("BACKGROUND",  (0,0), (-1,0), NAVY),
    ("TEXTCOLOR",   (0,0), (-1,0), WHITE),
    ("FONTNAME",    (0,0), (-1,0), "Helvetica-Bold"),
]))
story += [st, SP(16)]

# ── Footer note ──────────────────────────────────────────────────────────
story += [
    HR(),
    Paragraph("This report was generated automatically from E:\\claude\\Pst2Thunder\\make_summary_pdf.py", note),
    Paragraph(f"Report date: {date.today().isoformat()}  |  Output: E:\\TB_Mail_v2\\  |  Scripts: E:\\claude\\Pst2Thunder\\", note),
]

# ── Build PDF ────────────────────────────────────────────────────────────
doc = SimpleDocTemplate(
    OUT_PATH,
    pagesize=A4,
    leftMargin=1.8*cm,
    rightMargin=1.8*cm,
    topMargin=2.0*cm,
    bottomMargin=2.0*cm,
    title="PST to Thunderbird Conversion Summary",
    author="Tan Chan Lik",
)
doc.build(story)
print(f"PDF saved: {OUT_PATH}")

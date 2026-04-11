"""
make_ole_pdf.py — Generate PDF documenting the RTF/OLE inline image solution.
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

OUT = Path(r"E:\claude\Pst2Thunder\PST_OLE_Image_Solution.pdf")

doc = SimpleDocTemplate(
    str(OUT),
    pagesize=A4,
    leftMargin=2*cm, rightMargin=2*cm,
    topMargin=2*cm, bottomMargin=2*cm,
    title="PST→Thunderbird: RTF/OLE Inline Image Solution",
    author="PST2Thunderbird Project",
)

styles = getSampleStyleSheet()

# Custom styles
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

def h1(text): return Paragraph(text, H1)
def h2(text): return Paragraph(text, H2)
def h3(text): return Paragraph(text, H3)
def p(text):  return Paragraph(text, BODY)
def b(text):  return Paragraph(f"• {text}", BULLET)
def code(text): return Preformatted(text, MONO)
def note(text): return Paragraph(f"<i>ℹ {text}</i>", NOTE)
def sp(n=1):  return Spacer(1, n * 0.35 * cm)
def hr():     return HRFlowable(width="100%", thickness=0.5,
                                color=colors.HexColor("#cccccc"), spaceAfter=6)

def tbl(data, col_widths=None, header=True):
    t = Table(data, colWidths=col_widths, hAlign="LEFT")
    style = [
        ("FONTNAME",    (0, 0), (-1, -1), "Helvetica"),
        ("FONTSIZE",    (0, 0), (-1, -1), 9),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1),
         [colors.white, colors.HexColor("#f0f4f8")]),
        ("GRID",        (0, 0), (-1, -1), 0.4, colors.HexColor("#cccccc")),
        ("VALIGN",      (0, 0), (-1, -1), "TOP"),
        ("TOPPADDING",  (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING",(0, 0), (-1, -1), 3),
        ("LEFTPADDING", (0, 0), (-1, -1), 5),
    ]
    if header:
        style += [
            ("BACKGROUND",  (0, 0), (-1, 0), colors.HexColor("#1a3a5c")),
            ("FONTNAME",    (0, 0), (-1, 0), "Helvetica-Bold"),
            ("TEXTCOLOR",   (0, 0), (-1, 0), colors.white),
        ]
    t.setStyle(TableStyle(style))
    return t

# ── Document flow ────────────────────────────────────────────────────────────
story = []

# Title
story += [
    sp(1),
    Paragraph("PST → Thunderbird", ParagraphStyle(
        "Title", parent=styles["Title"], fontSize=22,
        textColor=colors.HexColor("#1a3a5c"), alignment=TA_CENTER)),
    Paragraph("RTF / OLE Inline Image Solution", ParagraphStyle(
        "Sub", parent=styles["Normal"], fontSize=14,
        textColor=colors.HexColor("#2d5986"), alignment=TA_CENTER,
        spaceAfter=4)),
    Paragraph("Technical Reference — April 2026", ParagraphStyle(
        "Date", parent=styles["Normal"], fontSize=10,
        textColor=colors.grey, alignment=TA_CENTER, spaceAfter=16)),
    hr(), sp(),
]

# ── 1. Problem Statement ─────────────────────────────────────────────────────
story += [h1("1. Problem Statement"), sp()]
story += [p(
    "Outlook PST files from approximately 2003–2012 store embedded images as "
    "OLE compound documents attached to the email, rather than as standard MIME "
    "inline attachments with Content-ID references. When converted naively to "
    "mbox format, these images are either lost or appear as placeholder text:"
), sp(0.5)]
story += [code('    << OLE Object: Picture (Device Independent Bitmap) >>')]
story += [sp(0.5), p(
    "Additionally, emails in this era often have only an RTF body (no HTML), "
    "and the RTF contains structural groups (font tables, stylesheets, color "
    "tables) that leaked as visible garbage text in earlier conversion attempts:"
), sp(0.5)]
story += [code('    Arial;Tahoma;Comic Sans MS;\n'
               '    \\*Riched20 5.50.99.2017;\n'
               '    Normal;\\*Default Paragraph Font;\\*Hyperlink;')]
story += [sp()]

# ── 2. Root Cause ────────────────────────────────────────────────────────────
story += [h1("2. Root Cause Analysis"), sp()]
story += [
    h3("2.1  OLE Embedded Objects"),
    p("MAPI attach_method = 6 (ATTACH_OLE) stores the object payload as an OLE2 "
      "compound document. The compound document contains a <b>CONTENTS</b> stream "
      "which holds the raw image as a Windows BMP file (<tt>BM</tt> magic)."),
    p("In the RTF body, the placeholder <tt>\\objattph</tt> control word marks the "
      "exact position where the image should appear inline. Without handling this "
      "control word, the image position information is lost."),
    sp(0.5),
    h3("2.2  RTF Header Group Leakage"),
    p("The old conversion stripped RTF braces with <tt>text.replace('{','').replace('}','')</tt>, "
      "leaving the text content of structural groups (font table, color table, stylesheet, "
      "Word namespace tables) visible as plain text."),
    sp(0.5),
    h3("2.3  Color / Formatting Loss"),
    p("Character formatting (<tt>\\cfN</tt> color, <tt>\\b</tt> bold, <tt>\\i</tt> italic, "
      "<tt>\\ul</tt> underline) was discarded by a regex that stripped all control words, "
      "and the color table (<tt>{\\colortbl}</tt>) was never parsed."),
    sp(),
]

# ── 3. Solution Architecture ─────────────────────────────────────────────────
story += [h1("3. Solution Architecture"), sp()]
story += [p("The fix is implemented in <b>reconvert.py</b> as three cooperating phases "
            "within the RTF rendering pipeline:"), sp(0.5)]

story += [tbl(
    [["Phase", "Handles", "Key Function"],
     ["Phase 1", "Body selection — HTML > RTF-rendered > plain",
      "should_try_rtf_fallback()"],
     ["Phase 2", "\\\\pict group extraction (inline images in RTF body)",
      "_extract_pict_groups()"],
     ["Phase 3", "ATTACH_OLE CONTENTS BMP → PNG, placed at \\\\objattph",
      "_extract_ole_images_list()"],
    ],
    col_widths=[2*cm, 8*cm, 6*cm],
), sp()]

story += [
    h3("3.1  Phase 1 — Body Selection"),
    p("The <tt>should_try_rtf_fallback()</tt> function decides when to activate the "
      "RTF rendering pipeline. It triggers when:"),
    b("The HTML body contains <tt>&lt;&lt; OLE Object &gt;&gt;</tt> placeholder text, <b>or</b>"),
    b("The plain-text body contains the placeholder, <b>or</b>"),
    b("No HTML body exists at all."),
    p("HTML bodies with valid CID inline image references are always kept as-is."),
    sp(0.5),

    h3("3.2  Phase 2 — \\\\pict Image Extraction"),
    p("RTF <tt>{\\pict ...}</tt> groups contain hex-encoded image data with format "
      "markers (<tt>\\pngblip</tt>, <tt>\\jpegblip</tt>, <tt>\\dibitmap</tt>). "
      "Each group is extracted, decoded, and tokenized before RTF body rendering. "
      "Tokens are replaced with <tt>&lt;img src=\"data:...\"&gt;</tt> tags at the "
      "correct position in the output HTML."),
    sp(0.5),

    h3("3.3  Phase 3 — OLE CONTENTS BMP Extraction"),
    p("For each attachment, <tt>_extract_ole_images_list()</tt> checks:"),
    b("Does the attachment data start with OLE2 magic bytes "
      "<tt>D0 CF 11 E0 A1 B1 1A E1</tt>?"),
    b("Does the compound document contain a <b>CONTENTS</b> stream?"),
    b("Does the CONTENTS stream start with <tt>BM</tt> (Windows BMP)?"),
    p("A slot is added to the list for every OLE compound doc attachment — "
      "<tt>(mime, bytes)</tt> for image slots, <tt>None</tt> for non-image OLE "
      "objects (e.g. embedded Word/Excel). This slot-preserving approach prevents "
      "positional drift when non-image OLE objects precede image ones."),
    p("BMP images are converted to PNG via Pillow and embedded as "
      "<tt>data:image/png;base64,...</tt> URIs so they are self-contained in the mbox."),
    sp(0.5),

    h3("3.4  Inline Placement via \\\\objattph"),
    p("The <tt>_rtf_body_to_html()</tt> state-machine parser handles "
      "<tt>\\objattph</tt> by consuming the next slot from the OLE image list "
      "at that exact position in the text flow. Non-image slots (None) advance "
      "the counter without emitting any HTML, maintaining alignment."),
    note("This replaces the earlier approach of appending all OLE images at the "
         "end of the email body, which lost the original inline position."),
    sp(),
]

# ── 4. RTF Rendering Improvements ───────────────────────────────────────────
story += [h1("4. RTF Rendering Improvements"), sp()]

story += [
    h3("4.1  Header Group Stripping"),
    p("A brace-balanced group removal function (<tt>_remove_balanced_groups()</tt>) "
      "strips known structural RTF groups before any text extraction:"),
    sp(0.5),
]
story += [tbl(
    [["Pattern stripped", "Content it removes"],
     ["{\\\\*\\\\keyword ...}", "All destination groups incl. \\\\*\\\\generator, \\\\*\\\\fldinst"],
     ["{\\\\fonttbl ...}",  "Font names and PANOSE data"],
     ["{\\\\colortbl ...}", "Color definitions (parsed separately before stripping)"],
     ["{\\\\stylesheet ...}", "Paragraph/character style names"],
     ["{\\\\info ...}",    "Document metadata"],
     ["{\\\\latentstyles ...}", "Word latent style tables"],
     ["{\\\\xmlnstbl ...}", "Word XML namespace tables"],
    ],
    col_widths=[6*cm, 10*cm],
), sp()]

story += [
    h3("4.2  Color Preservation"),
    p("<tt>_parse_color_table()</tt> extracts the <tt>{\\colortbl}</tt> group "
      "<b>before</b> header stripping, building a mapping of color index → "
      "hex color string (e.g. <tt>{1: '#ff0000', 2: '#0000ff'}</tt>)."),
    p("The RTF body renderer applies these colors as CSS: "
      "<tt>&lt;span style=\"color:#rrggbb\"&gt;</tt>. The color table is "
      "never modified — colors in Thunderbird exactly match the PST originals."),
    sp(0.5),

    h3("4.3  Full Formatting State Machine"),
    p("<tt>_rtf_body_to_html()</tt> is a proper token-stream parser with a "
      "formatting state stack. RTF groups (<tt>{}</tt>) push and restore "
      "formatting state, so character formatting is correctly scoped. "
      "Supported properties:"),
    sp(0.5),
]

story += [tbl(
    [["RTF Control",    "HTML Output"],
     ["\\\\cfN",        "<span style=\"color:#rrggbb\"> (from color table)"],
     ["\\\\b / \\\\b0", "<strong> / </strong>"],
     ["\\\\i / \\\\i0", "<em> / </em>"],
     ["\\\\ul / \\\\ulnone", "<u> / </u>"],
     ["\\\\par",        "</p><p> (paragraph break, formatting preserved)"],
     ["\\\\plain",      "Reset all character formatting"],
     ["\\\\'XX",        "cp1252 hex character decode"],
     ["\\\\uN",         "Unicode character (with fallback skip)"],
    ],
    col_widths=[5*cm, 11*cm],
), sp()]

# ── 5. OLE Duplicate Suppression ─────────────────────────────────────────────
story += [h1("5. OLE Duplicate Suppression"), sp()]
story += [
    p("Without suppression, OLE compound doc attachments appear twice: once as an "
      "inline image (placed by Phase 3) and once as a file attachment in the "
      "message pane."),
    p("In the attachment processing loop, any attachment that meets all three "
      "conditions is silently dropped from the regular attachment list:"),
    b("The email used the RTF-rendered HTML path (Phase 3 was active)."),
    b("The attachment data starts with OLE2 magic bytes."),
    b("The CONTENTS stream starts with <tt>BM</tt> (confirmed image — not a "
      "non-image OLE that was never inlined)."),
    sp(),
]

# ── 6. Archive Read Hardening ─────────────────────────────────────────────────
story += [h1("6. Archive Read Hardening"), sp()]
story += [
    p("Several PSTs from 2001–2004 have broken or ANSI-format roots that caused "
      "the folder traversal to crash with <tt>NoneType has no attribute sub_folders</tt>. "
      "Five hardening changes were applied to <tt>reconvert()</tt>:"),
    sp(0.5),
]
story += [tbl(
    [["Change", "Before", "After"],
     ["Root detection",
      "root is None → ANSI",
      "root is None OR not hasattr(root, 'sub_folders') → fallback"],
     ["folders() call",
      "for folder in archive.folders():",
      "folders = archive.folders() or []; for folder in folders:"],
     ["Null folder guard",
      "if folder is None: continue",
      "if not folder: continue"],
     ["sub_messages access",
      "msgs = list(folder.sub_messages)",
      "getattr(folder, 'sub_messages', None) with try/except per folder"],
     ["Failure logging",
      "print only",
      "Appends to failed_psts.log for retry"],
    ],
    col_widths=[4*cm, 5.5*cm, 6.5*cm],
), sp()]

story += [
    p("PSTs that fail the hardened traversal fall back to <tt>archive.messages()</tt>, "
      "which dumps all messages into a single Inbox folder. Truly corrupt PSTs "
      "(zero header, 0-byte) still fail and are logged for manual review."),
    sp(),
]

# ── 7. Retry Infrastructure ───────────────────────────────────────────────────
story += [h1("7. Retry Infrastructure"), sp()]
story += [
    p("<b>retry_failed.py</b> provides automatic retry of failed PSTs after any batch run:"),
    b("Reads failure list from <tt>failed_psts.log</tt> (hardened runs) or "
      "parses batch output files for <tt>FAILED</tt> lines (legacy runs)."),
    b("Skips PSTs whose <tt>.sbd</tt> output folder already exists and is "
      "non-empty — avoids re-processing already-converted PSTs."),
    b("Appends retry failures back to <tt>failed_psts.log</tt> so runs can "
      "be chained: run → retry → retry until stable."),
    b("Truncates large skip/missing lists to 20 entries with <tt>+N more</tt>."),
    b("Prints a final summary: Retried / OK / Failed / Skipped."),
    sp(),
    p("Expected outcome after retry:"),
]
story += [tbl(
    [["Category",             "Expected result"],
     ["Broken-root PSTs",     "✓ Converted via fallback path"],
     ["Partial folder errors", "✓ Mostly converted (per-folder guard)"],
     ["Truly corrupt PSTs",   "✗ Still fail — logged for manual review"],
    ],
    col_widths=[6*cm, 10*cm],
), sp()]

# ── 8. File Inventory ─────────────────────────────────────────────────────────
story += [h1("8. File Inventory"), sp()]
story += [tbl(
    [["File",                 "Purpose"],
     ["reconvert.py",         "Main converter — all phases, hardening, MIME build"],
     ["retry_failed.py",      "Retry helper — reads failure log, skips existing output"],
     ["gap_check.py",         "Scan mbox Date headers to find coverage gaps"],
     ["sort_folders.py",      "Rename unsorted output folders to YYYY-MM Name format"],
     ["failed_psts.log",      "Auto-generated — lists PSTs that failed with error message"],
    ],
    col_widths=[5*cm, 11*cm],
), sp()]

# ── Footer ────────────────────────────────────────────────────────────────────
story += [
    hr(),
    Paragraph(
        "PST2Thunderbird Project · github.com/chanlikt-lgtm/PST2Thunderbird · April 2026",
        ParagraphStyle("Footer", parent=styles["Normal"], fontSize=8,
                       textColor=colors.grey, alignment=TA_CENTER)
    ),
]

doc.build(story)
print(f"Written: {OUT}")

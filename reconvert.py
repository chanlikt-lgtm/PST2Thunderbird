"""
reconvert.py — Reconvert specific PSTs with full MIME support:
  - Plain text + HTML body (whichever is available)
  - Attachments preserved (ppt, xlsx, dat, pdf, etc.)

Usage:
    python reconvert.py "Dec 2021" "Nov 2021"
    python reconvert.py --all        (reconvert everything in E:\PST\)
"""

import re
import sys
import io
import html as _html
import base64
import binascii
import struct
import datetime
from io import BytesIO

try:
    import olefile as _olefile
    OleFileIO = _olefile.OleFileIO
    _isOleFile = _olefile.isOleFile
except Exception:
    OleFileIO = None
    _isOleFile = None
import shutil
import mailbox
import mimetypes
import email
import email.mime.multipart
import email.mime.text
import email.mime.base
import email.utils
import time as _time
from email import encoders
from pathlib import Path
from libratom.lib.pff import PffArchive

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

PST_DIR    = Path(r"E:\PST")
OUT_BASE   = Path(r"E:\TB_Mail_v3")
FAILED_LOG = Path(r"E:\claude\Pst2Thunder\failed_psts.log")
_OLE_MAGIC = b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1'

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


def extract_cids(html: str) -> set:
    if not html:
        return set()
    return {m.strip().strip("<>").strip()
            for m in re.findall(r'cid:([^"\'> ]+)', html, re.IGNORECASE)}


def looks_like_ole_placeholder_only(text: str) -> bool:
    if not text:
        return False
    t = text.lower()
    return "<< ole object:" in t or "<<ole object:" in t


def should_try_rtf_fallback(html: str, plain: str, rtf_raw) -> bool:
    if not rtf_raw:
        return False
    # HTML already has CID inline images — keep HTML path
    if html and "cid:" in html.lower():
        return False
    # Either body shows OLE placeholder — RTF has the real images
    if (plain and looks_like_ole_placeholder_only(plain)) or \
       (html and looks_like_ole_placeholder_only(html)):
        return True
    # No HTML at all but RTF exists
    if not html:
        return True
    return False


def get_message_rtf_data(message):
    """Return raw RTF bytes from the PST message. libratom decompresses internally."""
    for attr in ("rtf_body", "compressed_rtf", "rtf_compressed"):
        try:
            value = getattr(message, attr, None)
            if value:
                return value
        except Exception:
            pass
    return b""


def _remove_balanced_groups(text: str, start_literal: str) -> str:
    """Remove all brace-balanced groups beginning with start_literal."""
    out = []
    i = 0
    n = len(text)
    while i < n:
        idx = text.find(start_literal, i)
        if idx == -1:
            out.append(text[i:])
            break
        out.append(text[i:idx])
        depth = 0
        j = idx
        while j < n:
            c = text[j]
            if c == '\\' and j + 1 < n:
                j += 2
                continue
            if c == '{':
                depth += 1
            elif c == '}':
                depth -= 1
                if depth == 0:
                    j += 1
                    break
            j += 1
        i = j
    return ''.join(out)


def _strip_rtf_header_groups(rtf_text: str) -> str:
    """Strip RTF structural/destination groups that should never appear as text."""
    for pat in (
        '{\\*\\',           # all {\*\keyword ...} destination groups (incl. \*\generator)
        r'{\fonttbl',
        r'{\colortbl',
        r'{\stylesheet',
        r'{\info',
        r'{\themedata',
        r'{\colorschememapping',
        r'{\listtable',
        r'{\listoverridetable',
        r'{\rsidtbl',
        r'{\latentstyles',
        r'{\datastore',
        r'{\xmlnstbl',
    ):
        rtf_text = _remove_balanced_groups(rtf_text, pat)
    return rtf_text


def _parse_color_table(rtf_text: str) -> dict:
    """Return {1: '#rrggbb', ...} from {\\colortbl...}. Must be called before header stripping."""
    start = rtf_text.find('{\\colortbl')
    if start == -1:
        return {}
    depth, j, n = 0, start, len(rtf_text)
    while j < n:
        c = rtf_text[j]
        if c == '\\' and j + 1 < n:
            j += 2
            continue
        if c == '{':
            depth += 1
        elif c == '}':
            depth -= 1
            if depth == 0:
                inner = rtf_text[start + 10: j]   # strip '{\colortbl'
                break
        j += 1
    else:
        return {}
    colors = {}
    idx = 1
    for entry in inner.split(';'):
        r_m = re.search(r'\\red(\d+)', entry)
        g_m = re.search(r'\\green(\d+)', entry)
        b_m = re.search(r'\\blue(\d+)', entry)
        if r_m and g_m and b_m:
            colors[idx] = '#{:02x}{:02x}{:02x}'.format(
                int(r_m.group(1)), int(g_m.group(1)), int(b_m.group(1)))
        idx += 1
    return colors


def _rtf_body_to_html(rtf_text: str, color_table: dict, token_to_img_html: dict,
                      ole_img_tags: list = None) -> str:
    """
    Convert RTF body (header groups already stripped) to HTML.
    Preserves: color (\\cfN), bold (\\b), italic (\\i), underline (\\ul),
               paragraphs (\\par), line breaks (\\line), tabs (\\tab),
               hex escapes (\\'XX), unicode (\\uN).
    Group scoping: { pushes formatting state, } restores it.
    \\objattph consumes the next entry from ole_img_tags inline (not appended at end).
    """
    _NULL = {'color': None, 'bold': False, 'italic': False, 'ul': False}

    def close_fmt(st, out):
        if st['color']: out.append('</span>')
        if st['ul']:    out.append('</u>')
        if st['italic']: out.append('</em>')
        if st['bold']:  out.append('</strong>')

    def open_fmt(st, out):
        if st['bold']:   out.append('<strong>')
        if st['italic']: out.append('<em>')
        if st['ul']:     out.append('<u>')
        if st['color']:  out.append(f'<span style="color:{st["color"]}">')

    def transition(old, new_props, out):
        new = {**old, **new_props}
        if old == new:
            return old
        close_fmt(old, out)
        open_fmt(new, out)
        return new

    out = ['<p>']
    stack = [dict(_NULL)]   # saved states at each { level
    cur = dict(_NULL)
    i, n = 0, len(rtf_text)
    ole_list = ole_img_tags or []
    ole_idx  = 0

    while i < n:
        c = rtf_text[i]

        if c == '{':
            stack.append(dict(cur))
            i += 1

        elif c == '}':
            saved = stack.pop() if len(stack) > 1 else dict(_NULL)
            if cur != saved:
                close_fmt(cur, out)
                open_fmt(saved, out)
                cur = dict(saved)
            i += 1

        elif c == '\\':
            if i + 1 >= n:
                i += 1
                continue
            nc = rtf_text[i + 1]

            if nc in ('{', '}', '\\'):
                out.append(_html.escape(nc))
                i += 2
                continue

            if nc == "'":                           # \'XX hex char
                try:
                    ch = bytes.fromhex(rtf_text[i+2:i+4]).decode('cp1252', errors='replace')
                    out.append(_html.escape(ch))
                except Exception:
                    pass
                i += 4
                continue

            if nc == '~':                           # non-breaking space
                out.append('&nbsp;')
                i += 2
                continue

            if nc == '*':                           # destination marker (should be stripped)
                i += 2
                continue

            if nc in ('-', '_'):                    # optional / non-breaking hyphen
                if nc == '_': out.append('\u2011')
                i += 2
                continue

            if nc == 'u' and i + 2 < n and (rtf_text[i+2].isdigit() or rtf_text[i+2] == '-'):
                m = re.match(r'u(-?\d+) ?', rtf_text[i+1:])
                if m:
                    num = int(m.group(1))
                    if num < 0: num += 65536
                    out.append(_html.escape(chr(num)))
                    j = i + 1 + len(m.group(0))
                    # skip one fallback char (\'XX or plain)
                    if j < n and rtf_text[j] == '\\' and j+1 < n and rtf_text[j+1] == "'":
                        j += 4
                    i = j
                    continue

            if nc.isalpha():
                m = re.match(r'([a-zA-Z]+)(-?\d*) ?', rtf_text[i+1:])
                if not m:
                    i += 2
                    continue
                word, num_s = m.group(1), m.group(2)
                num = int(num_s) if num_s else None
                i += 1 + len(m.group(0))

                if word == 'par':
                    close_fmt(cur, out)
                    out.append('</p>\n<p>')
                    open_fmt(cur, out)
                elif word == 'line':
                    out.append('<br>')
                elif word == 'tab':
                    out.append('&nbsp;&nbsp;&nbsp;&nbsp;')
                elif word == 'plain':
                    cur = transition(cur, dict(_NULL), out)
                elif word == 'cf':
                    clr = color_table.get(num) if (num is not None and num > 0) else None
                    cur = transition(cur, {'color': clr}, out)
                elif word == 'b':
                    cur = transition(cur, {'bold': num is None or num != 0}, out)
                elif word == 'i':
                    cur = transition(cur, {'italic': num is None or num != 0}, out)
                elif word in ('ul', 'uld', 'uldb', 'ulwave', 'uldash'):
                    cur = transition(cur, {'ul': num is None or num != 0}, out)
                elif word == 'ulnone':
                    cur = transition(cur, {'ul': False}, out)
                elif word == 'objattph':
                    # OLE object placeholder — consume next slot regardless
                    if ole_idx < len(ole_list):
                        entry = ole_list[ole_idx]
                        ole_idx += 1
                        if entry is not None:          # image slot
                            mime, img_bytes = entry
                            data_uri = _make_data_uri(mime, img_bytes)
                            close_fmt(cur, out)
                            out.append(f'</p><p><img src="{data_uri}" alt="inline image"></p><p>')
                            open_fmt(cur, out)
                        # else: non-image OLE slot — advance counter, emit nothing
                # pard, ltrch, rtlch, lang, fs, f, cs, etc. — silently ignored
                continue

            i += 2  # unknown control symbol

        elif c in ('\r', '\n'):
            i += 1

        else:
            out.append(_html.escape(c))
            i += 1

    close_fmt(cur, out)
    out.append('</p>')

    html = ''.join(out)
    # Replace image tokens (they passed through _html.escape unchanged)
    for token, img_html in token_to_img_html.items():
        html = html.replace(token, f'</p>{img_html}<p>')
    # Clean empty paragraphs
    html = re.sub(r'<p>\s*</p>', '', html)
    return html


def _rtf_to_text_basic(rtf_text: str) -> str:
    text = _strip_rtf_header_groups(rtf_text)
    text = re.sub(r'\\par[d]? ?', '\n', text)
    text = re.sub(r'\\line ?', '\n', text)
    text = re.sub(r'\\tab ?', '\t', text)
    def _hex_char(m):
        try:
            return bytes.fromhex(m.group(1)).decode("cp1252", errors="replace")
        except Exception:
            return ""
    text = re.sub(r"\\'([0-9a-fA-F]{2})", _hex_char, text)
    def _unicode_char(m):
        try:
            n = int(m.group(1))
            if n < 0:
                n += 65536
            return chr(n)
        except Exception:
            return ""
    text = re.sub(r'\\u(-?\d+)\??', _unicode_char, text)
    text = re.sub(r'\\[a-zA-Z]+-?\d* ?', '', text)
    text = text.replace(r'\{', '{').replace(r'\}', '}').replace(r'\\', '\\')
    text = text.replace('{', '').replace('}', '')
    text = re.sub(r'\n{3,}', '\n\n', text)
    text = re.sub(r'[ \t]+\n', '\n', text)
    text = re.sub(r'[ \t]{2,}', ' ', text)
    return text.strip()


def _extract_pict_groups(rtf_text: str):
    results = []
    i = 0
    n = len(rtf_text)
    while i < n:
        start = rtf_text.find(r'{\pict', i)
        if start == -1:
            break
        depth = 0
        j = start
        while j < n:
            ch = rtf_text[j]
            if ch == '{':
                depth += 1
            elif ch == '}':
                depth -= 1
                if depth == 0:
                    j += 1
                    break
            j += 1
        group = rtf_text[start:j]
        mime, data = _decode_pict_group(group)
        if data:
            results.append({"start": start, "end": j, "group": group,
                            "mime": mime, "bytes": data})
        i = j if j > start else start + 6
    return results


def _decode_pict_group(group: str):
    lower = group.lower()
    if r'\pngblip' in lower:
        mime = 'image/png'
    elif r'\jpegblip' in lower or r'\jpgblip' in lower:
        mime = 'image/jpeg'
    elif r'\dibitmap' in lower or r'\wbitmap' in lower:
        mime = 'image/bmp'
    else:
        mime = 'application/octet-stream'
    hex_chunks = re.findall(r'(?<!\\)([0-9a-fA-F]{2,})', group)
    hex_data = ''.join(hex_chunks)
    if len(hex_data) % 2 == 1:
        hex_data = hex_data[:-1]
    if not hex_data:
        return mime, b""
    try:
        data = binascii.unhexlify(hex_data)
    except Exception:
        return mime, b""
    if mime == 'image/bmp':
        mime, data = _bmp_to_png(_wrap_dib_as_bmp(data))
    return mime, data


def _bmp_to_png(bmp_data: bytes) -> tuple:
    """Convert BMP bytes to PNG. Returns (mime, bytes) — falls back to image/bmp if Pillow fails."""
    try:
        from PIL import Image
        with Image.open(BytesIO(bmp_data)) as im:
            out = BytesIO()
            im.save(out, format="PNG")
            return "image/png", out.getvalue()
    except Exception:
        return "image/bmp", bmp_data


def _wrap_dib_as_bmp(data: bytes) -> bytes:
    if len(data) >= 2 and data[:2] == b'BM':
        return data
    if len(data) < 40:
        return data
    dib_header_size = int.from_bytes(data[0:4], 'little', signed=False)
    if dib_header_size < 40:
        return data
    try:
        bpp = int.from_bytes(data[14:16], 'little', signed=False)
    except Exception:
        bpp = 24
    colors_used = 0
    if len(data) >= 36:
        colors_used = int.from_bytes(data[32:36], 'little', signed=False)
    palette_colors = (colors_used or (1 << bpp)) if bpp <= 8 else 0
    bfOffBits = 14 + dib_header_size + palette_colors * 4
    bfSize = 14 + len(data)
    bmp_header = b'BM' + struct.pack('<IHHI', bfSize, 0, 0, bfOffBits)
    return bmp_header + data


def _make_data_uri(mime: str, data: bytes) -> str:
    return f"data:{mime};base64,{base64.b64encode(data).decode('ascii')}"


def _replace_pict_groups_with_tokens(rtf_text: str, picts):
    if not picts:
        return rtf_text, []
    out, tokens, last = [], [], 0
    for idx, pict in enumerate(picts):
        token = f"[[[RTF_INLINE_IMAGE_{idx}]]]"
        out.append(rtf_text[last:pict["start"]])
        out.append(token)
        tokens.append(token)
        last = pict["end"]
    out.append(rtf_text[last:])
    return ''.join(out), tokens


def _text_with_tokens_to_html(text: str, token_to_img_html: dict) -> str:
    escaped = _html.escape(text)
    parts = escaped.split('\n')
    html_lines = []
    for line in parts:
        stripped = line.strip()
        if stripped.startswith('[[[RTF_INLINE_IMAGE_') and stripped.endswith(']]]'):
            html_lines.append(token_to_img_html.get(stripped, ''))
        elif stripped:
            html_lines.append(f"<p>{stripped}</p>")
        else:
            html_lines.append("")
    html_out = '\n'.join(html_lines)
    for token, img_html in token_to_img_html.items():
        html_out = html_out.replace(_html.escape(token), img_html)
    return f"<html>\n  <body>\n    {html_out}\n  </body>\n</html>"


def _extract_ole_contents_bmp(att) -> bytes:
    """Phase 3: extract CONTENTS stream from an ATTACH_OLE (method=6) attachment."""
    if OleFileIO is None:
        return b""
    try:
        att_size = att.size or 0
        if att_size <= 0:
            return b""
        raw = att.read_buffer(att_size)
        if not raw:
            return b""
        bio = BytesIO(raw)
        if not _isOleFile(bio):
            return b""
        bio.seek(0)
        with OleFileIO(bio) as ole:
            if not ole.exists("CONTENTS"):
                return b""
            data = ole.openstream("CONTENTS").read()
            if data and data[:2] == b"BM":
                return data
    except Exception:
        pass
    return b""


def _extract_ole_images_list(message) -> list:
    """
    Return slot-preserving list aligned with \\objattph placeholders.
    Each OLE compound-doc attachment gets one slot:
      (mime, bytes)  — if CONTENTS stream is a BMP image
      None           — if it's a non-image OLE (e.g. embedded Word/Excel)
    Non-OLE attachments are skipped entirely (no slot).
    """
    images = []
    try:
        for att in message.attachments:
            try:
                att_size = att.size or 0
                if att_size <= 0:
                    continue
                raw = att.read_buffer(att_size)
                if not raw or raw[:8] != _OLE_MAGIC:
                    continue          # plain file attachment — no slot
                # OLE compound doc → always add a slot
                bmp = b""
                if OleFileIO is not None and _isOleFile is not None:
                    try:
                        bio = BytesIO(raw)
                        if _isOleFile(bio):
                            bio.seek(0)
                            with OleFileIO(bio) as ole:
                                if ole.exists("CONTENTS"):
                                    data = ole.openstream("CONTENTS").read()
                                    if data and data[:2] == b"BM":
                                        bmp = data
                    except Exception:
                        pass
                images.append(_bmp_to_png(bmp) if bmp else None)
            except Exception:
                images.append(None)   # error reading att — preserve slot
    except Exception:
        pass
    return images


def render_rtf_to_html_with_images(message, rtf_raw) -> str:
    """
    Phase 2 + 3:
      - Phase 2: extract \\pict images from RTF as data: URIs
      - Phase 3: extract OLE ATTACH_OLE (method=6) CONTENTS BMP from attachments
    Falls back to plain-text HTML if RTF is unrecognised.
    """
    if not rtf_raw:
        return ""
    if isinstance(rtf_raw, bytes):
        try:
            rtf_text = rtf_raw.decode("latin-1", errors="replace")
        except Exception:
            return ""
    else:
        rtf_text = str(rtf_raw)
    if r'{\rtf' not in rtf_text.lower():
        return ""
    # Extract color table BEFORE stripping header groups
    color_table = _parse_color_table(rtf_text)
    picts = _extract_pict_groups(rtf_text)
    tokenized_rtf, _tokens = _replace_pict_groups_with_tokens(rtf_text, picts)
    tokenized_rtf = _strip_rtf_header_groups(tokenized_rtf)
    token_to_img_html = {
        f"[[[RTF_INLINE_IMAGE_{idx}]]]":
            f'<p><img src="{_make_data_uri(p["mime"], p["bytes"])}" alt="inline image"></p>'
        for idx, p in enumerate(picts)
    }
    ole_images = _extract_ole_images_list(message)
    return _rtf_body_to_html(tokenized_rtf, color_table, token_to_img_html, ole_images)


_PR_ATTACH_CONTENT_ID = 0x3712   # entry_type is prop ID only (no type suffix)


def get_attachment_cid(att):
    """Read PR_ATTACH_CONTENT_ID (0x3712) from pypff attachment record_sets."""
    # Layer 1: direct attribute (future-proof)
    for attr in ("content_id", "attach_content_id", "mime_content_id"):
        try:
            v = getattr(att, attr, None)
            if v:
                return str(v).strip().strip("<>")
        except Exception:
            pass
    # Layer 2: scan MAPI record_sets for property 0x3712
    try:
        for rs in att.record_sets:
            for j in range(rs.number_of_entries):
                entry = rs.get_entry(j)
                if getattr(entry, "entry_type", None) == _PR_ATTACH_CONTENT_ID:
                    try:
                        v = entry.data_as_string
                    except Exception:
                        try:
                            v = entry.get_data_as_string()
                        except Exception:
                            v = None
                    if v:
                        return str(v).strip().strip("<>")
    except Exception:
        pass
    return None


def build_mime_message(message) -> email.message.Message:
    """Build a MIME message with HTML/CID inline images, RTF fallback, and regular attachments."""

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

    # ── Extract RTF and decide whether to use it ─────────────────────────────
    rtf_raw = get_message_rtf_data(message)
    rendered_rtf_html = ""
    if should_try_rtf_fallback(html, plain, rtf_raw):
        try:
            rendered_rtf_html = render_rtf_to_html_with_images(message, rtf_raw) or ""
        except Exception:
            rendered_rtf_html = ""

    # ── Final body choice ────────────────────────────────────────────────────
    if html and not looks_like_ole_placeholder_only(html):
        chosen_html = html
    elif rendered_rtf_html:
        chosen_html = rendered_rtf_html
    else:
        chosen_html = ""
    chosen_plain = plain or ""

    # ── Prepare CID set from chosen HTML ─────────────────────────────────────
    cids = extract_cids(chosen_html)
    cids_lower = {c.lower() for c in cids}

    # ── Extract attachments — split inline vs regular ─────────────────────────
    inline_parts  = []
    regular_parts = []
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
                # Skip OLE compound-doc attachments already inlined as images by Phase 3
                if (chosen_html and chosen_html == rendered_rtf_html
                        and OleFileIO is not None and _isOleFile is not None
                        and att_data[:8] == _OLE_MAGIC):
                    try:
                        _bio = BytesIO(att_data)
                        with OleFileIO(_bio) as _ole:
                            if _ole.exists("CONTENTS"):
                                _cont = _ole.openstream("CONTENTS").read()
                                if _cont and _cont[:2] == b"BM":
                                    continue  # already inlined — suppress duplicate
                    except Exception:
                        pass

                cid = get_attachment_cid(att)
                cid_norm = cid.strip().strip("<>").strip() if cid else ""
                is_inline = bool(cid_norm and cid_norm.lower() in cids_lower)
                if is_inline:
                    part["Content-ID"] = f"<{cid_norm}>"
                    part.add_header("Content-Disposition", "inline", filename=att_name)
                    inline_parts.append(part)
                else:
                    part.add_header("Content-Disposition", "attachment", filename=att_name)
                    regular_parts.append(part)
            except Exception:
                pass
    except Exception:
        pass

    # ── Build body MIME part ──────────────────────────────────────────────────
    if chosen_html:
        html_part = email.mime.text.MIMEText(chosen_html, "html", "utf-8")
        if inline_parts:
            related_part = email.mime.multipart.MIMEMultipart("related")
            related_part.attach(html_part)
            for p in inline_parts:
                related_part.attach(p)
            html_container = related_part
        else:
            html_container = html_part
        if chosen_plain:
            body_part = email.mime.multipart.MIMEMultipart("alternative")
            body_part.attach(email.mime.text.MIMEText(chosen_plain, "plain", "utf-8"))
            body_part.attach(html_container)
        else:
            body_part = html_container
    elif chosen_plain:
        body_part = email.mime.text.MIMEText(chosen_plain, "plain", "utf-8")
    else:
        body_part = email.mime.text.MIMEText("", "plain", "utf-8")

    # ── Wrap in multipart/mixed only if regular attachments exist ─────────────
    if regular_parts:
        msg = email.mime.multipart.MIMEMultipart("mixed")
        msg.attach(body_part)
        for part in regular_parts:
            msg.attach(part)
    else:
        msg = body_part

    # ── Fallback date from PST timestamp properties ───────────────────────────
    if "Date" not in header_data:
        for attr in ("delivery_time", "client_submit_time", "creation_time"):
            try:
                dt = getattr(message, attr, None)
                if dt:
                    if dt.tzinfo is None:
                        dt = dt.replace(tzinfo=datetime.timezone.utc)
                    header_data["Date"] = email.utils.format_datetime(dt)
                    break
            except Exception:
                pass

    # ── Fallback From/To from PST sender/recipient properties ─────────────────
    if "From" not in header_data:
        try:
            name  = getattr(message, "sender_name", None) or ""
            email_addr = getattr(message, "sender_email_address", None) or ""
            if email_addr:
                header_data["From"] = email.utils.formataddr((name, email_addr))
            elif name:
                header_data["From"] = name
        except Exception:
            pass

    if "To" not in header_data:
        try:
            recips = []
            for r in (message.recipients or []):
                rname  = getattr(r, "display_name", "") or ""
                raddr  = getattr(r, "email_address", "") or ""
                if raddr:
                    recips.append(email.utils.formataddr((rname, raddr)))
                elif rname:
                    recips.append(rname)
            if recips:
                header_data["To"] = ", ".join(recips)
        except Exception:
            pass

    # ── Copy headers ──────────────────────────────────────────────────────────
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
                    mbox_msg = mailbox.mboxMessage(msg)
                    # Stamp From_ line with original email date (not conversion date)
                    date_str = msg.get("Date", "")
                    if date_str:
                        try:
                            t = email.utils.parsedate_tz(date_str)
                            if t:
                                ts = email.utils.mktime_tz(t)
                                mbox_msg.set_from("MAILER-DAEMON", _time.gmtime(ts))
                        except Exception:
                            pass
                    mbox.add(mbox_msg)
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

    def _fallback_all_messages(reason: str):
        nonlocal total
        print(f"  {reason} → fallback to archive.messages()", flush=True)
        try:
            all_msgs = list(archive.messages())
            count = write_messages(all_msgs, sbd_dir / "Inbox")
            print(f"  Inbox (all): {count} messages", flush=True)
            total = count
        except Exception as e:
            print(f"  FAILED fallback: {e}", flush=True)
            _log_failure(e)

    def _log_failure(e):
        try:
            with FAILED_LOG.open("a", encoding="utf-8") as lf:
                lf.write(f"{pst_path} | {e}\n")
        except Exception:
            pass

    try:
        with PffArchive(str(pst_path)) as archive:
            # Detect broken or ANSI root → fallback
            try:
                root = archive._data.get_root_folder()
                is_bad_root = (root is None or not hasattr(root, "sub_folders"))
            except Exception:
                is_bad_root = True

            if is_bad_root:
                _fallback_all_messages("Broken/ANSI PST")
            else:
                try:
                    folders = archive.folders() or []
                except Exception:
                    folders = []

                for folder in folders:
                    if not folder:
                        continue
                    try:
                        name = folder.name or ""
                    except Exception:
                        continue
                    if not name or name.lower() in SKIP:
                        continue
                    try:
                        msgs_iter = getattr(folder, "sub_messages", None)
                        if not msgs_iter:
                            continue
                        msgs = list(msgs_iter)
                    except Exception as e:
                        print(f"  Folder traversal failed ({name}): {e}", flush=True)
                        continue
                    if not msgs:
                        continue
                    count = write_messages(msgs, sbd_dir / sanitize(name))
                    if count:
                        print(f"  {name}: {count} messages", flush=True)
                        total += count

    except Exception as e:
        print(f"  FAILED: {e}", flush=True)
        _log_failure(e)
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

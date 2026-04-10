# Pst2Thunder

Convert Microsoft Outlook PST archives to Thunderbird Local Folders (MBOX format),
with duplicate detection, attachment preservation, and chronological folder sorting.

## Requirements

- Python 3.11
- `libratom` library:  `py -3.11 -m pip install libratom`

## Folder Structure

```
E:\PST\              ← master archive of all .pst files
E:\TB_Mail_v2\       ← Thunderbird Local Folders output
E:\claude\Pst2Thunder\  ← these scripts
```

Thunderbird must point its Local Folders to `E:\TB_Mail_v2\`
(set in `prefs.js`: `mail.server.server3.directory`).

---

## Scripts

### 1. `pst_to_tb_v2.py` — Bulk convert all PSTs
Converts every `.pst` in `E:\PST\` that hasn't been converted yet.
Skips system folders (Calendar, Contacts, Tasks, etc.).
Outputs one MBOX file per mail folder under `E:\TB_Mail_v2\<stem>.sbd\`.

```bat
py -3.11 pst_to_tb_v2.py
```

---

### 2. `reconvert.py` — Reconvert with full MIME + attachments
Re-does conversion with proper HTML body support and attachment extraction
(PPT, XLSX, PDF, ZIP, etc. appear as email attachments in Thunderbird).

```bat
py -3.11 reconvert.py "Dec 2021" "Nov 2021"   # specific PSTs
py -3.11 reconvert.py --all                    # all PSTs
```

---

### 3. `add_pst.py` — Add a single new PST file
Checks for duplicates (filename + MD5 hash), copies to `E:\PST\`, converts.

```bat
py -3.11 add_pst.py "D:\Downloads\NewArchive.pst"
```

**Duplicate detection:**
- Filename match (case-insensitive) → abort
- MD5 hash match (catches renames of same file) → abort

---

### 4. `scan_new_drive.py` — Scan a drive/folder for new PSTs
Scans a source folder, identifies genuinely new PSTs (by name + hash),
copies and converts them.

```bat
py -3.11 scan_new_drive.py "G:\pst"            # scan and convert
py -3.11 scan_new_drive.py "G:\pst" --dry-run  # preview only
```

---

### 5. `sort_folders.py` — Sort Thunderbird folders chronologically
Prefixes folder names with `YYYY-MM` so Thunderbird's alphabetical sort
becomes chronological (2002 → 2003 → ... → 2024).
Undated misc folders get prefix `zz` and sort to the bottom.

```bat
py -3.11 sort_folders.py --dry-run   # preview
py -3.11 sort_folders.py             # apply (close Thunderbird first)
```

**Note:** `Mac` is treated as March (Malay month name).

---

### 6. `fix_failed_psts.py` — Recovery for encoding failures
Fixes PSTs that failed due to Chinese/non-ASCII folder names (e.g. Aug 2013.pst).
Forces UTF-8 stdout and uses per-folder iteration.

```bat
py -3.11 fix_failed_psts.py
```

---

### 7. `fix_old_format.py` — Recovery for ANSI-format PSTs
Copies pre-existing readpst MBOX output (from `E:\Mail_mbox\`) into
Thunderbird `.sbd` structure for old ANSI-format PSTs (pre-2003)
that libpff cannot traverse.

```bat
py -3.11 fix_old_format.py
```

---

### 8. `raw_to_tb_v2.py` — Convert readpst raw output to Thunderbird format
Maps the `<period>/<root>/<subfolder>/mbox` structure from readpst
into Thunderbird `.sbd` hierarchy.

```bat
py -3.11 raw_to_tb_v2.py
```

---

## Workflow: Adding New PSTs

### Single file:
```bat
py -3.11 add_pst.py "path\to\file.pst"
```

### Entire drive or folder:
```bat
py -3.11 scan_new_drive.py "G:\pst_new" --dry-run   # check first
py -3.11 scan_new_drive.py "G:\pst_new"              # then convert
```

### After conversion, re-sort folders:
```bat
py -3.11 sort_folders.py --dry-run
py -3.11 sort_folders.py
```

---

## PST Format Notes

| Format | Years | Handling |
|--------|-------|----------|
| Unicode (modern) | 2003+ | libratom folder-by-folder, full MIME + attachments |
| ANSI (old) | pre-2003 | `archive.messages()` flat dump → Inbox, or use readpst output |

## Known Issues

- ` Sep 2007.pst` — corrupted, no recoverable data
- `mailbox.pst` — corrupted (`libpff` name-to-id map error)
- `Feb 2020.pst`, `Dec 2023.pst` — 0-byte files, empty output

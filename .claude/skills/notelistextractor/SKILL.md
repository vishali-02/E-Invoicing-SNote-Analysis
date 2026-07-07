---
name: notelistextractor
description: User provides a parent SAP OSS note number or link. Skill opens it in SAP for Me, extracts all child/sequence notes in listed order, and writes them to an Excel file.
argument-hint: parent-note-number-or-url
---

## Overview

Given a parent SAP OSS note number or URL, this skill:
1. Opens the parent note in SAP for Me
2. Finds the **Sequence of Notes** section and extracts all child/sequence notes in order
3. Writes the ordered list to `SAP_Note_Sequence.xlsx`

---

## Step 1 — Open the parent note

If the user gave a full URL, open it directly. If they gave only a note number, construct the URL:

```
playwright-cli open https://me.sap.com/notes/<PARENT_NOTE_NUMBER>
playwright-cli snapshot
```

> If SAP for Me requires login, attach to an already-authenticated Edge session first:
> `playwright-cli attach --cdp=msedge`

---

## Step 2 — Locate the Sequence of Notes section

Take a snapshot and look for a tab or section labelled any of:
- **"Sequence of Notes"**
- **"Prerequisites"**
- **"Related Notes"**
- A numbered list of note links

If the section is inside a tab, click it:
```
playwright-cli click "getByText('Sequence of Notes')"
playwright-cli snapshot
```

---

## Step 3 — Extract child/sequence notes in order

Run this JS eval to pull all note numbers and their adjacent title text:
```
playwright-cli eval "(() => { const anchors = Array.from(document.querySelectorAll('a[href*=\"/notes/\"]')); const seen = new Set(); return anchors.filter(a => { const n = (a.textContent.trim().match(/\d{7,10}/) || [])[0]; if (!n || seen.has(n)) return false; seen.add(n); return true; }).map((a, i) => { const n = (a.textContent.trim().match(/\d{7,10}/) || [])[0]; const row = a.closest('tr,li,div[role],div[class]'); const title = row ? row.textContent.replace(n,'').trim().substring(0,120) : ''; return {seq: i+1, note: n, title}; }); })()"
```

If that returns an empty array, try a broader selector:
```
playwright-cli eval "Array.from(document.querySelectorAll('*')).filter(el => el.childElementCount===0 && /^\d{7,10}$/.test(el.textContent.trim())).map((el,i)=>({seq:i+1,note:el.textContent.trim(),title:el.closest('tr,li')?.textContent?.trim()?.substring(0,120)||''}))"
```

Record the full list:
| # | Note Number | Note Title |
|---|-------------|------------|
| 1 | 1234567     | …          |
| 2 | 1234568     | …          |

> If no sequence notes are found at all, record the parent note itself as the single row.

---

## Step 4 — Write the Excel file

Paste the collected note data into a Python script and run it:

```python
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment

# ---- populate this list from Step 3 ----
rows = [
    {"seq": 1, "note": "1234567", "title": "Example note title"},
    # … add all rows here
]
# ----------------------------------------

wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Note Sequence"

headers = ["#", "Note Number", "Note Title", "SAP for Me Link"]
header_fill = PatternFill("solid", fgColor="4472C4")

for col, h in enumerate(headers, 1):
    cell = ws.cell(row=1, column=col, value=h)
    cell.font = Font(bold=True, color="FFFFFF")
    cell.fill = header_fill
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

for i, row in enumerate(rows, 2):
    ws.cell(row=i, column=1, value=row["seq"])
    ws.cell(row=i, column=2, value=row["note"])
    ws.cell(row=i, column=3, value=row["title"])
    link = f"https://me.sap.com/notes/{row['note']}"
    cell = ws.cell(row=i, column=4, value=link)
    cell.hyperlink = link
    cell.font = Font(color="0563C1", underline="single")
    ws.row_dimensions[i].height = 20

col_widths = [5, 15, 70, 45]
for col, width in enumerate(col_widths, 1):
    ws.column_dimensions[ws.cell(row=1, column=col).column_letter].width = width

ws.freeze_panes = "A2"
wb.save("SAP_Note_Sequence.xlsx")
print("Saved: SAP_Note_Sequence.xlsx")
```

Run:
```
python note_sequence_writer.py
```

---

## Output

`SAP_Note_Sequence.xlsx` with four columns:

| Column | Name | Description |
|--------|------|-------------|
| A | # | Sequential order (1, 2, 3…) |
| B | Note Number | SAP OSS note number |
| C | Note Title | Title/description from SAP for Me |
| D | SAP for Me Link | Clickable hyperlink to `me.sap.com/notes/<number>` |

---

## Tips

- The parent note URL format is: `https://me.sap.com/notes/<NUMBER>`
- Sequence/child notes appear in the **"Sequence of Notes"** section; some notes list them under a **"Prerequisites"** or **"Related Notes"** tab instead.
- If a note is listed multiple times (e.g. as both a prerequisite and a sequence note), keep only the first occurrence to preserve order and avoid duplicates.
- If SAP for Me shows a login wall, run `playwright-cli attach --cdp=msedge` to reuse your existing Edge session.

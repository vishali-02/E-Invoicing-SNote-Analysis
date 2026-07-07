---
name: noteattachmentcheck
description: User provides the SAP_Note_Sequence.xlsx path. For each sequence note, skill checks SAP for Me for Attachments, Manual Activities, and Post Implement steps, then appends three new columns (Attachments, Manual Activities, Post Implement) with Yes/No values and colour coding.
argument-hint: path/to/SAP_Note_Sequence.xlsx
---

## Overview

Given the Excel file produced by previous skills, this skill:
1. Reads all note numbers from column B
2. For each note in SAP for Me:
   - Clicks the **Attachments** tab → `Yes` if any file is listed, `No` otherwise
   - Checks the **Description/Solution** section for a **Manual Activities** heading → `Yes` / `No`
   - Within the same section checks for a **Post Implement** heading → `Yes` / `No`
3. Appends three new columns at the end of the existing file:
   - Column H — **Attachments**
   - Column I — **Manual Activities**
   - Column J — **Post Implement**
4. Applies green/red conditional formatting on all three columns

---

## Current Excel Layout (input)

| A | B | C | D | E | F | G |
|---|---|---|---|---|---|---|
| # | Note Number | Component | Note Title | SAP for Me Link | S4 Core 109 Supported | Parent Note Number |

---

## Step 1 — Read note numbers from the Excel file

```python
import openpyxl

path = r"<EXCEL_FILE_PATH>"
wb = openpyxl.load_workbook(path)
ws = wb.active

notes = []
for row in ws.iter_rows(min_row=2, values_only=True):
    note_number = str(row[1]).strip() if row[1] else ""
    if note_number and note_number.isdigit():
        notes.append(note_number)

print(notes)
```

Run:
```
python read_notes.py
```

---

## Step 2 — Add the three new column headers

```python
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment

path = r"<EXCEL_FILE_PATH>"
wb = openpyxl.load_workbook(path)
ws = wb.active

header_fill = PatternFill("solid", fgColor="4472C4")
new_headers = ["Attachments", "Manual Activities", "Post Implement"]

start_col = ws.max_column + 1
for i, h in enumerate(new_headers):
    cell = ws.cell(row=1, column=start_col + i, value=h)
    cell.font = Font(bold=True, color="FFFFFF")
    cell.fill = header_fill
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    col_letter = cell.column_letter
    ws.column_dimensions[col_letter].width = 20

wb.save(path)
print("Headers added.")
```

Run:
```
python add_headers.py
```

---

## Step 3 — Attach to the authenticated Edge session (do once)

```
playwright-cli attach --cdp=msedge
```

---

## Step 4 — For each note, check Attachments, Manual Activities, and Post Implement

Repeat the sub-steps below for **every note number** from Step 1.

### 4a — Open the note

```
playwright-cli goto https://me.sap.com/notes/<NOTE_NUMBER>
playwright-cli snapshot
```

---

### 4b — Check Attachments

Click the Attachments tab:
```
playwright-cli click "getByText('Attachments')"
playwright-cli snapshot
```

Check whether any file is listed (looks for common attachment extensions):
```
playwright-cli eval "Array.from(document.querySelectorAll('*')).filter(el => el.childElementCount === 0 && /\.(bcs|zip|txt|pdf|xml|xlsx|doc|docx)/i.test(el.textContent.trim())).map(el => el.textContent.trim())"
```

- Array is **non-empty** → `Attachments = Yes`
- Array is **empty `[]`** → `Attachments = No`

If clicking "Attachments" fails (tab not found), try:
```
playwright-cli eval "Array.from(document.querySelectorAll('.sapMITHTextContent')).find(el => /attachments/i.test(el.textContent.trim()))?.closest('[role=tab]')?.id"
```
Then click that element by its ID.

---

### 4c — Check Manual Activities

Navigate back to the note description (click the **"Note"** or **"Description"** tab if needed):
```
playwright-cli click "getByText('Note')"
playwright-cli snapshot
```

Search for a heading that contains "Manual Activities", "Manual Steps", or "Manual Implementation":
```
playwright-cli eval "(() => { const h = Array.from(document.querySelectorAll('h1,h2,h3,h4,strong,b,span')).find(el => /manual\s*(activ|step|impl)/i.test(el.textContent.trim())); return h ? h.parentElement?.textContent?.trim()?.substring(0, 500) : 'NOT_FOUND'; })()"
```

- Returns text content → `Manual Activities = Yes`
- Returns `NOT_FOUND` or empty → `Manual Activities = No`

---

### 4d — Check Post Implement

On the same page (no need to navigate away), search for a heading that contains "Post Implement", "Post-Implementation", or "After Implementation":
```
playwright-cli eval "(() => { const h = Array.from(document.querySelectorAll('h1,h2,h3,h4,strong,b,span')).find(el => /post[\s\-]*(impl|install|activat)|after[\s\-]*impl/i.test(el.textContent.trim())); return h ? h.parentElement?.textContent?.trim()?.substring(0, 500) : 'NOT_FOUND'; })()"
```

- Returns text content → `Post Implement = Yes`
- Returns `NOT_FOUND` or empty → `Post Implement = No`

Record all three values for this note, then move to the next.

---

## Step 5 — Write results back to the Excel file

```python
import openpyxl

path = r"<EXCEL_FILE_PATH>"

# Fill this dict after completing Step 4 for all notes
# Format: { "note_number": {"attachments": "Yes"/"No", "manual": "Yes"/"No", "post": "Yes"/"No"} }
results = {
    "1234567": {"attachments": "Yes", "manual": "No",  "post": "No"},
    "1234568": {"attachments": "No",  "manual": "Yes", "post": "Yes"},
    # … add all notes here
}

wb = openpyxl.load_workbook(path)
ws = wb.active

# Locate the three new columns by header name
col_map = {}
for cell in ws[1]:
    if cell.value in ("Attachments", "Manual Activities", "Post Implement"):
        col_map[cell.value] = cell.column

for row in ws.iter_rows(min_row=2):
    note_number = str(row[1].value).strip() if row[1].value else ""
    if note_number not in results:
        continue
    r = results[note_number]
    if "Attachments" in col_map:
        ws.cell(row=row[0].row, column=col_map["Attachments"]).value = r["attachments"]
    if "Manual Activities" in col_map:
        ws.cell(row=row[0].row, column=col_map["Manual Activities"]).value = r["manual"]
    if "Post Implement" in col_map:
        ws.cell(row=row[0].row, column=col_map["Post Implement"]).value = r["post"]

wb.save(path)
print(f"Updated: {path}")
```

Run:
```
python write_attachment_results.py
```

---

## Step 6 — Apply conditional formatting (green = Yes, red = No)

```python
import openpyxl
from openpyxl.styles import PatternFill, Font
from openpyxl.formatting.rule import CellIsRule

path = r"<EXCEL_FILE_PATH>"
wb = openpyxl.load_workbook(path)
ws = wb.active

last_row = ws.max_row
green      = PatternFill("solid", fgColor="C6EFCE")
red        = PatternFill("solid", fgColor="FFC7CE")
green_font = Font(color="276221")
red_font   = Font(color="9C0006")

# Find columns for the three new headers
col_map = {}
for cell in ws[1]:
    if cell.value in ("Attachments", "Manual Activities", "Post Implement"):
        col_map[cell.value] = cell.column_letter

for col_letter in col_map.values():
    cell_range = f"{col_letter}2:{col_letter}{last_row}"
    ws.conditional_formatting.add(
        cell_range,
        CellIsRule(operator="equal", formula=['"Yes"'], fill=green, font=green_font)
    )
    ws.conditional_formatting.add(
        cell_range,
        CellIsRule(operator="equal", formula=['"No"'],  fill=red,   font=red_font)
    )

wb.save(path)
print("Formatting applied.")
```

Run:
```
python apply_formatting.py
```

---

## Final Excel Layout (output)

| Col | Name | Values |
|-----|------|--------|
| A | # | Row index |
| B | Note Number | SAP OSS note number |
| C | Component | SAP component (e.g. `FI-GL`) |
| D | Note Title | Title from SAP for Me |
| E | SAP for Me Link | Clickable hyperlink |
| F | S4 Core 109 Supported | `Yes` / `No` |
| G | Parent Note Number | Parent note number |
| H | **Attachments** | `Yes` / `No` (green / red) |
| I | **Manual Activities** | `Yes` / `No` (green / red) |
| J | **Post Implement** | `Yes` / `No` (green / red) |

---

## Tips

- Always click the **Attachments** tab before checking for files — the count in the tab label (e.g. `Attachments (3)`) also confirms files exist if the eval is slow.
- The **Manual Activities** and **Post Implement** sections are inside the note body (Description tab), not inside tabs — no extra navigation needed between 4c and 4d.
- If a note body uses collapsible sections, click to expand them before running the eval so hidden headings are visible in the DOM.
- If SAP for Me shows a login wall, run `playwright-cli attach --cdp=msedge` to reuse the existing authenticated Edge session.
- The Excel file is updated in-place — columns A–G from previous skills are untouched.

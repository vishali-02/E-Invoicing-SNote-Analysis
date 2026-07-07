---
name: noteenrichment
description: User provides the SAP_Note_Sequence.xlsx path. Skill reads each sequence note from column B, opens it in SAP for Me, extracts the Component value, inserts it as a new column C (adjacent to Note Number), and appends a Parent Note Number column at the end. Existing columns shift right accordingly.
argument-hint: path/to/SAP_Note_Sequence.xlsx
---

## Overview

Given the Excel file produced by previous skills, this skill:
1. Reads all note numbers from column B
2. Opens each note in SAP for Me and extracts the **Component** field (e.g. `FI-GL`, `SD-BIL`, `MM-IV`)
3. Inserts a new **Component** column at **column C** — immediately to the right of Note Number
4. Appends a **Parent Note Number** column at the end, filled with the same parent note number for every row
5. Saves the updated file in-place

---

## Resulting Excel Layout

After this skill runs the column order will be:

| Col | Name | Description |
|-----|------|-------------|
| A | # | Row index |
| B | Note Number | SAP OSS note number |
| C | Component | SAP component extracted from the note header (e.g. `FI-GL`, `SD-BIL`) |
| D | Note Title | Title from SAP for Me |
| E | SAP for Me Link | Clickable hyperlink |
| F | S4 Core 109 Supported | `Yes` / `No` |
| G | Parent Note Number | The parent note number supplied by the user |

---

## Step 1 — Ask for the parent note number

Before opening SAP, ask the user:
> "What is the parent note number for this analysis?"

Store the answer as `PARENT_NOTE`. This will be written into every row in column G.

---

## Step 2 — Read note numbers from the Excel file

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

## Step 3 — For each note, extract the Component from SAP for Me

Attach to the existing authenticated Edge session once:
```
playwright-cli attach --cdp=msedge
```

For **each note number** in the list:

### 3a — Open the note
```
playwright-cli goto https://me.sap.com/notes/<NOTE_NUMBER>
playwright-cli snapshot
```

### 3b — Extract the Component field

The Component appears in the note header area, typically next to a label like "Component", "Application Component", or "BC Area". Run:

```
playwright-cli eval "(() => { const labels = Array.from(document.querySelectorAll('*')).filter(el => el.childElementCount === 0 && /^component$/i.test(el.textContent.trim())); if (!labels.length) return 'NOT_FOUND'; const parent = labels[0].closest('tr,li,div[class]'); if (!parent) return 'NOT_FOUND'; const siblings = Array.from(parent.querySelectorAll('*')).filter(el => el.childElementCount === 0); const val = siblings.find(el => el !== labels[0] && el.textContent.trim().length > 0 && /[A-Z]{2,}-?[A-Z0-9]*/i.test(el.textContent.trim())); return val ? val.textContent.trim() : 'NOT_FOUND'; })()"
```

If the above returns `NOT_FOUND`, try the broader fallback:
```
playwright-cli eval "Array.from(document.querySelectorAll('.sapMObjStatusTitle, .sapMText, [class*=Header], [class*=header], [class*=field]')).filter(el => el.childElementCount===0 && /^[A-Z]{2,}(-[A-Z0-9]+){1,3}$/.test(el.textContent.trim())).map(el=>el.textContent.trim()).slice(0,3)"
```

Record the component value for this note (e.g. `FI-GL`, `SD-BIL`, `NOT_FOUND`).

---

## Step 4 — Insert Component column and add Parent Note Number column

After collecting all component values, run this Python script to update the Excel file:

```python
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.formatting.rule import CellIsRule

path = r"<EXCEL_FILE_PATH>"
PARENT_NOTE = "<PARENT_NOTE_NUMBER>"   # from Step 1

# component_map: { "note_number": "component_value" }
component_map = {
    "1234567": "FI-GL",
    "1234568": "SD-BIL",
    # … fill from Step 3
}

wb = openpyxl.load_workbook(path)
ws = wb.active

# ── Step 4a: insert a blank column at position C (column 3) ──
ws.insert_cols(3)

# ── Step 4b: write Component header ──
header_fill = PatternFill("solid", fgColor="4472C4")
hcell = ws.cell(row=1, column=3, value="Component")
hcell.font = Font(bold=True, color="FFFFFF")
hcell.fill = header_fill
hcell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
ws.column_dimensions["C"].width = 20

# ── Step 4c: fill Component values ──
for row in ws.iter_rows(min_row=2):
    note_cell = row[1]   # column B (unchanged)
    comp_cell = row[2]   # column C (newly inserted)
    note_number = str(note_cell.value).strip() if note_cell.value else ""
    if note_number in component_map:
        comp_cell.value = component_map[note_number]
        comp_cell.alignment = Alignment(horizontal="center")

# ── Step 4d: append Parent Note Number column at the end ──
last_col = ws.max_column + 1
phcell = ws.cell(row=1, column=last_col, value="Parent Note Number")
phcell.font = Font(bold=True, color="FFFFFF")
phcell.fill = header_fill
phcell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
col_letter = ws.cell(row=1, column=last_col).column_letter
ws.column_dimensions[col_letter].width = 22

for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
    note_cell = row[1]
    if note_cell.value:
        parent_cell = ws.cell(row=row[0].row, column=last_col)
        parent_cell.value = PARENT_NOTE
        parent_cell.alignment = Alignment(horizontal="center")

wb.save(path)
print(f"Updated: {path}")
```

Run:
```
python note_enrichment_writer.py
```

---

## Step 5 — Verify the final layout

Open the saved Excel file and confirm columns match:

| A | B | C | D | E | F | G |
|---|---|---|---|---|---|---|
| # | Note Number | Component | Note Title | SAP for Me Link | S4 Core 109 Supported | Parent Note Number |

---

## Tips

- The **Component** label in SAP for Me appears in the note header (top section of the page), not inside a tab — no tab click is needed.
- Common component formats: `FI-GL`, `SD-BIL-IV`, `MM-IV`, `BC-MID`. If the extracted value looks like a full path (e.g. `FI > General Ledger > ...`), take only the shortcode part before the first space or `>`.
- If `NOT_FOUND` is returned for a note, leave the Component cell blank — do not write `NOT_FOUND` into the Excel.
- If SAP for Me shows a login wall, run `playwright-cli attach --cdp=msedge` to reuse the existing authenticated Edge session.
- The Excel file is updated in-place — all existing columns are preserved; column C is inserted and column G is appended.

# E-Invoicing SAP Note Analysis — Claude Skills

This repository contains three Claude Code skills for automated analysis of SAP OSS notes and their prerequisite hierarchies. The skills work together as a pipeline, but can be invoked independently.

---

## Skill Overview

| Skill | Purpose | Input | Output |
|-------|---------|-------|--------|
| [sap-note-analyzer](#sap-note-analyzer) | Extract SAP note numbers from a Jira ticket or parent SAP note | Jira URL or parent note number | `.txt` file of note numbers |
| [oss-note-processor](#oss-note-processor) | Enrich note numbers with metadata from me.sap.com | `.txt` file of note numbers | `OSS_Notes_Report.xlsx` |
| [prereq-tree-walker](#prereq-tree-walker) | Walk the full SP109 prerequisite tree for each note | `.xlsx` or `.txt` of notes | `Prereq_Tree_Report.xlsx` |

### Pipeline Flow

```
Jira ticket / Parent SAP note
         |
         v
  sap-note-analyzer  -->  notes_to_process.txt
         |
         v
  oss-note-processor  -->  OSS_Notes_Report.xlsx
         |
         v
  prereq-tree-walker  -->  Prereq_Tree_Report.xlsx
```

> `oss-note-processor` automatically hands off to `prereq-tree-walker` when it completes.
> `sap-note-analyzer` automatically hands off to `oss-note-processor` when it completes.

---

## sap-note-analyzer

**Description:** Extracts SAP note numbers from a Jira ticket or a parent SAP note, applies an optional keyword filter, preserves source order, and feeds them into `oss-note-processor` for full enrichment.

**Trigger:** Use when the user provides a Jira link or a parent SAP note number and wants to process the SAP notes mentioned in it.

### What it does

1. Opens the Jira ticket or SAP note page in a browser session
2. Captures the full page content
3. Extracts all 7-digit SAP note numbers (optionally filtered by keyword)
4. Shows the user the extracted list for confirmation
5. Hands off to `oss-note-processor`

### Scripts

| Script | Purpose |
|--------|---------|
| `scripts/extract_notes.py` | Extracts 7-digit SAP note numbers from a text/HTML dump, filters by keyword context window, preserves order, deduplicates |
| `scripts/read_notes.py` | Reads note numbers from an Excel file (column B), returns JSON array |
| `scripts/setup_columns.py` | Ensures output columns exist in the active Excel sheet |
| `scripts/write_results.py` | Writes enriched data back into the Excel file from `note_state.json` |
| `scripts/update_state.py` | Persists one note analysis results into `note_state.json` |
| `scripts/reset_state.py` | Clears `note_state.json` for a fresh run |
| `scripts/parse_and_save.py` | Parses playwright eval output and saves to state |
| `scripts/process_note.py` | Processes a single note eval result and updates state |

### Usage

Invoke with `/sap-note-analyzer`. The skill will ask for:
- **Source**: a Jira ticket URL or a parent SAP note number (7 digits)
- **Keywords** (optional): words to narrow which notes are extracted

---

## oss-note-processor

**Description:** Reads SAP note numbers from a plain-text file, builds an Excel workbook, then visits each note on `me.sap.com` to enrich with metadata including SP109 support status, manual activities, and attachments.

**Trigger:** Use when the user has a text file of SAP note numbers and wants an enriched Excel output.

### What it does

1. Reads note numbers from the input `.txt` file
2. Creates `OSS_Notes_Report.xlsx` with all headers
3. For each note, visits `https://me.sap.com/notes/<NOTE>` and collects:
   - **Component** (e.g. `BC-FES-WGU`)
   - **Description** (note title)
   - **Link** (clickable hyperlink)
   - **SP109 Supported** (`Yes`/`No`) — does the note support S4CORE 109?
   - **Manual Activity (S4CORE 109)** (`Yes`/`No`)
   - **Activity Timing (Pre/Post)** (`Pre`/`Post`/`Pre & Post`/`None`)
   - **Attachment** (`Yes`/`No`)
4. Writes all results to Excel with conditional formatting (green = Yes, red = No)
5. Automatically invokes `prereq-tree-walker` when done

### Output columns

| Column | Content |
|--------|---------|
| S.No | Sequential row number |
| Note Number | SAP note number |
| Component | SAP component code (e.g. `BC-FES-WGU`) |
| Description | Note title/subject |
| Link | Clickable hyperlink to the note on me.sap.com |
| SP109 Supported | `Yes` / `No` / `N/A` |
| Manual Activity (S4CORE 109) | `Yes` / `No` / `N/A` |
| Activity Timing (Pre/Post) | `Pre` / `Post` / `Pre & Post` / `None` |
| Attachment | `Yes` / `No` / `N/A` |

### Default paths

| Item | Default |
|------|---------|
| Input text file | Ask user if not specified |
| Output Excel | `C:\Users\<user>\Downloads\OSS_Notes_Report.xlsx` |
| State file | `scripts/note_data.json` (runtime-only, gitignored) |

### Scripts

| Script | Purpose |
|--------|---------|
| `scripts/init_excel.py <txt> <excel>` | Parses note numbers from txt, creates Excel with all headers |
| `scripts/update_note.py` | Persists one note results into `note_data.json` |
| `scripts/write_results.py <excel>` | Reads `note_data.json` and writes all enriched data to Excel |
| `scripts/reset_state.py` | Deletes `note_data.json` for a clean run |

### Usage

Invoke with `/oss-note-processor`, or pass a text file path as an argument.

---

## prereq-tree-walker

**Description:** For each SAP note in an input file (or the `OSS_Notes_Report.xlsx` produced by `oss-note-processor`), walks the full prerequisite tree on `me.sap.com`. At every level it collects prerequisites that have S4CORE covering SP109, then recurses into those prerequisites until no further SP109 prerequisites exist.

**Trigger:** Use when the user provides a list of SAP notes (txt or xlsx) and wants the complete SP109-supported prerequisite chain — or when invoked automatically after `oss-note-processor` completes.

### What it does

1. Reads root note numbers from the input file
2. For each note, visits its Prerequisites section on `me.sap.com`
3. Collects all prerequisites that support S4CORE 109
4. For each SP109 prerequisite, collects: Link, Software Component, Attachment, Manual Activities, Pre/Post timing, SP109 Support Package
5. Recursively processes prerequisites until no further SP109 prereqs exist (BFS traversal)
6. Writes a two-section Excel report

### SP109 filter rule

A prerequisite counts as SP109-supported if the S4CORE row shows version `109`, or a range that includes 109 (e.g. `105-112`, `109-current`). Ranges that explicitly exclude 109 (e.g. `100-108`) do not count.

### Output Excel — Section 1: SP109 Prerequisite Tree (horizontal)

One column per depth level, colour-coded:

| Colour | Depth Level |
|--------|-------------|
| Grey | Level 0 — root notes |
| Blue | Level 1 — 1st prerequisites |
| Green | Level 2 |
| Yellow | Level 3 |
| Orange | Level 4+ |

Parent cells are merged vertically. When a branch ends, the next column shows **End** (red cell).

### Output Excel — Section 2: Prerequisite Detail (stacked per-level sub-tables)

| Column | Content |
|--------|---------|
| Sequence Note (Parent) | Parent note number |
| Prerequisite Note | Child note number |
| Link | Hyperlink to the prerequisite on me.sap.com |
| Software Component | SAP component code |
| Attachment | Yes / No |
| Manual Activities | Activity description text |
| Pre/Post Implementation | Pre / Post / Pre & Post |
| SP109 Support Package | Package identifier (e.g. `SAPK-10902INS4CORE`) |

### Default paths

| Item | Default |
|------|---------|
| Input (txt or xlsx) | `OSS_Notes_Report.xlsx` when invoked from `oss-note-processor`; ask user otherwise |
| Output Excel | `C:\Users\<user>\Downloads\Prereq_Tree_Report.xlsx` |
| State file | `scripts/tree_state.json` (runtime-only, gitignored) |

### Scripts

| Script | Purpose |
|--------|---------|
| `scripts/reset_state.py` | Deletes `tree_state.json` for a clean run |
| `scripts/read_input.py <input>` | Parses root notes from txt/xlsx; seeds `tree_state.json` |
| `scripts/update_note.py` | Saves one note prereqs, parent, level to `tree_state.json` |
| `scripts/update_prereq_detail.py` | Saves detail fields for one prerequisite note |
| `scripts/get_queue.py` | Returns `{queue, total_remaining}` from state |
| `scripts/write_results.py <excel>` | Builds the two-section Excel report from `tree_state.json` |

### State file schema (`tree_state.json`)

```json
{
  "root_notes": ["3499841"],
  "visited": ["3499841", "3400001"],
  "processed_notes": ["3499841", "3400001"],
  "notes": {
    "3499841": {
      "level": 0,
      "parent": "",
      "prereqs_109": ["3400001"],
      "prereq_details": {
        "3400001": {
          "link": "https://me.sap.com/notes/3400001",
          "component": "CA-GTF-CSC-EDO-PAP",
          "attachment": "No",
          "manual": "Execute program NOTE_3400001 after import",
          "timing": "Post",
          "sp109_sp": "SAPK-10902INS4CORE"
        }
      }
    }
  }
}
```

### Usage

Invoke with `/prereq-tree-walker`, or pass an input file path directly:

```
/prereq-tree-walker C:\path\to\OSS_Notes_Report.xlsx
```

---

## Requirements

- **Python 3.10+** with `openpyxl` installed: `pip install openpyxl`
- **Claude Code** with the `playwright-cli` tool enabled
- **SAP credentials** for `me.sap.com` (browser login required at first run)

---

## Repository Structure

```
.
+-- oss-note-processor/
|   +-- SKILL.md              # Skill definition and step-by-step instructions
|   +-- scripts/
|       +-- init_excel.py     # Create Excel workbook with headers
|       +-- update_note.py    # Save note results to note_data.json
|       +-- write_results.py  # Write note_data.json to Excel
|       +-- reset_state.py    # Clear note_data.json
+-- sap-note-analyzer/
|   +-- SKILL.md              # Skill definition and step-by-step instructions
|   +-- scripts/
|       +-- extract_notes.py  # Extract 7-digit note numbers from text/HTML
|       +-- read_notes.py     # Read notes from Excel (col B)
|       +-- setup_columns.py  # Ensure output columns exist in Excel
|       +-- write_results.py  # Write note_state.json to Excel
|       +-- update_state.py   # Persist one note results to note_state.json
|       +-- reset_state.py    # Clear note_state.json
|       +-- parse_and_save.py # Parse playwright eval output
|       +-- process_note.py   # Process a single note playwright result
+-- prereq-tree-walker/
|   +-- SKILL.md              # Skill definition and step-by-step instructions
|   +-- scripts/
|       +-- read_input.py           # Parse root notes from txt/xlsx
|       +-- update_note.py          # Save note prereqs to tree_state.json
|       +-- update_prereq_detail.py # Save prereq detail fields
|       +-- reset_state.py          # Clear tree_state.json
|       +-- get_queue.py            # Get remaining queue from state
|       +-- write_results.py        # Build two-section Excel report
+-- .gitignore
+-- README.md
```

---

## Installation

```bash
# Clone this repo
git clone https://github.com/vishali-02/E-Invoicing-SNote-Analysis.git

# Windows: copy each skill folder to your Claude skills directory
xcopy /E /I oss-note-processor %USERPROFILE%\.claude\skills\oss-note-processor
xcopy /E /I sap-note-analyzer %USERPROFILE%\.claude\skills\sap-note-analyzer
xcopy /E /I prereq-tree-walker %USERPROFILE%\.claude\skills\prereq-tree-walker

# Linux/Mac
cp -r oss-note-processor sap-note-analyzer prereq-tree-walker ~/.claude/skills/
```

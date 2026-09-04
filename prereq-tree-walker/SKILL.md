---
name: prereq-tree-walker
description: For each SAP note in an input file (or the OSS_Notes_Report.xlsx produced by oss-note-processor), walks the full prerequisite tree on me.sap.com. At every level it collects prerequisites that have S4CORE covering SP109, then recurses into those prerequisites — repeating until a note has no further SP109 prerequisites. For each node it also collects link, software component, attachments, manual activities (pre/post), and SP109 support package. Outputs an Excel report with a horizontal summary tree table and per-level detail sub-tables, both colour-coded by depth level. Use this skill when the user provides a list of SAP notes (txt or xlsx) and wants the complete SP109-supported prerequisite chain — or when invoked automatically after oss-note-processor completes.
argument-hint: <input_path>
allowed-tools: Bash(playwright-cli:*) Bash(python:*)
---

# Prerequisite Tree Walker

Reads root note numbers from a text or Excel file (including the `OSS_Notes_Report.xlsx` output of oss-note-processor), then walks the full prerequisite tree on **me.sap.com** — finding SP109 prerequisites at every level until no further SP109 prerequisites exist.

**SP109 filter:** a prerequisite counts if S4CORE appears with a version that includes 109 (e.g. `109`, `105 – 109`, `109 – current`).

**Termination per branch:** when a note yields no SP109 prerequisites, that branch stops. Other branches continue independently.

---

## Output Excel layout

Two sections on one sheet, colour-coded by depth (grey = level 0, blue = level 1, green = level 2, yellow = level 3, orange = level 4+):

**Section 1 — SP109 Prerequisite Tree (horizontal)**

One column per depth level:
| Sequence Note | 1st Pre Note | 2nd Pre Note | 3rd Pre Note | … |

- Each root note occupies a block of rows (one per leaf path). The root cell is **merged vertically** across all its rows.
- Each 1st pre note cell is merged vertically across its own children rows (when it has multiple children).
- When a branch has no further SP109 prereqs, the next column gets **"End"** (red cell, italic).
- Root notes with no SP109 prereqs at all: col A = root, col B = "End".
- Duplicate parent values in the same column are **merged** — identical adjacent cells in the same root-block become one cell.

**Section 2 — Prerequisite Detail (stacked per-level sub-tables)**

One sub-table per depth level, only for levels that had ≥1 note with SP109 prereqs. Notes with no prereqs are excluded entirely.

Each sub-table:
```
[ Title: "Prerequisite Detail — Level N  (Sequence Note → Nth Pre Note)" ]
[ Header: Sequence Note (Parent) | Prerequisite Note | Link | Software Component |
          Attachment | Manual Activities | Pre/Post Implementation | SP109 Support Package ]
[ Data rows — one per (parent, child) pair ]
```
Sub-tables stacked vertically with one blank row between them. Col A merged per parent group within each sub-table.

---

## Default paths

| Item | Default |
|------|---------|
| Input (txt or xlsx) | `C:\Users\C5404594\Downloads\OSS_Notes_Report.xlsx` when invoked from oss-note-processor; ask user if invoked standalone and not given |
| Output Excel | `C:\Users\C5404594\Downloads\Prereq_Tree_Report.xlsx` |
| Scripts | `C:\Users\C5404594\.claude\skills\prereq-tree-walker\scripts\` |
| State file | `C:\Users\C5404594\.claude\skills\prereq-tree-walker\scripts\tree_state.json` |

**When invoked from oss-note-processor:** the argument `$1` is already set to the oss output Excel path — use it directly without asking the user.

---

## Step 0 — Setup

The input file path is provided as the argument `$1`. If `$1` is not supplied, ask the user for the input file path.

Reset any previous run:

```bash
python "C:\Users\C5404594\.claude\skills\prereq-tree-walker\scripts\reset_state.py"
```

Seed root notes and get the initial queue:

```bash
python "C:\Users\C5404594\.claude\skills\prereq-tree-walker\scripts\read_input.py" "<input_path>"
```

Parse the printed JSON to get `root_notes`. Build the initial processing queue:

```
queue = [ { note: N, parent: "", level: 0 } for N in root_notes ]
```

Open the browser session once:

```bash
playwright-cli -s=ptw open "https://me.sap.com/notes/0000001"
```

If a login prompt appears, stop and tell the user. Wait for them to confirm login before continuing.

---

## Processing loop — BFS queue

Maintain a queue of items `{ note, parent, level }`. Start with all root notes at level 0 with no parent.

Work through the queue **front-to-back**. For each item:

---

### A. Navigate to the note

```bash
playwright-cli -s=ptw goto "https://me.sap.com/notes/<NOTE>"
```

Confirm the note title is visible. If 404 → `prereqs_109 = []`, call `update_note.py`, do **not** enqueue anything, move to next item.

---

### B. Find SP109 prerequisites

Check the Prerequisites section (inline grid or dedicated tab labelled "Prerequisites"). For each prerequisite row:

**Count as SP109-supported if:**
- S4CORE row has version `109`, OR range includes 109 (e.g. `105–112`, `109–current`), OR no version restriction shown.

**Exclude if:**
- Range explicitly excludes 109 (e.g. `100–108`, `110–current`).

If version details are not shown inline, navigate to each candidate prereq note and check its Software Component Versions section for an S4CORE row covering 109.

Collect all matching note numbers → `prereqs_109`.

---

### C. Save the note

```bash
python "C:\Users\C5404594\.claude\skills\prereq-tree-walker\scripts\update_note.py" \
  --note        "<NOTE>"       \
  --prereqs-109 "<csv_list>"   \
  --parent      "<PARENT>"     \
  --level       <LEVEL>
```

`--parent` is `""` for root notes; for all deeper notes it is the note number of the item being processed from the queue.
`--level` is 0 for root notes, 1 for their prereqs, 2 for the next level, and so on.

---

### D. Collect detail for each SP109 prerequisite

For every note in `prereqs_109`, navigate to it and collect:

```bash
playwright-cli -s=ptw goto "https://me.sap.com/notes/<PREREQ>"
```

**1. Link** — `https://me.sap.com/notes/<PREREQ>`

**2. Software Component** — value shown next to "Component:" in the note header.

**3. Attachment** — `"Yes"` if an Attachments tab exists with files, else `"No"`.

**4. Manual Activities** — text from the Manual Activities section. Join multiple with ` | `. Record `""` if absent.

**5. Pre/Post Implementation** — classify from activity heading/text:
- "Manual Pre-Implement." or keywords "before import / prior to" → **Pre**
- "Manual Post-Implement." or keywords "after import / after applying" → **Post**
- Both present → **Pre & Post**
- Undetermined → `""`

**6. SP109 Support Package** — from the Support Package section, find the S4CORE 109 row and extract the package identifier (e.g. `SAPK-10902INS4CORE`). Record `""` if not found.

```bash
python "C:\Users\C5404594\.claude\skills\prereq-tree-walker\scripts\update_prereq_detail.py" \
  --root-note   "<NOTE>"    \
  --prereq-note "<PREREQ>"  \
  --link        "https://me.sap.com/notes/<PREREQ>" \
  --component   "<component>"   \
  --attachment  "<Yes|No>"      \
  --manual      "<activity_text>" \
  --timing      "<Pre|Post|Pre & Post|>" \
  --sp109-sp    "<support_package>"
```

---

### E. Enqueue children

If `prereqs_109` is **non-empty**, add each prereq to the **back** of the queue:

```
for prereq in prereqs_109:
    if prereq not already visited:
        queue.append({ note: prereq, parent: <NOTE>, level: <LEVEL+1> })
```

If `prereqs_109` is **empty** → this branch terminates. Do not enqueue anything.

Skip enqueueing a note that is already in the visited set (prevents cycles).

Repeat from A with the next item in the queue.

---

## Step — Write Results to Excel

```bash
python "C:\Users\C5404594\.claude\skills\prereq-tree-walker\scripts\write_results.py" "C:\Users\C5404594\Downloads\Prereq_Tree_Report.xlsx"
```

---

## Step — Close browser and report

```bash
playwright-cli -s=ptw close
```

Report:
- Total notes processed (all levels).
- Total SP109 prerequisite links found.
- Deepest level reached.
- Path to the saved Excel file.

---

## Error handling

| Situation | Behaviour |
|-----------|-----------|
| Note page 404 | `prereqs_109 = []`. Do not enqueue. Continue. |
| Page load > 15 s | Retry once. If still fails, `prereqs_109 = []`. |
| Prerequisites section absent | `prereqs_109 = []`. Branch terminates. |
| No SP109 match | Branch terminates for that note. |
| Prereq detail page not found | Record all fields as `""`. Still enqueue if parent had SP109 prereqs. |
| Manual Activities absent | `manual = ""`, `timing = ""`. |
| Attachment absent/empty | `"No"`. |
| Cycle detected (note already visited) | Skip — do not re-enqueue. |
| Login wall | Pause, notify user, wait for confirmation. |

---

## Key navigation hints for me.sap.com

- Note URL: `https://me.sap.com/notes/<NOTE_NUMBER>`
- **Prerequisites**: inline grid in Description tab or dedicated "Prerequisites" tab. Rows show Software Component / From / To / Note number / Title / Component.
- **Software Component Versions**: table with S4CORE rows showing version ranges.
- **Manual Activities**: section under Correction tab or Description; heading "Manual Activities"; contains `VALID FOR` preformatted blocks.
- **Attachments**: "Attachments" tab in tab bar; count badge or file list.
- Version range inclusive check: `105–112` includes 109. `100–108` excludes 109. `109–current` includes 109.

---

## Scripts reference

All scripts: `C:\Users\C5404594\.claude\skills\prereq-tree-walker\scripts\`

| Script | Purpose |
|--------|---------|
| `reset_state.py` | Deletes `tree_state.json` for a clean run |
| `read_input.py <input>` | Parses root notes; seeds `tree_state.json`; prints `{root_notes, queue}` |
| `update_note.py --note … --prereqs-109 … --parent … --level …` | Saves one note's prereqs, parent, level; appends to `processed_notes` |
| `update_prereq_detail.py --root-note … --prereq-note … --link … --component … --attachment … --manual … --timing … --sp109-sp …` | Saves detail fields for one prerequisite note under its parent |
| `get_queue.py` | Returns `{queue, total_remaining}` |
| `write_results.py <excel>` | Builds the two-section Excel report from `tree_state.json` |

---

## State file schema (tree_state.json)

```json
{
  "root_notes":      ["3499841"],
  "visited":         ["3499841", "3400001"],
  "processed_notes": ["3499841", "3400001"],
  "notes": {
    "3499841": {
      "level":       0,
      "parent":      "",
      "prereqs_109": ["3400001"],
      "prereq_details": {
        "3400001": {
          "link":       "https://me.sap.com/notes/3400001",
          "component":  "CA-GTF-CSC-EDO-PAP",
          "attachment": "No",
          "manual":     "Execute program NOTE_3400001 after import",
          "timing":     "Post",
          "sp109_sp":   "SAPK-10902INS4CORE"
        }
      }
    },
    "3400001": {
      "level":       1,
      "parent":      "3499841",
      "prereqs_109": ["3300005"],
      "prereq_details": {
        "3300005": {
          "link":       "https://me.sap.com/notes/3300005",
          "component":  "CA-GTF-CSC-EDO",
          "attachment": "No",
          "manual":     "",
          "timing":     "",
          "sp109_sp":   "SAPK-10901INS4CORE"
        }
      }
    },
    "3300005": {
      "level":       2,
      "parent":      "3400001",
      "prereqs_109": []
    }
  }
}
```

`processed_notes` is the BFS traversal order. `write_results.py` renders groups in this order.
A note with empty `prereqs_109` terminates its branch.

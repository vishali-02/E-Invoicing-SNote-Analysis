---
name: oss-note-processor
description: Process a text file of SAP OSS note numbers into an enriched Excel report. Performs four steps: (1) list note numbers in Excel in order, (2) fetch component, description, and link for each note from me.sap.com, (3) check whether the note supports S4CORE 109 (SP109) and record the result, (4) check attachments and manual activities — classify pre/post-implementation and confirm they apply to S4CORE 109. Use this skill whenever the user has a text file of SAP note numbers and wants an enriched Excel output.
allowed-tools: Bash(playwright-cli:*) Bash(python:*)
---

# OSS Note Processor

Reads SAP note numbers from a plain-text file, builds an Excel workbook, then visits each note on **me.sap.com** to enrich four columns:

| Column | Content |
|--------|---------|
| Component | SAP component of the note (e.g. `BC-FES-WGU`) |
| Description | Short title / description from the note header |
| Link | Clickable hyperlink to `https://me.sap.com/notes/<NOTE>` |
| SP109 Supported | `Yes` / `No` — does the note support S4CORE 109? |
| Manual Activity (S4CORE 109) | `Yes` / `No` — is there a manual activity entry for S4CORE 109? |
| Activity Timing (Pre/Post) | `Pre` / `Post` / `Pre & Post` / `None` — when the manual activity runs |
| Attachment | `Yes` / `No` — does the note have an attachment? |

---

## Default file paths

| Item | Default path |
|------|-------------|
| Input text file | Ask user if not specified |
| Output Excel | `C:\Users\C5404594\Downloads\OSS_Notes_Report.xlsx` |
| Scripts dir | `C:\Users\C5404594\.claude\skills\oss-note-processor\scripts\` |
| State file | `C:\Users\C5404594\.claude\skills\oss-note-processor\scripts\note_data.json` |

---

## Step 0 — Identify files and initialize

Ask the user for the text file path if they have not provided one.  
Use `C:\Users\C5404594\Downloads\OSS_Notes_Report.xlsx` as the Excel output path unless the user specifies otherwise.

Reset any previous run's state:

```bash
python "C:\Users\C5404594\.claude\skills\oss-note-processor\scripts\reset_state.py"
```

Create the Excel file with all headers and populate column B with note numbers in order:

```bash
python "C:\Users\C5404594\.claude\skills\oss-note-processor\scripts\init_excel.py" "<txt_path>" "<excel_path>"
```

The script prints a line starting with `JSON:` followed by a JSON object — parse it to get the `notes` array (the ordered list of note numbers to process).

Open the browser session:

```bash
playwright-cli -s=oss open "https://me.sap.com/notes/3499841"
```

Wait for the page to load fully (check the snapshot shows a note title, not a login wall or error). If a login prompt appears, inform the user that they need to log in and then ask them to confirm before continuing.

---

## Step 1 — For each note: fetch Component, Description, and Link

Navigate to the note:

```bash
playwright-cli -s=oss goto "https://me.sap.com/notes/<NOTE>"
```

Take a snapshot and locate:

- **Component** — Look in the note header area for a label like "Primary Component", "Component", or "Application Component". The value is a dot-path or slash-path code such as `BC-FES-WGU` or `MM-PUR-GF`. Extract only the component code, not its description text.
- **Description** — The main note title/subject line shown prominently at the top of the note (e.g., "ABAP test cockpit — missing function"). Do **not** include the note number in the description.
- **Link** — Always `https://me.sap.com/notes/<NOTE>`.

If the page returns a "Note not found" or HTTP 404 error, set all three fields to `N/A` and continue.

---

## Step 2 — Check SP109 (S4CORE 109) support

Still on the same note page (from Step 1), find the **"Software Component Versions"** section (also labelled "Valid Releases" or "Software Components" on some notes).

Scan the table rows for a row where:
- The software component column contains `S4CORE` (case-insensitive), AND
- The version column contains `109`

Decision:
- Row found → `SP109 Supported = Yes`
- Row not found → `SP109 Supported = No`
- Section missing entirely → `SP109 Supported = No`
- Page not found → `SP109 Supported = N/A`

---

## Step 3 — Check Manual Activities for S4CORE 109

Click the **"Manual Activities"** tab or section on the note page.  
(Try clicking a tab labelled "Manual Activiti
es", "Manual Steps", or "Implementation Info". If no such tab exists, the note has no manual activity.)

If the tab does not exist → `Manual Activity = No`, `Activity Timing = None`.

If the tab exists:
1. Look at the entries listed. Each entry may be tagged with a software component version range (e.g., "Valid for S4CORE 109" or "From 105 to 109").
2. Filter to entries that apply to S4CORE version **109** (either an explicit "109" tag, or an entry with no version restriction that applies to all versions).
3. If no entries apply to 109 → `Manual Activity = No`, `Activity Timing = None`.
4. If entries exist that apply to 109:
   - `Manual Activity = Yes`
   - Check each entry's timing label:
     - Contains "Before", "Pre-", or "Pre " → counts as **Pre**
     - Contains "After", "Post-", or "Post " → counts as **Post**
     - No timing indicator or applies always → treat as **Post** (implementation step)
   - If only Pre entries: `Activity Timing = Pre`
   - If only Post entries: `Activity Timing = Post`
   - If both: `Activity Timing = Pre & Post`

---

## Step 4 — Check Attachments

Still on the same note page (or navigate back if needed), look for an **"Attachments"** section or tab.

- If the section/tab exists and contains one or more attached files → `Attachment = Yes`
- If the section is empty or absent → `Attachment = No`
- Page not found → `Attachment = N/A`

---

## Step 5 — Save and continue

After collecting all four data points for the current note, save:

```bash
python "C:\Users\C5404594\.claude\skills\oss-note-processor\scripts\update_note.py" \
  --note        "<NOTE>" \
  --component   "<COMPONENT>" \
  --description "<DESCRIPTION>" \
  --link        "https://me.sap.com/notes/<NOTE>" \
  --sp109       "<Yes|No|N/A|Error>" \
  --manual      "<Yes|No|N/A|Error>" \
  --timing      "<Pre|Post|Pre & Post|None|N/A|Error>" \
  --attachment  "<Yes|No|N/A|Error>"
```

Repeat Steps 1–5 for every note in the queue. Keep the browser session open across all notes.

---

## Step 6 — Write results to Excel

After all notes are processed:

```bash
python "C:\Users\C5404594\.claude\skills\oss-note-processor\scripts\write_results.py" "<excel_path>"
```

This script:
1. Reads `note_data.json`.
2. Updates Component, Description, Link, SP109 Supported, Manual Activity, Activity Timing, and Attachment columns for each note row in-place.
3. Applies green (Yes) / red (No) conditional formatting to the SP109 Supported column.
4. Writes hyperlinks in the Link column.
5. Saves the file.

---

## Step 7 — Close browser and report

```bash
playwright-cli -s=oss close
```

Tell the user:
- Total notes processed.
- How many are SP109-supported (Yes count).
- How many have manual activities for S4CORE 109.
  - Of those, how many are Pre-only, Post-only, Pre & Post.
- How many have attachments.
- Full path of the saved Excel file.

---

## Step 8 — Hand off to prereq-tree-walker

After completing Step 7, automatically invoke the **prereq-tree-walker** skill using the Skill tool.

Pass the Excel file produced in this run as the argument — no user prompt needed:

```
Skill("prereq-tree-walker", args="<excel_path>")
```

Where `<excel_path>` is the same path used in Step 6 (default `C:\Users\C5404594\Downloads\OSS_Notes_Report.xlsx`).

- The prereq-tree-walker skill reads the note numbers from the `Note Number` column (col B) of that Excel automatically.
- Its output Excel is saved to `C:\Users\C5404594\Downloads\Prereq_Tree_Report.xlsx`.
- Do **not** ask the user for a file path — use the path already known from this run.
- Do **not** reset the browser session between skills — the `ptw` session will be opened fresh by prereq-tree-walker.

---

## Error handling

| Situation | Action |
|-----------|--------|
| Note page returns 404 / "Note not found" | Set all fields to `N/A`, continue. |
| Page takes > 15 s to load | Retry once; if it still fails, set all fields to `Error`. |
| Software Component Versions section missing | `SP109 Supported = No` |
| Manual Activities tab missing | `Manual Activity = No`, `Activity Timing = None` |
| Attachments section missing | `Attachment = No` |
| Login wall on me.sap.com | Pause, tell the user to log in, wait for confirmation before continuing. |

---

## SAP Help Portal navigation hints

- Note URL: `https://me.sap.com/notes/<NOTE_NUMBER>`
- **Component** is usually near the top of the note header, next to "Primary Component" or "Application Component".
- **Software Component Versions** table: look for a section heading containing "Software Component" or "Valid Releases"; the table has rows with a component name column and a version column.
- **Manual Activities** tab: usually a horizontal tab near the top of the note body. May also appear as "Manual Steps" or inside "Implementation Info".
- **Attachments** tab: often the last tab in the note's tab bar. Look for a paperclip icon or "Attachments" label.
- When scanning for S4CORE 109: look for the text `S4CORE` with version `109` in the same row. Versions may be shown as ranges like `109 – current`; if 109 falls in range, it counts.
- When reading Manual Activity timing: look for keywords "Before", "After", "Pre", "Post" in the activity description or its timing label.

---

## Scripts reference

All scripts live in `C:\Users\C5404594\.claude\skills\oss-note-processor\scripts\`:

| Script | Purpose |
|--------|---------|
| `reset_state.py` | Deletes `note_data.json` for a clean run |
| `init_excel.py <txt> <excel>` | Parses note numbers from txt, creates Excel with all headers, populates S.No and Note Number columns |
| `update_note.py --note … --component … --description … --link … --sp109 … --manual … --timing … --attachment …` | Persists one note's results into `note_data.json` |
| `write_results.py <excel>` | Reads `note_data.json` and writes all enriched data to the Excel file |

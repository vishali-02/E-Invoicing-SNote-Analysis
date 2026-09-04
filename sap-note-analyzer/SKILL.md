
---
name: sap-note-analyzer
description: Extract SAP note numbers from a Jira ticket or a parent SAP note, filter by keyword, preserve source order, and feed them into oss-note-processor for enrichment. Use this skill when the user provides a Jira link or a parent SAP note number and wants to process the SAP notes mentioned in it.
allowed-tools: Bash(playwright-cli:*) Bash(python:*) Bash(python3:*)
---

# SAP Note Analyzer

Extracts SAP note numbers from a **Jira ticket** or a **parent SAP note**, applies an optional keyword filter, preserves the order the notes appear in the source, and hands the resulting list to the **oss-note-processor** skill for full enrichment.

---

## Inputs

Ask the user for the following before starting:

| Input | Required? | Description |
|-------|-----------|-------------|
| **Source** | Yes | A Jira ticket URL **or** a parent SAP note number (7 digits) |
| **Keyword(s)** | Optional | One or more words to narrow which notes are extracted. A note is included only if it appears near the keyword in the source text. Leave blank to extract all notes. |

---

## Step 0 — Open the source page

### If the user gave a Jira URL

```bash
playwright-cli -s=sna open "<JIRA_URL>"
```

Take a snapshot and confirm the page loaded (ticket title is visible). If a login wall appears, pause and ask the user to log in, then confirm before continuing.

### If the user gave a parent SAP note number

```bash
playwright-cli -s=sna open "https://me.sap.com/notes/<NOTE_NUMBER>"
```

Take a snapshot and confirm the note page loaded. If a login wall appears, pause and ask the user to log in, then confirm before continuing.

---

## Step 1 — Save the full page content

Capture the complete visible text of the page (all sections including description, comments, linked items) using a snapshot or inner-text extraction:

```bash
playwright-cli -s=sna innerText --selector="body" --output="C:\Users\C5404594\.claude\skills\sap-note-analyzer\scripts\source_dump.txt"
```

> **If `innerText` is unavailable**, use a snapshot and pipe the text result to the file, or use:
> ```bash
> playwright-cli -s=sna snapshot --output="C:\Users\C5404594\.claude\skills\sap-note-analyzer\scripts\source_dump.txt"
> ```

### For Jira: scroll to load all content

Jira pages may lazy-load comments and description sections. Before saving, scroll to the bottom to ensure all content is rendered:

```bash
playwright-cli -s=sna evaluate --js="window.scrollTo(0, document.body.scrollHeight)"
playwright-cli -s=sna snapshot --output="C:\Users\C5404594\.claude\skills\sap-note-analyzer\scripts\source_dump.txt"
```

### For SAP notes: capture description + all tabs

The parent SAP note's description is the primary source. Also check:
- The **"Description"** section (default visible tab)
- Any **"Correction Instructions"** or **"Related Notes"** section that may list additional note numbers

Take a snapshot after the page loads — the description area contains the note list. Save the snapshot text to `source_dump.txt`.

---

## Step 2 — Extract note numbers

Run the extractor script against the saved content:

```bash
python "C:\Users\C5404594\.claude\skills\sap-note-analyzer\scripts\extract_notes.py" \
  --input  "C:\Users\C5404594\.claude\skills\sap-note-analyzer\scripts\source_dump.txt" \
  --keywords "<KEYWORDS_COMMA_SEPARATED_OR_EMPTY>" \
  --output "C:\Users\C5404594\.claude\skills\sap-note-analyzer\scripts\notes_to_process.txt"
```

- If the user provided keywords (e.g. `"Basis,BC-FES"`), pass them as `--keywords "Basis,BC-FES"`.
- If the user said "all notes" or gave no keyword, omit `--keywords` or pass an empty string.

The script prints a line starting with `JSON:` — parse it to get:
- `notes` — ordered, deduplicated list of 7-digit note numbers
- `output_path` — path to the generated `.txt` file

Show the user the extracted note list and ask for confirmation before proceeding:

> "Found **N** notes: `<list>`. Proceed to process them with oss-note-processor?"

If the user wants to adjust the keyword and re-run, repeat Step 2 with the new keyword without re-fetching the page.

---

## Step 3 — Close the browser session

```bash
playwright-cli -s=sna close
```

---

## Step 4 — Hand off to oss-note-processor

Invoke the **oss-note-processor** skill, passing the path to the text file produced in Step 2:

```
Skill("oss-note-processor", args="C:\\Users\\C5404594\\.claude\\skills\\sap-note-analyzer\\scripts\\notes_to_process.txt")
```

The oss-note-processor skill will:
1. Build an Excel workbook from the note list (preserving the extracted order).
2. Visit each note on me.sap.com and fill Component, Description, Link, SP109 support, Manual Activity timing, and Attachment columns.
3. Hand off to prereq-tree-walker automatically after it finishes.

---

## Error handling

| Situation | Action |
|-----------|--------|
| Jira login wall | Pause, ask user to log in, wait for confirmation |
| SAP note login wall | Pause, ask user to log in, wait for confirmation |
| Page 404 / not found | Tell the user the source could not be loaded and ask for an alternative |
| No notes found after extraction | Tell the user; ask if they want to relax/change the keyword or verify the source URL |
| User wants to adjust keywords | Re-run Step 2 only (no need to re-fetch the page) |
| `source_dump.txt` is empty | Retry the snapshot; if still empty, ask the user to verify the page is accessible |

---

## Scripts reference

All scripts live in `C:\Users\C5404594\.claude\skills\sap-note-analyzer\scripts\`:

| Script | Purpose |
|--------|---------|
| `extract_notes.py --input <file> [--keywords <kw1,kw2>] [--output <out.txt>]` | Extracts 7-digit SAP note numbers from a text/HTML dump, filters by keyword context window, preserves order, deduplicates, writes one note per line to output file |

"""
read_input.py  <input_path>

Reads root note numbers from a .txt or .xlsx file and seeds them
into tree_state.json as root-level notes (depth=0, root=<self>).

Prints one JSON line:
  {"root_notes": ["3499841", ...], "queue": ["3499841", ...]}

Recognises 7-digit SAP note numbers in any format (one per line,
comma/space separated, mixed text, etc.).
"""
import sys
import re
import json
import os
import openpyxl

STATE_FILE = os.path.join(os.path.dirname(__file__), "tree_state.json")


def parse_txt(path):
    with open(path, encoding="utf-8", errors="replace") as f:
        content = f.read()
    seen, notes = set(), []
    for m in re.finditer(r"\b(\d{7})\b", content):
        n = m.group(1)
        if n not in seen:
            seen.add(n)
            notes.append(n)
    return notes


def parse_xlsx(path):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb.active

    # Detect oss-note-processor format: header row has "Note Number" in col B
    # and "SP109 Supported" somewhere. If so, only take notes where SP109 = Yes.
    rows = list(ws.iter_rows(values_only=True))
    wb.close()

    if not rows:
        return []

    header = [str(c).strip() if c else "" for c in rows[0]]
    note_col   = next((i for i, h in enumerate(header) if "Note Number" in h), None)
    sp109_col  = next((i for i, h in enumerate(header) if "SP109" in h and "Supported" in h), None)

    if note_col is not None and sp109_col is not None:
        # oss-note-processor format: filter to SP109 Supported = Yes
        seen, notes = set(), []
        for row in rows[1:]:
            if len(row) <= max(note_col, sp109_col):
                continue
            note_val  = row[note_col]
            sp109_val = str(row[sp109_col]).strip() if row[sp109_col] else ""
            if note_val is None:
                continue
            for m in re.finditer(r"\b(\d{7})\b", str(note_val)):
                n = m.group(1)
                if n not in seen and sp109_val.lower() == "yes":
                    seen.add(n)
                    notes.append(n)
        return notes

    # Generic xlsx: scan all cells for 7-digit numbers
    seen, notes = set(), []
    for row in rows:
        for cell in row:
            if cell is None:
                continue
            for m in re.finditer(r"\b(\d{7})\b", str(cell)):
                n = m.group(1)
                if n not in seen:
                    seen.add(n)
                    notes.append(n)
    return notes


def load_state():
    if os.path.exists(STATE_FILE):
        with open(STATE_FILE, encoding="utf-8") as f:
            return json.load(f)
    return {"notes": {}, "visited": [], "root_notes": []}


def save_state(state):
    with open(STATE_FILE, "w", encoding="utf-8") as f:
        json.dump(state, f, indent=2, ensure_ascii=False)


def main():
    if len(sys.argv) < 2:
        print("Usage: python read_input.py <txt_or_xlsx_path>", file=sys.stderr)
        sys.exit(1)

    path = sys.argv[1]
    ext = os.path.splitext(path)[1].lower()

    if ext in (".xlsx", ".xls", ".xlsm"):
        root_notes = parse_xlsx(path)
    else:
        root_notes = parse_txt(path)

    if not root_notes:
        print("WARNING: No 7-digit note numbers found in the input file.", file=sys.stderr)

    state = load_state()
    state["root_notes"] = root_notes

    # Seed each root note into the notes dict if not already present
    for note in root_notes:
        if note not in state["notes"]:
            state["notes"][note] = {
                "depth": 0,
                "root": note,       # root note is itself for top-level notes
                "parent": "",       # no parent
                "sp109": "",
                "prereqs_109": [],  # only prereqs that support S4CORE 109
                "all_prereqs": [],  # all prereqs found on the page
                "component": "",
                "description": "",
                "manual": "",
                "timing": "",
                "attachment": "",
            }

    # Queue = root notes not yet visited
    queue = [n for n in root_notes if n not in state["visited"]]
    save_state(state)

    print(json.dumps({"root_notes": root_notes, "queue": queue}))


if __name__ == "__main__":
    main()

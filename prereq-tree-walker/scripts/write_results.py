"""
write_results.py <excel_path>

Reads tree_state.json and writes a two-section Excel report.

SECTION 1 — SP109 Prerequisite Tree (horizontal expansion)
  Each depth level occupies one column.
  Col A = root (Sequence Note), Col B = 1st Pre Note, Col C = 2nd Pre Note, …
  When a branch has no further SP109 prereqs the next column gets "End".
  Parent cells are merged vertically across all rows they span.
  Root notes with no prereqs: col A = root, col B = "End".

SECTION 2 — Prerequisite Detail (stacked per-level sub-tables)
  One sub-table per depth level, only for levels that had at least one note
  with non-empty prereqs_109.  Notes with no prereqs are excluded entirely.
  Each sub-table has its own header row:
    Sequence Note (Parent) | Prerequisite Note | Link | Software Component |
    Attachment | Manual Activities | Pre/Post Implementation | SP109 Support Package
  Sub-tables stacked vertically with a blank row between them.
  Col A merged vertically per parent group within each sub-table.
"""
import sys
import json
import os
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

STATE_FILE = os.path.join(os.path.dirname(__file__), "tree_state.json")

# ── Styles ────────────────────────────────────────────────────────────────────
LEVEL_FILLS = [
    PatternFill("solid", fgColor="D9D9D9"),  # level 0 — grey
    PatternFill("solid", fgColor="BDD7EE"),  # level 1 — blue
    PatternFill("solid", fgColor="C6EFCE"),  # level 2 — green
    PatternFill("solid", fgColor="FFEB9C"),  # level 3 — yellow
    PatternFill("solid", fgColor="FCE4D6"),  # level 4+ — orange
]
END_FILL    = PatternFill("solid", fgColor="FFD7D7")   # light red for "End"
GREEN_FILL  = PatternFill("solid", fgColor="375623")
DARK_HDR    = PatternFill("solid", fgColor="595959")
WHITE_BOLD  = Font(bold=True, color="FFFFFF")
DARK_BOLD   = Font(bold=True, color="000000")
END_FONT    = Font(italic=True, color="990000")
LINK_FONT   = Font(color="0563C1", underline="single")
CENTER      = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT_WRAP   = Alignment(horizontal="left",   vertical="center", wrap_text=True)


def level_fill(level):
    return LEVEL_FILLS[min(level, len(LEVEL_FILLS) - 1)]


def load_state():
    if not os.path.exists(STATE_FILE):
        print("ERROR: tree_state.json not found. Run the analysis first.")
        sys.exit(1)
    with open(STATE_FILE, encoding="utf-8") as f:
        return json.load(f)


# ── Section 1 helpers ─────────────────────────────────────────────────────────

def build_tree_rows(notes, root_notes):
    """
    Return a flat list of row-paths.  Each path is a list of (value, level)
    tuples representing one leaf-to-root journey through the tree.

    Example for root 3773500 → child 3587967 → no further prereqs:
      [("3773500", 0), ("3587967", 1), ("End", None)]

    For a root with no prereqs:
      [("3773500", 0), ("End", None)]
    """
    rows = []

    def expand(note, path, visited):
        info    = notes.get(note, {})
        prereqs = info.get("prereqs_109", [])
        level   = info.get("level", 0)
        path    = path + [(note, level)]

        # filter out cycles / already-visited deeper nodes
        fresh = [p for p in prereqs if p not in visited]

        if not fresh:
            rows.append(path + [("End", None)])
        else:
            new_visited = visited | set(fresh)
            for child in fresh:
                expand(child, path, new_visited)

    for root in root_notes:
        expand(root, [], {root})

    return rows


def render_section1(ws, tree_rows, start_row):
    """
    Write Section 1 horizontally.  Returns the last used row number.
    """
    if not tree_rows:
        return start_row

    max_cols = max(len(r) for r in tree_rows)

    # Header labels
    col_labels = ["Sequence Note"]
    ordinals   = ["1st", "2nd", "3rd"] + [f"{n}th" for n in range(4, max_cols + 1)]
    for i in range(1, max_cols):
        col_labels.append(f"{ordinals[i-1]} Pre Note")

    title_row  = start_row
    header_row = start_row + 1
    data_start = start_row + 2

    # Section title
    title_cell = ws.cell(row=title_row, column=1, value="SP109 Prerequisite Tree")
    title_cell.font = Font(bold=True, size=12)
    if max_cols > 1:
        ws.merge_cells(start_row=title_row, start_column=1,
                       end_row=title_row, end_column=max_cols)

    # Header row
    col_widths = [22] + [20] * (max_cols - 1)
    for col_idx, (label, width) in enumerate(zip(col_labels, col_widths), start=1):
        cell = ws.cell(row=header_row, column=col_idx, value=label)
        cell.font  = WHITE_BOLD
        cell.fill  = DARK_HDR
        cell.alignment = CENTER
        ws.column_dimensions[get_column_letter(col_idx)].width = width
    ws.row_dimensions[header_row].height = 28

    # Write data rows
    for row_idx, path in enumerate(tree_rows):
        excel_row = data_start + row_idx
        ws.row_dimensions[excel_row].height = 20
        for col_idx, (value, level) in enumerate(path, start=1):
            cell = ws.cell(row=excel_row, column=col_idx, value=value)
            cell.alignment = CENTER
            if value == "End":
                cell.fill = END_FILL
                cell.font = END_FONT
            elif level is not None:
                cell.fill = level_fill(level)

    # Merge consecutive identical non-empty, non-End values in each column.
    # We also track the root-note row-block boundaries (from col A) so that
    # "End" cells in col B for different root notes never merge together.
    data_end = data_start + len(tree_rows) - 1

    # Build root-block boundaries: each root note starts a new non-mergeable group
    # across ALL columns so that "End" or blank cells from different root blocks
    # are never merged together.
    block_breaks = set()   # row indices where a new root block begins
    prev_root = None
    for row_idx, path in enumerate(tree_rows):
        root_val = path[0][0] if path else None
        if root_val != prev_root:
            block_breaks.add(data_start + row_idx)
            prev_root = root_val

    for col_idx in range(1, max_cols + 1):
        merge_start = data_start
        prev_val    = ws.cell(row=data_start, column=col_idx).value
        for r in range(data_start + 1, data_end + 2):  # +2 to flush last group
            curr_val = ws.cell(row=r, column=col_idx).value if r <= data_end else object()
            # Break merge group on: value change, block boundary, empty/End value
            force_break = (r in block_breaks) or (prev_val in (None, "", "End"))
            if curr_val != prev_val or force_break:
                if r - 1 > merge_start and prev_val not in (None, "", "End"):
                    ws.merge_cells(start_row=merge_start, start_column=col_idx,
                                   end_row=r - 1, end_column=col_idx)
                merge_start = r
                prev_val    = curr_val

    return data_end


# ── Section 2 helpers ─────────────────────────────────────────────────────────

DETAIL_HEADERS = [
    ("Sequence Note (Parent)", 22),
    ("Prerequisite Note",      20),
    ("Link",                   40),
    ("Software Component",     22),
    ("Attachment",             16),
    ("Manual Activities",      40),
    ("Pre/Post Implementation",22),
    ("SP109 Support Package",  24),
]


def render_detail_subtable(ws, row, level, parent_child_pairs, notes):
    """
    Write one per-level detail sub-table starting at `row`.
    parent_child_pairs: list of (parent_note, child_note) tuples.
    Returns next available row.
    """
    ordinals = ["1st", "2nd", "3rd"] + [f"{n}th" for n in range(4, 20)]
    child_ord = ordinals[level] if level < len(ordinals) else f"level-{level+1}"

    title = f"Prerequisite Detail — Level {level}  (Sequence Note → {child_ord} Pre Note)"
    title_cell = ws.cell(row=row, column=1, value=title)
    title_cell.font = Font(bold=True, size=11)
    ws.merge_cells(start_row=row, start_column=1,
                   end_row=row, end_column=len(DETAIL_HEADERS))
    row += 1

    # Header
    fill = level_fill(level)
    for col_idx, (label, width) in enumerate(DETAIL_HEADERS, start=1):
        cell = ws.cell(row=row, column=col_idx, value=label)
        cell.font  = WHITE_BOLD if level == 0 else DARK_BOLD
        cell.fill  = GREEN_FILL if level == 0 else fill
        if level != 0:
            cell.font = Font(bold=True, color="000000")
        cell.alignment = CENTER
        ws.column_dimensions[get_column_letter(col_idx)].width = width
    ws.row_dimensions[row].height = 28
    row += 1

    # Data rows — group by parent for col-A merging
    # Collect groups: list of (parent, [children])
    groups = []
    for parent, child in parent_child_pairs:
        if groups and groups[-1][0] == parent:
            groups[-1][1].append(child)
        else:
            groups.append((parent, [child]))

    for parent, children in groups:
        detail_map = notes.get(parent, {}).get("prereq_details", {})
        group_start = row
        for child in children:
            d         = detail_map.get(child, {})
            link      = d.get("link",       "")
            component = d.get("component",  "")
            attach    = d.get("attachment", "")
            manual    = d.get("manual",     "")
            timing    = d.get("timing",     "")
            sp109_sp  = d.get("sp109_sp",   "")

            ws.cell(row=row, column=2, value=child).alignment   = CENTER
            link_cell = ws.cell(row=row, column=3, value=link)
            if link.startswith("http"):
                link_cell.hyperlink = link
                link_cell.font      = LINK_FONT
            link_cell.alignment = LEFT_WRAP
            ws.cell(row=row, column=4, value=component).alignment = CENTER
            ws.cell(row=row, column=5, value=attach).alignment    = CENTER
            ws.cell(row=row, column=6, value=manual).alignment    = LEFT_WRAP
            ws.cell(row=row, column=7, value=timing).alignment    = CENTER
            ws.cell(row=row, column=8, value=sp109_sp).alignment  = CENTER
            ws.row_dimensions[row].height = 40
            row += 1

        # Write + colour + merge col-A for this group
        parent_cell       = ws.cell(row=group_start, column=1, value=parent)
        parent_cell.alignment = CENTER
        parent_cell.fill  = level_fill(level)
        if row - 1 > group_start:
            ws.merge_cells(start_row=group_start, start_column=1,
                           end_row=row - 1,       end_column=1)

    return row


def render_section2(ws, notes, processed_notes, root_notes, start_row):
    """
    Write per-level detail sub-tables.  Only levels with ≥1 note having
    non-empty prereqs_109 produce a sub-table.
    Returns last used row.
    """
    # Group notes by their level
    by_level = {}
    all_note_keys = list(root_notes) + [n for n in processed_notes if n not in root_notes]
    for note in all_note_keys:
        info    = notes.get(note, {})
        level   = info.get("level", 0)
        prereqs = info.get("prereqs_109", [])
        if not prereqs:
            continue
        by_level.setdefault(level, []).append(note)

    if not by_level:
        return start_row

    row = start_row
    first = True
    for level in sorted(by_level.keys()):
        if not first:
            row += 1   # blank row between sub-tables
        first = False

        # Build (parent, child) pairs in order
        pairs = []
        for parent in by_level[level]:
            for child in notes[parent]["prereqs_109"]:
                pairs.append((parent, child))

        row = render_detail_subtable(ws, row, level, pairs, notes)

    return row - 1


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    if len(sys.argv) < 2:
        print("Usage: python write_results.py <excel_path>")
        sys.exit(1)

    excel_path      = sys.argv[1]
    state           = load_state()
    notes           = state.get("notes", {})
    root_notes      = state.get("root_notes", [])
    processed_notes = state.get("processed_notes", [])

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "SP109 Prereqs"

    # ── Section 1 ─────────────────────────────────────────────────────────────
    tree_rows = build_tree_rows(notes, root_notes)
    s1_last   = render_section1(ws, tree_rows, start_row=1)

    # ── Section 2 ─────────────────────────────────────────────────────────────
    render_section2(ws, notes, processed_notes, root_notes,
                    start_row=s1_last + 2)

    wb.save(excel_path)

    total = sum(
        len(notes.get(n, {}).get("prereqs_109", []))
        for n in processed_notes
    )
    max_depth = max(
        (notes.get(n, {}).get("level", 0) for n in processed_notes),
        default=0
    )
    print(f"Done. {len(processed_notes)} notes processed. "
          f"{total} SP109 links. Max depth: {max_depth}.")
    print(f"Saved: {excel_path}")


if __name__ == "__main__":
    main()

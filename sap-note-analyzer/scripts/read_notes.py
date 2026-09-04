"""
read_notes.py <excel_path>

Reads note numbers from column B (skipping the header row) and prints a
JSON array of note-number strings.  Also prints the column index map so
the caller knows where each output column sits.
"""
import sys
import json
import openpyxl

REQUIRED_HEADERS = ["SP109 Supported", "Manual Activity", "Prerequisites", "Root Note"]


def get_col_map(ws):
    col_map = {}
    for col in range(1, (ws.max_column or 1) + 1):
        val = ws.cell(row=1, column=col).value
        if val and str(val).strip() in REQUIRED_HEADERS:
            col_map[str(val).strip()] = col
    return col_map


def main():
    if len(sys.argv) < 2:
        print("Usage: python read_notes.py <excel_path>", file=sys.stderr)
        sys.exit(1)

    path = sys.argv[1]
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb.active

    notes = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        note = str(row[1]).strip() if row[1] is not None else ""
        if note and note.isdigit():
            notes.append(note)

    wb.close()

    # Re-open for col map (read_only wb doesn't expose column_dimensions reliably)
    wb2 = openpyxl.load_workbook(path, data_only=True)
    ws2 = wb2.active
    col_map = get_col_map(ws2)
    wb2.close()

    print(json.dumps({"notes": notes, "col_map": col_map}))


if __name__ == "__main__":
    main()

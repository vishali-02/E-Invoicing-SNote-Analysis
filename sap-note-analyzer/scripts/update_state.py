"""
update_state.py

Persists one note's analysis results into note_state.json.

Usage:
  python update_state.py \
    --note  3499841 \
    --sp109 Yes \
    --manual Pre \
    --prereqs "3400001,3400002" \
    --root  3672304

  --prereqs accepts a comma-separated string or empty string.
  --root is empty for notes that were in the original Excel list.

note_state.json schema:
{
  "results": {
    "<note>": {
      "sp109": "Yes|No|N/A|Error",
      "manual": "Pre|Post|Pre & Post|None|N/A|Error",
      "prereqs": ["note1", "note2"],
      "root": "parent_note_or_empty"
    },
    ...
  },
  "visited": ["note1", "note2", ...]
}
"""
import sys
import json
import argparse
import os

STATE_FILE = os.path.join(os.path.dirname(__file__), "note_state.json")


def load_state():
    if os.path.exists(STATE_FILE):
        with open(STATE_FILE, encoding="utf-8") as f:
            return json.load(f)
    return {"results": {}, "visited": []}


def save_state(state):
    with open(STATE_FILE, "w", encoding="utf-8") as f:
        json.dump(state, f, indent=2)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--note",    required=True, help="Note number")
    parser.add_argument("--sp109",   required=True, help="Yes|No|N/A|Error")
    parser.add_argument("--manual",  required=True, help="Pre|Post|Pre & Post|None|N/A|Error")
    parser.add_argument("--prereqs", default="",    help="Comma-separated prereq note numbers")
    parser.add_argument("--root",    default="",    help="Root/parent note number (empty for original notes)")
    args = parser.parse_args()

    prereqs = [p.strip() for p in args.prereqs.split(",") if p.strip()]

    state = load_state()

    state["results"][args.note] = {
        "sp109":   args.sp109,
        "manual":  args.manual,
        "prereqs": prereqs,
        "root":    args.root,
    }

    if args.note not in state["visited"]:
        state["visited"].append(args.note)

    save_state(state)
    print(f"Saved state for note {args.note}. Total visited: {len(state['visited'])}")


if __name__ == "__main__":
    main()

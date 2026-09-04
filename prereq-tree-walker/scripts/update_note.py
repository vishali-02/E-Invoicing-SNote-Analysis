"""
update_note.py  -- save one note's SP109 prerequisites into tree_state.json.

Usage:
  python update_note.py \
    --note        3499841                    \
    --prereqs-109 "3400001,3400002"          \
    --parent      "3499000"                  \
    --level       1

--prereqs-109 : SP109-covered prerequisites found for this note
--parent      : note number of this note's parent in the tree (empty for root notes)
--level       : depth in the tree (0 = root notes from input file)

Prints JSON: {"total_visited": N}
"""
import sys
import json
import argparse
import os

STATE_FILE = os.path.join(os.path.dirname(__file__), "tree_state.json")


def load_state():
    if os.path.exists(STATE_FILE):
        with open(STATE_FILE, encoding="utf-8") as f:
            return json.load(f)
    return {"notes": {}, "visited": [], "root_notes": [], "processed_notes": []}


def save_state(state):
    with open(STATE_FILE, "w", encoding="utf-8") as f:
        json.dump(state, f, indent=2, ensure_ascii=False)


def split_csv(s):
    return [x.strip() for x in s.split(",") if x.strip()] if s.strip() else []


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--note",        required=True)
    parser.add_argument("--prereqs-109", default="", dest="prereqs_109")
    parser.add_argument("--parent",      default="", dest="parent")
    parser.add_argument("--level",       default=0,  dest="level", type=int)
    args = parser.parse_args()

    prereqs_109 = split_csv(args.prereqs_109)

    state = load_state()

    # Ensure top-level keys exist (backwards-compat with old state files)
    state.setdefault("processed_notes", [])

    existing = state["notes"].get(args.note, {})
    existing["prereqs_109"] = prereqs_109
    existing["parent"]      = args.parent
    existing["level"]       = args.level
    state["notes"][args.note] = existing

    if args.note not in state["visited"]:
        state["visited"].append(args.note)

    # processed_notes tracks BFS traversal order for Excel rendering
    if args.note not in state["processed_notes"]:
        state["processed_notes"].append(args.note)

    save_state(state)

    print(json.dumps({"total_visited": len(state["visited"])}))


if __name__ == "__main__":
    main()

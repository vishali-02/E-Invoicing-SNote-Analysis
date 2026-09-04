"""
update_prereq_detail.py  -- save detail fields for one prerequisite note.

Usage:
  python update_prereq_detail.py \
    --root-note   3773500          \
    --prereq-note 3587967          \
    --link        "https://me.sap.com/notes/3587967" \
    --component   "CA-GTF-CSC-EDO-PAP"               \
    --attachment  "Yes"                               \
    --manual      "Run report ZRPT after import"      \
    --timing      "Post"                              \
    --sp109-sp    "SAPK-10902INS4CORE"

All fields except --root-note and --prereq-note are optional (default "").
Timing should be "Pre", "Post", or "" if not determinable.
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
    return {"notes": {}, "visited": [], "root_notes": []}


def save_state(state):
    with open(STATE_FILE, "w", encoding="utf-8") as f:
        json.dump(state, f, indent=2, ensure_ascii=False)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--root-note",   required=True,  dest="root_note")
    parser.add_argument("--prereq-note", required=True,  dest="prereq_note")
    parser.add_argument("--link",        default="",     dest="link")
    parser.add_argument("--component",   default="",     dest="component")
    parser.add_argument("--attachment",  default="",     dest="attachment")
    parser.add_argument("--manual",      default="",     dest="manual")
    parser.add_argument("--timing",      default="",     dest="timing")
    parser.add_argument("--sp109-sp",    default="",     dest="sp109_sp")
    args = parser.parse_args()

    state = load_state()

    # Ensure root note entry exists
    if args.root_note not in state["notes"]:
        state["notes"][args.root_note] = {}

    # Store detail under prereq_details dict keyed by prereq note number
    if "prereq_details" not in state["notes"][args.root_note]:
        state["notes"][args.root_note]["prereq_details"] = {}

    state["notes"][args.root_note]["prereq_details"][args.prereq_note] = {
        "link":       args.link,
        "component":  args.component,
        "attachment": args.attachment,
        "manual":     args.manual,
        "timing":     args.timing,
        "sp109_sp":   args.sp109_sp,
    }

    save_state(state)
    print(json.dumps({"saved": args.prereq_note, "root": args.root_note}))


if __name__ == "__main__":
    main()

"""
update_note.py  -- persist one note's scraped data into note_data.json.

Usage:
  python update_note.py \
    --note     3499841 \
    --component "BC-FES-WGU" \
    --description "Short text of the note" \
    --link "https://me.sap.com/notes/3499841" \
    --sp109    "Yes" \
    --manual   "Yes" \
    --timing   "Post" \
    --attachment "Yes"

Values:
  --sp109       Yes | No | N/A | Error
  --manual      Yes | No | N/A | Error   (whether a manual activity exists for S4CORE 109)
  --timing      Pre | Post | Pre & Post | None | N/A | Error
  --attachment  Yes | No | N/A | Error

note_data.json schema:
{
  "results": {
    "<note>": {
      "component":   "...",
      "description": "...",
      "link":        "https://me.sap.com/notes/<note>",
      "sp109":       "Yes|No|N/A|Error",
      "manual":      "Yes|No|N/A|Error",
      "timing":      "Pre|Post|Pre & Post|None|N/A|Error",
      "attachment":  "Yes|No|N/A|Error"
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

STATE_FILE = os.path.join(os.path.dirname(__file__), "note_data.json")


def load_state():
    if os.path.exists(STATE_FILE):
        with open(STATE_FILE, encoding="utf-8") as f:
            return json.load(f)
    return {"results": {}, "visited": []}


def save_state(state):
    with open(STATE_FILE, "w", encoding="utf-8") as f:
        json.dump(state, f, indent=2, ensure_ascii=False)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--note",        required=True)
    parser.add_argument("--component",   default="")
    parser.add_argument("--description", default="")
    parser.add_argument("--link",        default="")
    parser.add_argument("--sp109",       required=True)
    parser.add_argument("--manual",      required=True)
    parser.add_argument("--timing",      required=True)
    parser.add_argument("--attachment",  required=True)
    args = parser.parse_args()

    state = load_state()
    state["results"][args.note] = {
        "component":   args.component,
        "description": args.description,
        "link":        args.link or f"https://me.sap.com/notes/{args.note}",
        "sp109":       args.sp109,
        "manual":      args.manual,
        "timing":      args.timing,
        "attachment":  args.attachment,
    }
    if args.note not in state["visited"]:
        state["visited"].append(args.note)

    save_state(state)
    print(f"Saved note {args.note}. Total visited: {len(state['visited'])}")


if __name__ == "__main__":
    main()

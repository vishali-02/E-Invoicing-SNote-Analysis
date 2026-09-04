"""
process_note.py - Parse playwright --raw eval output and save to state
Usage: echo '<raw_result>' | python process_note.py --note NOTE [--root ROOT]
Or: python process_note.py --note NOTE --result '<raw_result>' [--root ROOT]
"""
import sys
import json
import argparse
import os
import subprocess

SCRIPTS_DIR = os.path.dirname(os.path.abspath(__file__))


def parse_result(raw):
    raw = raw.strip()
    # playwright --raw wraps in outer double quotes
    if raw.startswith('"') and raw.endswith('"'):
        raw = json.loads(raw)
    return json.loads(raw)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--note", required=True)
    parser.add_argument("--result", default=None, help="JSON result string from playwright --raw eval")
    parser.add_argument("--root", default="")
    args = parser.parse_args()

    if args.result:
        raw = args.result
    else:
        raw = sys.stdin.read()

    try:
        data = parse_result(raw)
    except Exception as e:
        print(f"ERROR parsing result for note {args.note}: {e}", file=sys.stderr)
        print(f"Raw was: {raw[:200]}", file=sys.stderr)
        data = {"sp109": "Error", "manual": "Error", "prereqs": []}

    note = args.note
    sp109 = data.get("sp109", "Error")
    manual = data.get("manual", "Error")
    # Remove self-references from prereqs
    prereqs = [p for p in data.get("prereqs", []) if p != note]

    update_script = os.path.join(SCRIPTS_DIR, "update_state.py")
    cmd = [
        sys.executable, update_script,
        "--note", note,
        "--sp109", sp109,
        "--manual", manual,
        "--prereqs", ",".join(prereqs),
        "--root", args.root,
    ]
    result = subprocess.run(cmd, capture_output=True, text=True)
    print(f"  {note}: sp109={sp109}, manual={manual}, prereqs={prereqs}")
    if result.returncode != 0:
        print(f"  ERROR: {result.stderr.strip()}", file=sys.stderr)
    return result.returncode


if __name__ == "__main__":
    sys.exit(main())

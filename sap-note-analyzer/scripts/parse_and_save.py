"""
parse_and_save.py --note <NOTE> --json '<json_string>' [--root <ROOT>]

Parses the playwright --raw eval output (a double-quoted JSON string)
and calls update_state.py with the extracted values.
"""
import sys
import json
import argparse
import os
import subprocess

SCRIPTS_DIR = os.path.dirname(os.path.abspath(__file__))


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--note", required=True)
    parser.add_argument("--json", required=True, dest="json_str")
    parser.add_argument("--root", default="")
    args = parser.parse_args()

    raw = args.json_str.strip()
    # playwright --raw eval wraps the result in double quotes
    if raw.startswith('"') and raw.endswith('"'):
        raw = json.loads(raw)  # un-escape the outer string
    data = json.loads(raw)

    sp109 = data.get("sp109", "Error")
    manual = data.get("manual", "Error")
    prereqs = ",".join(data.get("prereqs", []))

    update_script = os.path.join(SCRIPTS_DIR, "update_state.py")
    cmd = [
        sys.executable, update_script,
        "--note", args.note,
        "--sp109", sp109,
        "--manual", manual,
        "--prereqs", prereqs,
        "--root", args.root,
    ]
    result = subprocess.run(cmd, capture_output=True, text=True)
    print(result.stdout.strip())
    if result.returncode != 0:
        print("ERROR:", result.stderr.strip(), file=sys.stderr)
    return result.returncode


if __name__ == "__main__":
    sys.exit(main())

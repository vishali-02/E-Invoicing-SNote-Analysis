"""
get_queue.py

Reads tree_state.json and prints the notes that still need to be visited
(present in state["notes"] but not yet in state["visited"]).

Prints JSON: {"queue": ["note1", "note2", ...], "total_remaining": N}
"""
import json
import os
import sys

STATE_FILE = os.path.join(os.path.dirname(__file__), "tree_state.json")


def main():
    if not os.path.exists(STATE_FILE):
        print(json.dumps({"queue": [], "total_remaining": 0}))
        return

    with open(STATE_FILE, encoding="utf-8") as f:
        state = json.load(f)

    visited = set(state.get("visited", []))
    queue = [n for n in state.get("notes", {}) if n not in visited]

    print(json.dumps({"queue": queue, "total_remaining": len(queue)}))


if __name__ == "__main__":
    main()

"""
reset_state.py

Clears note_state.json to start a fresh analysis run.
"""
import os
import json

STATE_FILE = os.path.join(os.path.dirname(__file__), "note_state.json")


def main():
    fresh = {"results": {}, "visited": []}
    with open(STATE_FILE, "w", encoding="utf-8") as f:
        json.dump(fresh, f, indent=2)
    print(f"Reset: {STATE_FILE}")


if __name__ == "__main__":
    main()

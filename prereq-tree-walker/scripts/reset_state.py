"""
reset_state.py

Deletes tree_state.json so the next run starts clean.
Safe to call multiple times.
"""
import os

STATE_FILE = os.path.join(os.path.dirname(__file__), "tree_state.json")


def main():
    if os.path.exists(STATE_FILE):
        os.remove(STATE_FILE)
        print(f"Deleted {STATE_FILE}. Ready for a fresh run.")
    else:
        print("Nothing to reset — tree_state.json did not exist.")


if __name__ == "__main__":
    main()

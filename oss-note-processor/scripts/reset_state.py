"""
reset_state.py

Deletes note_data.json so a fresh run can start.
Safe to run multiple times.
"""
import os

STATE_FILE = os.path.join(os.path.dirname(__file__), "note_data.json")


def main():
    if os.path.exists(STATE_FILE):
        os.remove(STATE_FILE)
        print(f"Deleted {STATE_FILE}")
    else:
        print("Nothing to reset — note_data.json did not exist.")


if __name__ == "__main__":
    main()

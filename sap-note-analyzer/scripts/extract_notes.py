"""
extract_notes.py

Extract 7-digit SAP note numbers from a text/HTML content source, optionally
filtered by one or more keywords. Preserves the order they first appear in
the source text and deduplicates.

Usage:
    python extract_notes.py --input <path_to_text_file> [--keywords <kw1,kw2,...>] [--output <notes.txt>]

Arguments:
    --input      Path to a .txt or .html file containing the source content
                 (Jira page export or SAP note description dump).
    --keywords   Comma-separated keywords. A note number is included only if
                 the surrounding text context (±300 chars) contains at least
                 one of the keywords (case-insensitive). Omit to include all notes.
    --output     Path for the output text file (one note per line).
                 Defaults to notes_to_process.txt in the same directory as this script.

Outputs:
    - Writes one SAP note number per line to --output.
    - Prints a JSON summary line: JSON:{"notes": [...], "output_path": "..."}
"""

import sys
import re
import argparse
import json
import os
from pathlib import Path


def strip_html(text: str) -> str:
    """Remove HTML tags, replace common entities, normalise whitespace."""
    text = re.sub(r"<[^>]+>", " ", text)
    text = text.replace("&nbsp;", " ").replace("&amp;", "&").replace(
        "&lt;", "<").replace("&gt;", ">").replace("&#39;", "'")
    text = re.sub(r"\s+", " ", text)
    return text


def extract_notes(content: str, keywords: list[str]) -> list[str]:
    """
    Find all 7-digit SAP note numbers in *content* that appear in a context
    that contains at least one keyword (or all notes if keywords is empty).
    Returns an ordered, deduplicated list.
    """
    CONTEXT_WINDOW = 300  # characters around each match to search for keywords

    seen: set[str] = set()
    result: list[str] = []

    for m in re.finditer(r"\b(\d{7})\b", content):
        note = m.group(1)
        if note in seen:
            continue

        if keywords:
            start = max(0, m.start() - CONTEXT_WINDOW)
            end = min(len(content), m.end() + CONTEXT_WINDOW)
            context = content[start:end].lower()
            if not any(kw.lower() in context for kw in keywords):
                continue

        seen.add(note)
        result.append(note)

    return result


def main():
    parser = argparse.ArgumentParser(description="Extract SAP note numbers from a source file")
    parser.add_argument("--input", required=True, help="Path to source text/HTML file")
    parser.add_argument("--keywords", default="", help="Comma-separated keyword filter (optional)")
    parser.add_argument("--output", default="", help="Output .txt path (one note per line)")
    args = parser.parse_args()

    input_path = Path(args.input)
    if not input_path.exists():
        print(f"ERROR: Input file not found: {input_path}", file=sys.stderr)
        sys.exit(1)

    content = input_path.read_text(encoding="utf-8", errors="replace")

    # Strip HTML if the file looks like HTML
    if input_path.suffix.lower() in (".html", ".htm") or content.lstrip().startswith("<"):
        content = strip_html(content)

    keywords = [k.strip() for k in args.keywords.split(",") if k.strip()] if args.keywords else []

    notes = extract_notes(content, keywords)

    if not notes:
        print("WARNING: No 7-digit SAP note numbers found matching the criteria.")
    else:
        print(f"Found {len(notes)} note(s):{'' if keywords else ' (no keyword filter)'}")
        for i, n in enumerate(notes, 1):
            print(f"  {i}. {n}")

    # Determine output path
    if args.output:
        output_path = Path(args.output)
    else:
        output_path = Path(__file__).parent / "notes_to_process.txt"

    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text("\n".join(notes) + ("\n" if notes else ""), encoding="utf-8")
    print(f"Saved to: {output_path}")
    print("JSON:" + json.dumps({"notes": notes, "output_path": str(output_path)}))


if __name__ == "__main__":
    main()

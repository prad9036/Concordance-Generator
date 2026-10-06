#!/usr/bin/env python3
"""Convert a stop-word DOCX into a plain-text, one-entry-per-line file."""

import argparse
import re
from pathlib import Path

from docx import Document


SCRIPT_DIR = Path(__file__).resolve().parent
DEFAULT_INPUT = SCRIPT_DIR / "stopwords (3).docx"
DEFAULT_OUTPUT = SCRIPT_DIR / "temp_stop_words_list.txt"
ENTRY_RE = re.compile(
    r"^(.+?)\s{2,}\d+(?:\.\d+)?(?:,\s*\d+(?:\.\d+)?)*\s*$"
)


def term_only(text):
    """Remove trailing concordance page references, if present."""
    text = text.strip()
    match = ENTRY_RE.fullmatch(text)
    return match.group(1).strip() if match else text


def extract_lines(document):
    """Return non-empty paragraph and table-cell text in document order."""
    lines = []
    for paragraph in document.paragraphs:
        text = paragraph.text.strip()
        if text:
            lines.extend(term_only(part) for part in text.splitlines() if part.strip())

    for table in document.tables:
        for row in table.rows:
            # Concordance tables put the indexed term in the first cell;
            # later cells contain page references.
            text = row.cells[0].text.strip() if row.cells else ""
            if text:
                lines.extend(term_only(part) for part in text.splitlines() if part.strip())

    return lines


def main():
    parser = argparse.ArgumentParser(
        description="Convert a DOCX stop-word list to a UTF-8 text file."
    )
    parser.add_argument("input", nargs="?", type=Path, default=DEFAULT_INPUT)
    parser.add_argument("output", nargs="?", type=Path, default=DEFAULT_OUTPUT)
    args = parser.parse_args()

    input_path = args.input.expanduser().resolve()
    output_path = args.output.expanduser().resolve()
    if input_path == output_path:
        parser.error("input and output must be different files; the DOCX will not be overwritten")
    if not input_path.is_file():
        parser.error(f"input DOCX not found: {input_path}")

    lines = extract_lines(Document(input_path))
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text("\n".join(lines) + "\n", encoding="utf-8")
    print(f"Wrote {len(lines)} line(s) to {output_path}")


if __name__ == "__main__":
    main()

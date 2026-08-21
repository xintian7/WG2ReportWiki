#!/usr/bin/env python3
"""Convert glossary markdown text into a two-column XLSX file.

The parser expects lines in this pattern for new entries:
    <term><2+ spaces><definition>
Continuation lines are appended to the previous definition.

Example:
    /opt/anaconda3/envs/tsu/bin/python script/glossary_md_to_xlsx.py \
      --input data/IPCC_AR6_WGI_FGD_AnnexVII_Glossary.md \
      --output data/IPCC_AR6_WGI_FGD_AnnexVII_Glossary.xlsx
"""

from __future__ import annotations

import argparse
import re
from pathlib import Path

from openpyxl import Workbook


ENTRY_SPLIT_RE = re.compile(r"^(?P<term>\S.*?\S)\s{2,}(?P<definition>\S.*)$")


def normalize_space(text: str) -> str:
    return re.sub(r"\s+", " ", text).strip()


def parse_glossary_entries(lines: list[str], start_term: str, stop_prefix: str) -> list[tuple[str, str]]:
    """Parse markdown lines into (term, definition) pairs."""
    entries: list[tuple[str, str]] = []

    started = False
    current_term: str | None = None
    current_def_parts: list[str] = []

    start_key = normalize_space(start_term)
    stop_key = stop_prefix.strip().lower()

    for raw_line in lines:
        line = raw_line.rstrip()
        stripped = line.strip()

        if not started:
            # Start when we find the requested first glossary term line.
            if start_key and normalize_space(stripped).startswith(start_key):
                started = True
            else:
                continue

        if not stripped:
            continue

        if stop_key and stripped.lower().startswith(stop_key):
            break

        split_match = ENTRY_SPLIT_RE.match(stripped)
        if split_match:
            if current_term is not None and current_def_parts:
                entries.append((current_term, normalize_space(" ".join(current_def_parts))))

            current_term = split_match.group("term").strip()
            current_def_parts = [split_match.group("definition").strip()]
            continue

        # If a line has no entry split and we are in an active entry, treat it as definition continuation.
        if current_term is not None:
            current_def_parts.append(stripped)

    if current_term is not None and current_def_parts:
        entries.append((current_term, normalize_space(" ".join(current_def_parts))))

    return entries


def write_xlsx(entries: list[tuple[str, str]], output_path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Glossary"

    ws.append(["Term", "Explanation"])
    for term, explanation in entries:
        ws.append([term, explanation])

    # Improve readability in spreadsheet apps.
    ws.column_dimensions["A"].width = 44
    ws.column_dimensions["B"].width = 120

    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Convert glossary markdown to two-column XLSX.")
    parser.add_argument(
        "--input",
        "-i",
        type=Path,
        default=Path("data/IPCC_AR6_WGI_FGD_AnnexVII_Glossary.md"),
        help="Input markdown file path.",
    )
    parser.add_argument(
        "--output",
        "-o",
        type=Path,
        default=Path("data/IPCC_AR6_WGI_FGD_AnnexVII_Glossary.xlsx"),
        help="Output XLSX file path.",
    )
    parser.add_argument(
        "--start-term",
        default="1.5°C pathway",
        help="Start parsing when a line beginning with this term is found.",
    )
    parser.add_argument(
        "--stop-prefix",
        default="Appendix",
        help="Stop parsing when a line begins with this prefix.",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()

    if not args.input.exists():
        raise FileNotFoundError(f"Input markdown not found: {args.input}")

    lines = args.input.read_text(encoding="utf-8").splitlines()
    entries = parse_glossary_entries(lines, start_term=args.start_term, stop_prefix=args.stop_prefix)

    if not entries:
        raise RuntimeError(
            "No glossary entries were parsed. Check --start-term/--stop-prefix and input formatting."
        )

    write_xlsx(entries, args.output)
    print(f"Wrote {len(entries)} entries to {args.output}")


if __name__ == "__main__":
    main()

#!/usr/bin/env python3
"""Convert a DOCX file to Markdown while preserving heading levels.

Example:
    /opt/anaconda3/envs/tsu/bin/python script/docx_to_md.py \
      --input data/SRCities_FOD_SPM_Final.docx \
      --output data/SRCities_FOD_SPM_Final.md
"""

from __future__ import annotations

import argparse
import re
from pathlib import Path

from docx import Document


def heading_level_from_style(style_name: str) -> int | None:
    """Map DOCX style names to Markdown heading levels."""
    name = (style_name or "").strip().lower()

    # Built-in Word headings: Heading 1, Heading 2, ...
    match = re.search(r"heading\s*(\d+)", name)
    if match:
        level = int(match.group(1))
        return min(max(level, 1), 6)

    # Custom styles seen in this repo's SRCities document.
    custom_map = {
        "0th level chapter heading": 1,
        "1st level heading": 2,
        "2nd level heading": 3,
        "3rd level heading": 4,
        "4th level heading": 5,
    }

    for key, level in custom_map.items():
        if key in name:
            return level

    return None


def clean_text(text: str) -> str:
    """Normalize whitespace but keep paragraph semantics."""
    text = text.replace("\u00a0", " ")
    text = re.sub(r"\s+", " ", text).strip()
    return text


def docx_to_markdown(input_docx: Path, output_md: Path) -> int:
    doc = Document(str(input_docx))

    lines: list[str] = []
    non_empty_count = 0

    for para in doc.paragraphs:
        raw = para.text or ""
        text = clean_text(raw)
        if not text:
            continue

        style_name = para.style.name if para.style is not None else ""
        level = heading_level_from_style(style_name)

        if level is not None:
            lines.append(f"{'#' * level} {text}")
        else:
            lines.append(text)

        non_empty_count += 1

    output_md.parent.mkdir(parents=True, exist_ok=True)
    output_md.write_text("\n\n".join(lines).strip() + "\n", encoding="utf-8")

    return non_empty_count


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Convert DOCX to Markdown with heading levels.")
    parser.add_argument(
        "--input",
        "-i",
        type=Path,
        default=Path("data/SRCities_FOD_SPM_Final.docx"),
        help="Path to input DOCX.",
    )
    parser.add_argument(
        "--output",
        "-o",
        type=Path,
        default=Path("data/SRCities_FOD_SPM_Final.md"),
        help="Path to output Markdown.",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()

    if not args.input.exists():
        raise FileNotFoundError(f"Input DOCX not found: {args.input}")

    count = docx_to_markdown(args.input, args.output)
    print(f"Converted {args.input} -> {args.output} ({count} non-empty paragraphs)")


if __name__ == "__main__":
    main()

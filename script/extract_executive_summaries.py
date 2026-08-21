#!/usr/bin/env python3
"""Extract the executive summary from each SRCities chapter DOCX as Markdown.

Example:
    /opt/anaconda3/envs/tsu/bin/python script/extract_executive_summaries.py
"""

from __future__ import annotations

import argparse
import re
from pathlib import Path

from docx import Document


EXECUTIVE_SUMMARY = "executive summary"
FIRST_LEVEL_STYLE = "srcities 1st level heading"
CHAPTER_NUMBER_RE = re.compile(r"_ch(\d+)_", re.IGNORECASE)
STYLE_HEADING_LEVELS = {
    "srcities 1st level heading": 1,
    "srcities 2nd level heading": 2,
    "srcities 3rd level heading": 3,
    "srcities 4th level heading": 4,
}


def clean_text(text: str) -> str:
    """Normalize DOCX whitespace while preserving paragraph boundaries."""
    return re.sub(r"\s+", " ", text.replace("\u00a0", " ")).strip()


def is_first_level_heading(paragraph) -> bool:
    """Identify the custom top-level heading style used in the source DOCX files."""
    style_name = paragraph.style.name if paragraph.style is not None else ""
    return style_name.strip().casefold() == FIRST_LEVEL_STYLE


def heading_level(paragraph) -> int | None:
    """Return the Markdown level represented by a DOCX paragraph, if any."""
    style_name = paragraph.style.name if paragraph.style is not None else ""
    level = STYLE_HEADING_LEVELS.get(style_name.strip().casefold())
    if level is not None:
        return level

    text = clean_text(paragraph.text)
    non_empty_runs = [run for run in paragraph.runs if clean_text(run.text)]
    if 0 < len(text) <= 100 and non_empty_runs and all(run.bold is True for run in non_empty_runs):
        return 2

    return None


def markdown_text(paragraph) -> str:
    """Render paragraph text while retaining direct bold and italic run formatting."""
    characters: list[tuple[str, bool, bool]] = []
    pending_space = False
    for run in paragraph.runs:
        for character in run.text.replace("\u00a0", " "):
            if character.isspace():
                pending_space = True
                continue

            if pending_space and characters:
                previous_bold, previous_italic = characters[-1][1:]
                space_style = (run.bold is True, run.italic is True)
                characters.append((" ", *(space_style if space_style == (previous_bold, previous_italic) else (False, False))))
            characters.append((character, run.bold is True, run.italic is True))
            pending_space = False

    if not characters:
        return ""

    index = 0
    while index < len(characters):
        if not characters[index][2]:
            index += 1
            continue

        end_index = index + 1
        while end_index < len(characters) and characters[end_index][2]:
            end_index += 1

        leading_index = index
        while leading_index < end_index and not characters[leading_index][0].isalnum():
            character, is_bold, _ = characters[leading_index]
            characters[leading_index] = (character, is_bold, False)
            leading_index += 1

        trailing_index = end_index - 1
        while trailing_index >= leading_index and not characters[trailing_index][0].isalnum():
            character, is_bold, _ = characters[trailing_index]
            characters[trailing_index] = (character, is_bold, False)
            trailing_index -= 1

        index = end_index

    uses_bold = any(is_bold for _, is_bold, _ in characters)
    uses_italic = any(is_italic for _, _, is_italic in characters)
    bold_marker = "__" if uses_bold and uses_italic else "**"
    fragments: list[str] = []
    current_markers: list[str] = []
    index = 0
    while index < len(characters):
        character, is_bold, is_italic = characters[index]
        end_index = index + 1
        while end_index < len(characters) and characters[end_index][1:] == (is_bold, is_italic):
            end_index += 1

        text = "".join(item[0] for item in characters[index:end_index])
        desired_markers = ([bold_marker] if is_bold else []) + (["*"] if is_italic else [])
        if not any(item.isalnum() for item in text) and not any(item.isspace() for item in text):
            desired_markers = current_markers.copy()
        shared_marker_count = 0
        for current_marker, desired_marker in zip(current_markers, desired_markers):
            if current_marker != desired_marker:
                break
            shared_marker_count += 1

        fragments.extend(reversed(current_markers[shared_marker_count:]))
        fragments.extend(desired_markers[shared_marker_count:])
        fragments.append(text)
        current_markers = desired_markers
        index = end_index

    fragments.extend(reversed(current_markers))
    return "".join(fragments).strip()


def markdown_line(paragraph) -> str:
    """Convert a source paragraph to either a Markdown heading or body text."""
    text = markdown_text(paragraph)
    level = heading_level(paragraph)
    return f"{'#' * level} {text}" if level is not None else text


def chapter_number(input_docx: Path) -> int:
    """Extract the chapter number from a SRCities chapter filename."""
    match = CHAPTER_NUMBER_RE.search(input_docx.name)
    if match is None:
        raise ValueError(f"Chapter number not found in {input_docx}")
    return int(match.group(1))


def number_executive_summary_lines(lines: list[str], chapter: int) -> list[str]:
    """Prefix ES headings and source paragraphs with stable chapter identifiers."""
    numbered_lines: list[str] = []
    section_number = 0
    paragraph_number = 0

    for line in lines:
        heading_match = re.match(r"^(#{1,6})\s+(.*)$", line)
        if heading_match:
            heading_marks, heading_text = heading_match.groups()
            if len(heading_marks) == 1:
                numbered_lines.append(line)
                continue

            section_number += 1
            paragraph_number = 0
            numbered_lines.append(f"{heading_marks} [ES{chapter}.{section_number}] {heading_text}")
            continue

        paragraph_number += 1
        numbered_lines.append(f"[ES{chapter}.{section_number}.{paragraph_number}] {line}")

    return numbered_lines


def executive_summary_lines(input_docx: Path) -> list[str]:
    """Return Markdown lines from the real summary heading through its end."""
    paragraphs = Document(str(input_docx)).paragraphs
    start_index = next(
        (
            index
            for index, paragraph in enumerate(paragraphs)
            if is_first_level_heading(paragraph)
            and clean_text(paragraph.text).casefold() == EXECUTIVE_SUMMARY
        ),
        None,
    )
    if start_index is None:
        raise ValueError(f"Executive Summary heading not found in {input_docx}")

    end_index = next(
        (
            index
            for index in range(start_index + 1, len(paragraphs))
            if is_first_level_heading(paragraphs[index])
        ),
        len(paragraphs),
    )

    lines = [markdown_line(paragraphs[start_index])]
    lines.extend(
        markdown_line(paragraph)
        for paragraph in paragraphs[start_index + 1 : end_index]
        if clean_text(paragraph.text)
    )
    return number_executive_summary_lines(lines, chapter_number(input_docx))


def extract_summaries(input_dir: Path, output_dir: Path) -> list[Path]:
    """Write an executive-summary Markdown file for every DOCX in input_dir."""
    source_files = sorted(input_dir.glob("*.docx"))
    if not source_files:
        raise FileNotFoundError(f"No DOCX files found in {input_dir}")

    output_dir.mkdir(parents=True, exist_ok=True)
    output_files: list[Path] = []
    for input_docx in source_files:
        output_md = output_dir / f"{input_docx.stem}_executive_summary.md"
        output_md.write_text("\n\n".join(executive_summary_lines(input_docx)) + "\n", encoding="utf-8")
        output_files.append(output_md)

    return output_files


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Extract SRCities executive summaries to Markdown.")
    parser.add_argument("--input-dir", type=Path, default=Path("data/ES"), help="Directory containing chapter DOCX files.")
    parser.add_argument("--output-dir", type=Path, default=Path("data/ES"), help="Directory for extracted Markdown files.")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    for output_file in extract_summaries(args.input_dir, args.output_dir):
        print(f"Wrote {output_file}")


if __name__ == "__main__":
    main()
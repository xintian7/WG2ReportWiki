#!/usr/bin/env python3
"""Convert DOCX or PDF files to Markdown.

Examples:
    /opt/anaconda3/envs/tsu/bin/python script/converters/convert_to_md.py \
      --input data/SRCities_FOD_SPM_Final.docx \
      --output data/SRCities_FOD_SPM_Final.md

    /opt/anaconda3/envs/tsu/bin/python script/converters/convert_to_md.py \
      --input data/IPCC_AR6_WGI_FGD_AnnexVII_Glossary.pdf \
      --output data/IPCC_AR6_WGI_FGD_AnnexVII_Glossary.md
"""

from __future__ import annotations

import argparse
import re
from pathlib import Path

from docx import Document
from pypdf import PdfReader


BOILERPLATE_PATTERNS = [
    re.compile(r"final\s+government\s+distribution\s+glossary\s+ipcc\s+ar6\s+wgi", re.IGNORECASE),
    re.compile(r"do\s+not\s+cite,\s*quote\s+or\s+distribute", re.IGNORECASE),
    re.compile(r"total\s+pages\s*:\s*\d+", re.IGNORECASE),
    re.compile(r"\bavii-\d+\b", re.IGNORECASE),
]


def heading_level_from_style(style_name: str) -> int | None:
    """Map DOCX style names to Markdown heading levels."""
    name = (style_name or "").strip().lower()

    match = re.search(r"heading\s*(\d+)", name)
    if match:
        level = int(match.group(1))
        return min(max(level, 1), 6)

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


def clean_docx_text(text: str) -> str:
    """Normalize DOCX text while preserving paragraph semantics."""
    text = text.replace("\u00a0", " ")
    return re.sub(r"\s+", " ", text).strip()


def should_drop_pdf_line(line: str) -> bool:
    """Return True for lines that are likely page artifacts/boilerplate."""
    stripped = line.strip()
    if not stripped:
        return False

    if re.fullmatch(r"\d{1,3}", stripped):
        return True

    return any(pattern.search(stripped) for pattern in BOILERPLATE_PATTERNS)


def clean_pdf_text(text: str) -> str:
    """Apply cleanup so PDF extracted text reads better in Markdown."""
    if not text:
        return ""

    text = text.replace("\r\n", "\n").replace("\r", "\n")
    text = "\n".join(line.rstrip() for line in text.split("\n"))

    lines = [line for line in text.split("\n") if not should_drop_pdf_line(line)]
    text = "\n".join(lines)

    text = re.sub(r"(?<=\w)-\n(?=\w)", "", text)
    text = re.sub(r"(?m)(\S(?:.*\S)?)\s+\d{1,3}$", r"\1", text)
    text = re.sub(r"\n{3,}", "\n\n", text)

    return text.strip()


def docx_to_markdown(input_docx: Path, output_md: Path) -> int:
    """Convert DOCX paragraphs to Markdown and return written paragraph count."""
    doc = Document(str(input_docx))

    lines: list[str] = []
    non_empty_count = 0

    for para in doc.paragraphs:
        text = clean_docx_text(para.text or "")
        if not text:
            continue

        style_name = para.style.name if para.style is not None else ""
        level = heading_level_from_style(style_name)
        lines.append(f"{'#' * level} {text}" if level is not None else text)
        non_empty_count += 1

    output_md.parent.mkdir(parents=True, exist_ok=True)
    output_md.write_text("\n\n".join(lines).strip() + "\n", encoding="utf-8")
    return non_empty_count


def pdf_to_markdown(input_pdf: Path, output_md: Path, include_page_headers: bool = False) -> None:
    """Extract text from every page in a PDF and write to a Markdown file."""
    reader = PdfReader(str(input_pdf))

    sections: list[str] = [f"# {input_pdf.stem}"]

    for index, page in enumerate(reader.pages, start=1):
        page_text = clean_pdf_text(page.extract_text() or "")

        if include_page_headers:
            sections.append(f"\n## Page {index}\n")

        sections.append(page_text if page_text else "_No extractable text on this page._")

    output_md.parent.mkdir(parents=True, exist_ok=True)
    output_md.write_text("\n\n".join(sections).strip() + "\n", encoding="utf-8")


def detect_format(input_path: Path, requested_format: str) -> str:
    """Resolve converter format from explicit flag or file extension."""
    if requested_format != "auto":
        return requested_format

    suffix = input_path.suffix.lower()
    if suffix == ".docx":
        return "docx"
    if suffix == ".pdf":
        return "pdf"

    raise ValueError(
        "Could not infer input format from extension. Use --format docx or --format pdf."
    )


def convert_to_markdown(
    input_path: Path,
    output_path: Path,
    requested_format: str = "auto",
    include_page_headers: bool = False,
) -> str:
    """Convert one file to Markdown and return a summary line."""
    if not input_path.exists():
        raise FileNotFoundError(f"Input file not found: {input_path}")

    resolved_format = detect_format(input_path, requested_format)
    if resolved_format == "docx":
        count = docx_to_markdown(input_path, output_path)
        return f"Converted {input_path} -> {output_path} ({count} non-empty paragraphs)"

    pdf_to_markdown(input_path, output_path, include_page_headers=include_page_headers)
    return f"Converted {input_path} -> {output_path}"


def parse_args() -> argparse.Namespace:
    """Parse CLI arguments for the unified converter."""
    parser = argparse.ArgumentParser(description="Convert DOCX/PDF to Markdown.")
    parser.add_argument("--input", "-i", type=Path, required=True, help="Input .docx or .pdf file.")
    parser.add_argument("--output", "-o", type=Path, required=True, help="Output .md file.")
    parser.add_argument(
        "--format",
        choices=("auto", "docx", "pdf"),
        default="auto",
        help="Input format (default: auto from extension).",
    )
    parser.add_argument(
        "--page-headers",
        action="store_true",
        help="For PDF conversion, insert a 'Page N' heading before each page.",
    )
    return parser.parse_args()


def main() -> None:
    """Run unified DOCX/PDF conversion."""
    args = parse_args()
    result = convert_to_markdown(
        input_path=args.input,
        output_path=args.output,
        requested_format=args.format,
        include_page_headers=args.page_headers,
    )
    print(result)


if __name__ == "__main__":
    main()

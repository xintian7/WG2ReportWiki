#!/usr/bin/env python3
"""Convert a PDF file to Markdown.

Example:
    python script/pdf_to_md.py \
      --input data/IPCC_AR6_WGI_FGD_AnnexVII_Glossary.pdf \
      --output data/IPCC_AR6_WGI_FGD_AnnexVII_Glossary.md
"""

from __future__ import annotations

import argparse
import re
from pathlib import Path

from pypdf import PdfReader


BOILERPLATE_PATTERNS = [
    re.compile(r"final\s+government\s+distribution\s+glossary\s+ipcc\s+ar6\s+wgi", re.IGNORECASE),
    re.compile(r"do\s+not\s+cite,\s*quote\s+or\s+distribute", re.IGNORECASE),
    re.compile(r"total\s+pages\s*:\s*\d+", re.IGNORECASE),
    re.compile(r"\bavii-\d+\b", re.IGNORECASE),
]


def should_drop_line(line: str) -> bool:
    """Return True for lines that are likely page artifacts/boilerplate."""
    stripped = line.strip()
    if not stripped:
        return False

    # Standalone page numbers like "1" or " 2".
    if re.fullmatch(r"\d{1,3}", stripped):
        return True

    return any(pattern.search(stripped) for pattern in BOILERPLATE_PATTERNS)


def clean_text(text: str) -> str:
    """Apply lightweight cleanup so extracted text reads better in Markdown."""
    if not text:
        return ""

    # Normalize line endings and trim trailing spaces.
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    text = "\n".join(line.rstrip() for line in text.split("\n"))

    # Drop known recurring header/footer lines and bare page-number lines.
    lines = [line for line in text.split("\n") if not should_drop_line(line)]
    text = "\n".join(lines)

    # Join soft hyphenated line breaks: "exam-\nple" -> "example".
    text = re.sub(r"(?<=\w)-\n(?=\w)", "", text)

    # Remove trailing page-number artifacts: "... definition text 5" -> "... definition text".
    text = re.sub(r"(?m)(\S(?:.*\S)?)\s+\d{1,3}$", r"\1", text)

    # Collapse excessive blank lines.
    text = re.sub(r"\n{3,}", "\n\n", text)

    return text.strip()


def pdf_to_markdown(input_pdf: Path, output_md: Path, include_page_headers: bool = False) -> None:
    """Extract text from every page in a PDF and write to a Markdown file."""
    reader = PdfReader(str(input_pdf))

    sections: list[str] = [f"# {input_pdf.stem}"]

    for i, page in enumerate(reader.pages, start=1):
        page_text = clean_text(page.extract_text() or "")

        if include_page_headers:
            sections.append(f"\n## Page {i}\n")

        if page_text:
            sections.append(page_text)
        else:
            sections.append("_No extractable text on this page._")

    output_md.parent.mkdir(parents=True, exist_ok=True)
    output_md.write_text("\n\n".join(sections).strip() + "\n", encoding="utf-8")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Convert a PDF file to a Markdown file.")
    parser.add_argument(
        "--input",
        "-i",
        type=Path,
        default=Path("data/IPCC_AR6_WGI_FGD_AnnexVII_Glossary.pdf"),
        help="Path to the source PDF (default: data/IPCC_AR6_WGI_FGD_AnnexVII_Glossary.pdf)",
    )
    parser.add_argument(
        "--output",
        "-o",
        type=Path,
        default=Path("data/IPCC_AR6_WGI_FGD_AnnexVII_Glossary.md"),
        help="Path to the output Markdown file (default: data/IPCC_AR6_WGI_FGD_AnnexVII_Glossary.md)",
    )
    parser.add_argument(
        "--page-headers",
        action="store_true",
        help="Insert a 'Page N' markdown heading before each page.",
    )

    return parser.parse_args()


def main() -> None:
    args = parse_args()

    if not args.input.exists():
        raise FileNotFoundError(f"Input PDF not found: {args.input}")

    pdf_to_markdown(
        input_pdf=args.input,
        output_md=args.output,
        include_page_headers=args.page_headers,
    )

    print(f"Converted {args.input} -> {args.output}")


if __name__ == "__main__":
    main()

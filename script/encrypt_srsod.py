#!/usr/bin/env python3
"""Encrypt the SRCities report, glossary, network, executive summaries, LLM prompt, and summaries.

Set FERNET_KEY to a valid Fernet key before running:
    FERNET_KEY="..." /opt/anaconda3/envs/tsu/bin/python script/encrypt_srsod.py
"""

from __future__ import annotations

import argparse
from io import BytesIO
import os
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

from cryptography.fernet import Fernet
from dotenv import load_dotenv


SCRIPT_DIR = Path(__file__).resolve().parent
REPO_ROOT = SCRIPT_DIR.parent
EXECUTIVE_SUMMARY_INPUT_PATHS = tuple(
    REPO_ROOT / f"data/ES/SRCities_SOD_Ch{chapter:02d}_Final_executive_summary.md"
    for chapter in range(1, 6)
)
TERM_USAGE_SUMMARY_PROMPT_INPUT_PATH = REPO_ROOT / "data/prompt/llm_term_usage_summary_prompt.md"
TERM_USAGE_SUMMARIES_INPUT_PATH = REPO_ROOT / "data/analysis/term_usage_summaries.json"
INPUT_PATHS = (
    REPO_ROOT / "data/Glossary/AR6FGD_Glossary.xlsx",
    REPO_ROOT / "data/Glossary/AR7SOD_Glossary.xlsx",
    REPO_ROOT / "data/report/SRCities_FOD_SPM_Final.md",
    REPO_ROOT / "data/network/SRCities_glossary_network.cypher",
    *EXECUTIVE_SUMMARY_INPUT_PATHS,
    TERM_USAGE_SUMMARY_PROMPT_INPUT_PATH,
    TERM_USAGE_SUMMARIES_INPUT_PATH,
)
DEFAULT_OUTPUT = REPO_ROOT / "data/encrypted/SRSOD.enc"


def get_fernet() -> Fernet:
    """Load and validate the Fernet key without exposing it in error output."""
    load_dotenv(REPO_ROOT / ".env", override=False)
    key = os.environ.get("FERNET_KEY")
    if not key:
        raise RuntimeError("FERNET_KEY is not set in the environment or project .env file.")

    try:
        return Fernet(key.encode("ascii"))
    except (UnicodeEncodeError, ValueError) as error:
        raise ValueError("FERNET_KEY must be a valid URL-safe base64 Fernet key.") from error


def build_archive(paths: tuple[Path, ...]) -> bytes:
    """Package the source files into an in-memory ZIP archive."""
    missing_paths = [path for path in paths if not path.is_file()]
    if missing_paths:
        missing = ", ".join(str(path) for path in missing_paths)
        raise FileNotFoundError(f"Encryption input file not found: {missing}")

    archive_buffer = BytesIO()
    with ZipFile(archive_buffer, "w", compression=ZIP_DEFLATED) as archive:
        for path in paths:
            archive.writestr(path.relative_to(REPO_ROOT).as_posix(), path.read_bytes())
    return archive_buffer.getvalue()


def verify_archive(encrypted_payload: bytes, fernet: Fernet, paths: tuple[Path, ...]) -> None:
    """Confirm that decryption reproduces the exact expected source files."""
    with ZipFile(BytesIO(fernet.decrypt(encrypted_payload))) as archive:
        expected_names = {path.relative_to(REPO_ROOT).as_posix() for path in paths}
        if set(archive.namelist()) != expected_names:
            raise RuntimeError("Decrypted archive manifest does not match the requested input files.")
        for path in paths:
            archive_name = path.relative_to(REPO_ROOT).as_posix()
            if archive.read(archive_name) != path.read_bytes():
                raise RuntimeError(f"Decrypted content does not match {archive_name}.")


def encrypt_files(output_path: Path) -> None:
    """Create and verify the encrypted archive at output_path."""
    fernet = get_fernet()
    encrypted_payload = fernet.encrypt(build_archive(INPUT_PATHS))
    verify_archive(encrypted_payload, fernet, INPUT_PATHS)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_bytes(encrypted_payload)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Encrypt SRCities data and executive summaries with FERNET_KEY.")
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT, help="Path for the encrypted archive.")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    encrypt_files(args.output)
    print(f"Wrote encrypted archive to {args.output}")


if __name__ == "__main__":
    main()
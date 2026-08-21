#!/usr/bin/env python3
"""Generate resumable GPT-5.4 term-usage summaries for glossary terms.

The script reads the report corpus and approved prompt from the encrypted
SRSOD archive, writes a plaintext JSON build artifact, and saves after every
completed term. Terms used once or not found receive a local occurrence notice
without an Azure request. Run encrypt_srsod.py after this script completes to
package the JSON into the deployment archive.
"""

from __future__ import annotations

import argparse
from datetime import datetime, timezone
import json
import os
from pathlib import Path
import sys
import time
from typing import Any

import dotenv
from openai import AzureOpenAI

from srcities_streamlit_app import (
    ENCRYPTED_REPORT_PATH,
    Glossary,
    LLM_SUMMARY_EXCLUDED_TERM_KEYS,
    build_term_usage_summary_prompt,
    llm_summary_not_applied_message,
    load_encrypted_assets,
    parse_markdown_sections,
    term_occurrences,
)


SCRIPT_DIR = Path(__file__).resolve().parent
REPO_ROOT = SCRIPT_DIR.parent
DEFAULT_OUTPUT_PATH = REPO_ROOT / "data" / "analysis" / "term_usage_summaries.json"
FORMAT_VERSION = 1
AZURE_OPENAI_DEPLOYMENT = "gpt-5.4"
AZURE_OPENAI_API_VERSION = "2025-04-01-preview"
SYSTEM_INSTRUCTIONS = "Follow the supplied IPCC terminology-review instructions exactly and do not use external sources."


def utc_timestamp() -> str:
    """Return a compact UTC timestamp for the generated artifact."""
    return datetime.now(timezone.utc).replace(microsecond=0).isoformat().replace("+00:00", "Z")


def new_results_artifact() -> dict[str, Any]:
    """Create the versioned JSON structure consumed by the Streamlit app."""
    timestamp = utc_timestamp()
    return {
        "format_version": FORMAT_VERSION,
        "model": AZURE_OPENAI_DEPLOYMENT,
        "created_at": timestamp,
        "updated_at": timestamp,
        "summaries": {},
        "failures": {},
    }


def load_results_artifact(output_path: Path) -> dict[str, Any]:
    """Load and validate an existing resumable JSON artifact when available."""
    if not output_path.exists():
        return new_results_artifact()

    try:
        payload = json.loads(output_path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as error:
        raise ValueError(f"Unable to read the existing summary artifact: {output_path}") from error

    if not isinstance(payload, dict) or payload.get("format_version") != FORMAT_VERSION:
        raise ValueError(
            f"{output_path} does not use supported format version {FORMAT_VERSION}. Use --overwrite to replace it."
        )
    if payload.get("model") != AZURE_OPENAI_DEPLOYMENT:
        raise ValueError(
            f"{output_path} was generated with {payload.get('model')!r}. Use --overwrite to regenerate with "
            f"{AZURE_OPENAI_DEPLOYMENT}."
        )
    if not isinstance(payload.get("summaries"), dict) or not isinstance(payload.get("failures"), dict):
        raise ValueError(f"{output_path} must contain summaries and failures objects.")

    return payload


def write_results_artifact(output_path: Path, payload: dict[str, Any]) -> None:
    """Atomically persist progress so interrupted runs can resume safely."""
    output_path.parent.mkdir(parents=True, exist_ok=True)
    payload["updated_at"] = utc_timestamp()
    temporary_path = output_path.with_suffix(f"{output_path.suffix}.tmp")
    temporary_path.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2, sort_keys=True) + "\n",
        encoding="utf-8",
    )
    temporary_path.replace(output_path)


def has_summary(record: object) -> bool:
    """Return whether an existing artifact entry has usable summary text."""
    return isinstance(record, dict) and isinstance(record.get("summary"), str) and bool(record["summary"].strip())


def local_occurrence_summary(term: str, occurrence_count: int) -> str:
    """Return an accurate non-LLM message for terms with zero or one occurrence."""
    if occurrence_count == 1:
        return f"The term '{term}' is only used once in the provided texts."
    if occurrence_count == 0:
        return f"The term '{term}' is not used in the provided texts."
    raise ValueError("Local occurrence summaries are only valid for zero or one occurrence.")


def get_azure_openai_client() -> AzureOpenAI:
    """Create the offline GPT-5.4 client from local environment configuration."""
    dotenv.load_dotenv(REPO_ROOT / ".env", override=False)
    api_key = os.getenv("AZURE_API_KEY", "").strip()
    endpoint = os.getenv("AZURE_OPENAI_ENDPOINT", "").strip()
    if not api_key:
        raise RuntimeError("AZURE_API_KEY is missing. Set it in the local .env file before generating summaries.")
    if not endpoint:
        raise RuntimeError("AZURE_OPENAI_ENDPOINT is missing. Set it in the local .env file before generating summaries.")

    return AzureOpenAI(
        api_version=AZURE_OPENAI_API_VERSION,
        azure_endpoint=endpoint,
        api_key=api_key,
    )


def request_summary(client: AzureOpenAI, prompt: str) -> str:
    """Request one evidence-bound summary from the GPT-5.4 deployment."""
    response = client.chat.completions.create(
        messages=[
            {"role": "system", "content": SYSTEM_INSTRUCTIONS},
            {"role": "user", "content": prompt},
        ],
        max_completion_tokens=20000,
        reasoning_effort="high",
        model=AZURE_OPENAI_DEPLOYMENT,
    )
    summary = response.choices[0].message.content
    if not isinstance(summary, str) or not summary.strip():
        raise RuntimeError("Azure OpenAI returned no summary text.")
    return summary.strip()


def remove_stale_records(payload: dict[str, Any], glossary: Glossary) -> bool:
    """Drop persisted entries for terms that no longer exist in the glossary."""
    changed = False
    current_terms = set(glossary)
    for record_collection in (payload["summaries"], payload["failures"]):
        for term_key in list(record_collection):
            if term_key not in current_terms:
                del record_collection[term_key]
                changed = True
    return changed


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Generate resumable GPT-5.4 terminology-use summaries from the encrypted SRCities archive."
    )
    parser.add_argument(
        "--archive",
        type=Path,
        default=ENCRYPTED_REPORT_PATH,
        help="Encrypted source archive to analyze.",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=DEFAULT_OUTPUT_PATH,
        help="Ignored plaintext JSON artifact written before encryption.",
    )
    parser.add_argument(
        "--overwrite",
        action="store_true",
        help="Regenerate selected terms even when a summary is already present.",
    )
    parser.add_argument(
        "--limit",
        type=int,
        help="Process only the first COUNT terms in alphabetical order; use --limit 5 for a small test run.",
        metavar="COUNT",
    )
    parser.add_argument(
        "--term",
        action="append",
        metavar="TERM",
        help="Process only the named glossary term; may be supplied more than once.",
    )
    parser.add_argument(
        "--pause-seconds",
        type=float,
        default=0.0,
        help="Seconds to wait between completed API requests.",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    if args.pause_seconds < 0:
        print("--pause-seconds must be zero or greater.", file=sys.stderr)
        return 2
    if args.limit is not None and args.limit < 1:
        print("--limit must be at least 1.", file=sys.stderr)
        return 2

    try:
        (
            report_text,
            glossary,
            _,
            executive_summaries,
            prompt_template,
            _,
        ) = load_encrypted_assets(str(args.archive), args.archive.stat().st_mtime_ns)
        _, sections = parse_markdown_sections(report_text)
        payload = load_results_artifact(args.output)
    except (FileNotFoundError, OSError, RuntimeError, ValueError) as error:
        print(str(error), file=sys.stderr)
        return 1

    if remove_stale_records(payload, glossary):
        write_results_artifact(args.output, payload)

    summaries: dict[str, Any] = payload["summaries"]
    failures: dict[str, Any] = payload["failures"]
    all_terms = sorted(glossary.items(), key=lambda item: item[1][0][0].casefold())
    requested_term_keys = {term.casefold() for term in args.term} if args.term else set()
    available_term_keys = {term_key for term_key, _ in all_terms}
    unknown_term_keys = requested_term_keys - available_term_keys
    if unknown_term_keys:
        print(f"Unknown glossary term: {', '.join(sorted(unknown_term_keys))}", file=sys.stderr)
        return 2
    terms = (
        [item for item in all_terms if item[0] in requested_term_keys]
        if requested_term_keys
        else all_terms[: args.limit] if args.limit is not None else all_terms
    )
    client: AzureOpenAI | None = None
    completed = 0
    skipped = 0
    failed = 0

    print(
        f"Preparing {len(terms)} of {len(all_terms)} glossary terms; "
        f"{AZURE_OPENAI_DEPLOYMENT} is used only when occurrence count is greater than one.",
        flush=True,
    )
    for index, (term_key, definitions) in enumerate(terms, start=1):
        term = definitions[0][0]
        print(f"Processing {index}/{len(terms)}: {term}", flush=True)
        matches = term_occurrences(term, "", sections, executive_summaries)
        occurrence_count = len(matches)
        existing_record = summaries.get(term_key)

        if term_key in LLM_SUMMARY_EXCLUDED_TERM_KEYS:
            summary = llm_summary_not_applied_message(term)
            if not args.overwrite and has_summary(existing_record) and existing_record.get("summary") == summary:
                skipped += 1
                print("  Reusing LLM-not-applied notice.", flush=True)
                continue

            summaries[term_key] = {
                "term": term,
                "occurrence_count": occurrence_count,
                "summary": summary,
                "summary_source": "LLM not applied",
                "generated_at": utc_timestamp(),
            }
            failures.pop(term_key, None)
            write_results_artifact(args.output, payload)
            completed += 1
            print("  Saved LLM-not-applied notice.", flush=True)
            continue

        if occurrence_count <= 1:
            summary = local_occurrence_summary(term, occurrence_count)
            if not args.overwrite and has_summary(existing_record) and existing_record.get("summary") == summary:
                skipped += 1
                print("  Reusing local occurrence notice.", flush=True)
                continue

            summaries[term_key] = {
                "term": term,
                "occurrence_count": occurrence_count,
                "summary": summary,
                "summary_source": "local occurrence notice",
                "generated_at": utc_timestamp(),
            }
            failures.pop(term_key, None)
            write_results_artifact(args.output, payload)
            completed += 1
            print("  Saved local occurrence notice.", flush=True)
            continue

        if not args.overwrite and has_summary(existing_record):
            skipped += 1
            print("  Reusing existing GPT summary.", flush=True)
            continue

        try:
            if not prompt_template.strip():
                raise ValueError("The approved term-usage summary prompt is empty.")
            if client is None:
                client = get_azure_openai_client()
            prompt = build_term_usage_summary_prompt(prompt_template, term, definitions, matches)
            summary = request_summary(client, prompt)
        except Exception as error:
            failures[term_key] = {
                "term": term,
                "error": f"{type(error).__name__}: {error}",
                "updated_at": utc_timestamp(),
            }
            write_results_artifact(args.output, payload)
            failed += 1
            print(f"  Failed: {error}", file=sys.stderr)
            continue

        summaries[term_key] = {
            "term": term,
            "occurrence_count": occurrence_count,
            "summary": summary,
            "summary_source": AZURE_OPENAI_DEPLOYMENT,
            "generated_at": utc_timestamp(),
        }
        failures.pop(term_key, None)
        write_results_artifact(args.output, payload)
        completed += 1

        if args.pause_seconds:
            time.sleep(args.pause_seconds)

    print(
        f"Finished: {completed} generated, {skipped} already present, {failed} failed. "
        f"Artifact: {args.output}"
    )
    if failed:
        print("Rerun the command to retry failed terms; completed summaries will be reused.", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

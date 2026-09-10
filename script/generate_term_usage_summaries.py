#!/usr/bin/env python3
"""Generate GPT-5.4 terminology-consistency summaries.

This script analyzes canonical terms from
``data/Glossary/AR6_AR7SOD_Glossary_AO.xlsx`` and searches the full report
corpus (Chapters 1-5, SPM, TS) using all searchable forms from columns A through E.

Each standard run replaces the JSON results artifact before processing terms while
the status log remains append-only. Use ``--resume`` to reuse fresh prior results.

For high-frequency terms, analysis is exhaustive but chunked: each term's matched
occurrences are deduplicated, split into token-bounded batches, mapped to
structured consistency findings, and reduced to one final evidence-backed summary.

The output JSON remains format version 1 so the Streamlit parser can continue
reading summaries without changes.
"""

# Key workflow steps:
# 1. Load core inputs.
#    Input: prompt template file, source JSON, reconstructed report HTML, AO glossary XLSX,
#    and existing summaries JSON.
#    Output: in-memory prompt text, canonical report payload, glossary entries, prior artifact.
# 2. Build alias-aware occurrence index.
#    Input: glossary terms and aliases (A-E), canonical report payload, HTML-derived node codes.
#    Output: per-term raw occurrence list with source label, node code/id, matched names, sentence.
# 3. Normalize and deduplicate evidence.
#    Input: raw occurrence list for one term.
#    Output: stable aggregated occurrence records with deduplicated text and frequency counters.
# 4. Compute cache keys and select work.
#    Input: term metadata, aggregated occurrences, prompt text hash, corpus hash, prior summaries.
#    Output: per-term input hash plus decision to reuse or regenerate the summary.
# 5. Run chunked map analysis for terms with occurrences >= 1.
#    Input: aggregated occurrences, chunk size, map prompt schema, Azure OpenAI client settings.
#    Output: structured JSON map results per chunk and chunk-level coverage metadata.
#    High-frequency handling: for very large occurrence sets, process all evidence (no sampling)
#    by splitting into report-balanced chunks, recording per-chunk hashes/coverage, and carrying
#    every chunk result forward so reduce synthesis remains exhaustive and auditable.
# 6. Run reduce synthesis.
#    Input: term definition/aliases, occurrence statistics, all map JSON outputs.
#    Output: one structured term-level assessment (contexts, issues, alignment, conclusion label).
# 7. Persist final records incrementally.
#    Input: reduce result (or local zero-occurrence notice), hashes, chunk stats, prior artifact.
#    Output: updated summaries/failures JSON artifact written after each processed term.

from __future__ import annotations

import argparse
from collections import Counter, defaultdict
from dataclasses import asdict, dataclass
from datetime import datetime, timezone
import hashlib
import json
import os
from pathlib import Path
import re
import sys
import time
from typing import Any

import dotenv
from openai import AzureOpenAI

from reconstruct_srcities_report import (
    DEFAULT_GLOSSARY_PATH,
    DEFAULT_OUTPUT_HTML,
    DEFAULT_SOURCE_JSON,
    GlossaryOccurrence,
    build_glossary_match_map,
    canonical_report_data,
    full_report_term_occurrences,
    load_revised_glossary,
    report_node_codes,
)
SCRIPT_DIR = Path(__file__).resolve().parent
REPO_ROOT = SCRIPT_DIR.parent
DEFAULT_PROMPT_PATH = REPO_ROOT / "data" / "prompt" / "llm_term_usage_summary_prompt.md"
DEFAULT_OUTPUT_PATH = REPO_ROOT / "data" / "analysis" / "llm_term_check.json"
DEFAULT_LOG_PATH = REPO_ROOT / "data" / "analysis" / "llm_term_check.log"
FORMAT_VERSION = 1
PIPELINE_VERSION = "2.1"
AZURE_OPENAI_DEPLOYMENT = "gpt-5.4"
AZURE_OPENAI_API_VERSION = "2025-04-01-preview"
AZURE_OPENAI_REASONING_EFFORT = "high"
AZURE_OPENAI_MAP_REASONING_EFFORT = "low"
DEFAULT_REDUCE_MAX_COMPLETION_TOKENS = 10_000
FINAL_REDUCE_RETRY_MAX_COMPLETION_TOKENS = 20_000
DEFAULT_TARGET_REDUCE_PROMPT_TOKENS = 20_000
DEFAULT_REDUCE_PROMPT_SAFETY_TOKENS = 1_500
DEFAULT_SEGMENT_MAX_COMPLETION_TOKENS = 3_000
MAX_REDUCTION_LEVELS = 8
DEFAULT_INPUT_COST_PER_1M = 2.50
DEFAULT_OUTPUT_COST_PER_1M = 15.00
DEFAULT_CACHED_INPUT_COST_PER_1M = 0.25
SYSTEM_INSTRUCTIONS = (
    "You are an IPCC terminology consistency analyst. "
    "Use only provided definitions and evidence. "
    "Do not use external sources."
)
CONCLUSION_LABELS = (
    "Consistent use",
    "Mostly consistent with minor ambiguity",
    "Potentially inconsistent, needs substantive review",
)
SUMMARY_CONTEXTS_HEADING = "Contexts of use"
SUMMARY_POTENTIAL_ISSUES_HEADING = "Potential issues needing substantive review"
SUMMARY_CONCLUSION_HEADING = "Conclusion"


@dataclass(frozen=True)
class AggregatedOccurrence:
    """One deduplicated evidence unit with frequency and provenance."""

    report: str
    source_label: str
    node_code: str
    node_id: str
    sentence: str
    matched_names: tuple[str, ...]
    frequency: int


@dataclass(frozen=True)
class TermInput:
    """All deterministic inputs required for one term's analysis."""

    term_key: str
    term: str
    aliases: tuple[str, ...]
    explanation: str
    source: str
    parent_terms: tuple[str, ...]
    child_terms: tuple[str, ...]
    occurrences: tuple[AggregatedOccurrence, ...]


@dataclass
class UsageTotals:
    """Accumulate model token usage for one run."""

    prompt_tokens: int = 0
    completion_tokens: int = 0
    cached_prompt_tokens: int = 0
    reasoning_tokens: int = 0

    def add(self, usage: dict[str, int]) -> None:
        self.prompt_tokens += usage.get("prompt_tokens", 0)
        self.completion_tokens += usage.get("completion_tokens", 0)
        self.cached_prompt_tokens += usage.get("cached_prompt_tokens", 0)
        self.reasoning_tokens += usage.get("reasoning_tokens", 0)

    @property
    def uncached_prompt_tokens(self) -> int:
        return max(0, self.prompt_tokens - self.cached_prompt_tokens)


class EmptyModelOutputError(ValueError):
    """An empty model response with metadata needed for a safe retry decision."""

    def __init__(self, message: str, finish_reason: str | None, usage: dict[str, int]) -> None:
        super().__init__(message)
        self.finish_reason = finish_reason
        self.usage = usage


class StatusLogger:
    """Write detailed run events to the log with optional concise console progress."""

    def __init__(self, log_path: Path) -> None:
        self._log_path = log_path
        self._log_path.parent.mkdir(parents=True, exist_ok=True)
        has_existing_content = self._log_path.is_file() and self._log_path.stat().st_size > 0
        self._handle = self._log_path.open("a", encoding="utf-8")
        if has_existing_content:
            self._handle.write("\n-----\n")
            self._handle.flush()

    def log(
        self,
        message: str,
        *,
        is_error: bool = False,
        event: str = "log",
        fields: dict[str, Any] | None = None,
        console_message: str | None = None,
    ) -> None:
        timestamp = utc_timestamp()
        level = "ERROR" if is_error else "INFO"
        extras = ""
        if fields:
            serialized = ", ".join(f"{key}={json.dumps(value, ensure_ascii=False)}" for key, value in fields.items())
            extras = f" | {serialized}"
        line = f"[{timestamp}] {level} {event}: {message}{extras}"
        if console_message is not None:
            print(console_message, flush=True)
        self._handle.write(f"{line}\n")
        self._handle.flush()

    def close(self) -> None:
        self._handle.close()


def utc_timestamp() -> str:
    return datetime.now(timezone.utc).replace(microsecond=0).isoformat().replace("+00:00", "Z")


def normalize_text(value: object) -> str:
    return re.sub(r"\s+", " ", str(value or "")).strip()


def sha256_text(value: str) -> str:
    return hashlib.sha256(value.encode("utf-8")).hexdigest()


def stable_json(value: object) -> str:
    return json.dumps(value, ensure_ascii=False, sort_keys=True, separators=(",", ":"))


def parse_response_json(text: str) -> dict[str, Any]:
    """Parse model output and tolerate accidental leading or trailing text."""
    raw = normalize_text(text)
    if not raw:
        raise ValueError("Model returned empty output.")

    try:
        payload = json.loads(raw)
        if isinstance(payload, dict):
            return payload
    except json.JSONDecodeError:
        pass

    start = raw.find("{")
    end = raw.rfind("}")
    if start != -1 and end != -1 and end > start:
        payload = json.loads(raw[start : end + 1])
        if isinstance(payload, dict):
            return payload
    raise ValueError("Model output is not a valid JSON object.")


def usage_from_response(response: Any) -> dict[str, int]:
    """Extract token usage counters from one Azure OpenAI response."""
    usage = getattr(response, "usage", None)
    if usage is None:
        return {
            "prompt_tokens": 0,
            "completion_tokens": 0,
            "cached_prompt_tokens": 0,
            "reasoning_tokens": 0,
        }

    prompt_details = getattr(usage, "prompt_tokens_details", None)
    completion_details = getattr(usage, "completion_tokens_details", None)
    return {
        "prompt_tokens": int(getattr(usage, "prompt_tokens", 0) or 0),
        "completion_tokens": int(getattr(usage, "completion_tokens", 0) or 0),
        "cached_prompt_tokens": int(getattr(prompt_details, "cached_tokens", 0) or 0),
        "reasoning_tokens": int(getattr(completion_details, "reasoning_tokens", 0) or 0),
    }


def estimate_cost_usd(
    usage: UsageTotals,
    input_cost_per_1m: float,
    output_cost_per_1m: float,
    cached_input_cost_per_1m: float,
) -> dict[str, float]:
    """Estimate run cost from token usage and per-1M token prices in USD."""
    input_cost = (usage.uncached_prompt_tokens / 1_000_000) * input_cost_per_1m
    cached_input_cost = (usage.cached_prompt_tokens / 1_000_000) * cached_input_cost_per_1m
    output_cost = (usage.completion_tokens / 1_000_000) * output_cost_per_1m
    total_cost = input_cost + cached_input_cost + output_cost
    return {
        "input_cost_usd": input_cost,
        "cached_input_cost_usd": cached_input_cost,
        "output_cost_usd": output_cost,
        "total_cost_usd": total_cost,
    }


def new_results_artifact() -> dict[str, Any]:
    timestamp = utc_timestamp()
    return {
        "format_version": FORMAT_VERSION,
        "pipeline_version": PIPELINE_VERSION,
        "model": AZURE_OPENAI_DEPLOYMENT,
        "created_at": timestamp,
        "updated_at": timestamp,
        "summaries": {},
        "failures": {},
    }


def load_results_artifact(output_path: Path) -> dict[str, Any]:
    """Load an existing artifact for an explicit resumable run."""
    if not output_path.exists():
        return new_results_artifact()

    try:
        payload = json.loads(output_path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as error:
        raise ValueError(f"Unable to read existing summary artifact: {output_path}") from error

    if not isinstance(payload, dict) or payload.get("format_version") != FORMAT_VERSION:
        raise ValueError(
            f"{output_path} does not use supported format version {FORMAT_VERSION}. "
            "Run without --resume to replace it."
        )
    if payload.get("model") != AZURE_OPENAI_DEPLOYMENT:
        raise ValueError(
            f"{output_path} was generated with {payload.get('model')!r}. "
            f"Run without --resume to regenerate with {AZURE_OPENAI_DEPLOYMENT}."
        )
    if not isinstance(payload.get("summaries"), dict) or not isinstance(payload.get("failures"), dict):
        raise ValueError(f"{output_path} must contain summaries and failures objects.")

    payload.setdefault("pipeline_version", PIPELINE_VERSION)
    return payload


def write_results_artifact(output_path: Path, payload: dict[str, Any]) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    payload["updated_at"] = utc_timestamp()
    temp_path = output_path.with_suffix(f"{output_path.suffix}.tmp")
    temp_path.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2, sort_keys=True) + "\n",
        encoding="utf-8",
    )
    temp_path.replace(output_path)


def has_fresh_summary(record: object, input_hash: str) -> bool:
    if not isinstance(record, dict):
        return False
    summary = record.get("summary")
    if not isinstance(summary, str) or not summary.strip():
        return False
    return record.get("input_hash") == input_hash and record.get("pipeline_version") == PIPELINE_VERSION


def local_occurrence_summary(term: str, occurrence_count: int) -> str:
    if occurrence_count == 0:
        return f"The term '{term}' is not used in the provided report corpus."
    raise ValueError("Local occurrence summaries are only valid for zero occurrences.")


def get_azure_openai_client() -> AzureOpenAI:
    dotenv.load_dotenv(REPO_ROOT / ".env", override=False)
    api_key = os.getenv("AZURE_API_KEY", "").strip()
    endpoint = os.getenv("AZURE_OPENAI_ENDPOINT", "").strip()
    if not api_key:
        raise RuntimeError("AZURE_API_KEY is missing. Set it in the local .env file before generating summaries.")
    if not endpoint:
        raise RuntimeError(
            "AZURE_OPENAI_ENDPOINT is missing. Set it in the local .env file before generating summaries."
        )

    return AzureOpenAI(
        api_version=AZURE_OPENAI_API_VERSION,
        azure_endpoint=endpoint,
        api_key=api_key,
    )


def request_json(
    client: AzureOpenAI,
    prompt: str,
    max_completion_tokens: int,
    reasoning_effort: str,
    request_label: str,
) -> tuple[dict[str, Any], dict[str, int]]:
    response = client.chat.completions.create(
        messages=[
            {"role": "system", "content": SYSTEM_INSTRUCTIONS},
            {"role": "user", "content": prompt},
        ],
        model=AZURE_OPENAI_DEPLOYMENT,
        reasoning_effort=reasoning_effort,
        max_completion_tokens=max_completion_tokens,
        response_format={"type": "json_object"},
    )
    choices = getattr(response, "choices", None)
    if not choices:
        raise RuntimeError(f"{request_label} returned no completion choices.")

    choice = choices[0]
    message = getattr(choice, "message", None)
    content = getattr(message, "content", None)
    if not isinstance(content, str) or not content.strip():
        usage = usage_from_response(response)
        finish_reason = getattr(choice, "finish_reason", None)
        details = [
            f"finish_reason={finish_reason!r}",
            f"prompt_tokens={usage['prompt_tokens']}",
            f"completion_tokens={usage['completion_tokens']}",
            f"reasoning_tokens={usage['reasoning_tokens']}",
        ]
        refusal = getattr(message, "refusal", None)
        if refusal:
            details.append(f"refusal={refusal!r}")
        raise EmptyModelOutputError(
            f"{request_label} returned empty output ({', '.join(details)}).",
            finish_reason=finish_reason,
            usage=usage,
        )

    try:
        parsed = parse_response_json(content)
    except ValueError as error:
        raise ValueError(f"{request_label}: {error}") from error
    return parsed, usage_from_response(response)


def report_name_from_source_label(source_label: str) -> str:
    return source_label.split(",", maxsplit=1)[0].strip()


def deduplicate_occurrences(matches: list[GlossaryOccurrence]) -> list[AggregatedOccurrence]:
    grouped: dict[tuple[str, str, tuple[str, ...]], AggregatedOccurrence] = {}
    frequencies: Counter[tuple[str, str, tuple[str, ...]]] = Counter()

    for match in matches:
        normalized_names = tuple(dict.fromkeys(name.casefold() for name in match.matched_names))
        key = (match.node_id, normalize_text(match.sentence).casefold(), normalized_names)
        frequencies[key] += 1
        if key in grouped:
            continue
        grouped[key] = AggregatedOccurrence(
            report=report_name_from_source_label(match.source_label),
            source_label=match.source_label,
            node_code=match.node_code,
            node_id=match.node_id,
            sentence=normalize_text(match.sentence),
            matched_names=tuple(dict.fromkeys(normalize_text(name) for name in match.matched_names if normalize_text(name))),
            frequency=1,
        )

    deduplicated = []
    for key, occurrence in grouped.items():
        deduplicated.append(
            AggregatedOccurrence(
                report=occurrence.report,
                source_label=occurrence.source_label,
                node_code=occurrence.node_code,
                node_id=occurrence.node_id,
                sentence=occurrence.sentence,
                matched_names=occurrence.matched_names,
                frequency=frequencies[key],
            )
        )

    deduplicated.sort(key=lambda item: (item.report, item.node_code, item.node_id, item.sentence.casefold()))
    return deduplicated


def chunk_occurrences(occurrences: list[AggregatedOccurrence], chunk_size: int) -> list[list[AggregatedOccurrence]]:
    by_report: dict[str, list[AggregatedOccurrence]] = defaultdict(list)
    for occurrence in occurrences:
        by_report[occurrence.report].append(occurrence)

    report_names = sorted(by_report)
    chunks: list[list[AggregatedOccurrence]] = []
    while any(by_report.values()):
        chunk: list[AggregatedOccurrence] = []
        for report_name in report_names:
            report_items = by_report[report_name]
            if not report_items:
                continue
            budget = max(1, chunk_size // max(1, len(report_names)))
            take = min(len(report_items), budget, chunk_size - len(chunk))
            if take <= 0:
                continue
            chunk.extend(report_items[:take])
            del report_items[:take]
            if len(chunk) >= chunk_size:
                break
        if not chunk:
            break
        chunks.append(chunk)
    return chunks


def occurrence_statistics(term_input: TermInput) -> dict[str, Any]:
    by_report = Counter(item.report for item in term_input.occurrences)
    by_name = Counter(name for item in term_input.occurrences for name in item.matched_names)
    return {
        "total_occurrences": sum(item.frequency for item in term_input.occurrences),
        "unique_occurrences": len(term_input.occurrences),
        "reports": dict(sorted(by_report.items())),
        "matched_name_distribution": dict(by_name.most_common()),
    }


def term_input_hash(term_input: TermInput, prompt_hash: str, corpus_hash: str) -> str:
    payload = {
        "pipeline_version": PIPELINE_VERSION,
        "model": AZURE_OPENAI_DEPLOYMENT,
        "term": term_input.term,
        "aliases": term_input.aliases,
        "definition": term_input.explanation,
        "source": term_input.source,
        "parents": term_input.parent_terms,
        "children": term_input.child_terms,
        "occurrences": [asdict(item) for item in term_input.occurrences],
        "prompt_hash": prompt_hash,
        "corpus_hash": corpus_hash,
    }
    return sha256_text(stable_json(payload))


def format_occurrence_rows(occurrences: list[AggregatedOccurrence]) -> str:
    rows = [
        {
            "report": item.report,
            "source_label": item.source_label,
            "node_code": item.node_code,
            "node_id": item.node_id,
            "sentence": item.sentence,
            "matched_names": list(item.matched_names),
            "frequency": item.frequency,
        }
        for item in occurrences
    ]
    return json.dumps(rows, ensure_ascii=False, indent=2)


def estimate_tokens_from_text(text: str) -> int:
    """Approximate token count using a conservative chars-per-token heuristic."""
    return max(1, (len(text) + 3) // 4)


def reduce_input_token_budget(
    target_reduce_prompt_tokens: int,
    reduce_max_completion_tokens: int,
    reduce_prompt_safety_tokens: int,
) -> int:
    """Return the portion of a reduce request available for prompt content."""
    return max(1, target_reduce_prompt_tokens - reduce_max_completion_tokens - reduce_prompt_safety_tokens)


def chunk_analysis_records_for_reduce(
    term_input: TermInput,
    stats: dict[str, Any],
    analysis_records: list[dict[str, Any]],
    target_reduce_prompt_tokens: int,
    reduce_max_completion_tokens: int,
    reduce_prompt_safety_tokens: int,
) -> list[list[dict[str, Any]]]:
    """Group analysis records so every reduce prompt fits its token budget."""
    if not analysis_records:
        return []

    input_budget = reduce_input_token_budget(
        target_reduce_prompt_tokens,
        reduce_max_completion_tokens,
        reduce_prompt_safety_tokens,
    )
    empty_prompt_tokens = estimate_tokens_from_text(build_reduce_prompt(term_input, stats, []))
    if empty_prompt_tokens > input_budget:
        raise ValueError("The configured reduce prompt budget is too small for the reduce instructions.")

    batches: list[list[dict[str, Any]]] = []
    batch: list[dict[str, Any]] = []
    for record in analysis_records:
        candidate = [*batch, record]
        candidate_tokens = estimate_tokens_from_text(build_reduce_prompt(term_input, stats, candidate))
        if batch and candidate_tokens > input_budget:
            batches.append(batch)
            batch = [record]
            candidate_tokens = estimate_tokens_from_text(build_reduce_prompt(term_input, stats, batch))

        if candidate_tokens > input_budget:
            raise ValueError("A single analysis record exceeds the configured reduce prompt budget.")
        if not batch or batch[-1] is not record:
            batch = candidate

    if batch:
        batches.append(batch)
    return batches


def choose_dynamic_chunk_size(
    term_input: TermInput,
    base_chunk_size: int,
    target_map_prompt_tokens: int,
    map_max_completion_tokens: int,
    map_prompt_safety_tokens: int,
    min_dynamic_chunk_size: int,
) -> int:
    """Choose chunk size so map prompts stay within a safe token budget.

    This keeps full evidence coverage while adapting per term to avoid
    prompt overflow for extremely high-frequency terms.
    """
    empty_prompt = build_map_prompt(term_input, 1, 1, [])
    empty_tokens = estimate_tokens_from_text(empty_prompt)

    sample_size = min(20, len(term_input.occurrences))
    if sample_size == 0:
        return base_chunk_size

    sample_prompt = build_map_prompt(term_input, 1, 1, list(term_input.occurrences[:sample_size]))
    sample_tokens = estimate_tokens_from_text(sample_prompt)
    variable_tokens = max(1, sample_tokens - empty_tokens)
    tokens_per_row = max(1, variable_tokens // sample_size)

    available_input_tokens = max(
        1,
        target_map_prompt_tokens - map_max_completion_tokens - map_prompt_safety_tokens,
    )
    estimated_rows = max(1, (available_input_tokens - empty_tokens) // tokens_per_row)

    if estimated_rows < min_dynamic_chunk_size:
        estimated_rows = min_dynamic_chunk_size

    return max(1, min(base_chunk_size, estimated_rows))


def build_map_prompt(
    term_input: TermInput,
    chunk_index: int,
    chunk_total: int,
    chunk: list[AggregatedOccurrence],
) -> str:
    return f"""
Analyze terminology consistency for one chunk. Return JSON only.

Term: {term_input.term}
Definition: {term_input.explanation or 'Definition not available.'}
Source: {term_input.source or 'Not specified'}
Parent terms: {', '.join(term_input.parent_terms) if term_input.parent_terms else 'None'}
Child terms: {', '.join(term_input.child_terms) if term_input.child_terms else 'None'}
Aliases to treat as equivalent names: {', '.join(term_input.aliases) if term_input.aliases else 'None'}
Chunk: {chunk_index}/{chunk_total}

Occurrence rows JSON:
{format_occurrence_rows(chunk)}

Output schema (JSON object):
{{
  "chunk_index": {chunk_index},
  "coverage": {{"rows": <int>, "frequency_sum": <int>}},
  "contexts": [
    {{"name": "<string>", "description": "<string>", "reports": ["Chapter 1"], "evidence": [{{"node_code": "<string>", "quote": "<string>"}}]}}
  ],
  "alignment_signals": ["<string>"],
  "issue_candidates": [
    {{"type": "contradiction|ambiguity|scope-shift|alias-risk", "description": "<string>", "severity": "low|medium|high", "evidence": [{{"node_code": "<string>", "quote": "<string>"}}]}}
  ],
  "ambiguous_matches": [
    {{"matched_name": "<string>", "node_code": "<string>", "reason": "<string>"}}
  ]
}}

Rules:
- Use only provided rows and definition.
- Different context alone is not inconsistency.
- Report a candidate issue only with explicit evidence.
- Keep evidence quotes brief and exact.
""".strip()


def build_reduce_prompt(
    term_input: TermInput,
    stats: dict[str, Any],
    analysis_records: list[dict[str, Any]],
    stage: str = "final synthesis",
) -> str:
    return f"""
Synthesize a terminology-consistency assessment from analysis records.
Return JSON only.

Term: {term_input.term}
Definition: {term_input.explanation or 'Definition not available.'}
Aliases considered equivalent: {', '.join(term_input.aliases) if term_input.aliases else 'None'}
Stage: {stage}
Statistics JSON:
{json.dumps(stats, ensure_ascii=False, indent=2)}

Analysis record JSON list:
{json.dumps(analysis_records, ensure_ascii=False, indent=2)}

Output schema (JSON object):
{{
  "contexts": [
    {{"name": "<string>", "description": "<string>", "reports": ["Chapter 1"], "sample_ids": ["<node_code>"]}}
  ],
  "issues": [
    {{"label": "<string>", "severity": "low|medium|high", "description": "<string>", "evidence": [{{"node_code": "<string>", "quote": "<string>"}}]}}
    ],
  "conclusion": {{
    "label": "Consistent use|Mostly consistent with minor ambiguity|Potentially inconsistent, needs substantive review",
    "rationale": "<string>"
  }}
}}

Rules:
- Respect exhaustive coverage metrics.
- Records can be individual evidence-chunk analyses or earlier segment syntheses.
- Keep issues only when evidence is concrete.
- Avoid claiming inconsistency from style or context differences alone.
- Return at most 6 contexts, 8 issues, 3 sample IDs per context, and 3 evidence items per issue.
""".strip()


def evidence_line(item: dict[str, Any]) -> str:
    node_code = normalize_text(item.get("node_code"))
    quote = normalize_text(item.get("quote"))
    if node_code and quote:
        return f"- [{node_code}] \"{quote}\""
    if node_code:
        return f"- [{node_code}]"
    if quote:
        return f"- \"{quote}\""
    return ""


def format_context_statement_links(sample_ids: list[str]) -> str:
    """Render sample IDs as plain references in one parenthesized list."""
    cleaned_ids = [normalize_text(item) for item in sample_ids if normalize_text(item)]
    if not cleaned_ids:
        return ""
    labels = [f"[{item}]" for item in cleaned_ids]
    return f"({', '.join(labels)})"


def render_summary_markdown(term: str, stats: dict[str, Any], result: dict[str, Any]) -> str:
    contexts = result.get("contexts") if isinstance(result.get("contexts"), list) else []
    issues = result.get("issues") if isinstance(result.get("issues"), list) else []
    conclusion = result.get("conclusion") if isinstance(result.get("conclusion"), dict) else {}

    lines: list[str] = []
    lines.append(f"### {SUMMARY_CONTEXTS_HEADING}")
    if contexts:
        for context in contexts[:6]:
            if not isinstance(context, dict):
                continue
            name = normalize_text(context.get("name")) or "Context"
            description = normalize_text(context.get("description"))
            sample_ids = context.get("sample_ids") if isinstance(context.get("sample_ids"), list) else []
            id_links = format_context_statement_links(sample_ids)
            if id_links and description:
                lines.append(f"- **{name}** {id_links}: {description}")
            elif id_links:
                lines.append(f"- **{name}** {id_links}")
            elif description:
                lines.append(f"- **{name}**: {description}")
            else:
                lines.append(f"- **{name}**")
    else:
        lines.append("- No recurring contexts extracted.")

    lines.append("")
    lines.append(f"### {SUMMARY_POTENTIAL_ISSUES_HEADING}")
    if issues:
        for issue in issues[:8]:
            if not isinstance(issue, dict):
                continue
            label = normalize_text(issue.get("label")) or "Potential issue"
            severity = normalize_text(issue.get("severity"))
            description = normalize_text(issue.get("description"))
            heading = f"- **{label}**"
            if severity:
                heading += f" ({severity})"
            if description:
                heading += f": {description}"
            lines.append(heading)
            evidence = issue.get("evidence") if isinstance(issue.get("evidence"), list) else []
            for evidence_item in evidence[:5]:
                if isinstance(evidence_item, dict):
                    line = evidence_line(evidence_item)
                    if line:
                        lines.append(f"  {line}")
    else:
        lines.append("- None identified.")

    lines.append("")
    lines.append(f"### {SUMMARY_CONCLUSION_HEADING}")
    label = normalize_text(conclusion.get("label"))
    if label not in CONCLUSION_LABELS:
        label = "Mostly consistent with minor ambiguity"
    rationale = normalize_text(conclusion.get("rationale")) or "No rationale returned by model."
    lines.append(f"**{label}**")
    lines.append(rationale)

    return "\n".join(lines).strip()


def remove_stale_records(payload: dict[str, Any], current_terms: set[str]) -> bool:
    changed = False
    for key in ("summaries", "failures"):
        records = payload.get(key)
        if not isinstance(records, dict):
            continue
        for term_key in list(records):
            if term_key not in current_terms:
                del records[term_key]
                changed = True
    return changed


def load_term_inputs(
    source_json_path: Path,
    report_html_path: Path,
    glossary_path: Path,
) -> tuple[dict[str, TermInput], str]:
    payload = json.loads(source_json_path.read_text(encoding="utf-8"))
    if not isinstance(payload, dict):
        raise ValueError("The source JSON root must be an object.")
    canonical_report_data(payload)

    glossary = load_revised_glossary(glossary_path)
    match_map = build_glossary_match_map(glossary)
    report_html = report_html_path.read_text(encoding="utf-8")
    node_codes = report_node_codes(report_html)
    occurrences_by_term = full_report_term_occurrences(payload, glossary, match_map, node_codes)

    term_inputs: dict[str, TermInput] = {}
    for term_key, entry in glossary.items():
        occurrences = tuple(deduplicate_occurrences(occurrences_by_term.get(term_key, [])))
        term_inputs[term_key] = TermInput(
            term_key=term_key,
            term=entry.term,
            aliases=entry.aliases,
            explanation=entry.explanation,
            source=entry.source,
            parent_terms=entry.parent_terms,
            child_terms=entry.child_terms,
            occurrences=occurrences,
        )

    corpus_hash = sha256_text(stable_json({
        "source_json": sha256_text(source_json_path.read_text(encoding="utf-8")),
        "report_html": sha256_text(report_html),
        "glossary_path": str(glossary_path),
        "glossary_terms": len(glossary),
    }))
    return term_inputs, corpus_hash


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Generate GPT-5.4 terminology-consistency summaries from full report occurrences."
    )
    parser.add_argument("--source-json", type=Path, default=DEFAULT_SOURCE_JSON)
    parser.add_argument("--report-html", type=Path, default=DEFAULT_OUTPUT_HTML)
    parser.add_argument("--glossary", type=Path, default=DEFAULT_GLOSSARY_PATH)
    parser.add_argument("--prompt", type=Path, default=DEFAULT_PROMPT_PATH)
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT_PATH)
    parser.add_argument("--log", type=Path, default=DEFAULT_LOG_PATH)
    result_mode = parser.add_mutually_exclusive_group()
    result_mode.add_argument(
        "--resume",
        action="store_true",
        help="Reuse fresh summaries from the existing JSON artifact instead of replacing it.",
    )
    result_mode.add_argument("--overwrite", action="store_true", help=argparse.SUPPRESS)
    parser.add_argument("--limit", type=int, metavar="COUNT")
    parser.add_argument("--term", action="append", metavar="TERM")
    parser.add_argument("--pause-seconds", type=float, default=0.0)
    parser.add_argument("--chunk-size", type=int, default=80, metavar="COUNT")
    parser.add_argument("--map-max-completion-tokens", type=int, default=5000)
    parser.add_argument("--reduce-max-completion-tokens", type=int, default=DEFAULT_REDUCE_MAX_COMPLETION_TOKENS)
    parser.add_argument("--segment-max-completion-tokens", type=int, default=DEFAULT_SEGMENT_MAX_COMPLETION_TOKENS)
    parser.add_argument("--target-map-prompt-tokens", type=int, default=14000)
    parser.add_argument("--map-prompt-safety-tokens", type=int, default=1500)
    parser.add_argument("--target-reduce-prompt-tokens", type=int, default=DEFAULT_TARGET_REDUCE_PROMPT_TOKENS)
    parser.add_argument("--reduce-prompt-safety-tokens", type=int, default=DEFAULT_REDUCE_PROMPT_SAFETY_TOKENS)
    parser.add_argument("--min-dynamic-chunk-size", type=int, default=10, metavar="COUNT")
    parser.add_argument("--input-cost-per-1m", type=float, default=DEFAULT_INPUT_COST_PER_1M)
    parser.add_argument("--output-cost-per-1m", type=float, default=DEFAULT_OUTPUT_COST_PER_1M)
    parser.add_argument("--cached-input-cost-per-1m", type=float, default=DEFAULT_CACHED_INPUT_COST_PER_1M)
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    logger = StatusLogger(args.log.expanduser())
    logger.log(
        "Run started",
        event="run_start",
        fields={
            "log_path": str(args.log.expanduser()),
            "output_path": str(args.output.expanduser()),
            "source_json": str(args.source_json.expanduser()),
            "report_html": str(args.report_html.expanduser()),
            "glossary": str(args.glossary.expanduser()),
            "prompt": str(args.prompt.expanduser()),
            "model": AZURE_OPENAI_DEPLOYMENT,
            "api_version": AZURE_OPENAI_API_VERSION,
            "map_reasoning_effort": AZURE_OPENAI_MAP_REASONING_EFFORT,
            "reduce_reasoning_effort": AZURE_OPENAI_REASONING_EFFORT,
            "result_mode": "resume" if args.resume else "replace",
        },
    )
    logger.log(
        "Pricing configured",
        event="pricing",
        fields={
            "input_cost_per_1m": args.input_cost_per_1m,
            "cached_input_cost_per_1m": args.cached_input_cost_per_1m,
            "output_cost_per_1m": args.output_cost_per_1m,
        },
    )
    if args.pause_seconds < 0:
        logger.log("--pause-seconds must be zero or greater.", is_error=True)
        logger.close()
        return 2
    if args.limit is not None and args.limit < 1:
        logger.log("--limit must be at least 1.", is_error=True)
        logger.close()
        return 2
    if args.chunk_size < 1:
        logger.log("--chunk-size must be at least 1.", is_error=True)
        logger.close()
        return 2
    if args.map_max_completion_tokens < 1:
        logger.log("--map-max-completion-tokens must be at least 1.", is_error=True)
        logger.close()
        return 2
    if args.reduce_max_completion_tokens < 1:
        logger.log("--reduce-max-completion-tokens must be at least 1.", is_error=True)
        logger.close()
        return 2
    if args.segment_max_completion_tokens < 1:
        logger.log("--segment-max-completion-tokens must be at least 1.", is_error=True)
        logger.close()
        return 2
    if args.target_map_prompt_tokens < 1:
        logger.log("--target-map-prompt-tokens must be at least 1.", is_error=True)
        logger.close()
        return 2
    if args.map_prompt_safety_tokens < 0:
        logger.log("--map-prompt-safety-tokens must be zero or greater.", is_error=True)
        logger.close()
        return 2
    if args.target_reduce_prompt_tokens < 1:
        logger.log("--target-reduce-prompt-tokens must be at least 1.", is_error=True)
        logger.close()
        return 2
    if args.reduce_prompt_safety_tokens < 0:
        logger.log("--reduce-prompt-safety-tokens must be zero or greater.", is_error=True)
        logger.close()
        return 2
    if args.min_dynamic_chunk_size < 1:
        logger.log("--min-dynamic-chunk-size must be at least 1.", is_error=True)
        logger.close()
        return 2
    if args.input_cost_per_1m < 0 or args.output_cost_per_1m < 0 or args.cached_input_cost_per_1m < 0:
        logger.log("Token cost parameters must be zero or greater.", is_error=True)
        logger.close()
        return 2

    try:
        prompt_template = args.prompt.read_text(encoding="utf-8")
        if not prompt_template.strip():
            raise ValueError("The approved term-usage summary prompt is empty.")
        term_inputs, corpus_hash = load_term_inputs(
            args.source_json.expanduser(),
            args.report_html.expanduser(),
            args.glossary.expanduser(),
        )
        output_path = args.output.expanduser()
        payload = load_results_artifact(output_path) if args.resume else new_results_artifact()
        payload["pipeline_version"] = PIPELINE_VERSION
        if args.resume:
            remove_stale_records(payload, set(term_inputs))
        else:
            write_results_artifact(output_path, payload)
    except (FileNotFoundError, OSError, RuntimeError, ValueError, json.JSONDecodeError) as error:
        logger.log(str(error), is_error=True)
        logger.close()
        return 1

    summaries: dict[str, Any] = payload["summaries"]
    failures: dict[str, Any] = payload["failures"]
    logger.log(
        "Resuming existing results artifact" if args.resume else "Replaced results artifact",
        event="artifact_mode",
        fields={"output_path": str(output_path), "result_mode": "resume" if args.resume else "replace"},
    )
    all_terms = sorted(term_inputs.items(), key=lambda item: item[1].term.casefold())

    requested_term_keys: set[str]
    if args.term:
        requested_term_keys = {normalize_text(term).casefold() for term in args.term}
    else:
        requested_term_keys = set()

    unknown_term_keys = requested_term_keys - set(term_inputs)
    if unknown_term_keys:
        logger.log(f"Unknown glossary term: {', '.join(sorted(unknown_term_keys))}", is_error=True)
        logger.close()
        return 2

    terms = (
        [item for item in all_terms if item[0] in requested_term_keys]
        if requested_term_keys
        else all_terms[: args.limit] if args.limit is not None else all_terms
    )

    prompt_hash = sha256_text(prompt_template)
    client: AzureOpenAI | None = None
    completed = 0
    failed = 0
    run_usage = UsageTotals()

    logger.log(
        "Prepared run scope",
        event="scope",
        fields={
            "selected_terms": len(terms),
            "all_terms": len(all_terms),
            "requested_terms": sorted(requested_term_keys),
            "limit": args.limit,
        },
    )

    for index, (term_key, term_input) in enumerate(terms, start=1):
        progress = f"{index}/{len(terms)}"
        occurrence_count = sum(item.frequency for item in term_input.occurrences)
        logger.log(
            f"Processing {progress}: {term_input.term}",
            event="term_start",
            fields={
                "term": term_input.term,
                "term_key": term_key,
                "term_index": index,
                "term_total": len(terms),
            },
            console_message=(
                f"Processing {progress} (Term: {term_input.term}, Count: {occurrence_count})"
            ),
        )
        input_hash = term_input_hash(term_input, prompt_hash=prompt_hash, corpus_hash=corpus_hash)

        if args.resume and has_fresh_summary(summaries.get(term_key), input_hash):
            completed += 1
            logger.log(
                "Reused current term summary",
                event="term_reused",
                fields={"term": term_input.term, "input_hash": input_hash},
                console_message=f"{progress} successful",
            )
            continue

        if occurrence_count == 0:
            summary = local_occurrence_summary(term_input.term, occurrence_count)
            summaries[term_key] = {
                "term": term_input.term,
                "occurrence_count": occurrence_count,
                "summary": summary,
                "summary_source": "local occurrence notice",
                "generated_at": utc_timestamp(),
                "pipeline_version": PIPELINE_VERSION,
                "input_hash": input_hash,
                "corpus_hash": corpus_hash,
                "prompt_hash": prompt_hash,
            }
            failures.pop(term_key, None)
            write_results_artifact(args.output, payload)
            completed += 1
            logger.log(
                "Saved local occurrence notice",
                event="term_complete",
                fields={
                    "term": term_input.term,
                    "occurrence_count": occurrence_count,
                    "summary_source": "local occurrence notice",
                    "input_hash": input_hash,
                },
                console_message=f"{progress} successful",
            )
            continue

        effective_chunk_size = 0
        chunks: list[list[AggregatedOccurrence]] = []
        chunk_records: list[dict[str, Any]] = []
        reduction_stages: list[dict[str, Any]] = []
        current_stage = "preparing map chunks"
        try:
            if client is None:
                client = get_azure_openai_client()

            effective_chunk_size = choose_dynamic_chunk_size(
                term_input=term_input,
                base_chunk_size=args.chunk_size,
                target_map_prompt_tokens=args.target_map_prompt_tokens,
                map_max_completion_tokens=args.map_max_completion_tokens,
                map_prompt_safety_tokens=args.map_prompt_safety_tokens,
                min_dynamic_chunk_size=args.min_dynamic_chunk_size,
            )
            chunks = chunk_occurrences(list(term_input.occurrences), effective_chunk_size)
            map_outputs: list[dict[str, Any]] = []
            for chunk_index, chunk in enumerate(chunks, start=1):
                current_stage = f"map chunk {chunk_index}/{len(chunks)}"
                map_prompt = build_map_prompt(term_input, chunk_index, len(chunks), chunk)
                map_result = request_json(
                    client,
                    map_prompt,
                    max_completion_tokens=args.map_max_completion_tokens,
                    reasoning_effort=AZURE_OPENAI_MAP_REASONING_EFFORT,
                    request_label=current_stage,
                )
                map_json, map_usage = map_result
                run_usage.add(map_usage)
                logger.log(
                    "Map chunk analyzed",
                    event="map_chunk",
                    fields={
                        "term": term_input.term,
                        "chunk_index": chunk_index,
                        "chunk_total": len(chunks),
                        "rows": len(chunk),
                        "frequency_sum": sum(item.frequency for item in chunk),
                        "prompt_tokens": map_usage.get("prompt_tokens", 0),
                        "completion_tokens": map_usage.get("completion_tokens", 0),
                        "cached_prompt_tokens": map_usage.get("cached_prompt_tokens", 0),
                        "reasoning_tokens": map_usage.get("reasoning_tokens", 0),
                    },
                )
                map_outputs.append(map_json)
                chunk_records.append(
                    {
                        "chunk_index": chunk_index,
                        "rows": len(chunk),
                        "frequency_sum": sum(item.frequency for item in chunk),
                        "chunk_hash": sha256_text(stable_json([asdict(item) for item in chunk])),
                    }
                )
                if args.pause_seconds:
                    time.sleep(args.pause_seconds)

            stats = occurrence_statistics(term_input)
            analysis_records = map_outputs
            reduction_level = 0
            while True:
                current_stage = "planning final reduction"
                try:
                    final_batches = chunk_analysis_records_for_reduce(
                        term_input,
                        stats,
                        analysis_records,
                        target_reduce_prompt_tokens=args.target_reduce_prompt_tokens,
                        reduce_max_completion_tokens=args.reduce_max_completion_tokens,
                        reduce_prompt_safety_tokens=args.reduce_prompt_safety_tokens,
                    )
                    final_planning_error = ""
                except ValueError as error:
                    final_batches = []
                    final_planning_error = str(error)
                if len(final_batches) == 1:
                    current_stage = "final reduction"
                    reduce_prompt = build_reduce_prompt(
                        term_input,
                        stats,
                        final_batches[0],
                        stage="final synthesis",
                    )
                    try:
                        reduce_result = request_json(
                            client,
                            reduce_prompt,
                            max_completion_tokens=args.reduce_max_completion_tokens,
                            reasoning_effort=AZURE_OPENAI_REASONING_EFFORT,
                            request_label=current_stage,
                        )
                    except EmptyModelOutputError as error:
                        run_usage.add(error.usage)
                        if (
                            args.reduce_max_completion_tokens == DEFAULT_REDUCE_MAX_COMPLETION_TOKENS
                            and error.finish_reason == "length"
                        ):
                            current_stage = "final reduction retry (20000 tokens)"
                            logger.log(
                                "Final reduce exhausted the 10000-token completion limit; retrying at 20000 tokens",
                                event="reduce_retry",
                                fields={
                                    "term": term_input.term,
                                    "initial_finish_reason": error.finish_reason,
                                    "initial_prompt_tokens": error.usage.get("prompt_tokens", 0),
                                    "initial_completion_tokens": error.usage.get("completion_tokens", 0),
                                    "initial_reasoning_tokens": error.usage.get("reasoning_tokens", 0),
                                    "retry_max_completion_tokens": FINAL_REDUCE_RETRY_MAX_COMPLETION_TOKENS,
                                },
                            )
                            reduce_result = request_json(
                                client,
                                reduce_prompt,
                                max_completion_tokens=FINAL_REDUCE_RETRY_MAX_COMPLETION_TOKENS,
                                reasoning_effort=AZURE_OPENAI_REASONING_EFFORT,
                                request_label=current_stage,
                            )
                        else:
                            raise
                    reduce_json, reduce_usage = reduce_result
                    run_usage.add(reduce_usage)
                    logger.log(
                        "Final reduce synthesis completed",
                        event="reduce",
                        fields={
                            "term": term_input.term,
                            "input_records": len(final_batches[0]),
                            "prompt_tokens": reduce_usage.get("prompt_tokens", 0),
                            "completion_tokens": reduce_usage.get("completion_tokens", 0),
                            "cached_prompt_tokens": reduce_usage.get("cached_prompt_tokens", 0),
                            "reasoning_tokens": reduce_usage.get("reasoning_tokens", 0),
                        },
                    )
                    break

                if reduction_level >= MAX_REDUCTION_LEVELS:
                    detail = f" ({final_planning_error})" if final_planning_error else ""
                    raise ValueError(
                        f"Exceeded {MAX_REDUCTION_LEVELS} segment reduction levels before final synthesis{detail}"
                    )

                current_stage = f"planning segment reduction level {reduction_level + 1}"
                segment_batches = chunk_analysis_records_for_reduce(
                    term_input,
                    stats,
                    analysis_records,
                    target_reduce_prompt_tokens=args.target_reduce_prompt_tokens,
                    reduce_max_completion_tokens=args.segment_max_completion_tokens,
                    reduce_prompt_safety_tokens=args.reduce_prompt_safety_tokens,
                )

                reduction_level += 1
                reduction_stages.append(
                    {
                        "level": reduction_level,
                        "input_record_count": len(analysis_records),
                        "segment_count": len(segment_batches),
                        "segment_record_counts": [len(batch) for batch in segment_batches],
                        "max_completion_tokens": args.segment_max_completion_tokens,
                    }
                )
                segment_outputs: list[dict[str, Any]] = []
                for segment_index, segment_batch in enumerate(segment_batches, start=1):
                    current_stage = (
                        f"segment reduction {reduction_level}, batch {segment_index}/{len(segment_batches)}"
                    )
                    segment_prompt = build_reduce_prompt(
                        term_input,
                        stats,
                        segment_batch,
                        stage=(
                            f"segment synthesis level {reduction_level}, "
                            f"batch {segment_index}/{len(segment_batches)}"
                        ),
                    )
                    segment_result = request_json(
                        client,
                        segment_prompt,
                        max_completion_tokens=args.segment_max_completion_tokens,
                        reasoning_effort=AZURE_OPENAI_MAP_REASONING_EFFORT,
                        request_label=current_stage,
                    )
                    segment_json, segment_usage = segment_result
                    run_usage.add(segment_usage)
                    logger.log(
                        "Segment reduce synthesis completed",
                        event="reduce_segment",
                        fields={
                            "term": term_input.term,
                            "level": reduction_level,
                            "segment_index": segment_index,
                            "segment_total": len(segment_batches),
                            "input_records": len(segment_batch),
                            "prompt_tokens": segment_usage.get("prompt_tokens", 0),
                            "completion_tokens": segment_usage.get("completion_tokens", 0),
                            "cached_prompt_tokens": segment_usage.get("cached_prompt_tokens", 0),
                            "reasoning_tokens": segment_usage.get("reasoning_tokens", 0),
                        },
                    )
                    segment_outputs.append(segment_json)
                analysis_records = segment_outputs

            summary = render_summary_markdown(term_input.term, stats, reduce_json)
        except Exception as error:
            failures[term_key] = {
                "term": term_input.term,
                "error": f"{type(error).__name__}: {error}",
                "updated_at": utc_timestamp(),
                "pipeline_version": PIPELINE_VERSION,
                "input_hash": input_hash,
                "analysis": {
                    "stage": current_stage,
                    "configured_chunk_size": args.chunk_size,
                    "effective_chunk_size": effective_chunk_size,
                    "chunk_count": len(chunks),
                    "completed_map_chunks": len(chunk_records),
                    "reduction_stages": reduction_stages,
                },
            }
            write_results_artifact(args.output, payload)
            failed += 1
            logger.log(
                f"Failed: {error}",
                is_error=True,
                event="term_failed",
                fields={
                    "term": term_input.term,
                    "occurrence_count": occurrence_count,
                    "input_hash": input_hash,
                },
                console_message=f"{progress} failed",
            )
            continue

        summaries[term_key] = {
            "term": term_input.term,
            "occurrence_count": occurrence_count,
            "summary": summary,
            "summary_source": AZURE_OPENAI_DEPLOYMENT,
            "generated_at": utc_timestamp(),
            "pipeline_version": PIPELINE_VERSION,
            "input_hash": input_hash,
            "corpus_hash": corpus_hash,
            "prompt_hash": prompt_hash,
            "analysis": {
                "configured_chunk_size": args.chunk_size,
                "effective_chunk_size": effective_chunk_size,
                "chunk_count": len(chunks),
                "chunk_records": chunk_records,
                "stats": occurrence_statistics(term_input),
                "token_budget": {
                    "target_map_prompt_tokens": args.target_map_prompt_tokens,
                    "map_max_completion_tokens": args.map_max_completion_tokens,
                    "map_prompt_safety_tokens": args.map_prompt_safety_tokens,
                    "min_dynamic_chunk_size": args.min_dynamic_chunk_size,
                },
                "reduction": {
                    "target_reduce_prompt_tokens": args.target_reduce_prompt_tokens,
                    "reduce_max_completion_tokens": args.reduce_max_completion_tokens,
                    "segment_max_completion_tokens": args.segment_max_completion_tokens,
                    "reduce_prompt_safety_tokens": args.reduce_prompt_safety_tokens,
                    "stages": reduction_stages,
                },
            },
        }
        failures.pop(term_key, None)
        write_results_artifact(args.output, payload)
        completed += 1
        logger.log(
            "Term summary saved",
            event="term_complete",
            fields={
                "term": term_input.term,
                "occurrence_count": occurrence_count,
                "chunk_count": len(chunks),
                "effective_chunk_size": effective_chunk_size,
                "input_hash": input_hash,
            },
            console_message=f"{progress} successful",
        )

        if args.pause_seconds:
            time.sleep(args.pause_seconds)

    estimated_cost = estimate_cost_usd(
        run_usage,
        input_cost_per_1m=args.input_cost_per_1m,
        output_cost_per_1m=args.output_cost_per_1m,
        cached_input_cost_per_1m=args.cached_input_cost_per_1m,
    )

    payload["run_usage"] = {
        "prompt_tokens": run_usage.prompt_tokens,
        "completion_tokens": run_usage.completion_tokens,
        "cached_prompt_tokens": run_usage.cached_prompt_tokens,
        "uncached_prompt_tokens": run_usage.uncached_prompt_tokens,
        "reasoning_tokens": run_usage.reasoning_tokens,
    }
    payload["run_cost_estimate_usd"] = {
        "input_cost_per_1m": args.input_cost_per_1m,
        "cached_input_cost_per_1m": args.cached_input_cost_per_1m,
        "output_cost_per_1m": args.output_cost_per_1m,
        **estimated_cost,
    }
    write_results_artifact(args.output, payload)

    logger.log(
        "Token usage summary",
        event="usage",
        fields={
            "prompt_tokens": run_usage.prompt_tokens,
            "cached_prompt_tokens": run_usage.cached_prompt_tokens,
            "uncached_prompt_tokens": run_usage.uncached_prompt_tokens,
            "completion_tokens": run_usage.completion_tokens,
            "reasoning_tokens": run_usage.reasoning_tokens,
        },
    )
    logger.log(
        "Estimated run cost",
        event="cost",
        fields={
            "input_cost_usd": round(estimated_cost["input_cost_usd"], 6),
            "cached_input_cost_usd": round(estimated_cost["cached_input_cost_usd"], 6),
            "output_cost_usd": round(estimated_cost["output_cost_usd"], 6),
            "total_cost_usd": round(estimated_cost["total_cost_usd"], 6),
            "input_cost_per_1m": args.input_cost_per_1m,
            "cached_input_cost_per_1m": args.cached_input_cost_per_1m,
            "output_cost_per_1m": args.output_cost_per_1m,
        },
    )
    if args.input_cost_per_1m == 0 and args.output_cost_per_1m == 0 and args.cached_input_cost_per_1m == 0:
        logger.log(
            "Cost rates are set to 0. Pass --input-cost-per-1m, --output-cost-per-1m, and optionally "
            "--cached-input-cost-per-1m for non-zero estimates."
        )

    logger.log(
        "Run finished",
        event="run_finish",
        fields={
            "completed": completed,
            "failed": failed,
            "artifact": str(args.output),
            "log": str(args.log),
        },
    )
    if failed:
        logger.log("Rerun the command to retry failed terms.", is_error=True)
        logger.close()
        return 1
    logger.close()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

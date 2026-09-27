#!/usr/bin/env python3
"""Build the ordered SRCities report viewer from the inspection JSON.

The reference reconstruction supplies the established node markup and client
behavior. The inspection JSON remains canonical for the expected report and
node identities, so this generator refuses to write an export if they diverge.
"""

from __future__ import annotations

import argparse
from collections import Counter
from dataclasses import dataclass
import html
from html.parser import HTMLParser
import json
from pathlib import Path
import re
from typing import Any

from openpyxl import load_workbook


REPO_ROOT = Path(__file__).resolve().parents[2]
DEFAULT_SOURCE_JSON = REPO_ROOT / "data" / "srsod-structure.json"
DEFAULT_REFERENCE_HTML = REPO_ROOT / "data" / "export" / "SRCities_consistencycheck.html"
DEFAULT_OUTPUT_HTML = REPO_ROOT / "data" / "export" / "SRCities_consistencycheck.html"
DEFAULT_GLOSSARY_PATH = REPO_ROOT / "data" / "Glossary" / "AR6_AR7SOD_Glossary_AO.xlsx"
DEFAULT_TERM_SUMMARIES_PATH = REPO_ROOT / "data" / "analysis" / "llm_term_check.json"
REPORT_HEADER_TITLE = "SRCities terminology review (version 10 Sep 2026 based on SOD)"
REFERENCE_KICKER = '<p class="kicker">Reconstructed report</p>'
OUTPUT_KICKER = '<p class="kicker">Reconstructed report in HTML</p>'
GLOSSARY_TAB_ID = "glossary-overview-tab"
GLOSSARY_PANEL_ID = "glossary-overview-panel"
CAE_TAB_ID = "cae-check-tab"
CAE_PANEL_ID = "cae-check-panel"
GLOSSARY_ISSUE_TAB_ID = "glossary-issue-table-tab"
GLOSSARY_ISSUE_PANEL_ID = "glossary-issue-table-panel"
REPORT_ORDER = (
    "Chapter 1",
    "Chapter 2",
    "Chapter 3",
    "Chapter 4",
    "Chapter 5",
    "SPM",
    "TS",
)
NODE_ID_RE = re.compile(r'\bdata-node-id="([^"]+)"')
DOCUMENT_NODE_ID_RE = re.compile(r'data-node-id="([^"]+:document:[^"]+)"')
IMAGE_SOURCE_RE = re.compile(r'<img\b[^>]*\bsrc="([^"]+)"', re.IGNORECASE)
SENTENCE_BOUNDARY_RE = re.compile(r"(?<=[.!?])\s+")
POTENTIAL_ISSUES_SECTION_RE = re.compile(
    r"^###\s+Potential issues(?:\s+(?:needing|for))?\s+substantive review[^\n]*\n"
    r"(?P<content>.*?)(?=\n###\s|\Z)",
    re.IGNORECASE | re.MULTILINE | re.DOTALL,
)
NO_POTENTIAL_ISSUE_RE = re.compile(r"^-?\s*none identified\b", re.IGNORECASE)
MARKDOWN_BULLET_RE = re.compile(r"^(?P<indent>[ \t]*)-\s+(?P<content>.*)$")
POTENTIAL_ISSUE_EVIDENCE_RE = re.compile(r'^(?:\[[^\]\n]+\](?:\s|$)|")')
LLM_CONTEXTS_HEADING = "Contexts of use"
LLM_POTENTIAL_ISSUES_HEADING = "Potential issues needing substantive review"
LLM_CONCLUSION_HEADING = "Conclusion"
AO_GLOSSARY_HEADERS = (
    "term",
    "equivalent terms",
    "equivalent terms 2",
    "equivalent terms -revised",
    "equivalent terms 2 - revised",
    "explanation (srcitiessod otherwise ar6)",
    "source",
    "parent terms",
    "child terms",
)
EXCLUDED_ALIASES_BY_TERM_KEY: dict[str, set[str]] = {
    "atmospheric rivers (ars)": {"ar"},
}
ALIAS_REPLACEMENTS_BY_TERM_KEY: dict[str, dict[str, str]] = {
    "atmospheric rivers (ars)": {"atmospheric rivers": "Atmospheric river"},
}
AGREEMENT_LEVELS = ("low", "medium", "high")
EVIDENCE_LEVELS = ("limited", "medium", "robust")
CONFIDENCE_LEVELS = ("very low", "low", "medium", "medium to high", "high", "very high")
CAE_PARENTHESES_RE = re.compile(r"\(([^()]*)\)")
CAE_ASSESSMENT_TOKEN_RE = re.compile(
    r"\b(?:very\s+low|very\s+high|low|medium|high|limited|robust)"
    r"(?:\s+to\s+(?:very\s+low|very\s+high|low|medium|high|limited|robust))?"
    r"\s+(?:confidence|agreement|evidence)\b",
    re.IGNORECASE,
)
CAE_CONFIDENCE_RE = re.compile(
    rf"^({'|'.join(CONFIDENCE_LEVELS)})\s+confidence$",
    re.IGNORECASE,
)
CAE_PAIR_SEPARATOR = r"(?:\s*,\s*(?:and\s+)?|\s*;\s*|\s+and\s+|\s*\+\s*)"
CAE_PAIR_RE = re.compile(
    rf"^(?:(?P<agreement>{'|'.join(AGREEMENT_LEVELS)})\s+agreement{CAE_PAIR_SEPARATOR}"
    rf"(?P<evidence>{'|'.join(EVIDENCE_LEVELS)})\s+evidence|"
    rf"(?P<evidence_first>{'|'.join(EVIDENCE_LEVELS)})\s+evidence{CAE_PAIR_SEPARATOR}"
    rf"(?P<agreement_second>{'|'.join(AGREEMENT_LEVELS)})\s+agreement)$",
    re.IGNORECASE,
)


@dataclass(frozen=True)
class CaeOccurrence:
    """One sentence-final confidence, agreement, or evidence assessment."""

    report_name: str
    source_label: str
    node_code: str
    node_id: str
    sentence: str
    assessment: str
    issue: str = ""


@dataclass(frozen=True)
class RevisedGlossaryEntry:
    """One canonical glossary entry with its AO workbook search aliases."""

    term: str
    aliases: tuple[str, ...]
    explanation: str
    source: str
    parent_terms: tuple[str, ...]
    child_terms: tuple[str, ...]


@dataclass(frozen=True)
class GlossaryOccurrence:
    """One report sentence matched by a canonical term or revised alias."""

    source_label: str
    node_code: str
    node_id: str
    sentence: str
    source_text: str
    matched_names: tuple[str, ...]


@dataclass(frozen=True)
class GlossaryIssueRow:
    """One issue-table row derived from a term summary's potential issue section."""

    section: str
    sentence: str
    term: str
    issue: str
    node_id: str = ""
    source_label: str = ""
    aliases: tuple[str, ...] = ()


RevisedGlossary = dict[str, RevisedGlossaryEntry]
GlossaryMatchMap = dict[str, tuple[str, ...]]
GlossaryOccurrences = dict[str, list[GlossaryOccurrence]]


def normalize_text(value: object) -> str:
    """Convert a source cell or report fragment to normalized plain text."""
    return re.sub(r"\s+", " ", str(value or "").replace("\u00a0", " ")).strip()


def normalize_section_identifier(value: object) -> str:
    """Normalize section labels so numeric codes and paragraph markers are consistently spaced."""
    text = normalize_text(value)
    text = re.sub(
        r"\bP[\s.\-]*(\d+)\b",
        lambda match: f"P{match.group(1)}",
        text,
        flags=re.IGNORECASE,
    )
    return re.sub(
        r"\b(\d+(?:\.\d+)+)\s*(P\d+)\b",
        lambda match: f"{match.group(1)} {match.group(2).upper()}",
        text,
        flags=re.IGNORECASE,
    )


def split_cell_lines(value: object) -> tuple[str, ...]:
    """Return unique non-empty newline-separated workbook values in source order."""
    values = [normalize_text(line) for line in str(value or "").splitlines()]
    return tuple(dict.fromkeys(value for value in values if value))


def load_revised_glossary(input_path: Path) -> RevisedGlossary:
    """Load AO terms from A, aliases from B-E, and detail fields from F-I."""
    workbook = load_workbook(input_path, read_only=True, data_only=True)
    try:
        worksheet = workbook.active
        headers = tuple(normalize_text(cell.value).casefold() for cell in worksheet[1])
        if headers[: len(AO_GLOSSARY_HEADERS)] != AO_GLOSSARY_HEADERS:
            raise ValueError(
                "AO glossary columns A-I must be Term, four Equivalent Terms columns, "
                "Explanation, Source, Parent Terms, and Child Terms."
            )

        glossary: RevisedGlossary = {}
        for row in worksheet.iter_rows(min_row=2, values_only=True):
            term = normalize_text(row[0])
            if not term:
                continue
            term_key = term.casefold()
            if term_key in glossary:
                raise ValueError(f"AO glossary contains duplicate term: {term}")

            alias_replacements = ALIAS_REPLACEMENTS_BY_TERM_KEY.get(term_key, {})
            aliases = tuple(
                alias
                for alias in dict.fromkeys(
                    alias_replacements.get(alias.casefold(), alias)
                    for alias in [
                        *split_cell_lines(row[1]),
                        *split_cell_lines(row[2]),
                        *split_cell_lines(row[3]),
                        *split_cell_lines(row[4]),
                    ]
                )
                if alias.casefold() != term_key
                and alias.casefold() not in EXCLUDED_ALIASES_BY_TERM_KEY.get(term_key, set())
            )
            glossary[term_key] = RevisedGlossaryEntry(
                term=term,
                aliases=aliases,
                explanation=normalize_text(row[5]),
                source=normalize_text(row[6]),
                parent_terms=split_cell_lines(row[7]),
                child_terms=split_cell_lines(row[8]),
            )
    finally:
        workbook.close()

    if not glossary:
        raise RuntimeError(f"No glossary terms were read from {input_path}.")
    return glossary


def build_glossary_match_map(glossary: RevisedGlossary) -> GlossaryMatchMap:
    """Map each canonical term or alias to every canonical entry it represents."""
    owners_by_name: dict[str, list[str]] = {}
    for term_key, entry in glossary.items():
        for name in (entry.term, *entry.aliases):
            name_key = name.casefold()
            owners = owners_by_name.setdefault(name_key, [])
            if term_key not in owners:
                owners.append(term_key)
    return {name_key: tuple(owners) for name_key, owners in owners_by_name.items()}


def build_glossary_match_pattern(match_map: GlossaryMatchMap) -> re.Pattern[str] | None:
    """Build a longest-first case-insensitive pattern for canonical terms and aliases."""
    if not match_map:
        return None
    return re.compile(
        "|".join(re.escape(name) for name in sorted(match_map, key=len, reverse=True)),
        re.IGNORECASE,
    )


@dataclass
class CaeCheckResult:
    """Counts of valid CAE statements and malformed cases requiring review."""

    agreement_evidence: Counter[tuple[str, str]]
    confidence: Counter[str]
    issues: list[CaeOccurrence]
    agreement_evidence_by_report: Counter[tuple[str, str, str]]
    confidence_by_report: Counter[tuple[str, str]]

    @property
    def valid_pair_count(self) -> int:
        return sum(self.agreement_evidence.values())

    @property
    def confidence_count(self) -> int:
        return sum(self.confidence.values())

    @property
    def valid_count(self) -> int:
        return self.valid_pair_count + self.confidence_count

    @property
    def candidate_count(self) -> int:
        return self.valid_count + len(self.issues)

PREVIOUS_HEADER_CSS = """
            :root {
                --coral: #009edb;
                --gold: #0076a8;
            }
            .report-panel > header {
                background: linear-gradient(135deg, #edf8fc, var(--paper) 60%);
            }
            .site-header {
                background: #ffffff;
                border-bottom: 4px solid var(--ipcc-blue);
                padding: 1.35rem 1.5rem 0;
                position: sticky;
                top: 0;
                z-index: 10;
            }
            .site-header__inner { margin: 0 auto; max-width: 1120px; }
            .site-header h1 {
                color: #1e2a30;
                font-family: "Iowan Old Style", "Palatino Linotype", Georgia, serif;
                font-size: 2rem;
                font-weight: 600;
                letter-spacing: 0;
                line-height: 1.15;
                margin: 0;
            }
            .report-panel > header h1 { font-size: 2rem; }
            .site-header .chapter-tabs {
                align-items: center;
                background: #edf5f5;
                border: 0;
                display: flex;
                flex-wrap: wrap;
                gap: .35rem .9rem;
                margin-top: .9rem;
                overflow: visible;
                padding: .7rem 1rem;
                position: static;
            }
            .site-header .chapter-tab,
            .site-header .metadata-toggle {
                background: transparent;
                border: 0;
                border-radius: 0;
                color: #004f6a;
                flex: 0 0 auto;
                font: 400 1rem/1.2 "Avenir Next", "Helvetica Neue", sans-serif;
                letter-spacing: 0;
                min-height: 0;
                padding: 0;
            }
            .site-header .chapter-tab:hover,
            .site-header .chapter-tab[aria-selected="true"],
            .site-header .metadata-toggle:hover {
                background: transparent;
                color: #006b8f;
                text-decoration: underline;
                text-decoration-thickness: 1px;
                text-underline-offset: .13em;
            }
            .site-header .metadata-toggle {
                margin-left: auto;
                position: static;
            }
            .site-header .chapter-tab:focus-visible,
            .site-header .metadata-toggle:focus-visible,
            .back-to-top:focus-visible {
                outline: 3px solid var(--ipcc-blue);
                outline-offset: 2px;
            }
            .chapter_information,
            .section,
            .box,
            .cross_chapter_box,
            figure {
                scroll-margin-top: 9rem;
            }
            .report-panel .section > h2 .node-code,
            .report-panel .section > h3 .node-code,
            .report-panel .node-code { color: var(--ipcc-blue); }
            @media (max-width: 42rem) {
                .site-header { padding: 1rem 1rem 0; }
                .site-header h1,
                .report-panel > header h1 { font-size: 1.5rem; }
                .site-header .metadata-toggle { margin-left: 0; }
                .chapter_information,
                .section,
                .box,
                .cross_chapter_box,
                figure {
                    scroll-margin-top: 11rem;
                }
            }
"""

GLOSSARY_CSS = """
            .glossary-overview { padding: clamp(1.25rem, 3vw, 3rem); }
            .glossary-workspace {
                --glossary-index-width: 30%;
                display: grid;
                grid-template-columns: minmax(16rem, var(--glossary-index-width)) .75rem minmax(22rem, 1fr);
            }
            .glossary-index-pane,
            .glossary-detail-pane {
                border: 1px solid var(--rule);
                height: 640px;
                min-width: 0;
                overflow-y: auto;
                padding: 1rem;
            }
            .glossary-divider {
                align-items: center;
                cursor: col-resize;
                display: flex;
                justify-content: center;
                touch-action: none;
            }
            .glossary-divider::before {
                background: #8aaeb9;
                content: "";
                height: 100%;
                transition: background-color .15s, width .15s;
                width: 1px;
            }
            .glossary-divider:hover::before,
            .glossary-divider:focus-visible::before,
            .glossary-divider.is-dragging::before {
                background: var(--ipcc-blue);
                width: 3px;
            }
            .glossary-divider:focus-visible { outline: 0; }
            .glossary-index-pane h2,
            .glossary-detail-pane h2 {
                color: var(--ipcc-blue);
                font-size: 1.22rem;
            }
            .glossary-unused-toggle {
                background: transparent;
                border: 1px solid #0076a8;
                border-radius: 3px;
                color: #004f6a;
                cursor: pointer;
                font: 600 .85rem/1.2 "Avenir Next", "Helvetica Neue", sans-serif;
                margin: .65rem 0 .1rem;
                padding: .45rem .65rem;
            }
            .glossary-unused-toggle:hover,
            .glossary-unused-toggle[aria-pressed="true"] {
                background: #e1f5fa;
                color: #006b8f;
            }
            .glossary-unused-toggle:focus-visible {
                outline: 3px solid var(--ipcc-blue);
                outline-offset: 2px;
            }
            .glossary-search-label {
                display: grid;
                font-size: .85rem;
                font-weight: 800;
                gap: .3rem;
                margin: .8rem 0 .35rem;
            }
            .glossary-search {
                border: 1px solid #8aaeb9;
                border-radius: 3px;
                color: var(--ink);
                font: inherit;
                padding: .5rem .6rem;
                width: 100%;
            }
            .glossary-result-count,
            .glossary-detail-placeholder {
                color: var(--muted);
                font-size: .85rem;
                margin: .35rem 0 .75rem;
            }
            .glossary-term-list { list-style: none; margin: 0; padding: 0; }
            .glossary-term-row { margin: .15rem 0; }
            .glossary-term-button {
                background: transparent;
                border: 0;
                color: var(--ink);
                cursor: pointer;
                font: inherit;
                padding: .25rem 0;
                text-align: left;
                width: 100%;
            }
            .glossary-term-button:not(:disabled):hover,
            .glossary-term-button[aria-pressed="true"] {
                color: var(--ipcc-blue);
                text-decoration: underline;
                text-decoration-thickness: 1px;
                text-underline-offset: .14em;
            }
            .glossary-term-button:disabled { color: var(--muted); cursor: default; }
            .glossary-term-count { color: #7a4b2a; font-weight: 400; }
            .glossary-term-source { color: #0076a8; font-weight: 400; }
            .glossary-term-issue {
                color: #c4271e;
                display: inline-block;
                font-size: 1rem;
                font-weight: 700;
                margin-left: .35rem;
            }
            .glossary-detail[hidden] { display: none; }
            .glossary-detail h3 { font-size: 1.45rem; font-weight: 400; margin-top: .8rem; }
            .glossary-detail-heading {
                align-items: baseline;
                display: flex;
                flex-wrap: wrap;
                gap: .3rem .45rem;
            }
            .glossary-canonical-term,
            .glossary-parent,
            .glossary-child,
            .glossary-aliases { color: var(--muted); margin: .4rem 0 0; }
            .glossary-definition {
                border-left: 3px solid #9fcdd8;
                margin-top: 1rem;
                padding-left: .8rem;
            }
            .glossary-definition h4 { color: var(--ipcc-blue); font-size: 1rem; }
            .glossary-definition p { margin: .35rem 0 0; overflow-wrap: anywhere; }
            .glossary-llm-check {
                border-top: 1px solid var(--rule);
                margin-top: 1rem;
                padding-top: .75rem;
            }
            .glossary-llm-check summary { color: var(--ipcc-blue); cursor: pointer; font-weight: 800; }
            .glossary-llm-check-notebar {
                align-items: center;
                background: #d7ecff;
                border: 1px solid #8bb7e0;
                border-radius: 6px;
                box-sizing: border-box;
                color: var(--ink);
                display: flex;
                font: inherit;
                margin-top: .55rem;
                padding: .5rem .65rem;
                width: 100%;
            }
            .glossary-llm-check-eye {
                color: #0b4f86;
                font-size: 1rem;
                line-height: 1;
                margin-right: .45rem;
            }
            .glossary-llm-check-note {
                line-height: 1.35;
            }
            .glossary-llm-check-content {
                margin-top: .65rem;
                overflow-wrap: anywhere;
            }
            .glossary-llm-check-content h4 {
                color: var(--ipcc-blue);
                font-size: .98rem;
                margin: .55rem 0 .35rem;
            }
            .glossary-llm-check-content p {
                margin: .35rem 0;
            }
            .glossary-llm-check-content ul {
                margin: .35rem 0 .5rem 1.25rem;
                padding: 0;
            }
            .glossary-evidence-code.glossary-inline-evidence-code {
                color: #0076a8;
                cursor: pointer;
                display: inline;
                font: inherit;
                margin: 0;
                padding: 0;
                text-decoration: underline;
                text-decoration-thickness: 1px;
                text-underline-offset: .14em;
                vertical-align: baseline;
            }
            .glossary-evidence-code.glossary-inline-evidence-code:hover { color: #004f6a; }
            .glossary-evidence-code.glossary-inline-evidence-code:focus-visible {
                outline: 3px solid var(--ipcc-blue);
                outline-offset: 2px;
            }
            .glossary-evidence {
                border-top: 1px solid var(--rule);
                margin-top: 1rem;
                padding-top: .75rem;
            }
            .glossary-evidence summary { color: var(--ipcc-blue); cursor: pointer; font-weight: 800; }
            .glossary-evidence table { margin-top: .7rem; }
            .glossary-evidence th {
                border: 1px solid var(--rule);
                padding: .5rem .65rem;
                text-align: left;
                vertical-align: top;
            }
            .glossary-evidence td:first-child { white-space: nowrap; }
            .glossary-evidence-location { display: block; }
            .glossary-evidence-code {
                background: transparent;
                border: 0;
                color: #0076a8;
                cursor: pointer;
                display: block;
                font: inherit;
                margin-top: .2rem;
                padding: 0;
                text-decoration: underline;
                text-decoration-thickness: 1px;
                text-underline-offset: .14em;
            }
            .glossary-evidence-code:hover { color: #004f6a; }
            .glossary-evidence-code:focus-visible {
                outline: 3px solid var(--ipcc-blue);
                outline-offset: 2px;
            }
            .report-glossary-link {
                color: #0076a8;
                font-weight: 700;
                text-decoration: underline;
                text-decoration-thickness: 1px;
                text-underline-offset: .14em;
            }
            .report-glossary-link:hover { color: #004f6a; }
            .report-glossary-link.has-issue {
                color: #8a2a16;
            }
            .report-glossary-link-issue-flag {
                color: #8a2a16;
                font-weight: 800;
                margin-left: .2em;
            }
            .glossary-definition-dialog {
                border: 1px solid var(--rule);
                border-radius: 4px;
                color: var(--ink);
                max-height: min(80vh, 46rem);
                max-width: min(92vw, 46rem);
                padding: 0;
                width: 100%;
            }
            .glossary-definition-dialog::backdrop { background: rgba(32, 39, 41, .45); }
            .glossary-dialog-header {
                align-items: center;
                border-bottom: 3px solid var(--ipcc-blue);
                display: flex;
                gap: 1rem;
                justify-content: space-between;
                padding: 1rem 1.2rem;
            }
            .glossary-dialog-header h2 { color: var(--ipcc-blue); font-size: 1.45rem; }
            .glossary-dialog-close {
                background: transparent;
                border: 0;
                color: var(--ink);
                cursor: pointer;
                font-size: 1.5rem;
                height: 2rem;
                line-height: 1;
                padding: 0;
                width: 2rem;
            }
            .glossary-dialog-content { overflow-y: auto; padding: 0 1.2rem 1.2rem; }
            .glossary-paragraph-location { color: var(--muted); font-size: .9rem; }
            .glossary-paragraph-text { line-height: 1.65; }
            mark { background: #e1f5fa; color: inherit; }
            @media (max-width: 52rem) {
                .glossary-workspace { grid-template-columns: 1fr; }
                .glossary-divider { display: none; }
                .glossary-index-pane,
                .glossary-detail-pane { height: min(60vh, 640px); }
            }
"""

GLOSSARY_ISSUE_TABLE_CSS = """
            .glossary-issue-overview { padding: clamp(1.25rem, 3vw, 3rem); }
            .glossary-issue-wrap { overflow-x: auto; }
            .glossary-issue-table {
                border-collapse: collapse;
                border-spacing: 0;
                table-layout: fixed;
                width: 100%;
            }
            .glossary-issue-table th:nth-child(1),
            .glossary-issue-table td:nth-child(1) { width: 15%; }
            .glossary-issue-table th:nth-child(2),
            .glossary-issue-table td:nth-child(2) { width: 15%; }
            .glossary-issue-table th:nth-child(3),
            .glossary-issue-table td:nth-child(3) { width: 30%; }
            .glossary-issue-table th:nth-child(4),
            .glossary-issue-table td:nth-child(4) { width: 40%; }
            .glossary-issue-table th,
            .glossary-issue-table td {
                border: 1px solid var(--rule);
                padding: .6rem .7rem;
                text-align: left;
                vertical-align: top;
                word-break: break-word;
            }
            .glossary-issue-table th {
                background: #f2f7f9;
                color: var(--ipcc-blue);
                font-weight: 800;
                white-space: normal;
            }
            .glossary-issue-empty {
                color: var(--muted);
                margin: .2rem 0 0;
            }

            .glossary-issue-search { margin: 1rem 0; }
            .glossary-issue-search label { display: block; font-weight: 600; }
            .glossary-issue-search-input { display: block; width: 100%; max-width: 42rem; box-sizing: border-box; margin-top: .35rem; padding: .45rem .6rem; }
            .glossary-issue-search-input:focus-visible { outline: 2px solid currentColor; outline-offset: 2px; }
            .glossary-issue-search-hint { margin: .35rem 0; font-size: .9em; }
            .glossary-issue-search-status { margin: .35rem 0; font-size: .9em; }
            .glossary-issue-search-status[data-error='true'] { color: #a00; }
            .glossary-issue-export { display: flex; flex-wrap: wrap; gap: .5rem; margin: .35rem 0; }
            .glossary-issue-export button { cursor: pointer; }
            .glossary-issue-term-button {
                background: none;
                border: 0;
                color: #005f86;
                cursor: pointer;
                font: inherit;
                font-weight: 700;
                padding: 0;
                text-align: left;
                text-decoration: underline;
                text-decoration-thickness: 1px;
                text-underline-offset: .12em;
            }
            .glossary-issue-term-button:hover { color: #004660; }
            .glossary-issue-term-button:focus-visible {
                outline: 2px solid currentColor;
                outline-offset: 2px;
            }
            .glossary-issue-table { table-layout: fixed; }

            .glossary-issue-controls {
                display: grid;
                gap: .75rem;
                grid-template-columns: minmax(11rem, 14rem) minmax(0, 1fr);
            }
            .glossary-issue-chapter {
                appearance: none;
                background: #fff;
                border: 1px solid #c7c7c7;
                border-radius: .25rem;
                box-sizing: border-box;
                color: inherit;
                font: inherit;
                min-height: 2.5rem;
                padding: .5rem .75rem;
                width: 100%;
            }
            .glossary-issue-chapter:focus-visible {
                outline: 2px solid currentColor;
                outline-offset: 2px;
            }
            @media (max-width: 40rem) {
                .glossary-issue-controls { grid-template-columns: 1fr; }
            }
"""

GLOSSARY_ISSUE_TABLE_JAVASCRIPT = """
        <script>
        // glossary-issue-search-script
        (() => {
            const panel = document.querySelector('#glossary-issue-table-panel');
            if (!panel) return;

            const chapterSelect = panel.querySelector('#glossary-issue-chapter');
            const searchInput = panel.querySelector('.glossary-issue-search-input');
            const status = panel.querySelector('.glossary-issue-search-status');
            const downloadHtml = panel.querySelector('.glossary-issue-download-html');
            const downloadPdf = panel.querySelector('.glossary-issue-download-pdf');
            const definitionDialog = document.getElementById('glossary-definition-dialog');
            const definitionDialogTitle = definitionDialog?.querySelector('.glossary-dialog-title');
            const definitionDialogContent = definitionDialog?.querySelector('.glossary-dialog-content');
            const rows = Array.from(panel.querySelectorAll('.glossary-issue-table tbody tr'));
            if (!chapterSelect || !searchInput || !status || rows.length === 0) return;

            const normalizeTerm = (value) => (value || '').replace(/\\s+/g, ' ').trim().toLowerCase();
            const sentenceBoundary = /(?<=[.!?])\\s+(?=[A-Z0-9"“(\\[])/;
            const escapeHtml = (value) => String(value)
                .replace(/&/g, '&amp;')
                .replace(/</g, '&lt;')
                .replace(/>/g, '&gt;')
                .replace(/"/g, '&quot;')
                .replace(/'/g, '&#39;');
            const escapeRegExp = (value) => value.replace(/[.*+?^${}()|[\\]\\]/g, '\\\\$&');

            const detailByTerm = new Map();
            const aliasesByCanonical = new Map();
            const canonicalByTerm = new Map();
            const addDetail = (rawTerm, htmlContent) => {
                const key = normalizeTerm(rawTerm);
                if (!key || detailByTerm.has(key) || !htmlContent) return;
                detailByTerm.set(key, htmlContent);
            };

            const parseAliasesFromDetail = (detail) => {
                const aliasesText = detail.querySelector('.glossary-aliases')?.textContent || '';
                return aliasesText
                    .replace(/^\\s*Alias\\(es\\):\\s*/i, '')
                    .split(',')
                    .map((item) => item.replace(/\\s*\\[[^\\]]+\\]\\s*$/, '').trim())
                    .filter(Boolean);
            };

            const detailPopupMarkup = (detail) => {
                const clone = detail.cloneNode(true);
                clone.removeAttribute('hidden');
                clone.removeAttribute('id');
                clone.querySelectorAll('[id]').forEach((element) => element.removeAttribute('id'));
                return clone.innerHTML;
            };

            document.querySelectorAll('#glossary-overview-panel .glossary-detail').forEach((detail) => {
                const canonicalTerm = detail.querySelector('h3')?.textContent?.trim() || '';
                const detailHtml = detailPopupMarkup(detail);
                if (!canonicalTerm || !detailHtml) return;

                const canonicalKey = normalizeTerm(canonicalTerm);
                const aliases = parseAliasesFromDetail(detail);
                aliasesByCanonical.set(canonicalKey, aliases);
                canonicalByTerm.set(canonicalKey, canonicalTerm);

                addDetail(canonicalTerm, detailHtml);
                aliases.forEach((alias) => {
                    canonicalByTerm.set(normalizeTerm(alias), canonicalTerm);
                    addDetail(alias, detailHtml);
                });
            });

            const termLookupVariants = (rawTerm) => {
                const source = (rawTerm || '').replace(/\\s+/g, ' ').trim();
                if (!source) return [];

                const variants = [source];
                const withoutParen = source
                    .replace(/\\s*\\([^)]*\\)\\s*/g, ' ')
                    .replace(/\\s+/g, ' ')
                    .trim();
                if (withoutParen && withoutParen !== source) {
                    variants.push(withoutParen);
                }

                const parenMatches = source.match(/\\(([^)]+)\\)/g) || [];
                parenMatches
                    .map((chunk) => chunk.replace(/[()]/g, '').trim())
                    .filter(Boolean)
                    .forEach((chunk) => {
                        variants.push(chunk);
                        if (/^[A-Za-z]{2,}s$/.test(chunk)) {
                            variants.push(chunk.slice(0, -1));
                        }
                    });

                return Array.from(new Set(variants));
            };

            const resolveCanonicalKey = (term) => {
                const variantKeys = termLookupVariants(term).map((variant) => normalizeTerm(variant));
                for (const key of variantKeys) {
                    const mappedCanonical = canonicalByTerm.get(key);
                    if (mappedCanonical) {
                        return normalizeTerm(mappedCanonical);
                    }
                    if (aliasesByCanonical.has(key)) {
                        return key;
                    }
                }
                return normalizeTerm(term || '');
            };

            const termHighlightCandidates = (term, includeAliases = false) => {
                const source = (term || '').replace(/\\s+/g, ' ').trim();
                if (!source) return [];

                const candidates = [source];
                const parenMatches = source.match(/\\(([^)]+)\\)/g) || [];
                parenMatches
                    .map((chunk) => chunk.replace(/[()]/g, '').trim())
                    .filter(Boolean)
                    .forEach((chunk) => candidates.push(chunk));

                if (includeAliases) {
                    const canonicalKey = resolveCanonicalKey(source);
                    const aliasCandidates = aliasesByCanonical.get(canonicalKey) || [];
                    aliasCandidates.forEach((alias) => {
                        candidates.push(alias);
                        const aliasParenMatches = alias.match(/\\(([^)]+)\\)/g) || [];
                        aliasParenMatches
                            .map((chunk) => chunk.replace(/[()]/g, '').trim())
                            .filter(Boolean)
                            .forEach((chunk) => candidates.push(chunk));
                    });
                }

                return Array.from(new Set(candidates.map((item) => item.trim()).filter(Boolean)))
                    .sort((a, b) => b.length - a.length);
            };

            const sentenceHasCandidate = (sentence, candidates) => candidates.some((candidate) => {
                const matcher = new RegExp(`(^|[^\\w])${escapeRegExp(candidate)}($|[^\\w])`, 'i');
                return matcher.test(sentence);
            });

            const rowHasTermMatch = (sentence, term) => {
                const canonicalCandidates = termHighlightCandidates(term, false);
                if (canonicalCandidates.length && sentenceHasCandidate(sentence, canonicalCandidates)) {
                    return true;
                }
                const aliasCandidates = termHighlightCandidates(term, true)
                    .filter((candidate) => !canonicalCandidates.includes(candidate));
                return aliasCandidates.length ? sentenceHasCandidate(sentence, aliasCandidates) : false;
            };

            const looksLikeWordCountPlaceholder = (value) => /^\\d+\\s*\\|\\s*\\d+\\s*words?$/i.test((value || '').trim());

            const sourceParagraphTextFromRow = (row) => {
                const sourceButton = row.querySelector('td:first-child .glossary-evidence-code[data-source-node-id]');
                const nodeId = sourceButton?.dataset.sourceNodeId || '';
                if (!nodeId) return '';
                const sourceNode = document.querySelector(`[data-node-id="${CSS.escape(nodeId)}"]`);

                const extractCleanText = (node) => {
                    if (!node) return '';
                    const clone = node.cloneNode(true);
                    clone.querySelectorAll('.report-glossary-link-issue-flag, p.source-reference').forEach((element) => element.remove());
                    const rawText = clone.textContent || node.textContent || '';
                    const normalized = rawText.replace(/\\s+/g, ' ').trim();
                    if (!normalized) return '';
                    return normalized
                        .replace(/^\\[[^\\]]+\\]\\s*/, '')
                        .replace(/\\s*↑\\s*/g, ' ')
                        .trim();
                };

                const paragraphText = extractCleanText(sourceNode?.querySelector('p.paragraph'));
                if (paragraphText) return paragraphText;

                const figureCaptionText = extractCleanText(sourceNode?.querySelector('figcaption'));
                if (figureCaptionText) return figureCaptionText;

                return extractCleanText(sourceNode);
            };

            const repairedSentenceFromSource = (row, term, currentSentence) => {
                const normalizedCurrent = (currentSentence || '').replace(/\\s+/g, ' ').trim();
                if (normalizedCurrent && rowHasTermMatch(normalizedCurrent, term)) {
                    return normalizedCurrent;
                }

                const sourceText = sourceParagraphTextFromRow(row);
                if (!sourceText) return normalizedCurrent;

                const sentences = sourceText
                    .split(sentenceBoundary)
                    .map((sentence) => sentence.replace(/\\s+/g, ' ').trim())
                    .filter(Boolean);

                const matchingSentence = sentences.find((sentence) => rowHasTermMatch(sentence, term));
                if (matchingSentence) return matchingSentence;

                if (looksLikeWordCountPlaceholder(normalizedCurrent)) {
                    return sentences[0] || sourceText || normalizedCurrent;
                }
                return normalizedCurrent;
            };

            const highlightTermOccurrences = (sentenceCell, term) => {
                const plainText = (sentenceCell.textContent || '').replace(/\\s+/g, ' ').trim();
                if (!plainText) {
                    sentenceCell.innerHTML = '';
                    return;
                }
                const canonicalCandidates = termHighlightCandidates(term, false);
                const canonicalHasMatch = canonicalCandidates.length && sentenceHasCandidate(plainText, canonicalCandidates);
                const aliasCandidates = termHighlightCandidates(term, true)
                    .filter((candidate) => !canonicalCandidates.includes(candidate));
                const activeCandidates = canonicalHasMatch
                    ? canonicalCandidates
                    : (aliasCandidates.length && sentenceHasCandidate(plainText, aliasCandidates)
                        ? aliasCandidates
                        : canonicalCandidates);

                if (!activeCandidates.length) {
                    sentenceCell.innerHTML = escapeHtml(plainText);
                    return;
                }
                const matcher = new RegExp(`(${activeCandidates.map(escapeRegExp).join('|')})`, 'gi');
                sentenceCell.innerHTML = plainText
                    .split(matcher)
                    .map((part, index) => (index % 2 === 1 ? `<mark>${escapeHtml(part)}</mark>` : escapeHtml(part)))
                    .join('');
            };

            const dedupeIssueRows = () => {
                const seen = new Set();
                rows.forEach((row) => {
                    const cells = row.querySelectorAll('td');
                    if (cells.length < 4) return;

                    const sectionText = (cells[0].textContent || '').replace(/\\s+/g, ' ').trim().toLowerCase();
                    const termText = (cells[1].textContent || '').replace(/\\s+/g, ' ').trim().toLowerCase();
                    const sentenceText = (cells[2].textContent || '').replace(/\\s+/g, ' ').trim().toLowerCase();
                    const issueText = (cells[3].textContent || '').replace(/\\s+/g, ' ').trim().toLowerCase();
                    const key = [sectionText, termText, sentenceText, issueText].join('||');

                    if (seen.has(key)) {
                        row.remove();
                        return;
                    }
                    seen.add(key);
                });
            };

            dedupeIssueRows();

            rows.forEach((row) => {
                if (!row.isConnected) return;
                const cells = row.querySelectorAll('td');
                const termCell = cells[1];
                const sentenceCell = cells[2];
                if (!termCell || !sentenceCell) return;
                const term = (termCell.textContent || '').replace(/\\s+/g, ' ').trim();
                const sentence = repairedSentenceFromSource(
                    row,
                    term,
                    (sentenceCell.textContent || '').replace(/\\s+/g, ' ').trim(),
                );
                sentenceCell.textContent = sentence;

                termCell.innerHTML = `<button class="glossary-issue-term-button" type="button" data-term="${escapeHtml(term)}">${escapeHtml(term)}</button>`;
                highlightTermOccurrences(sentenceCell, term);
            });

            panel.addEventListener('click', (event) => {
                const termButton = event.target.closest('.glossary-issue-term-button');
                if (!termButton || !definitionDialog || !definitionDialogTitle || !definitionDialogContent) return;
                const term = termButton.dataset.term || termButton.textContent || '';
                const detailHtml = detailByTerm.get(normalizeTerm(term));
                definitionDialogTitle.textContent = term || 'Glossary definition';
                definitionDialogContent.innerHTML = detailHtml
                    ? detailHtml
                    : '<p>Definition not available for this term in the glossary overview.</p>';
                definitionDialog.showModal();
            });

            const chapterFromSectionText = (sectionText) => {
                const chapterMatch = sectionText.match(/\\bchapter\\s+([1-5])\\b/i);
                if (chapterMatch) return `Chapter ${chapterMatch[1]}`;
                if (/\\bspm\\b/i.test(sectionText)) return 'SPM';
                if (/\\bts\\b/i.test(sectionText) || /technical\\s+summary/i.test(sectionText)) return 'TS';
                return 'Other';
            };

            const normalizeToken = (value) => value.trim().toLowerCase();

            const normalizeSectionToken = (value) => normalizeToken(value)
                .replace(/[\\[\\]]/g, '')
                .replace(/\\bp\\.\\s*(\\d+?)(\\d+\\.\\d+)/gi, 'p$1 $2')
                .replace(/\\bp(\\d+)(\\d+\\.\\d+)/gi, 'p$1 $2')
                .replace(/\\bp[\\s.\\-]*(\\d+)\\b/gi, 'p$1')
                .replace(/\\s+/g, ' ')
                .trim();

            const isSectionToken = (value) => {
                const token = normalizeSectionToken(value);
                return /^\\d+(?:\\.\\d+)*(?:\\s*p\\d+|p\\d+)?$/.test(token) || /^p\\d+$/.test(token);
            };

            const hasMatchingSectionCode = (sectionText, queryToken) => {
                const query = normalizeSectionToken(queryToken);
                if (!query) return false;
                const sectionMatches = Array.from(
                    normalizeSectionToken(sectionText).matchAll(/\\b(\\d+(?:\\.\\d+)+)\\s*(p\\d+)?\\b/gi),
                    (match) => ({
                        base: (match[1] || '').toLowerCase(),
                        paragraph: (match[2] || '').toLowerCase(),
                    }),
                );
                if (sectionMatches.length === 0) return false;

                if (/^p\\d+$/.test(query)) {
                    return sectionMatches.some((entry) => entry.paragraph === query);
                }

                const queryMatch = query.match(/^(\\d+(?:\\.\\d+)+)(?:\\s*(p\\d+))?$/i);
                if (!queryMatch) return false;
                const queryBase = (queryMatch[1] || '').toLowerCase();
                const queryParagraph = (queryMatch[2] || '').toLowerCase();

                return sectionMatches.some((entry) => {
                    if (entry.base === queryBase || entry.base.startsWith(`${queryBase}.`)) {
                        return !queryParagraph || entry.paragraph === queryParagraph;
                    }
                    return false;
                });
            };

            const activeRows = () => Array.from(panel.querySelectorAll('.glossary-issue-table tbody tr'))
                .filter((row) => row.isConnected);

            const currentRowData = () => activeRows().map((row) => {
                const cells = row.querySelectorAll('td');
                const values = Array.from(cells).map((cell) => (cell.textContent || '').replace(/\\s+/g, ' ').trim().toLowerCase());
                return {
                    row,
                    sectionText: values[0] || '',
                    termText: values[1] || '',
                    chapter: chapterFromSectionText(values[0] || ''),
                };
            });

            const tokenize = (query) => {
                const tokens = [];
                const pattern = /\\s*(\\(|\\)|\\bAND\\b|\\bOR\\b|[^\\s()]+)/gi;
                let match;
                while ((match = pattern.exec(query)) !== null) {
                    const raw = match[1];
                    if (!raw) continue;
                    const upper = raw.toUpperCase();
                    if (raw === '(' || raw === ')') tokens.push({ type: raw });
                    else if (upper === 'AND' || upper === 'OR') tokens.push({ type: upper });
                    else tokens.push({ type: 'TERM', value: normalizeToken(raw) });
                }
                return tokens;
            };

            const needsImplicitAnd = (prev, next) => {
                const leftTerm = prev.type === 'TERM' || prev.type === ')';
                const rightTerm = next.type === 'TERM' || next.type === '(';
                return leftTerm && rightTerm;
            };

            const addImplicitAnd = (tokens) => {
                if (tokens.length <= 1) return tokens;
                const output = [tokens[0]];
                for (let index = 1; index < tokens.length; index += 1) {
                    const previous = output[output.length - 1];
                    const current = tokens[index];
                    if (needsImplicitAnd(previous, current)) output.push({ type: 'AND' });
                    output.push(current);
                }
                return output;
            };

            const parseQuery = (rawQuery) => {
                const tokens = addImplicitAnd(tokenize(rawQuery));
                if (tokens.length === 0) return null;
                let pointer = 0;

                const current = () => tokens[pointer];
                const consume = (type) => {
                    if (!current() || current().type !== type) return null;
                    pointer += 1;
                    return tokens[pointer - 1];
                };

                const parsePrimary = () => {
                    const token = current();
                    if (!token) throw new Error('Unexpected end of query.');
                    if (consume('(')) {
                        const expression = parseOr();
                        if (!consume(')')) throw new Error('Missing closing parenthesis.');
                        return expression;
                    }
                    if (token.type === 'TERM') {
                        pointer += 1;
                        return {
                            type: 'TERM',
                            value: token.value,
                            field: isSectionToken(token.value) ? 'section' : 'term',
                        };
                    }
                    throw new Error(`Unexpected token '${token.type}'.`);
                };

                const parseAnd = () => {
                    let left = parsePrimary();
                    while (current() && current().type === 'AND') {
                        consume('AND');
                        const right = parsePrimary();
                        left = { type: 'AND', left, right };
                    }
                    return left;
                };

                const parseOr = () => {
                    let left = parseAnd();
                    while (current() && current().type === 'OR') {
                        consume('OR');
                        const right = parseAnd();
                        left = { type: 'OR', left, right };
                    }
                    return left;
                };

                const ast = parseOr();
                if (pointer < tokens.length) throw new Error('Unexpected trailing tokens.');
                return ast;
            };

            const evaluate = (node, row) => {
                if (!node) return true;
                if (node.type === 'TERM') {
                    if (node.field === 'section') return hasMatchingSectionCode(row.sectionText, node.value);
                    return row.termText.includes(node.value);
                }
                if (node.type === 'AND') return evaluate(node.left, row) && evaluate(node.right, row);
                if (node.type === 'OR') return evaluate(node.left, row) || evaluate(node.right, row);
                return true;
            };

            const updateStatus = (visibleCount, totalCount, message, isError) => {
                status.textContent = message || `Showing ${visibleCount} of ${totalCount} issue sentences`;
                status.dataset.error = isError ? 'true' : 'false';
            };

            const applyFilter = () => {
                const rowData = currentRowData();
                const selectedChapter = chapterSelect.value || 'All';
                const query = searchInput.value.trim();
                const chapterMatches = (entry) => selectedChapter === 'All' || entry.chapter === selectedChapter;

                if (!query) {
                    let visible = 0;
                    rowData.forEach((entry) => {
                        const match = chapterMatches(entry);
                        entry.row.hidden = !match;
                        if (match) visible += 1;
                    });
                    updateStatus(visible, rowData.length, '', false);
                    return;
                }

                let ast;
                try {
                    ast = parseQuery(query);
                } catch (error) {
                    let visible = 0;
                    rowData.forEach((entry) => {
                        const match = chapterMatches(entry);
                        entry.row.hidden = !match;
                        if (match) visible += 1;
                    });
                    updateStatus(
                        visible,
                        rowData.length,
                        `Invalid query: ${error.message} Use section numbers/terms with AND, OR, and parentheses.`,
                        true,
                    );
                    return;
                }

                let visible = 0;
                rowData.forEach((entry) => {
                    const match = chapterMatches(entry) && evaluate(ast, entry);
                    entry.row.hidden = !match;
                    if (match) visible += 1;
                });
                updateStatus(visible, rowData.length, '', false);
            };

            const visibleRows = () => activeRows().filter((row) => !row.hidden);

            const exportTable = () => {
                const table = panel.querySelector('.glossary-issue-table').cloneNode(true);
                const body = table.querySelector('tbody');
                body.replaceChildren(...visibleRows().map((row) => row.cloneNode(true)));
                return table.outerHTML;
            };

            const exportDocument = () => `<!doctype html>
<html lang="en"><head><meta charset="utf-8"><title>Glossary Issue Table</title>
<style>body{font-family:Arial,sans-serif;margin:2rem}table{border-collapse:collapse;width:100%}th,td{border:1px solid #999;padding:.4rem;text-align:left;vertical-align:top}th{background:#eee}</style>
</head><body><h1>Glossary Issue Table</h1>${exportTable()}</body></html>`;

            const downloadFile = (content, filename, type) => {
                const link = document.createElement('a');
                link.href = URL.createObjectURL(new Blob([content], { type }));
                link.download = filename;
                link.click();
                URL.revokeObjectURL(link.href);
            };

            const printVisibleRows = () => {
                const printWindow = window.open('', '_blank');
                if (!printWindow) return;
                printWindow.document.open();
                printWindow.document.write(exportDocument());
                printWindow.document.close();
                printWindow.focus();
                printWindow.print();
            };

            chapterSelect.addEventListener('change', applyFilter);
            searchInput.addEventListener('input', applyFilter);
            if (downloadHtml) downloadHtml.addEventListener('click', () => downloadFile(exportDocument(), 'glossary-issues.html', 'text/html'));
            if (downloadPdf) downloadPdf.addEventListener('click', printVisibleRows);
            applyFilter();
        })();
    </script>
"""

CAE_CSS = """
            .cae-overview { padding: clamp(1.25rem, 3vw, 3rem); }
            .cae-report-filter {
                margin-top: 1rem;
                max-width: 22rem;
                position: relative;
                z-index: 2;
            }
            .cae-report-filter summary {
                background: var(--paper);
                border: 1px solid #8aaeb9;
                border-radius: 3px;
                cursor: pointer;
                display: grid;
                gap: .1rem .75rem;
                grid-template-columns: minmax(0, 1fr) auto;
                list-style: none;
                padding: .55rem .7rem;
            }
            .cae-report-filter summary::-webkit-details-marker { display: none; }
            .cae-report-filter summary::after {
                align-self: center;
                border-left: .3rem solid transparent;
                border-right: .3rem solid transparent;
                border-top: .4rem solid currentColor;
                content: "";
                grid-column: 2;
                grid-row: 1 / 3;
                transition: transform .15s ease;
            }
            .cae-report-filter[open] summary::after { transform: rotate(180deg); }
            .cae-report-filter summary:focus-visible {
                outline: 3px solid var(--ipcc-blue);
                outline-offset: 2px;
            }
            .cae-report-filter-label {
                color: var(--muted);
                font-size: .75rem;
                font-weight: 700;
            }
            .cae-report-filter-value {
                font-size: .92rem;
                font-weight: 700;
                overflow-wrap: anywhere;
            }
            .cae-report-filter-menu {
                background: var(--paper);
                border: 1px solid #8aaeb9;
                box-shadow: 0 .45rem 1rem rgb(22 52 61 / 16%);
                box-sizing: border-box;
                left: 0;
                margin: .25rem 0 0;
                padding: .45rem .7rem .6rem;
                position: absolute;
                top: 100%;
                width: 100%;
            }
            .cae-report-filter-menu legend {
                color: var(--muted);
                font-size: .75rem;
                font-weight: 700;
                padding: 0 .2rem;
            }
            .cae-report-option {
                align-items: center;
                cursor: pointer;
                display: flex;
                gap: .55rem;
                min-height: 2rem;
            }
            .cae-report-option input {
                accent-color: var(--ipcc-blue);
                height: 1rem;
                margin: 0;
                width: 1rem;
            }
            .cae-report-option input:focus-visible {
                outline: 3px solid var(--ipcc-blue);
                outline-offset: 2px;
            }
            .cae-report-all-option {
                border-bottom: 1px solid var(--rule);
                font-weight: 700;
                margin-bottom: .25rem;
                padding-bottom: .25rem;
            }
            .cae-section + .cae-section {
                border-top: 1px solid var(--rule);
                margin-top: 2rem;
                padding-top: 1.5rem;
            }
            .cae-section h2 {
                color: var(--ipcc-blue);
                font-size: 1.35rem;
                margin-bottom: .75rem;
            }
            .cae-table-wrap { overflow-x: auto; }
            .cae-table {
                border-collapse: collapse;
                min-width: 42rem;
                width: 100%;
            }
            .cae-table th,
            .cae-table td {
                border: 1px solid var(--rule);
                padding: .65rem .75rem;
                text-align: left;
                vertical-align: top;
            }
            .cae-table thead th { background: #edf5f5; color: #004f6a; }
            .cae-count-table th,
            .cae-count-table td,
            .cae-matrix td:not(:first-child) { text-align: center; }
            .cae-matrix tbody th { white-space: nowrap; }
            .cae-review-table td:first-child { white-space: nowrap; width: 13rem; }
            .cae-issue-label {
                color: #7a4b2a;
                display: block;
                font-size: .85rem;
                margin-top: .4rem;
            }
            .cae-empty { color: var(--muted); }
            @media (max-width: 42rem) {
                .cae-report-filter { max-width: none; }
            }
"""

CAE_JAVASCRIPT = """
        <script>
            (() => {
                const panel = document.getElementById("cae-check-panel");
                const dataElement = panel?.querySelector(".cae-filter-data");
                if (!panel || !dataElement) return;

                const filterData = JSON.parse(dataElement.textContent);
                const reportFilter = panel.querySelector(".cae-report-filter");
                const allCheckbox = panel.querySelector(".cae-report-all");
                const reportCheckboxes = Array.from(panel.querySelectorAll(".cae-report-checkbox"));
                const filterValue = panel.querySelector(".cae-report-filter-value");
                const facts = panel.querySelector(".facts");
                const pairHeading = panel.querySelector("#cae-pair-heading");
                const confidenceHeading = panel.querySelector("#cae-confidence-heading");
                const reviewHeading = panel.querySelector("#cae-review-heading");
                const reviewRows = Array.from(panel.querySelectorAll(".cae-review-table tbody tr"));
                const reviewWrap = panel.querySelector(".cae-review-wrap");
                const reviewEmpty = panel.querySelector(".cae-review-empty");

                const selectedReports = () => reportCheckboxes
                    .filter((checkbox) => checkbox.checked)
                    .map((checkbox) => checkbox.value);

                const updateResults = () => {
                    const selected = selectedReports();
                    const selectedSet = new Set(selected);
                    const allSelected = selected.length === filterData.reports.length;
                    if (allCheckbox) {
                        allCheckbox.checked = allSelected;
                        allCheckbox.indeterminate = selected.length > 0 && !allSelected;
                    }
                    if (filterValue) {
                        filterValue.textContent = allSelected
                            ? "All reports"
                            : selected.length === 0
                                ? "No reports selected"
                                : selected.length === 1
                                    ? selected[0]
                                    : `${selected.length} reports selected`;
                    }

                    let pairCount = 0;
                    panel.querySelectorAll("[data-agreement][data-evidence]").forEach((cell) => {
                        const count = selected.reduce(
                            (total, report) => total + filterData.counts[report]
                                .agreementEvidence[cell.dataset.agreement][cell.dataset.evidence],
                            0,
                        );
                        cell.textContent = String(count);
                        pairCount += count;
                    });

                    let confidenceCount = 0;
                    panel.querySelectorAll("[data-confidence]").forEach((cell) => {
                        const count = selected.reduce(
                            (total, report) => total + filterData.counts[report].confidence[cell.dataset.confidence],
                            0,
                        );
                        cell.textContent = String(count);
                        confidenceCount += count;
                    });

                    let issueCount = 0;
                    reviewRows.forEach((row) => {
                        const visible = selectedSet.has(row.dataset.report);
                        row.hidden = !visible;
                        if (visible) issueCount += 1;
                    });
                    if (reviewWrap) reviewWrap.hidden = issueCount === 0;
                    if (reviewEmpty) reviewEmpty.hidden = issueCount !== 0;

                    if (pairHeading) pairHeading.textContent = `Agreement and evidence (${pairCount})`;
                    if (confidenceHeading) confidenceHeading.textContent = `Confidence (${confidenceCount})`;
                    if (reviewHeading) reviewHeading.textContent = `Cases requiring review (${issueCount})`;
                    if (facts) {
                        const validCount = pairCount + confidenceCount;
                        facts.textContent = `${validCount} valid sentence-final assessment${validCount === 1 ? "" : "s"}; `
                            + `${issueCount} case${issueCount === 1 ? "" : "s"} `
                            + `${issueCount === 1 ? "requires" : "require"} review`;
                    }
                };

                allCheckbox?.addEventListener("change", () => {
                    reportCheckboxes.forEach((checkbox) => { checkbox.checked = allCheckbox.checked; });
                    updateResults();
                });
                reportCheckboxes.forEach((checkbox) => checkbox.addEventListener("change", updateResults));
                document.addEventListener("click", (event) => {
                    if (reportFilter?.open && !reportFilter.contains(event.target)) reportFilter.open = false;
                });
                updateResults();
            })();
        </script>
"""

GLOSSARY_JAVASCRIPT = """
        <script>
            (() => {
                const panel = document.getElementById("glossary-overview-panel");
                if (!panel) return;
                const searchInput = panel.querySelector(".glossary-search");
                const unusedToggle = panel.querySelector(".glossary-unused-toggle");
                const resultCount = panel.querySelector(".glossary-result-count");
                const rows = Array.from(panel.querySelectorAll(".glossary-term-row"));
                const buttons = Array.from(panel.querySelectorAll(".glossary-term-button:not(:disabled)"));
                const details = Array.from(panel.querySelectorAll(".glossary-detail"));
                const placeholder = panel.querySelector(".glossary-detail-placeholder");
                const workspace = panel.querySelector(".glossary-workspace");
                const divider = panel.querySelector(".glossary-divider");
                const definitionDialog = document.getElementById("glossary-definition-dialog");
                const dialogTitle = definitionDialog?.querySelector(".glossary-dialog-title");
                const dialogContent = definitionDialog?.querySelector(".glossary-dialog-content");
                const dialogClose = definitionDialog?.querySelector(".glossary-dialog-close");
                const paragraphDialog = document.getElementById("glossary-paragraph-dialog");
                const paragraphDialogTitle = paragraphDialog?.querySelector(".glossary-dialog-title");
                const paragraphDialogLocation = paragraphDialog?.querySelector(".glossary-paragraph-location");
                const paragraphDialogText = paragraphDialog?.querySelector(".glossary-paragraph-text");
                const paragraphDialogClose = paragraphDialog?.querySelector(".glossary-dialog-close");
                const issuePanel = document.getElementById("glossary-issue-table-panel");

                const normalizeIssueTerm = (value) => value.replace(/\\s+/g, " ").trim().toLocaleLowerCase();

                const issueTargetsFromTable = () => {
                    if (!issuePanel) return [];
                    const targets = [];
                    issuePanel.querySelectorAll(".glossary-issue-table tbody tr").forEach((row) => {
                        const sourceButton = row.querySelector('td:first-child .glossary-evidence-code[data-source-node-id]');
                        const termCell = row.querySelector("td:nth-child(2)");
                        const key = normalizeIssueTerm(termCell?.textContent || "");
                        const nodeId = sourceButton?.dataset.sourceNodeId || "";
                        if (nodeId && key) {
                            targets.push({ nodeId, termKey: key });
                        }
                    });
                    return targets;
                };

                const flagIssueTermsInReportText = () => {
                    const targets = issueTargetsFromTable();
                    const links = document.querySelectorAll('section.report-panel[id^="chapter-panel-"] a.report-glossary-link');
                    links.forEach((link) => {
                        link.classList.remove("has-issue");
                        link.querySelector(".report-glossary-link-issue-flag")?.remove();
                    });

                    targets.forEach(({ nodeId, termKey }) => {
                        const sourceNode = document.querySelector(`[data-node-id="${CSS.escape(nodeId)}"]`);
                        if (!sourceNode) return;
                        if (!sourceNode.closest('section.report-panel[id^="chapter-panel-"]')) return;

                        sourceNode.querySelectorAll("a.report-glossary-link").forEach((link) => {
                            const linkTerm = normalizeIssueTerm(link.dataset.term || link.textContent || "");
                            if (linkTerm !== termKey) return;
                            if (link.querySelector(".report-glossary-link-issue-flag")) return;
                            link.classList.add("has-issue");
                            const flag = document.createElement("span");
                            flag.className = "report-glossary-link-issue-flag";
                            flag.setAttribute("aria-hidden", "true");
                            flag.textContent = "!";
                            link.append(flag);
                        });
                    });
                };

                const filterTerms = () => {
                    const query = searchInput?.value.trim().toLocaleLowerCase() || "";
                    const hideUnused = unusedToggle?.getAttribute("aria-pressed") === "true";
                    let visibleCount = 0;
                    rows.forEach((row) => {
                        const matchesQuery = !query || row.dataset.search.includes(query);
                        const visible = matchesQuery && (!hideUnused || row.dataset.usageCount !== "0");
                        row.hidden = !visible;
                        if (visible) visibleCount += 1;
                    });
                    if (resultCount) {
                        const qualifier = query && hideUnused
                            ? "matching used glossary"
                            : query
                                ? "matching glossary"
                                : hideUnused
                                    ? "used glossary"
                                    : "glossary";
                        resultCount.textContent = `${visibleCount} ${qualifier} term${visibleCount === 1 ? "" : "s"}`;
                    }
                };

                const showTerm = (button, updateHash = true) => {
                    const detailId = button.dataset.detailId;
                    buttons.forEach((candidate) => candidate.setAttribute("aria-pressed", String(candidate === button)));
                    details.forEach((detail) => { detail.hidden = detail.id !== detailId; });
                    if (placeholder) placeholder.hidden = true;
                    const detail = detailId ? document.getElementById(detailId) : null;
                    if (updateHash && detail && window.history.replaceState) {
                        window.history.replaceState(null, "", `#${detail.id}`);
                    }
                    detail?.scrollIntoView({ behavior: "smooth", block: "nearest" });
                };

                const createInlineEvidenceButton = (sectionCode) => {
                    const button = document.createElement("button");
                    button.className = "glossary-evidence-code glossary-inline-evidence-code";
                    button.type = "button";
                    button.dataset.sectionCode = sectionCode;
                    button.textContent = sectionCode;
                    return button;
                };

                const splitIdentifiers = (text) => text
                    .split(/\\s*[;,]\\s*/)
                    .map((value) => value.trim())
                    .filter(Boolean);

                const identifierLooksLikeSectionCode = (value) => {
                    if (!value || value.length > 64) return false;
                    if (!/^[A-Za-z0-9][A-Za-z0-9 .:/+-]*$/.test(value)) return false;
                    return /\\bP\\d+\\b/.test(value)
                        || /\\d/.test(value)
                        || /^(SPM|TS|ES|Box|Figure|D-Figure|C-Figure)/.test(value);
                };

                const findEvidenceButton = (sectionCode) => {
                    const activeDetail = details.find((detail) => !detail.hidden);
                    const scope = activeDetail || panel;
                    const primary = Array.from(scope.querySelectorAll("button.glossary-evidence-code:not(.glossary-inline-evidence-code)"));
                    const fallback = Array.from(panel.querySelectorAll("button.glossary-evidence-code:not(.glossary-inline-evidence-code)"));
                    const exact = primary.find((candidate) => candidate.textContent.trim() === sectionCode)
                        || fallback.find((candidate) => candidate.textContent.trim() === sectionCode);
                    if (exact) return exact;
                    return primary.find((candidate) => candidate.textContent.trim().startsWith(`${sectionCode} `))
                        || fallback.find((candidate) => candidate.textContent.trim().startsWith(`${sectionCode} `))
                        || null;
                };

                const formatLegacyContextLists = () => {
                    panel.querySelectorAll(".glossary-llm-check-content h4").forEach((heading) => {
                        if (heading.textContent.trim().toLocaleLowerCase() !== "contexts of use") return;
                        const contextList = heading.nextElementSibling;
                        if (!contextList || contextList.tagName !== "UL") return;
                        contextList.querySelectorAll(":scope > li").forEach((item) => {
                            if (item.dataset.contextFormatted === "true") return;
                            const title = item.querySelector(":scope > strong");
                            const text = item.textContent.replace(/\\s+/g, " ").trim();
                            const sampleMarker = "Sample IDs:";
                            const sampleIndex = text.indexOf(sampleMarker);
                            if (!title) return;

                            const titleText = title.textContent.trim().replace(/:\\s*$/, "");
                            let description = "";
                            let identifiers = [];

                            if (sampleIndex !== -1) {
                                const beforeSamples = text.slice(0, sampleIndex).trim()
                                    .replace(/\\s+Reports:\\s*[^.]+\\.?$/i, "");
                                description = beforeSamples.startsWith(title.textContent.trim())
                                    ? beforeSamples.slice(title.textContent.trim().length).replace(/^:\\s*/, "").trim()
                                    : beforeSamples;
                                identifiers = text.slice(sampleIndex + sampleMarker.length)
                                    .replace(/\\.$/, "")
                                    .split(/\\s*,\\s*/)
                                    .map((identifier) => identifier.trim())
                                    .filter(Boolean);
                            } else {
                                const parenthesizedMatch = text.match(/^[^()]+\\(([^)]+)\\):\\s*(.*)$/);
                                if (!parenthesizedMatch) return;
                                identifiers = splitIdentifiers(parenthesizedMatch[1]);
                                description = parenthesizedMatch[2].trim();
                            }

                            identifiers = identifiers.filter(identifierLooksLikeSectionCode);
                            if (!identifiers.length) return;

                            item.replaceChildren();
                            const strong = document.createElement("strong");
                            strong.textContent = titleText;
                            item.append(strong, document.createTextNode(" ("));
                            identifiers.forEach((identifier, index) => {
                                if (index) item.append(document.createTextNode(", "));
                                item.append(createInlineEvidenceButton(identifier));
                            });
                            item.append(document.createTextNode(`): ${description}`));
                            item.dataset.contextFormatted = "true";
                        });
                    });
                };

                const formatPotentialIssueEvidence = () => {
                    panel.querySelectorAll(".glossary-llm-check-content h4").forEach((heading) => {
                        const headingText = heading.textContent.trim().toLocaleLowerCase();
                        if (!headingText.startsWith("potential issues")) return;
                        const issueList = heading.nextElementSibling;
                        if (!issueList || issueList.tagName !== "UL") return;
                        issueList.querySelectorAll(":scope > li").forEach((item) => {
                            if (item.dataset.issueFormatted === "true") return;
                            if (item.querySelector(":scope > strong")) return;
                            const text = item.textContent.replace(/\\s+/g, " ").trim();

                            const leadingMatch = text.match(/^\[([^\]]+)\]\\s*(.*)$/);
                            if (leadingMatch) {
                                const sectionCode = leadingMatch[1].trim();
                                if (!identifierLooksLikeSectionCode(sectionCode)) return;
                                item.replaceChildren(createInlineEvidenceButton(sectionCode));
                                if (leadingMatch[2]) {
                                    item.append(document.createTextNode(` ${leadingMatch[2]}`));
                                }
                                item.dataset.issueFormatted = "true";
                                return;
                            }

                            const structuredMatch = text.match(/^(.*?ID\(s\):\\s*)\[([^\]]+)\](.*)$/i);
                            if (!structuredMatch) return;
                            const identifiers = splitIdentifiers(structuredMatch[2]).filter(identifierLooksLikeSectionCode);
                            if (!identifiers.length) return;

                            item.replaceChildren(document.createTextNode(structuredMatch[1]));
                            identifiers.forEach((identifier, index) => {
                                if (index) item.append(document.createTextNode(", "));
                                item.append(createInlineEvidenceButton(identifier));
                            });
                            if (structuredMatch[3]) {
                                item.append(document.createTextNode(structuredMatch[3]));
                            }
                            item.dataset.issueFormatted = "true";
                        });
                    });
                };

                const formatGlossaryNameRows = () => {
                    details.forEach((detail) => {
                        if (detail.querySelector(".glossary-canonical-term")) return;

                        const heading = detail.querySelector(".glossary-detail-heading");
                        const title = heading?.querySelector("h3");
                        const totalCount = heading?.querySelector(".glossary-term-count");
                        const aliases = detail.querySelector(".glossary-aliases");
                        const names = [];
                        let pendingText = "";

                        const normalizeName = (name) => name.trim().toLocaleLowerCase();
                        const evidenceMarks = Array.from(detail.querySelectorAll(".glossary-evidence mark"));
                        const useCounts = new Map();
                        evidenceMarks.forEach((mark) => {
                            const name = normalizeName(mark.textContent || "");
                            if (name) useCounts.set(name, (useCounts.get(name) || 0) + 1);
                        });
                        const nameUseCount = (name) => useCounts.get(normalizeName(name)) || 0;
                        const totalUseCount = evidenceMarks.length;
                        if (totalCount) totalCount.textContent = `[${totalUseCount}]`;
                        const button = buttons.find((candidate) => candidate.dataset.detailId === detail.id);
                        const buttonCount = button?.querySelector(".glossary-term-count");
                        if (buttonCount) buttonCount.textContent = `[${totalUseCount}]`;
                        const row = button?.closest(".glossary-term-row");
                        if (row) row.dataset.usageCount = String(totalUseCount);
                        const evidence = detail.querySelector(".glossary-evidence");
                        const evidenceSummary = evidence?.querySelector("summary");
                        if (evidenceSummary) {
                            const sentenceRowCount = evidence.querySelectorAll("tbody > tr").length;
                            const sentenceLabel = sentenceRowCount === 1 ? "sentence" : "sentences";
                            const useLabel = totalUseCount === 1 ? "use" : "uses";
                            evidenceSummary.textContent = `Term use overview table (${sentenceRowCount} ${sentenceLabel}; ${totalUseCount} ${useLabel})`;
                        }

                        aliases?.childNodes.forEach((node) => {
                            if (node.nodeType === Node.TEXT_NODE) {
                                pendingText += node.textContent || "";
                                return;
                            }
                            if (node.nodeType !== Node.ELEMENT_NODE
                                || !node.classList.contains("glossary-term-count")) {
                                pendingText += node.textContent || "";
                                return;
                            }

                            const rawName = pendingText.trim();
                            const name = rawName.startsWith("Also known as:")
                                ? rawName.slice("Also known as:".length).trim()
                                : rawName.startsWith(",")
                                    ? rawName.slice(1).trim()
                                    : rawName;
                            if (name) names.push(name);
                            pendingText = "";
                        });

                        const canonicalName = title?.textContent.trim();
                        if (!canonicalName || !heading) return;

                        const canonicalRow = document.createElement("p");
                        canonicalRow.className = "glossary-canonical-term";
                        const canonicalCount = document.createElement("span");
                        canonicalCount.className = "glossary-term-count";
                        canonicalCount.textContent = `[${nameUseCount(canonicalName)}]`;
                        canonicalRow.append(
                            document.createTextNode("Canonical term: "),
                            document.createTextNode(`${canonicalName} `),
                            canonicalCount,
                        );

                        const aliasRow = aliases || document.createElement("p");
                        aliasRow.className = "glossary-aliases";
                        aliasRow.replaceChildren(document.createTextNode("Alias(es): "));
                        if (names.length) {
                            names.forEach((name, index) => {
                                if (index) aliasRow.append(document.createTextNode(", "));
                                const aliasCount = document.createElement("span");
                                aliasCount.className = "glossary-term-count";
                                aliasCount.textContent = `[${nameUseCount(name)}]`;
                                aliasRow.append(
                                    document.createTextNode(`${name} `),
                                    aliasCount,
                                );
                            });
                        } else {
                            aliasRow.append(document.createTextNode("None"));
                        }

                        if (aliases) {
                            aliases.before(canonicalRow);
                        } else {
                            heading.after(canonicalRow, aliasRow);
                        }
                    });
                };

                const setDividerPosition = (percentage) => {
                    const constrained = Math.round(Math.min(60, Math.max(20, percentage)) * 10) / 10;
                    workspace?.style.setProperty("--glossary-index-width", `${constrained}%`);
                    divider?.setAttribute("aria-valuenow", String(Math.round(constrained)));
                    divider?.setAttribute("aria-valuetext", `${Math.round(constrained)}% glossary overview width`);
                };

                const resizeFromPointer = (event) => {
                    if (!workspace) return;
                    const bounds = workspace.getBoundingClientRect();
                    setDividerPosition(((event.clientX - bounds.left) / bounds.width) * 100);
                };

                divider?.addEventListener("pointerdown", (event) => {
                    divider.setPointerCapture(event.pointerId);
                    divider.classList.add("is-dragging");
                    resizeFromPointer(event);
                });
                divider?.addEventListener("pointermove", (event) => {
                    if (divider.hasPointerCapture(event.pointerId)) resizeFromPointer(event);
                });
                const finishResize = (event) => {
                    if (divider.hasPointerCapture(event.pointerId)) divider.releasePointerCapture(event.pointerId);
                    divider.classList.remove("is-dragging");
                };
                divider?.addEventListener("pointerup", finishResize);
                divider?.addEventListener("pointercancel", finishResize);
                divider?.addEventListener("keydown", (event) => {
                    if (event.key !== "ArrowLeft" && event.key !== "ArrowRight") return;
                    event.preventDefault();
                    const current = Number(divider.getAttribute("aria-valuenow")) || 30;
                    setDividerPosition(current + (event.key === "ArrowRight" ? 2 : -2));
                });

                searchInput?.addEventListener("input", filterTerms);
                unusedToggle?.addEventListener("click", () => {
                    const hideUnused = unusedToggle.getAttribute("aria-pressed") !== "true";
                    unusedToggle.setAttribute("aria-pressed", String(hideUnused));
                    unusedToggle.textContent = hideUnused
                        ? "Show terms used 0 times"
                        : "Hide terms used 0 times";
                    filterTerms();
                });
                buttons.forEach((button) => button.addEventListener("click", () => showTerm(button)));
                document.addEventListener("click", (event) => {
                    const inlineReferenceButton = event.target.closest("button.glossary-inline-evidence-code");
                    if (inlineReferenceButton && paragraphDialog && paragraphDialogTitle && paragraphDialogLocation && paragraphDialogText) {
                        const sectionCode = inlineReferenceButton.dataset.sectionCode || inlineReferenceButton.textContent.trim();
                        const evidenceButton = findEvidenceButton(sectionCode);
                        if (!evidenceButton) {
                            return;
                        }

                        const nodeId = evidenceButton.dataset.sourceNodeId;
                        const nodeIdAttribute = ["data", "node", "id"].join("-");
                        const sourceNode = nodeId
                            ? document.querySelector(`[${nodeIdAttribute}="${CSS.escape(nodeId)}"]`)
                            : null;
                        const sourceContent = sourceNode?.querySelector(".paragraph, .figure-explanation");
                        if (!sourceContent) {
                            return;
                        }
                        const sourceClone = sourceContent.cloneNode(true);
                        sourceClone.querySelectorAll(".node-code, .back-to-top").forEach((element) => element.remove());
                        paragraphDialogTitle.textContent = sectionCode;
                        paragraphDialogLocation.textContent = evidenceButton.closest("td")
                            ?.querySelector(".glossary-evidence-location")?.textContent || "";
                            paragraphDialogText.textContent = sourceClone.textContent.trim().replace(/\\s+/g, " ");
                        paragraphDialog.showModal();
                        return;
                    }

                    const codeButton = event.target.closest("button.glossary-evidence-code");
                    if (codeButton && paragraphDialog && paragraphDialogTitle && paragraphDialogLocation && paragraphDialogText) {
                        const nodeId = codeButton.dataset.sourceNodeId;
                        const nodeIdAttribute = ["data", "node", "id"].join("-");
                        const sourceNode = nodeId
                            ? document.querySelector(`[${nodeIdAttribute}="${CSS.escape(nodeId)}"]`)
                            : null;
                        const sourceContent = sourceNode?.querySelector(".paragraph, .figure-explanation");
                        if (sourceContent) {
                            const sourceClone = sourceContent.cloneNode(true);
                            sourceClone.querySelectorAll(".node-code, .back-to-top").forEach((element) => element.remove());
                            paragraphDialogTitle.textContent = codeButton.textContent.trim();
                            paragraphDialogLocation.textContent = codeButton.closest("td")
                                ?.querySelector(".glossary-evidence-location")?.textContent || "";
                            paragraphDialogText.textContent = sourceClone.textContent.trim().replace(/\\s+/g, " ");
                            paragraphDialog.showModal();
                        }
                        return;
                    }
                    const link = event.target.closest("a.report-glossary-link");
                    if (!link || !definitionDialog || !dialogTitle || !dialogContent) return;
                    event.preventDefault();
                    const sourceDetail = details.find((detail) => detail.dataset.term === link.dataset.term);
                    if (!sourceDetail) return;
                    dialogTitle.textContent = link.dataset.term;
                    dialogContent.replaceChildren();
                    sourceDetail.querySelectorAll(".glossary-canonical-term, .glossary-aliases, .glossary-parent, .glossary-child").forEach((detail) => {
                        dialogContent.append(detail.cloneNode(true));
                    });
                    sourceDetail.querySelectorAll(".glossary-definition").forEach((definition) => {
                        dialogContent.append(definition.cloneNode(true));
                    });
                    definitionDialog.showModal();
                });
                dialogClose?.addEventListener("click", () => definitionDialog.close());
                definitionDialog?.addEventListener("click", (event) => {
                    if (event.target === definitionDialog) definitionDialog.close();
                });
                paragraphDialogClose?.addEventListener("click", () => paragraphDialog.close());
                paragraphDialog?.addEventListener("click", (event) => {
                    if (event.target === paragraphDialog) paragraphDialog.close();
                });
                formatGlossaryNameRows();
                formatLegacyContextLists();
                formatPotentialIssueEvidence();
                flagIssueTermsInReportText();
                filterTerms();

                const hashDetail = window.location.hash ? document.getElementById(window.location.hash.slice(1)) : null;
                const hashButton = hashDetail?.classList.contains("glossary-detail")
                    ? buttons.find((button) => button.dataset.detailId === hashDetail.id)
                    : null;
                if (hashButton) showTerm(hashButton, false);
            })();
        </script>
"""

GLOSSARY_DIALOG_MARKUP = """
        <dialog class="glossary-definition-dialog" id="glossary-definition-dialog" aria-labelledby="glossary-dialog-title">
            <header class="glossary-dialog-header">
                <h2 class="glossary-dialog-title" id="glossary-dialog-title">Glossary definition</h2>
                <button class="glossary-dialog-close" type="button" aria-label="Close definitions" title="Close">&#215;</button>
            </header>
            <div class="glossary-dialog-content"></div>
        </dialog>
        <dialog class="glossary-definition-dialog glossary-paragraph-dialog" id="glossary-paragraph-dialog" aria-labelledby="glossary-paragraph-dialog-title">
            <header class="glossary-dialog-header">
                <h2 class="glossary-dialog-title" id="glossary-paragraph-dialog-title">Source paragraph</h2>
                <button class="glossary-dialog-close" type="button" aria-label="Close source paragraph" title="Close">&#215;</button>
            </header>
            <div class="glossary-dialog-content">
                <p class="glossary-paragraph-location"></p>
                <p class="glossary-paragraph-text"></p>
            </div>
        </dialog>
"""


@dataclass(frozen=True)
class ElementSpan:
    """The source range and attributes for a parsed HTML element."""

    start: int
    end: int
    attributes: dict[str, str | None]


def class_names(attributes: dict[str, str | None]) -> set[str]:
    """Return normalized CSS class names for an element."""
    return set((attributes.get("class") or "").split())


class ReportMarkupParser(HTMLParser):
    """Locate the reference document's tab strip and top-level report panels."""

    def __init__(self, markup: str) -> None:
        super().__init__(convert_charrefs=False)
        self.markup = markup
        self.line_offsets = [0]
        self.line_offsets.extend(match.end() for match in re.finditer("\n", markup))
        self.navigation: ElementSpan | None = None
        self.panels: list[ElementSpan] = []
        self._navigation_start: int | None = None
        self._navigation_depth = 0
        self._navigation_attributes: dict[str, str | None] | None = None
        self._panel_start: int | None = None
        self._panel_depth = 0
        self._panel_attributes: dict[str, str | None] | None = None

    def absolute_offset(self) -> int:
        line_number, column = self.getpos()
        return self.line_offsets[line_number - 1] + column

    def end_tag_offset(self) -> int:
        closing_bracket = self.markup.find(">", self.absolute_offset())
        if closing_bracket == -1:
            raise ValueError("Encountered an unterminated HTML closing tag.")
        return closing_bracket + 1

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        attributes = dict(attrs)
        start = self.absolute_offset()

        if tag == "nav":
            if self._navigation_start is not None:
                self._navigation_depth += 1
            elif "chapter-tabs" in class_names(attributes):
                self._navigation_start = start
                self._navigation_depth = 1
                self._navigation_attributes = attributes

        if tag == "section":
            if self._panel_start is not None:
                self._panel_depth += 1
            elif "report-panel" in class_names(attributes):
                self._panel_start = start
                self._panel_depth = 1
                self._panel_attributes = attributes

    def handle_endtag(self, tag: str) -> None:
        if tag == "nav" and self._navigation_start is not None:
            self._navigation_depth -= 1
            if self._navigation_depth == 0:
                self.navigation = ElementSpan(
                    self._navigation_start,
                    self.end_tag_offset(),
                    self._navigation_attributes or {},
                )
                self._navigation_start = None
                self._navigation_attributes = None

        if tag == "section" and self._panel_start is not None:
            self._panel_depth -= 1
            if self._panel_depth == 0:
                self.panels.append(
                    ElementSpan(
                        self._panel_start,
                        self.end_tag_offset(),
                        self._panel_attributes or {},
                    )
                )
                self._panel_start = None
                self._panel_attributes = None

    def result(self) -> tuple[ElementSpan, list[ElementSpan]]:
        if self._navigation_start is not None:
            raise ValueError("The chapter navigation is not closed.")
        if self._panel_start is not None:
            raise ValueError("A report panel is not closed.")
        if self.navigation is None:
            raise ValueError("Could not find the chapter navigation in the reference HTML.")
        if not self.panels:
            raise ValueError("Could not find report panels in the reference HTML.")
        return self.navigation, self.panels


class ChapterTabParser(HTMLParser):
    """Locate chapter-tab button ranges without reserializing their markup."""

    def __init__(self, markup: str) -> None:
        super().__init__(convert_charrefs=False)
        self.markup = markup
        self.line_offsets = [0]
        self.line_offsets.extend(match.end() for match in re.finditer("\n", markup))
        self.tabs: list[ElementSpan] = []
        self._open_tab_start: int | None = None
        self._open_tab_attributes: dict[str, str | None] | None = None

    def absolute_offset(self) -> int:
        line_number, column = self.getpos()
        return self.line_offsets[line_number - 1] + column

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        attributes = dict(attrs)
        if tag == "button" and "chapter-tab" in class_names(attributes):
            if self._open_tab_start is not None:
                raise ValueError("Chapter tab buttons must not be nested.")
            self._open_tab_start = self.absolute_offset()
            self._open_tab_attributes = attributes

    def handle_endtag(self, tag: str) -> None:
        if tag != "button" or self._open_tab_start is None:
            return
        closing_bracket = self.markup.find(">", self.absolute_offset())
        if closing_bracket == -1:
            raise ValueError("Encountered an unterminated chapter-tab button.")
        self.tabs.append(
            ElementSpan(
                self._open_tab_start,
                closing_bracket + 1,
                self._open_tab_attributes or {},
            )
        )
        self._open_tab_start = None
        self._open_tab_attributes = None

    def result(self) -> list[ElementSpan]:
        if self._open_tab_start is not None:
            raise ValueError("A chapter-tab button is not closed.")
        if not self.tabs:
            raise ValueError("Could not find chapter-tab buttons in the navigation.")
        return self.tabs


class NodeCodeParser(HTMLParser):
    """Map report node IDs to the codes displayed in their own headings."""

    VOID_TAGS = {"area", "base", "br", "col", "embed", "hr", "img", "input", "link", "meta", "source", "track", "wbr"}

    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self.node_stack: list[str | None] = []
        self.codes: dict[str, str] = {}
        self.capture_node_id: str | None = None
        self.capture_depth = 0
        self.capture_text: list[str] = []

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        attributes = dict(attrs)
        parent_node_id = self.node_stack[-1] if self.node_stack else None
        node_id = attributes.get("data-node-id") or parent_node_id
        if tag not in self.VOID_TAGS:
            self.node_stack.append(node_id)
        if "node-code" in class_names(attributes) and node_id and node_id not in self.codes:
            self.capture_node_id = node_id
            self.capture_depth = len(self.node_stack)
            self.capture_text = []

    def handle_data(self, data: str) -> None:
        if self.capture_node_id is not None:
            self.capture_text.append(data)

    def handle_endtag(self, tag: str) -> None:
        if self.capture_node_id is not None and len(self.node_stack) == self.capture_depth:
            code = "".join(self.capture_text).strip().strip("[]")
            if code:
                self.codes[self.capture_node_id] = code
            self.capture_node_id = None
            self.capture_depth = 0
            self.capture_text = []
        if tag not in self.VOID_TAGS and self.node_stack:
            self.node_stack.pop()


def report_node_codes(markup: str) -> dict[str, str]:
    """Return the canonical visible code for each encoded report node."""
    parser = NodeCodeParser()
    parser.feed(markup)
    parser.close()
    return parser.codes


def is_word_character(character: str) -> bool:
    """Return whether a character prevents a boundary-safe term match."""
    return character.isalnum() or character == "_"


def glossary_name_matches(
    text: str,
    match_map: GlossaryMatchMap,
    term_pattern: re.Pattern[str] | None,
) -> list[tuple[int, int, str, tuple[str, ...]]]:
    """Find boundary-safe canonical terms and revised aliases in source text."""
    if term_pattern is None:
        return []

    matches = []
    for match in term_pattern.finditer(text):
        start, end = match.span()
        matched_name = match.group(0)
        if matched_name and matched_name[0].isalnum() and start > 0 and is_word_character(text[start - 1]):
            continue
        if matched_name and matched_name[-1].isalnum() and end < len(text) and is_word_character(text[end]):
            continue

        owner_keys = match_map.get(normalize_text(matched_name).casefold(), ())
        if owner_keys:
            matches.append((start, end, matched_name, owner_keys))
    return matches


def linkify_revised_glossary_terms(
    text: str,
    glossary: RevisedGlossary,
    match_map: GlossaryMatchMap,
    term_pattern: re.Pattern[str] | None,
) -> str:
    """Link unambiguous canonical terms and aliases to their canonical detail panel."""
    if not glossary or term_pattern is None:
        return html.escape(text)

    output_parts: list[str] = []
    last_end = 0
    for start, end, matched_name, owner_keys in glossary_name_matches(text, match_map, term_pattern):
        name_key = normalize_text(matched_name).casefold()
        canonical_owners = [
            owner_key
            for owner_key in owner_keys
            if glossary[owner_key].term.casefold() == name_key
        ]
        if len(canonical_owners) == 1:
            owner_key = canonical_owners[0]
        elif len(owner_keys) == 1:
            owner_key = owner_keys[0]
        else:
            continue

        output_parts.append(html.escape(text[last_end:start]))
        output_parts.append(
            f'<a href="#" data-term="{html.escape(glossary[owner_key].term, quote=True)}">'
            f"{html.escape(matched_name)}</a>"
        )
        last_end = end

    output_parts.append(html.escape(text[last_end:]))
    return "".join(output_parts)


class GlossaryMarkupLinker(HTMLParser):
    """Link glossary terms in narrative report text while preserving source markup."""

    VOID_TAGS = {"area", "base", "br", "col", "embed", "hr", "img", "input", "link", "meta", "source", "track", "wbr"}
    PROTECTED_TAGS = {"a", "button", "code", "pre", "script", "style"}
    PROTECTED_CLASSES = {"node-code", "source-reference"}
    LINKABLE_CLASSES = {"paragraph", "figure-explanation"}

    def __init__(
        self,
        markup: str,
        glossary: RevisedGlossary,
        match_map: GlossaryMatchMap,
        excluded_root_ids: set[str],
    ) -> None:
        super().__init__(convert_charrefs=False)
        self.markup = markup
        self.glossary = glossary
        self.match_map = match_map
        self.excluded_root_ids = excluded_root_ids
        self.term_pattern = build_glossary_match_pattern(match_map)
        self.line_offsets = [0]
        self.line_offsets.extend(match.end() for match in re.finditer("\n", markup))
        self.states: list[tuple[bool, bool]] = [(False, False)]
        self.replacements: list[tuple[int, int, str]] = []

    def absolute_offset(self) -> int:
        line_number, column = self.getpos()
        return self.line_offsets[line_number - 1] + column

    def push_state(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        attributes = dict(attrs)
        classes = class_names(attributes)
        parent_linkable, parent_protected = self.states[-1]
        linkable = parent_linkable or bool(classes & self.LINKABLE_CLASSES)
        protected = (
            parent_protected
            or tag in self.PROTECTED_TAGS
            or bool(classes & self.PROTECTED_CLASSES)
            or attributes.get("data-node-id") in self.excluded_root_ids
        )
        self.states.append((linkable, protected))

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        if tag not in self.VOID_TAGS:
            self.push_state(tag, attrs)

    def handle_startendtag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        return

    def handle_endtag(self, tag: str) -> None:
        if tag not in self.VOID_TAGS and len(self.states) > 1:
            self.states.pop()

    def handle_data(self, data: str) -> None:
        linkable, protected = self.states[-1]
        if not linkable or protected or not data.strip():
            return
        linked_text = linkify_revised_glossary_terms(
            data,
            self.glossary,
            self.match_map,
            self.term_pattern,
        )
        if 'data-term="' not in linked_text:
            return
        linked_text = linked_text.replace(
            '<a href="#" data-term=',
            '<a class="report-glossary-link" href="#" data-term=',
        )
        start = self.absolute_offset()
        self.replacements.append((start, start + len(data), linked_text))

    def result(self) -> str:
        linked_markup = self.markup
        for start, end, replacement in reversed(self.replacements):
            linked_markup = f"{linked_markup[:start]}{replacement}{linked_markup[end:]}"
        return linked_markup


def linkify_report_markup(
    markup: str,
    glossary: RevisedGlossary,
    match_map: GlossaryMatchMap,
    excluded_root_ids: set[str],
) -> str:
    """Make glossary terms clickable in eligible paragraphs and figure explanations."""
    parser = GlossaryMarkupLinker(markup, glossary, match_map, excluded_root_ids)
    parser.feed(markup)
    parser.close()
    return parser.result()


def parse_report_markup(markup: str) -> tuple[ElementSpan, list[ElementSpan]]:
    """Parse source locations needed to rearrange the report without restyling it."""
    parser = ReportMarkupParser(markup)
    parser.feed(markup)
    parser.close()
    return parser.result()


def parse_chapter_tabs(markup: str) -> list[ElementSpan]:
    """Parse the chapter tabs from the already-isolated navigation markup."""
    parser = ChapterTabParser(markup)
    parser.feed(markup)
    parser.close()
    return parser.result()


def report_key(title: str) -> str:
    """Map a JSON document title to the requested navigation order key."""
    if title == "Summary for Policymakers":
        return "SPM"
    if title == "Technical Summary":
        return "TS"
    match = re.match(r"^Chapter ([1-5]):", title)
    if match:
        return f"Chapter {match.group(1)}"
    raise ValueError(f"Unsupported report title in inspection JSON: {title!r}")


def collect_node_ids(node: dict[str, Any], node_ids: Counter[str]) -> int:
    """Collect node identities and count figure nodes in one report tree."""
    node_id = node.get("id")
    if not isinstance(node_id, str):
        raise ValueError("A report tree node is missing a string id.")
    node_ids[node_id] += 1

    figure_count = 1 if node.get("kind") == "figure" else 0
    children = node.get("children", [])
    if not isinstance(children, list):
        raise ValueError(f"Node {node_id} has non-list children.")
    for child in children:
        if not isinstance(child, dict):
            raise ValueError(f"Node {node_id} has a non-object child.")
        figure_count += collect_node_ids(child, node_ids)
    return figure_count


def canonical_report_data(payload: dict[str, Any]) -> tuple[list[str], Counter[str], int]:
    """Return ordered document IDs, all node IDs, and the JSON figure count."""
    reports = payload.get("reports")
    if not isinstance(reports, list):
        raise ValueError("Inspection JSON must contain a reports list.")

    root_ids_by_key: dict[str, str] = {}
    node_ids: Counter[str] = Counter()
    figure_count = 0
    for report in reports:
        if not isinstance(report, dict) or not isinstance(report.get("tree"), dict):
            raise ValueError("Each inspection report must contain a tree object.")
        tree = report["tree"]
        title = tree.get("title")
        if not isinstance(title, str):
            raise ValueError("A report tree is missing its title.")
        key = report_key(title)
        root_id = tree.get("id")
        if not isinstance(root_id, str):
            raise ValueError(f"Report {title!r} is missing its document id.")
        if key in root_ids_by_key:
            raise ValueError(f"Inspection JSON contains duplicate {key} reports.")
        root_ids_by_key[key] = root_id
        figure_count += collect_node_ids(tree, node_ids)

    if set(root_ids_by_key) != set(REPORT_ORDER):
        missing = sorted(set(REPORT_ORDER) - set(root_ids_by_key))
        unexpected = sorted(set(root_ids_by_key) - set(REPORT_ORDER))
        raise ValueError(f"Unexpected report set. Missing: {missing}; unexpected: {unexpected}")
    return [root_ids_by_key[key] for key in REPORT_ORDER], node_ids, figure_count


def glossary_matches_by_term(
    text: str,
    match_map: GlossaryMatchMap,
    term_pattern: re.Pattern[str] | None,
) -> dict[str, tuple[str, ...]]:
    """Group every boundary-safe canonical-term or alias match by its canonical term."""
    matches_by_term: dict[str, list[str]] = {}
    for _, _, matched_name, owner_keys in glossary_name_matches(text, match_map, term_pattern):
        for owner_key in owner_keys:
            matches_by_term.setdefault(owner_key, []).append(matched_name)
    return {owner_key: tuple(matched_names) for owner_key, matched_names in matches_by_term.items()}


def excludes_glossary_occurrences(node: dict[str, Any]) -> bool:
    """Return whether a titled report subtree is outside the usage review."""
    title = node.get("title")
    if not isinstance(title, str):
        return False
    normalized_title = " ".join(title.split()).casefold()
    return normalized_title == "references" or normalized_title.startswith("supplementary material")


def excluded_glossary_root_ids(payload: dict[str, Any]) -> set[str]:
    """Return roots of reference and supplementary-material subtrees."""
    excluded_ids: set[str] = set()

    def collect(node: dict[str, Any]) -> None:
        if excludes_glossary_occurrences(node):
            node_id = node.get("id")
            if not isinstance(node_id, str):
                raise ValueError("An excluded report subtree is missing its node id.")
            excluded_ids.add(node_id)
            return
        for child in node.get("children", []):
            if isinstance(child, dict):
                collect(child)

    for report in payload.get("reports", []):
        if isinstance(report, dict) and isinstance(report.get("tree"), dict):
            collect(report["tree"])
    return excluded_ids


def terminal_cae_parentheticals(text: str) -> list[tuple[str, int, int]]:
    """Return assessment text and source spans for sentence-final parentheses."""
    assessments: list[tuple[str, int, int]] = []
    for match in CAE_PARENTHESES_RE.finditer(text):
        assessment = " ".join(match.group(1).split())
        if not CAE_ASSESSMENT_TOKEN_RE.search(assessment):
            continue
        suffix = text[match.end() :]
        suffix_match = re.match(r"\s*(?:(?P<punctuation>[.!?])\s*)?(?:\{[^{}]*\}\s*)?", suffix)
        if suffix_match is None:
            continue
        remainder = suffix[suffix_match.end() :]
        if suffix_match.group("punctuation") is None and remainder.strip():
            continue
        assessments.append((assessment, match.start(), match.end() + suffix_match.end()))
    return assessments


def assessed_sentence(text: str, assessment_start: int, sentence_end: int) -> str:
    """Return the statement associated with a terminal assessment parenthesis."""
    boundaries = list(SENTENCE_BOUNDARY_RE.finditer(text[:assessment_start]))
    sentence_start = boundaries[-1].end() if boundaries else 0
    if not text[sentence_start:assessment_start].strip() and boundaries:
        sentence_start = boundaries[-2].end() if len(boundaries) > 1 else 0
    return " ".join(text[sentence_start:sentence_end].split())


def classify_cae_assessment(assessment: str) -> tuple[str, str, str, str]:
    """Classify a terminal parenthesis as confidence, a pair, or an issue."""
    normalized = " ".join(assessment.casefold().split())
    confidence_match = CAE_CONFIDENCE_RE.fullmatch(normalized)
    if confidence_match:
        return "confidence", confidence_match.group(1), "", ""

    pair_match = CAE_PAIR_RE.fullmatch(normalized)
    if pair_match:
        agreement = pair_match.group("agreement") or pair_match.group("agreement_second")
        evidence = pair_match.group("evidence") or pair_match.group("evidence_first")
        return "pair", agreement.casefold(), evidence.casefold(), ""

    has_confidence = re.search(r"\bconfidence\b", normalized) is not None
    has_agreement = re.search(r"\bagreement\b", normalized) is not None
    has_evidence = re.search(r"\bevidence\b", normalized) is not None
    if has_confidence and (has_agreement or has_evidence):
        issue = "Confidence is mixed with agreement or evidence"
    elif has_agreement != has_evidence:
        issue = "Agreement/evidence pair is incomplete"
    elif has_agreement and has_evidence:
        issue = "Invalid agreement/evidence level or format"
    else:
        issue = "Invalid confidence level or format"
    return "issue", "", "", issue


def full_report_cae_check(payload: dict[str, Any], node_codes: dict[str, str]) -> CaeCheckResult:
    """Count valid CAE statements and collect malformed terminal assessments."""
    reports = payload.get("reports")
    if not isinstance(reports, list):
        raise ValueError("Inspection JSON must contain a reports list.")

    result = CaeCheckResult(Counter(), Counter(), [], Counter(), Counter())

    def scan_node(node: dict[str, Any], source_name: str) -> None:
        if excludes_glossary_occurrences(node):
            return
        kind = node.get("kind")
        text = node.get("text") if kind == "paragraph" else node.get("explanation") if kind == "figure" else None
        if isinstance(text, str) and text.strip():
            node_id = node.get("id")
            for assessment, assessment_start, sentence_end in terminal_cae_parentheticals(text):
                if not isinstance(node_id, str) or node_id not in node_codes:
                    raise ValueError("A CAE evidence node is missing its rendered code.")
                source_span = node.get("source_span", {})
                page_number = source_span.get("from_page") if isinstance(source_span, dict) else None
                source_label = f"{source_name}, p. {page_number}" if page_number else source_name
                category, first_level, second_level, issue = classify_cae_assessment(assessment)
                if category == "confidence":
                    result.confidence[first_level] += 1
                    result.confidence_by_report[(source_name, first_level)] += 1
                elif category == "pair":
                    result.agreement_evidence[(first_level, second_level)] += 1
                    result.agreement_evidence_by_report[(source_name, first_level, second_level)] += 1
                else:
                    result.issues.append(
                        CaeOccurrence(
                            source_name,
                            source_label,
                            node_codes[node_id],
                            node_id,
                            assessed_sentence(text, assessment_start, sentence_end),
                            assessment,
                            issue,
                        )
                    )
        for child in node.get("children", []):
            if isinstance(child, dict):
                scan_node(child, source_name)

    for report in reports:
        if not isinstance(report, dict) or not isinstance(report.get("tree"), dict):
            raise ValueError("Each inspection report must contain a tree object.")
        tree = report["tree"]
        title = tree.get("title")
        if not isinstance(title, str):
            raise ValueError("A report tree is missing its title.")
        scan_node(tree, report_key(title))

    return result


def full_report_term_occurrences(
    payload: dict[str, Any],
    glossary: RevisedGlossary,
    match_map: GlossaryMatchMap,
    node_codes: dict[str, str],
) -> GlossaryOccurrences:
    """Find glossary terms in report content outside references and supplements."""
    reports = payload.get("reports")
    if not isinstance(reports, list):
        raise ValueError("Inspection JSON must contain a reports list.")

    occurrences: GlossaryOccurrences = {term_key: [] for term_key in glossary}
    term_pattern = build_glossary_match_pattern(match_map)

    def scan_node(node: dict[str, Any], source_name: str) -> None:
        if excludes_glossary_occurrences(node):
            return
        kind = node.get("kind")
        text = node.get("text") if kind == "paragraph" else node.get("explanation") if kind == "figure" else None
        if isinstance(text, str) and text.strip():
            node_id = node.get("id")
            if not isinstance(node_id, str) or node_id not in node_codes:
                raise ValueError("A glossary evidence node is missing its rendered code.")
            source_span = node.get("source_span", {})
            page_number = source_span.get("from_page") if isinstance(source_span, dict) else None
            source_label = f"{source_name}, p. {page_number}" if page_number else source_name
            for sentence in SENTENCE_BOUNDARY_RE.split(text.strip()):
                sentence = sentence.strip()
                if not sentence:
                    continue
                for term_key, matched_names in glossary_matches_by_term(
                    sentence,
                    match_map,
                    term_pattern,
                ).items():
                    occurrences[term_key].append(
                        GlossaryOccurrence(
                            source_label,
                            node_codes[node_id],
                            node_id,
                            sentence,
                            text,
                            matched_names,
                        )
                    )
        for child in node.get("children", []):
            if isinstance(child, dict):
                scan_node(child, source_name)

    for report in reports:
        if not isinstance(report, dict) or not isinstance(report.get("tree"), dict):
            raise ValueError("Each inspection report must contain a tree object.")
        tree = report["tree"]
        title = tree.get("title")
        if not isinstance(title, str):
            raise ValueError("A report tree is missing its title.")
        scan_node(tree, report_key(title))

    return occurrences


def strip_markdown_inline(text: str) -> str:
    """Remove lightweight markdown markers from one summary line."""
    cleaned = re.sub(r"[*_`]+", "", text or "")
    cleaned = re.sub(r"\s+", " ", cleaned).strip()
    return cleaned


def markdown_bullet_parts(line: str) -> tuple[int, str] | None:
    """Return indentation and content for one Markdown bullet line."""
    match = MARKDOWN_BULLET_RE.match(line)
    if match is None:
        return None
    indentation = len(match.group("indent").expandtabs(4))
    return indentation, match.group("content").strip()


def load_term_usage_summaries(summary_path: Path) -> dict[str, str]:
    """Load precomputed term summaries keyed by glossary term key."""
    if not summary_path.is_file():
        return {}
    payload = json.loads(summary_path.read_text(encoding="utf-8"))
    if not isinstance(payload, dict):
        return {}
    records = payload.get("summaries")
    if not isinstance(records, dict):
        return {}

    summaries: dict[str, str] = {}
    for term_key, record in records.items():
        if not isinstance(term_key, str) or not isinstance(record, dict):
            continue
        summary = record.get("summary")
        if isinstance(summary, str) and summary.strip():
            summaries[term_key.casefold()] = summary
    return summaries


def summary_issue_section(summary: str) -> str:
    """Extract the Potential issues section body from a markdown summary."""
    match = POTENTIAL_ISSUES_SECTION_RE.search(summary or "")
    return match.group("content") if match else ""


def summary_has_potential_issue(summary: str) -> bool:
    """Return whether a summary reports a potential consistency issue."""
    content = summary_issue_section(summary)
    if not content:
        return False

    bullet_lines = [
        strip_markdown_inline(bullet_content)
        for line in content.splitlines()
        if (bullet := markdown_bullet_parts(line)) is not None
        for _, bullet_content in [bullet]
    ]
    if not bullet_lines:
        text = strip_markdown_inline(content)
        return bool(text) and NO_POTENTIAL_ISSUE_RE.match(text) is None

    verdict = bullet_lines[0].casefold()
    if verdict.startswith("inconsistency identified"):
        return True
    if verdict.startswith("consistency identified"):
        return False

    if all(NO_POTENTIAL_ISSUE_RE.match(line) for line in bullet_lines if line):
        return False
    return True


def parse_issue_rows_for_term(term: str, aliases: tuple[str, ...], summary: str) -> list[GlossaryIssueRow]:
    """Parse issue-table rows from a term's Potential issues section."""
    content = summary_issue_section(summary)
    if not content:
        return []

    bullets = [
        bullet_content
        for line in content.splitlines()
        if (bullet := markdown_bullet_parts(line)) is not None
        for _, bullet_content in [bullet]
    ]
    if not bullets:
        return []

    first = strip_markdown_inline(bullets[0]).casefold()
    if first.startswith("consistency identified"):
        return []

    rows: list[GlossaryIssueRow] = []
    current_issue = "Potential issue"
    structured_re = re.compile(
        r"Issue\s+type:\s*(?P<kind>[^;]+);\s*"
        r"ID\(s\):\s*\[(?P<ids>[^\]]*)\];\s*"
        r"Quote:\s*\"(?P<quote>[^\"]*)\";\s*"
        r"Why\s+conflicting:\s*(?P<why>.*)",
        re.IGNORECASE,
    )
    evidence_re = re.compile(r"\[(?P<section>[^\]]+)\]\s*\"?(?P<quote>.*?)(?:\"\s*)?$")

    for raw_bullet in bullets[1:]:
        cleaned = strip_markdown_inline(raw_bullet)
        if not cleaned:
            continue
        if NO_POTENTIAL_ISSUE_RE.match(cleaned):
            continue

        structured_match = structured_re.match(cleaned)
        if structured_match:
            issue_text = cleaned
            quote = strip_markdown_inline(structured_match.group("quote"))
            identifiers = [item.strip() for item in structured_match.group("ids").split(",") if item.strip()]
            if not identifiers:
                identifiers = [""]
            for identifier in identifiers:
                rows.append(
                    GlossaryIssueRow(
                        section=normalize_section_identifier(identifier),
                        sentence=quote,
                        term=term,
                        issue=issue_text,
                        aliases=aliases,
                    )
                )
            current_issue = issue_text or current_issue
            continue

        evidence_match = evidence_re.match(cleaned)
        if evidence_match:
            rows.append(
                GlossaryIssueRow(
                    section=normalize_section_identifier(strip_markdown_inline(evidence_match.group("section"))),
                    sentence=strip_markdown_inline(evidence_match.group("quote")),
                    term=term,
                    issue=current_issue,
                    aliases=aliases,
                )
            )
            continue

        current_issue = cleaned

    if rows:
        return rows

    if first.startswith("inconsistency identified"):
        return [GlossaryIssueRow(section="", sentence="", term=term, issue="Inconsistency identified", aliases=aliases)]
    return []


def build_glossary_issue_rows(
    glossary: RevisedGlossary,
    term_usage_summaries: dict[str, str],
) -> tuple[set[str], list[GlossaryIssueRow]]:
    """Return term keys with issues and table rows derived from term summaries."""
    terms_with_issues: set[str] = set()
    rows: list[GlossaryIssueRow] = []

    for term_key, entry in glossary.items():
        summary = term_usage_summaries.get(term_key)
        if not summary or not summary_has_potential_issue(summary):
            continue
        terms_with_issues.add(term_key)
        rows.extend(parse_issue_rows_for_term(entry.term, entry.aliases, summary))

    rows.sort(key=lambda item: (item.section.casefold(), item.term.casefold(), item.issue.casefold(), item.sentence.casefold()))
    return terms_with_issues, rows


def build_section_node_lookup(node_codes: dict[str, str]) -> dict[str, str]:
    """Map visible section code to source node id for clickable issue-table links."""
    lookup: dict[str, str] = {}
    for node_id, node_code in node_codes.items():
        key = normalize_section_identifier(node_code)
        if key and key not in lookup:
            lookup[key] = node_id
    return lookup


def build_section_source_label_lookup(occurrences: GlossaryOccurrences) -> dict[str, str]:
    """Map section code to a representative source label."""
    labels: dict[str, str] = {}
    for matches in occurrences.values():
        for item in matches:
            key = normalize_section_identifier(item.node_code)
            if key and key not in labels:
                labels[key] = item.source_label
    return labels


def enrich_glossary_issue_rows(
    rows: list[GlossaryIssueRow],
    section_node_lookup: dict[str, str],
    section_source_lookup: dict[str, str],
) -> list[GlossaryIssueRow]:
    """Attach node ids and source labels to parsed issue rows for clickable rendering."""
    enriched: list[GlossaryIssueRow] = []
    for row in rows:
        section_key = normalize_section_identifier(row.section)
        node_id = section_node_lookup.get(section_key, "")
        source_label = section_source_lookup.get(section_key, "")
        enriched.append(
            GlossaryIssueRow(
                section=row.section,
                sentence=row.sentence,
                term=row.term,
                issue=row.issue,
                node_id=node_id,
                source_label=source_label,
                aliases=row.aliases,
            )
        )
    return enriched


def document_id_from_panel(markup: str) -> str:
    """Return the document node ID encoded inside one report-panel fragment."""
    match = DOCUMENT_NODE_ID_RE.search(markup)
    if match is None:
        raise ValueError("A report panel does not contain a document node id.")
    return match.group(1)


def replace_attribute(markup: str, attribute: str, value: str) -> str:
    """Replace one double-quoted HTML attribute while retaining all other markup."""
    pattern = re.compile(rf'(\s{re.escape(attribute)}=")[^"]*(")')
    updated, replacements = pattern.subn(rf'\g<1>{value}\g<2>', markup, count=1)
    if replacements != 1:
        raise ValueError(f"Could not update {attribute!r} on a chapter tab.")
    return updated


def set_panel_visibility(markup: str, hidden: bool) -> str:
    """Set the top-level panel's boolean hidden attribute."""
    tag_end = markup.find(">")
    if tag_end == -1:
        raise ValueError("A report panel has an unterminated opening tag.")
    opening_tag = re.sub(r"\s+hidden(?=\s|>)", "", markup[: tag_end + 1])
    if hidden:
        opening_tag = f"{opening_tag[:-1]} hidden>"
    return f"{opening_tag}{markup[tag_end + 1:]}"


def reorder_panels(markup: str, root_ids: list[str]) -> str:
    """Rearrange full report panels and make only the first one visible."""
    _, panels = parse_report_markup(markup)
    panel_markup_by_root_id: dict[str, str] = {}
    for panel in panels:
        panel_markup = markup[panel.start : panel.end]
        root_id = document_id_from_panel(panel_markup)
        if root_id in panel_markup_by_root_id:
            raise ValueError(f"Reference HTML has duplicate panel root id {root_id!r}.")
        panel_markup_by_root_id[root_id] = panel_markup

    if set(panel_markup_by_root_id) != set(root_ids):
        missing = sorted(set(root_ids) - set(panel_markup_by_root_id))
        unexpected = sorted(set(panel_markup_by_root_id) - set(root_ids))
        raise ValueError(f"Reference panels do not match the JSON. Missing: {missing}; unexpected: {unexpected}")

    separator = markup[panels[0].end : panels[1].start] if len(panels) > 1 else "\n"
    ordered_panels = [
        set_panel_visibility(panel_markup_by_root_id[root_id], hidden=index != 0)
        for index, root_id in enumerate(root_ids)
    ]
    return f"{markup[:panels[0].start]}{separator.join(ordered_panels)}{markup[panels[-1].end:]}"


def reorder_navigation(markup: str, panel_ids: list[str]) -> str:
    """Put the chapter tabs in the same order as their panels."""
    navigation, _ = parse_report_markup(markup)
    navigation_markup = markup[navigation.start : navigation.end]
    tabs = parse_chapter_tabs(navigation_markup)
    tab_markup_by_panel_id: dict[str, str] = {}
    for tab in tabs:
        panel_id = tab.attributes.get("aria-controls")
        if not isinstance(panel_id, str):
            raise ValueError("A chapter tab is missing aria-controls.")
        if panel_id in tab_markup_by_panel_id:
            raise ValueError(f"Reference HTML has duplicate tab target {panel_id!r}.")
        tab_markup_by_panel_id[panel_id] = navigation_markup[tab.start : tab.end]

    if set(tab_markup_by_panel_id) != set(panel_ids):
        missing = sorted(set(panel_ids) - set(tab_markup_by_panel_id))
        unexpected = sorted(set(tab_markup_by_panel_id) - set(panel_ids))
        raise ValueError(f"Reference tabs do not match panels. Missing: {missing}; unexpected: {unexpected}")

    separator = navigation_markup[tabs[0].end : tabs[1].start] if len(tabs) > 1 else ""
    ordered_tabs = []
    for index, panel_id in enumerate(panel_ids):
        tab_markup = tab_markup_by_panel_id[panel_id]
        tab_markup = replace_attribute(tab_markup, "aria-selected", "true" if index == 0 else "false")
        tab_markup = replace_attribute(tab_markup, "tabindex", "0" if index == 0 else "-1")
        ordered_tabs.append(tab_markup)
    reordered_navigation = (
        f"{navigation_markup[:tabs[0].start]}{separator.join(ordered_tabs)}{navigation_markup[tabs[-1].end:]}"
    )
    return f"{markup[:navigation.start]}{reordered_navigation}{markup[navigation.end:]}"


def normalize_figure_sources(markup: str) -> str:
    """Resolve figure assets from data/export to the repository's artifacts directory."""
    return re.sub(
        r'(\bsrc=["\'])\.\./artifacts/images/',
        r'\g<1>../../artifacts/images/',
        markup,
    )


def update_report_kickers(markup: str, report_count: int) -> str:
    """Describe each reconstructed panel as an HTML report."""
    if markup.count(REFERENCE_KICKER) != report_count:
        raise ValueError(f"Expected {report_count} reference report kickers.")
    return markup.replace(REFERENCE_KICKER, OUTPUT_KICKER)


def apply_previous_header_style(markup: str) -> str:
    """Move the tab navigation into the former terminology-review header treatment."""
    navigation, _ = parse_report_markup(markup)
    navigation_markup = markup[navigation.start : navigation.end]
    markup_without_navigation = f"{markup[:navigation.start]}{markup[navigation.end:]}"

    body_match = re.search(r"<body\b[^>]*>", markup_without_navigation, re.IGNORECASE)
    if body_match is None:
        raise ValueError("Could not find the body opening tag in the reference HTML.")
    style_end = markup_without_navigation.find("</style>")
    if style_end == -1:
        raise ValueError("Could not find the reference HTML style block.")

    indented_navigation = "\n".join(f"                {line}" for line in navigation_markup.splitlines())
    header_markup = (
        '\n        <header class="site-header">\n'
        '            <div class="site-header__inner">\n'
        f"                <h1>{REPORT_HEADER_TITLE}</h1>\n"
        f"{indented_navigation}\n"
        "            </div>\n"
        "        </header>\n"
    )
    markup_with_header = (
        f"{markup_without_navigation[:body_match.end()]}{header_markup}{markup_without_navigation[body_match.end():]}"
    )
    return f"{markup_with_header[:style_end]}{PREVIOUS_HEADER_CSS}{markup_with_header[style_end:]}"


def render_cae_source_cell(item: CaeOccurrence) -> str:
    """Render report location and a clickable canonical node code."""
    return (
        '<td><span class="glossary-evidence-location">'
        f"{html.escape(item.source_label)}</span>"
        '<button class="glossary-evidence-code" type="button" '
        f'data-source-node-id="{html.escape(item.node_id, quote=True)}">'
        f"{html.escape(item.node_code)}</button></td>"
    )


def cae_filter_data(result: CaeCheckResult) -> dict[str, Any]:
    """Return structured CAE counts for client-side report filtering."""
    return {
        "reports": list(REPORT_ORDER),
        "counts": {
            report_name: {
                "agreementEvidence": {
                    agreement: {
                        evidence: result.agreement_evidence_by_report[(report_name, agreement, evidence)]
                        for evidence in EVIDENCE_LEVELS
                    }
                    for agreement in AGREEMENT_LEVELS
                },
                "confidence": {
                    level: result.confidence_by_report[(report_name, level)] for level in CONFIDENCE_LEVELS
                },
            }
            for report_name in REPORT_ORDER
        },
    }


def render_cae_report_filter() -> str:
    """Render the all-or-subset report selector for CAE results."""
    options = "".join(
        '<label class="cae-report-option">'
        f'<input class="cae-report-checkbox" type="checkbox" value="{html.escape(report_name, quote=True)}" checked>'
        f"<span>{html.escape(report_name)}</span></label>"
        for report_name in REPORT_ORDER
    )
    return (
        '<details class="cae-report-filter">'
        '<summary><span class="cae-report-filter-label">Reports</span>'
        '<span class="cae-report-filter-value">All reports</span></summary>'
        '<fieldset class="cae-report-filter-menu">'
        '<legend>Select reports</legend>'
        '<label class="cae-report-option cae-report-all-option">'
        '<input class="cae-report-all" type="checkbox" checked>'
        '<span>All reports</span></label>'
        f"{options}</fieldset></details>"
    )


def render_cae_panel(result: CaeCheckResult) -> str:
    """Render CAE matrices and malformed sentence review rows."""
    matrix_header = "".join(
        f'<th scope="col">{html.escape(level.title())} evidence</th>' for level in EVIDENCE_LEVELS
    )
    matrix_rows = []
    for agreement in reversed(AGREEMENT_LEVELS):
        counts = "".join(
            f'<td data-agreement="{agreement}" data-evidence="{evidence}">'
            f"{result.agreement_evidence[(agreement, evidence)]}</td>"
            for evidence in EVIDENCE_LEVELS
        )
        matrix_rows.append(f'<tr><th scope="row">{agreement.title()} agreement</th>{counts}</tr>')

    confidence_header = "".join(
        f'<th scope="col">{html.escape(level.capitalize())} confidence</th>' for level in CONFIDENCE_LEVELS
    )
    confidence_counts = "".join(
        f'<td data-confidence="{html.escape(level, quote=True)}">{result.confidence[level]}</td>'
        for level in CONFIDENCE_LEVELS
    )

    issue_rows = []
    for item in result.issues:
        issue_rows.append(
            f'<tr data-report="{html.escape(item.report_name, quote=True)}">'
            f"{render_cae_source_cell(item)}"
            f"<td>{html.escape(item.sentence)}"
            f'<span class="cae-issue-label">{html.escape(item.issue)}: '
            f"({html.escape(item.assessment)})</span></td>"
            "</tr>"
        )
    if issue_rows:
        review_markup = (
            '<div class="cae-table-wrap cae-review-wrap"><table class="cae-table cae-review-table">'
            "<thead><tr><th>Section</th><th>Sentence</th></tr></thead><tbody>"
            f'{"".join(issue_rows)}</tbody></table></div>'
            '<p class="cae-empty cae-review-empty" hidden>No cases requiring review for the selected reports.</p>'
        )
    else:
        review_markup = '<p class="cae-empty">No incorrect or incomplete CAE cases found.</p>'

    return (
        f'<section class="report-panel" id="{CAE_PANEL_ID}" role="tabpanel" '
        f'aria-labelledby="{CAE_TAB_ID}" tabindex="0" data-metadata-visible="true" hidden '
        f'data-pair-count="{result.valid_pair_count}" data-confidence-count="{result.confidence_count}" '
        f'data-issue-count="{len(result.issues)}">'
        "<header>"
        '<h1>Confidence, Agreement, and Evidence check (WGII TSU)'
        '<button class="back-to-top" type="button" aria-label="Back to top" '
        'title="Back to top">&#8593;</button></h1>'
        f'<p class="facts" aria-live="polite">{result.valid_count} valid sentence-final assessments; '
        f'{len(result.issues)} cases require review</p>'
        '<p class="source-reference">Chapters 1-5, SPM, and TS; references and supplementary material excluded</p>'
        f"{render_cae_report_filter()}"
        "</header>"
        '<article class="cae-overview">'
        '<section class="cae-section" aria-labelledby="cae-pair-heading">'
        f'<h2 id="cae-pair-heading">Agreement and evidence ({result.valid_pair_count})</h2>'
        '<div class="cae-table-wrap"><table class="cae-table cae-matrix">'
        f'<thead><tr><th scope="col">Agreement \\ Evidence levels</th>{matrix_header}</tr></thead>'
        f'<tbody>{"".join(matrix_rows)}</tbody></table></div></section>'
        '<section class="cae-section" aria-labelledby="cae-confidence-heading">'
        f'<h2 id="cae-confidence-heading">Confidence ({result.confidence_count})</h2>'
        '<div class="cae-table-wrap"><table class="cae-table cae-count-table">'
        f'<thead><tr>{confidence_header}</tr></thead><tbody><tr>{confidence_counts}</tr></tbody>'
        "</table></div></section>"
        '<section class="cae-section" aria-labelledby="cae-review-heading">'
        f'<h2 id="cae-review-heading">Cases requiring review ({len(result.issues)})</h2>'
        f'{review_markup}</section>'
        '<script class="cae-filter-data" type="application/json">'
        f'{json.dumps(cae_filter_data(result), ensure_ascii=True, separators=(",", ":"))}'
        "</script></article></section>"
    )


def highlight_glossary_matches(text: str, matched_names: tuple[str, ...]) -> str:
    """Escape an evidence sentence and highlight its canonical-term or alias matches."""
    names = tuple(dict.fromkeys(name.casefold() for name in matched_names))
    if not names:
        return html.escape(text)
    pattern = re.compile(
        r"(?<!\w)(?:" + "|".join(re.escape(name) for name in sorted(names, key=len, reverse=True)) + r")(?!\w)",
        re.IGNORECASE,
    )
    output_parts: list[str] = []
    last_end = 0
    for match in pattern.finditer(text):
        output_parts.append(html.escape(text[last_end : match.start()]))
        output_parts.append(f"<mark>{html.escape(match.group(0))}</mark>")
        last_end = match.end()
    output_parts.append(html.escape(text[last_end:]))
    return "".join(output_parts)


def render_glossary_evidence(matches: list[GlossaryOccurrence]) -> str:
    """Render section and sentence evidence without LLM content."""
    if not matches:
        return '<p class="glossary-result-count">No usage found in the report body.</p>'

    use_count = sum(len(occurrence.matched_names) for occurrence in matches)
    rows = [
        '<details class="glossary-evidence">',
        f"<summary>Term use overview table ({len(matches)} {'sentence' if len(matches) == 1 else 'sentences'}; {use_count} {'use' if use_count == 1 else 'uses'})</summary>",
        "<table>",
        "<thead><tr><th>Section</th><th>Sentence</th></tr></thead>",
        "<tbody>",
    ]
    for occurrence in matches:
        rows.append(
            "<tr>"
            '<td><span class="glossary-evidence-location">'
            f"{html.escape(occurrence.source_label)}</span>"
            '<button class="glossary-evidence-code" type="button" '
            f'data-source-node-id="{html.escape(occurrence.node_id, quote=True)}">'
            f"{html.escape(occurrence.node_code)}</button></td>"
            f"<td>{highlight_glossary_matches(occurrence.sentence, occurrence.matched_names)}</td>"
            "</tr>"
        )
    rows.extend(["</tbody>", "</table>", "</details>"])
    return "".join(rows)


def render_summary_inline_markdown(
    text: str,
    available_section_codes: set[str] | None = None,
) -> str:
    """Render a minimal inline markdown subset used in generated summaries."""
    placeholders: list[str] = []

    def is_probable_section_code(value: str) -> bool:
        token = normalize_text(value)
        if not token or len(token) > 64:
            return False
        if not re.fullmatch(r"[A-Za-z0-9][A-Za-z0-9 .:/+-]*", token):
            return False
        return bool(
            re.search(r"\bP\d+\b", token)
            or re.search(r"\d", token)
            or token.startswith(("SPM", "TS", "ES", "Box", "Figure", "D-Figure", "C-Figure"))
        )

    def split_section_codes(label: str) -> list[str]:
        candidates = [normalize_text(part) for part in re.split(r"\s*[;,]\s*", label) if normalize_text(part)]
        if not candidates:
            return []
        if all(is_probable_section_code(item) for item in candidates):
            return candidates
        if is_probable_section_code(label):
            return [normalize_text(label)]
        return []

    def add_button_placeholder(section_code: str) -> str:
        if available_section_codes is not None and section_code not in available_section_codes:
            return section_code
        token = f"@@LINK{len(placeholders)}@@"
        escaped_label = html.escape(section_code)
        section_attr = html.escape(section_code, quote=True)
        placeholders.append(
            '<button class="glossary-evidence-code glossary-inline-evidence-code" type="button" '
            f'data-section-code="{section_attr}">{escaped_label}</button>'
        )
        return token

    def replace_link(match: re.Match[str]) -> str:
        label = normalize_text(match.group("label"))
        if not label:
            return match.group(0)
        return add_button_placeholder(label)

    text_with_tokens = re.sub(
        r"\[(?P<label>[^\]]+)\]\((?P<url>https?://[^)\s]+)\)",
        replace_link,
        text,
    )

    def replace_section_brackets(match: re.Match[str]) -> str:
        section_codes = split_section_codes(match.group("label"))
        if not section_codes:
            return match.group(0)
        return ", ".join(add_button_placeholder(code) for code in section_codes)

    text_with_tokens = re.sub(
        r"\[(?P<label>[^\]\n]+)\]",
        replace_section_brackets,
        text_with_tokens,
    )

    escaped = html.escape(text_with_tokens)
    escaped = re.sub(r"\*\*(.+?)\*\*", r"<strong>\1</strong>", escaped)
    escaped = re.sub(r"\*(.+?)\*", r"<em>\1</em>", escaped)
    for index, link_html in enumerate(placeholders):
        escaped = escaped.replace(f"@@LINK{index}@@", link_html)
    return escaped


def strip_usage_summary_section(summary: str) -> str:
    """Remove Usage summary blocks from legacy term summaries before rendering."""
    return re.sub(
        r"(?ims)^#{3,4}\s*usage\s+summary\s*$.*?(?=^#{3,4}\s+|\Z)",
        "",
        summary,
    ).strip()


def strip_glossary_alignment_section(summary: str) -> str:
    """Remove Glossary alignment blocks from legacy term summaries before rendering."""
    return re.sub(
        r"(?ims)^#{3,4}\s*glossary\s+alignment\s*$.*?(?=^#{3,4}\s+|\Z)",
        "",
        summary,
    ).strip()


def render_summary_markdown_lines(
    lines: list[str],
    available_section_codes: set[str] | None = None,
) -> str:
    """Render one non-heading summary section with flat Markdown bullets."""
    parts: list[str] = []
    in_list = False

    def close_list() -> None:
        nonlocal in_list
        if in_list:
            parts.append("</ul>")
            in_list = False

    for raw_line in lines:
        line = raw_line.strip()
        if not line:
            close_list()
            continue

        bullet = markdown_bullet_parts(raw_line)
        if bullet is not None:
            if not in_list:
                parts.append("<ul>")
                in_list = True
            parts.append(
                f"<li>{render_summary_inline_markdown(bullet[1], available_section_codes)}</li>"
            )
            continue

        close_list()
        parts.append(f"<p>{render_summary_inline_markdown(line, available_section_codes)}</p>")

    close_list()
    return "".join(parts)


def render_potential_issue_markdown_lines(
    lines: list[str],
    available_section_codes: set[str] | None = None,
) -> str:
    """Render issue evidence as nested list items beneath its parent issue."""
    parts: list[str] = []
    in_issue_list = False
    in_issue_item = False
    in_evidence_list = False

    def open_issue_list() -> None:
        nonlocal in_issue_list
        if not in_issue_list:
            parts.append("<ul>")
            in_issue_list = True

    def close_issue_item() -> None:
        nonlocal in_issue_item, in_evidence_list
        if in_evidence_list:
            parts.append("</ul>")
            in_evidence_list = False
        if in_issue_item:
            parts.append("</li>")
            in_issue_item = False

    def close_issue_list() -> None:
        nonlocal in_issue_list
        close_issue_item()
        if in_issue_list:
            parts.append("</ul>")
            in_issue_list = False

    for raw_line in lines:
        line = raw_line.strip()
        if not line:
            continue

        bullet = markdown_bullet_parts(raw_line)
        if bullet is None:
            close_issue_list()
            parts.append(f"<p>{render_summary_inline_markdown(line, available_section_codes)}</p>")
            continue

        indentation, bullet_content = bullet
        is_evidence = indentation > 0 or POTENTIAL_ISSUE_EVIDENCE_RE.match(bullet_content) is not None
        open_issue_list()
        if is_evidence and in_issue_item:
            if not in_evidence_list:
                parts.append("<ul>")
                in_evidence_list = True
            parts.append(
                f"<li>{render_summary_inline_markdown(bullet_content, available_section_codes)}</li>"
            )
            continue

        close_issue_item()
        if is_evidence:
            parts.append(
                f"<li>{render_summary_inline_markdown(bullet_content, available_section_codes)}</li>"
            )
            continue

        parts.append(f"<li>{render_summary_inline_markdown(bullet_content, available_section_codes)}")
        in_issue_item = True

    close_issue_list()
    return "".join(parts)


def render_summary_markdown_block(
    summary: str,
    available_section_codes: set[str] | None = None,
) -> str:
    """Render summary markdown (### headings, bullets, paragraphs) to safe HTML."""
    summary = strip_usage_summary_section(summary)
    summary = strip_glossary_alignment_section(summary)
    parts: list[str] = []
    heading: str | None = None
    section_lines: list[str] = []

    def flush_section() -> None:
        nonlocal section_lines
        rendered_heading = heading
        if rendered_heading is not None:
            heading_key = rendered_heading.casefold()
            if heading_key.startswith("contexts of use"):
                rendered_heading = LLM_CONTEXTS_HEADING
            elif heading_key.startswith("potential issues"):
                rendered_heading = LLM_POTENTIAL_ISSUES_HEADING
            elif heading_key == "conclusion":
                rendered_heading = LLM_CONCLUSION_HEADING
        if heading is not None:
            parts.append(f"<h4>{render_summary_inline_markdown(rendered_heading, available_section_codes)}</h4>")
        if heading is not None and rendered_heading == LLM_POTENTIAL_ISSUES_HEADING:
            parts.append(render_potential_issue_markdown_lines(section_lines, available_section_codes))
        else:
            parts.append(render_summary_markdown_lines(section_lines, available_section_codes))
        section_lines = []

    for raw_line in summary.splitlines():
        line = raw_line.strip()
        if line.startswith("### "):
            flush_section()
            heading = line[4:].strip()
            continue
        section_lines.append(raw_line)

    flush_section()
    return "".join(parts)


def render_glossary_llm_check(
    term_key: str,
    term_usage_summaries: dict[str, str],
    available_section_codes: set[str],
) -> str:
    """Render one foldable LLM consistency check block for a glossary term."""
    summary = term_usage_summaries.get(term_key, "").strip()
    if not summary:
        summary = "LLM analysis has not been run yet."
    callout_text = (
        "This LLM-assisted summary, including the potential inconsistency flags, "
        "is provided only as a reference. It may contain errors and it is not intended "
        "to replace the human judgment for consistency checks."
    )
    return (
        '<details class="glossary-llm-check">'
        "<summary>LLM-assisted consistency check</summary>"
        '<p class="glossary-llm-check-notebar" role="note" aria-label="LLM summary caution">'
        '<span class="glossary-llm-check-eye" aria-hidden="true">⚠️</span>'
        f'<span class="glossary-llm-check-note">{html.escape(callout_text)}</span>'
        '</p>'
        '<div class="glossary-llm-check-content">'
        f"{render_summary_markdown_block(summary, available_section_codes)}"
        "</div>"
        "</details>"
    )


def highlight_issue_sentence(sentence: str, term: str, aliases: tuple[str, ...]) -> str:
    """Highlight canonical term matches, or alias matches when canonical is absent."""
    if not sentence:
        return ""

    def build_candidates(*names: str) -> tuple[str, ...]:
        candidates: list[str] = []
        for name in names:
            normalized_name = normalize_text(name)
            if not normalized_name:
                continue
            candidates.append(normalized_name)
            for paren in re.findall(r"\(([^)]+)\)", normalized_name):
                cleaned = normalize_text(paren)
                if cleaned:
                    candidates.append(cleaned)
        return tuple(
            sorted(
                dict.fromkeys(candidate for candidate in candidates if candidate),
                key=len,
                reverse=True,
            )
        )

    def highlight_with_candidates(raw_sentence: str, candidates: tuple[str, ...]) -> tuple[str, bool]:
        if not candidates:
            return html.escape(raw_sentence), False
        pattern = re.compile(
            r"(?<!\w)(?:" + "|".join(re.escape(candidate) for candidate in candidates) + r")(?!\w)",
            re.IGNORECASE,
        )
        matches = list(pattern.finditer(raw_sentence))
        if not matches:
            return html.escape(raw_sentence), False

        output_parts: list[str] = []
        last_end = 0
        for match in matches:
            output_parts.append(html.escape(raw_sentence[last_end : match.start()]))
            output_parts.append(f"<mark>{html.escape(match.group(0))}</mark>")
            last_end = match.end()
        output_parts.append(html.escape(raw_sentence[last_end:]))
        return "".join(output_parts), True

    canonical_candidates = build_candidates(term)
    highlighted, has_canonical_match = highlight_with_candidates(sentence, canonical_candidates)
    if has_canonical_match:
        return highlighted

    alias_candidates = tuple(
        candidate
        for candidate in build_candidates(*aliases)
        if candidate.casefold() not in {name.casefold() for name in canonical_candidates}
    )
    highlighted_alias, has_alias_match = highlight_with_candidates(sentence, alias_candidates)
    return highlighted_alias if has_alias_match else highlighted


def render_glossary_panel(
    glossary: RevisedGlossary,
    occurrences: GlossaryOccurrences,
    terms_with_issues: set[str],
    term_usage_summaries: dict[str, str],
) -> tuple[str, int]:
    """Render the AO glossary overview as a static tab panel."""
    usage_counts = {
        term_key: sum(len(occurrence.matched_names) for occurrence in matches)
        for term_key, matches in occurrences.items()
    }
    available_section_codes: set[str] = set()
    for matches in occurrences.values():
        for occurrence in matches:
            section_code = normalize_text(occurrence.node_code)
            if section_code:
                available_section_codes.add(section_code)
    used_term_count = sum(count > 0 for count in usage_counts.values())
    ordered_entries = sorted(glossary.values(), key=lambda entry: entry.term.casefold())

    term_rows: list[str] = []
    detail_panels: list[str] = []
    for index, entry in enumerate(ordered_entries, start=1):
        term = entry.term
        term_key = term.casefold()
        frequency = usage_counts[term_key]
        search_terms = " ".join((term, *entry.aliases)).casefold()
        detail_id = f"glossary-term-detail-{index}"
        disabled = " disabled" if frequency == 0 else ""
        detail_attribute = f' data-detail-id="{detail_id}"' if frequency else ""
        issue_indicator = (
            '<span class="glossary-term-issue" role="img" '
            'aria-label="Potential issue needing substantive review" '
            'title="Potential issue needing substantive review">!</span>'
            if term_key in terms_with_issues
            else ""
        )
        term_rows.append(
            f'<li class="glossary-term-row" data-search="{html.escape(search_terms, quote=True)}" '
            f'data-usage-count="{frequency}">'
            f'<button class="glossary-term-button" type="button" aria-pressed="false"{detail_attribute}{disabled}>'
            f"&#8226; {html.escape(term)} "
            f'<span class="glossary-term-count">[{frequency}]</span> '
            f'<span class="glossary-term-source">[{html.escape(entry.source)}]</span>'
            f"{issue_indicator}"
            "</button>"
            "</li>"
        )
        if not frequency:
            continue

        matches = occurrences[term_key]
        alias_usage_counts = Counter(
            matched_name.casefold()
            for occurrence in matches
            for matched_name in occurrence.matched_names
        )
        canonical_name = term
        canonical_count = alias_usage_counts[canonical_name.casefold()]
        alias_names = entry.aliases
        canonical_markup = (
            '<p class="glossary-canonical-term">Canonical term: '
            f'{html.escape(canonical_name)} '
            f'<span class="glossary-term-count">[{canonical_count}]</span>'
            "</p>"
        )
        alias_items = ", ".join(
            f'{html.escape(alias)} <span class="glossary-term-count">'
            f'[{alias_usage_counts[alias.casefold()]}]</span>'
            for alias in alias_names
        )
        aliases_markup = (
            f'<p class="glossary-aliases">Alias(es): {alias_items}</p>'
            if alias_items
            else '<p class="glossary-aliases">Alias(es): None</p>'
        )
        parent_markup = (
            f'<p class="glossary-parent">Parent terms: {html.escape("; ".join(entry.parent_terms))}</p>'
            if entry.parent_terms
            else ""
        )
        child_markup = (
            f'<p class="glossary-child">Child terms: {html.escape("; ".join(entry.child_terms))}</p>'
            if entry.child_terms
            else ""
        )
        definition_markup = (
            '<section class="glossary-definition">'
            "<h4>Explanation</h4>"
            f"<p>{html.escape(entry.explanation or 'Definition not available.')}</p>"
            "</section>"
        )
        detail_panels.append(
            f'<article class="glossary-detail" id="{detail_id}" '
            f'data-term="{html.escape(term, quote=True)}" hidden>'
            '<div class="glossary-detail-heading">'
            f"<h3>{html.escape(term)}</h3>"
            f'<span class="glossary-term-count">[{frequency}]</span>'
            f'<span class="glossary-term-source">[{html.escape(entry.source)}]</span>'
            "</div>"
            f"{canonical_markup}"
            f"{aliases_markup}"
            f"{parent_markup}"
            f"{child_markup}"
            f"{definition_markup}"
            f"{render_glossary_llm_check(term_key, term_usage_summaries, available_section_codes)}"
            f"{render_glossary_evidence(matches)}"
            "</article>"
        )

    panel = (
        f'<section class="report-panel" id="{GLOSSARY_PANEL_ID}" role="tabpanel" '
        f'aria-labelledby="{GLOSSARY_TAB_ID}" tabindex="0" data-metadata-visible="true" hidden>'
        "<header>"
        '<p class="kicker">Reference glossaries</p>'
        '<h1>Glossary Overview<button class="back-to-top" type="button" aria-label="Back to top" '
        'title="Back to top">&#8593;</button></h1>'
        f'<p class="facts">{used_term_count} of {len(glossary)} terms occur across Chapters 1-5, SPM, and TS</p>'
        '<p class="source-reference">AR6 and AR7 SOD glossary (AO)</p>'
        "</header>"
        '<article class="glossary-overview">'
        '<div class="glossary-workspace">'
        '<section class="glossary-index-pane" aria-labelledby="glossary-index-heading">'
        f'<h2 id="glossary-index-heading">Glossary Overview ({used_term_count}/{len(glossary)})</h2>'
        '<button class="glossary-unused-toggle" type="button" aria-pressed="false">'
        'Hide terms used 0 times</button>'
        '<label class="glossary-search-label" for="glossary-search">Search for a term'
        '<input class="glossary-search" id="glossary-search" type="search" autocomplete="off" '
        'placeholder="Type a term name...">'
        "</label>"
        '<p class="glossary-result-count" aria-live="polite"></p>'
        f'<ul class="glossary-term-list">{"".join(term_rows)}</ul>'
        "</section>"
        '<div class="glossary-divider" role="separator" aria-label="Resize glossary panels" '
        'aria-orientation="vertical" aria-valuemin="20" aria-valuemax="60" aria-valuenow="30" '
        'aria-valuetext="30% glossary overview width" tabindex="0"></div>'
        '<section class="glossary-detail-pane" aria-labelledby="glossary-detail-heading">'
        '<h2 id="glossary-detail-heading">Terms</h2>'
        '<p class="glossary-detail-placeholder">Click a term in the left panel to show its definition and related texts.</p>'
        f'{"".join(detail_panels)}'
        "</section>"
        "</div>"
        "</article>"
        "</section>"
    )
    return panel, used_term_count


def render_glossary_issue_table_panel(rows: list[GlossaryIssueRow], issue_term_count: int) -> str:
    """Render a dedicated tab panel listing glossary terms with potential issues."""
    if rows:
        table_rows = []
        for row in rows:
            if row.node_id:
                section_cell = (
                    '<span class="glossary-evidence-location">'
                    f"{html.escape(row.source_label)}</span>"
                    ' '
                    '<button class="glossary-evidence-code" type="button" '
                    f'data-source-node-id="{html.escape(row.node_id, quote=True)}">'
                    f"{html.escape(row.section)}</button>"
                )
            else:
                section_cell = html.escape(row.section)
            table_rows.append(
                "<tr>"
                f"<td>{section_cell}</td>"
                f"<td>{html.escape(row.term)}</td>"
                f"<td>{highlight_issue_sentence(row.sentence, row.term, row.aliases)}</td>"
                f"<td>{html.escape(row.issue)}</td>"
                "</tr>"
            )
        content = (
            '<div class="glossary-issue-wrap"><table class="glossary-issue-table">'
            '<colgroup><col style="width: 15%"><col style="width: 15%"><col style="width: 30%"><col style="width: 40%"></colgroup>'
            "<thead><tr><th>Section</th><th>Term with a potential issue</th><th>Sentence with a potential issue</th><th>Potential issues</th></tr></thead>"
            f'<tbody>{"".join(table_rows)}</tbody></table></div>'
        )
    else:
        content = '<p class="glossary-issue-empty">No potential consistency issues identified from available term summaries.</p>'

    return (
        f'<section class="report-panel" id="{GLOSSARY_ISSUE_PANEL_ID}" role="tabpanel" '
        f'aria-labelledby="{GLOSSARY_ISSUE_TAB_ID}" tabindex="0" data-metadata-visible="true" hidden>'
        "<header>"
        '<p class="kicker">Term consistency review</p>'
        '<h1>Glossary Issue Table<button class="back-to-top" type="button" aria-label="Back to top" '
        'title="Back to top">&#8593;</button></h1>'
        f'<p class="facts">{len(rows)} issue sentences across {issue_term_count} terms with potential issues</p>'
        '<p class="source-reference">Derived from llm_term_check.json potential-issue sections (section, sentence, term, issue)</p>'
        '<div class="glossary-issue-search">'
        '  <div class="glossary-issue-controls">'
        '    <label for="glossary-issue-chapter">Chapter<select class="glossary-issue-chapter" id="glossary-issue-chapter"><option value="All">All chapters</option><option value="Chapter 1">Chapter 1</option><option value="Chapter 2">Chapter 2</option><option value="Chapter 3">Chapter 3</option><option value="Chapter 4">Chapter 4</option><option value="Chapter 5">Chapter 5</option><option value="SPM">SPM</option><option value="TS">TS</option><option value="Other">Other</option></select></label>'
        '    <label for="glossary-issue-search">Search for terms in sections (Supports AND, OR, and parentheses ())<input autocomplete="off" class="glossary-issue-search-input" id="glossary-issue-search" placeholder="Example: (1.2 OR 1.3.2) AND confidence" type="search"></label><p aria-live="polite" class="glossary-issue-search-status"></p><div class="glossary-issue-export"><button class="glossary-issue-download-html" type="button">Download HTML</button><button class="glossary-issue-download-pdf" type="button">Download PDF</button></div>'
        '  </div>'
        '</div>'
        "</header>"
        f'<article class="glossary-issue-overview">{content}</article>'
        "</section>"
    )


def add_cae_check(markup: str, panel_markup: str) -> str:
    """Add the CAE navigation tab, panel, and styles."""
    metadata_button = markup.find('<button class="metadata-toggle"')
    if metadata_button == -1:
        raise ValueError("Could not find the metadata button in the report navigation.")
    cae_tab = (
        f'<button class="chapter-tab" type="button" id="{CAE_TAB_ID}" role="tab" '
        f'aria-selected="false" aria-controls="{CAE_PANEL_ID}" tabindex="-1" '
        'title="CAE check">CAE check</button>'
    )
    markup = f"{markup[:metadata_button]}{cae_tab}{markup[metadata_button:]}"

    main_end = markup.rfind("</main>")
    if main_end == -1:
        raise ValueError("Could not find the report main closing tag.")
    markup = f"{markup[:main_end]}\n{panel_markup}\n{markup[main_end:]}"

    style_end = markup.find("</style>")
    if style_end == -1:
        raise ValueError("Could not find the report style block.")
    markup = f"{markup[:style_end]}{CAE_CSS}{markup[style_end:]}"

    body_end = markup.rfind("</body>")
    if body_end == -1:
        raise ValueError("Could not find the report body closing tag.")
    return f"{markup[:body_end]}{CAE_JAVASCRIPT}{markup[body_end:]}"


def add_glossary_overview(markup: str, panel_markup: str) -> str:
    """Add the glossary tab, panel, CSS, and client-side interactions."""
    metadata_button = markup.find('<button class="metadata-toggle"')
    if metadata_button == -1:
        raise ValueError("Could not find the metadata button in the report navigation.")
    glossary_tab = (
        f'<button class="chapter-tab" type="button" id="{GLOSSARY_TAB_ID}" role="tab" '
        f'aria-selected="false" aria-controls="{GLOSSARY_PANEL_ID}" tabindex="-1" '
        'title="Glossary Overview">Glossary Overview</button>'
    )
    markup = f"{markup[:metadata_button]}{glossary_tab}{markup[metadata_button:]}"

    main_end = markup.rfind("</main>")
    if main_end == -1:
        raise ValueError("Could not find the report main closing tag.")
    markup = f"{markup[:main_end]}\n{panel_markup}\n{markup[main_end:]}"

    style_end = markup.find("</style>")
    if style_end == -1:
        raise ValueError("Could not find the report style block.")
    markup = f"{markup[:style_end]}{GLOSSARY_CSS}{markup[style_end:]}"

    body_end = markup.rfind("</body>")
    if body_end == -1:
        raise ValueError("Could not find the report body closing tag.")
    return f"{markup[:body_end]}{GLOSSARY_DIALOG_MARKUP}{GLOSSARY_JAVASCRIPT}{markup[body_end:]}"


def add_glossary_issue_table(markup: str, panel_markup: str) -> str:
    """Add the Glossary Issue Table tab, panel, and styles."""
    metadata_button = markup.find('<button class="metadata-toggle"')
    if metadata_button == -1:
        raise ValueError("Could not find the metadata button in the report navigation.")
    issue_tab = (
        f'<button class="chapter-tab" type="button" id="{GLOSSARY_ISSUE_TAB_ID}" role="tab" '
        f'aria-selected="false" aria-controls="{GLOSSARY_ISSUE_PANEL_ID}" tabindex="-1" '
        'title="Glossary Issue Table">Glossary Issue Table</button>'
    )
    markup = f"{markup[:metadata_button]}{issue_tab}{markup[metadata_button:]}"

    main_end = markup.rfind("</main>")
    if main_end == -1:
        raise ValueError("Could not find the report main closing tag.")
    markup = f"{markup[:main_end]}\n{panel_markup}\n{markup[main_end:]}"

    style_end = markup.find("</style>")
    if style_end == -1:
        raise ValueError("Could not find the report style block.")
    markup = f"{markup[:style_end]}{GLOSSARY_ISSUE_TABLE_CSS}{markup[style_end:]}"

    body_end = markup.rfind("</body>")
    if body_end == -1:
        raise ValueError("Could not find the report body closing tag.")
    return f"{markup[:body_end]}{GLOSSARY_ISSUE_TABLE_JAVASCRIPT}{markup[body_end:]}"


def validate_cae_output(markup: str, result: CaeCheckResult) -> None:
    """Verify CAE tab order, aggregate counts, and malformed-case rows."""
    navigation, panels = parse_report_markup(markup)
    if len(panels) != len(REPORT_ORDER) + 3:
        raise ValueError("Generated output must contain seven reports, CAE check, Glossary Overview, and Glossary Issue Table.")
    cae_panel = panels[-3]
    if cae_panel.attributes.get("id") != CAE_PANEL_ID or "hidden" not in cae_panel.attributes:
        raise ValueError("CAE check must be the penultimate, initially hidden panel.")
    expected_counts = {
        "data-pair-count": str(result.valid_pair_count),
        "data-confidence-count": str(result.confidence_count),
        "data-issue-count": str(len(result.issues)),
    }
    if any(cae_panel.attributes.get(attribute) != value for attribute, value in expected_counts.items()):
        raise ValueError("CAE panel aggregate counts differ from the corpus scan.")

    tabs = parse_chapter_tabs(markup[navigation.start : navigation.end])
    if len(tabs) != len(REPORT_ORDER) + 3 or tabs[-3].attributes.get("aria-controls") != CAE_PANEL_ID:
        raise ValueError("CAE check must be the third-from-last chapter-navigation tab.")

    panel_markup = markup[cae_panel.start : cae_panel.end]
    if panel_markup.count('class="glossary-evidence-code"') != len(result.issues):
        raise ValueError("CAE review row count differs from malformed corpus cases.")
    if panel_markup.count('class="cae-report-checkbox"') != len(REPORT_ORDER):
        raise ValueError("CAE report filter must contain all seven report options.")
    if panel_markup.count('class="cae-report-all"') != 1:
        raise ValueError("CAE report filter must contain one All reports option.")
    if panel_markup.count(' data-agreement="') != len(AGREEMENT_LEVELS) * len(EVIDENCE_LEVELS):
        raise ValueError("Every CAE matrix cell must expose its agreement and evidence levels.")
    if panel_markup.count(' data-confidence="') != len(CONFIDENCE_LEVELS):
        raise ValueError("Every CAE confidence cell must expose its confidence level.")

    issue_reports = re.findall(r'<tr data-report="([^"]+)">', panel_markup)
    if issue_reports != [item.report_name for item in result.issues]:
        raise ValueError("Every CAE review row must identify its source report.")

    filter_data_match = re.search(
        r'<script class="cae-filter-data" type="application/json">(.*?)</script>',
        panel_markup,
        re.DOTALL,
    )
    if filter_data_match is None or json.loads(filter_data_match.group(1)) != cae_filter_data(result):
        raise ValueError("CAE client-side filter data differs from the corpus scan.")
    if sum(result.agreement_evidence.values()) != result.valid_pair_count:
        raise ValueError("CAE agreement/evidence matrix does not match its total.")
    if sum(result.confidence.values()) != result.confidence_count:
        raise ValueError("CAE confidence table does not match its total.")
    for agreement in AGREEMENT_LEVELS:
        for evidence in EVIDENCE_LEVELS:
            report_total = sum(
                result.agreement_evidence_by_report[(report_name, agreement, evidence)]
                for report_name in REPORT_ORDER
            )
            if report_total != result.agreement_evidence[(agreement, evidence)]:
                raise ValueError("Per-report CAE matrix counts do not match their aggregate.")
    for level in CONFIDENCE_LEVELS:
        report_total = sum(
            result.confidence_by_report[(report_name, level)] for report_name in REPORT_ORDER
        )
        if report_total != result.confidence[level]:
            raise ValueError("Per-report CAE confidence counts do not match their aggregate.")


def validate_glossary_output(markup: str, glossary: RevisedGlossary, used_term_count: int) -> None:
    """Verify the glossary tab mirrors the revised workbook without LLM content."""
    _, panels = parse_report_markup(markup)
    if len(panels) != len(REPORT_ORDER) + 3:
        raise ValueError("Generated output must contain seven reports, CAE check, Glossary Overview, and Glossary Issue Table.")
    if panels[-2].attributes.get("id") != GLOSSARY_PANEL_ID or "hidden" not in panels[-2].attributes:
        raise ValueError("Glossary Overview must be the penultimate, initially hidden panel.")

    navigation, _ = parse_report_markup(markup)
    tabs = parse_chapter_tabs(markup[navigation.start : navigation.end])
    if len(tabs) != len(REPORT_ORDER) + 3 or tabs[-2].attributes.get("aria-controls") != GLOSSARY_PANEL_ID:
        raise ValueError("Glossary Overview must be the penultimate chapter-navigation tab.")

    glossary_panel = markup[panels[-2].start : panels[-2].end]
    if glossary_panel.count('class="glossary-term-row"') != len(glossary):
        raise ValueError("Generated glossary term count differs from the revised workbook.")
    if glossary_panel.count('class="glossary-detail"') != used_term_count:
        raise ValueError("Generated selectable glossary count differs from app usage counts.")

    contexts_heading = f"<h4>{LLM_CONTEXTS_HEADING}</h4>"
    issues_heading = f"<h4>{LLM_POTENTIAL_ISSUES_HEADING}</h4>"
    conclusion_heading = f"<h4>{LLM_CONCLUSION_HEADING}</h4>"
    if glossary_panel.count(contexts_heading) != used_term_count:
        raise ValueError("Each rendered term must include the default Contexts of use heading.")
    if glossary_panel.count(issues_heading) != used_term_count:
        raise ValueError("Each rendered term must include the default Potential issues heading.")
    if glossary_panel.count(conclusion_heading) != used_term_count:
        raise ValueError("Each rendered term must include the default Conclusion heading.")

    potential_issue_sections = re.findall(
        r"<h4>" + re.escape(LLM_POTENTIAL_ISSUES_HEADING) + r"</h4>(?P<body>.*?)(?=<h4>" + re.escape(LLM_CONCLUSION_HEADING) + r"</h4>)",
        glossary_panel,
        re.DOTALL,
    )
    if len(potential_issue_sections) != used_term_count:
        raise ValueError("Could not isolate all Potential issues sections in rendered summaries.")


def validate_glossary_issue_output(markup: str) -> None:
    """Verify the Glossary Issue Table panel exists as the final hidden tab panel."""
    navigation, panels = parse_report_markup(markup)
    if len(panels) != len(REPORT_ORDER) + 3:
        raise ValueError("Generated output must contain seven reports, CAE check, Glossary Overview, and Glossary Issue Table.")
    if panels[-1].attributes.get("id") != GLOSSARY_ISSUE_PANEL_ID or "hidden" not in panels[-1].attributes:
        raise ValueError("Glossary Issue Table must be the final, initially hidden panel.")

    tabs = parse_chapter_tabs(markup[navigation.start : navigation.end])
    if len(tabs) != len(REPORT_ORDER) + 3 or tabs[-1].attributes.get("aria-controls") != GLOSSARY_ISSUE_PANEL_ID:
        raise ValueError("Glossary Issue Table must be the final chapter-navigation tab.")


def validate_output(
    markup: str,
    output_path: Path,
    root_ids: list[str],
    node_ids: Counter[str],
    figure_count: int,
) -> None:
    """Verify canonical content, requested order, visibility, and local figures."""
    navigation, panels = parse_report_markup(markup)
    panel_ids = [panel.attributes.get("id") for panel in panels]
    if not all(isinstance(panel_id, str) for panel_id in panel_ids):
        raise ValueError("A generated report panel is missing its id.")
    panel_ids = [panel_id for panel_id in panel_ids if isinstance(panel_id, str)]
    panel_root_ids = [document_id_from_panel(markup[panel.start : panel.end]) for panel in panels]
    if panel_root_ids != root_ids:
        raise ValueError("Generated report panels are not in the requested JSON order.")

    navigation_markup = markup[navigation.start : navigation.end]
    tabs = parse_chapter_tabs(navigation_markup)
    tab_panel_ids = [tab.attributes.get("aria-controls") for tab in tabs]
    if tab_panel_ids != panel_ids:
        raise ValueError("Generated chapter tabs do not match the ordered panels.")
    selected_tabs = [tab.attributes.get("aria-selected") for tab in tabs]
    if selected_tabs != ["true", *["false"] * (len(tabs) - 1)]:
        raise ValueError("Only the first chapter tab must be selected initially.")
    visible_panel_ids = [panel.attributes.get("id") for panel in panels if "hidden" not in panel.attributes]
    if visible_panel_ids != [panel_ids[0]]:
        raise ValueError("Only the first report panel must be visible initially.")

    rendered_node_ids = Counter(NODE_ID_RE.findall(markup))
    if rendered_node_ids != node_ids:
        missing = list((node_ids - rendered_node_ids).elements())[:5]
        unexpected = list((rendered_node_ids - node_ids).elements())[:5]
        raise ValueError(f"Generated node ids differ from inspection JSON. Missing: {missing}; unexpected: {unexpected}")

    figure_sources = IMAGE_SOURCE_RE.findall(markup)
    if len(figure_sources) != figure_count:
        raise ValueError(f"Expected {figure_count} figure sources, found {len(figure_sources)}.")
    invalid_sources = [source for source in figure_sources if not source.startswith("../../artifacts/images/")]
    if invalid_sources:
        raise ValueError(f"Figure source does not target artifacts/images: {invalid_sources[0]!r}")
    missing_figures = [source for source in figure_sources if not (output_path.parent / source).is_file()]
    if missing_figures:
        raise ValueError(f"Figure asset does not exist: {missing_figures[0]!r}")


def parse_args() -> argparse.Namespace:
    """Parse command-line paths for the report reconstruction."""
    parser = argparse.ArgumentParser(description="Build the reordered SRCities reconstructed report.")
    parser.add_argument("--source-json", type=Path, default=DEFAULT_SOURCE_JSON)
    parser.add_argument("--reference-html", type=Path, default=DEFAULT_REFERENCE_HTML)
    parser.add_argument("--glossary", type=Path, default=DEFAULT_GLOSSARY_PATH)
    parser.add_argument("--term-summaries", type=Path, default=DEFAULT_TERM_SUMMARIES_PATH)
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT_HTML)
    return parser.parse_args()


def is_current_consistencycheck_markup(markup: str) -> bool:
    """Return whether markup is already the full consistencycheck export layout."""
    required_tokens = (
        'id="cae-check-panel"',
        'id="glossary-overview-panel"',
        'id="glossary-issue-table-panel"',
        'class="glossary-issue-search"',
        '// glossary-issue-search-script',
    )
    return all(token in markup for token in required_tokens)


def main() -> None:
    """Generate and validate the report viewer."""
    args = parse_args()
    source_json = args.source_json.expanduser()
    reference_html = args.reference_html.expanduser()
    glossary_path = args.glossary.expanduser()
    term_summaries_path = args.term_summaries.expanduser()
    output_path = args.output.expanduser()

    payload = json.loads(source_json.read_text(encoding="utf-8"))
    if not isinstance(payload, dict):
        raise ValueError("Inspection JSON must have an object at its root.")
    root_ids, node_ids, figure_count = canonical_report_data(payload)

    markup = reference_html.read_text(encoding="utf-8")

    if is_current_consistencycheck_markup(markup):
        output_path.parent.mkdir(parents=True, exist_ok=True)
        output_path.write_text(markup, encoding="utf-8")
        print(f"Wrote {output_path} by preserving current consistencycheck markup from {reference_html}.")
        return

    markup = reorder_panels(markup, root_ids)
    _, ordered_panels = parse_report_markup(markup)
    panel_ids = [panel.attributes.get("id") for panel in ordered_panels]
    if not all(isinstance(panel_id, str) for panel_id in panel_ids):
        raise ValueError("Reference HTML has a report panel without an id.")
    markup = reorder_navigation(markup, [panel_id for panel_id in panel_ids if isinstance(panel_id, str)])
    markup = normalize_figure_sources(markup)
    markup = update_report_kickers(markup, len(root_ids))
    markup = apply_previous_header_style(markup)
    validate_output(markup, output_path, root_ids, node_ids, figure_count)

    if not glossary_path.is_file():
        raise FileNotFoundError(f"Revised glossary workbook is unavailable: {glossary_path}")
    glossary = load_revised_glossary(glossary_path)
    match_map = build_glossary_match_map(glossary)
    node_codes = report_node_codes(markup)
    occurrences = full_report_term_occurrences(payload, glossary, match_map, node_codes)
    cae_result = full_report_cae_check(payload, node_codes)
    term_usage_summaries = load_term_usage_summaries(term_summaries_path)
    terms_with_issues, glossary_issue_rows = build_glossary_issue_rows(glossary, term_usage_summaries)
    section_node_lookup = build_section_node_lookup(node_codes)
    section_source_lookup = build_section_source_label_lookup(occurrences)
    glossary_issue_rows = enrich_glossary_issue_rows(glossary_issue_rows, section_node_lookup, section_source_lookup)
    markup = linkify_report_markup(markup, glossary, match_map, excluded_glossary_root_ids(payload))
    markup = add_cae_check(markup, render_cae_panel(cae_result))
    glossary_panel, used_term_count = render_glossary_panel(
        glossary,
        occurrences,
        terms_with_issues,
        term_usage_summaries,
    )
    markup = add_glossary_overview(markup, glossary_panel)
    issue_panel = render_glossary_issue_table_panel(glossary_issue_rows, len(terms_with_issues))
    markup = add_glossary_issue_table(markup, issue_panel)
    validate_cae_output(markup, cae_result)
    validate_glossary_output(markup, glossary, used_term_count)
    validate_glossary_issue_output(markup)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(markup, encoding="utf-8")
    print(
        f"Wrote {output_path} with {len(root_ids)} reports, {sum(node_ids.values())} nodes, "
        f"{figure_count} figures, {len(glossary)} glossary terms, {len(match_map)} glossary match names, and "
        f"{cae_result.candidate_count} CAE candidates, and {len(glossary_issue_rows)} glossary issue rows."
    )


if __name__ == "__main__":
    main()
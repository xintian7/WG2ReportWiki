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

from srcities_streamlit_app import (
    ENCRYPTED_REPORT_PATH,
    Glossary,
    build_glossary_pattern,
    glossary_parent_label,
    glossary_source_label,
    highlight_term,
    linkify_glossary_terms,
    load_encrypted_assets,
)


REPO_ROOT = Path(__file__).resolve().parents[1]
DEFAULT_SOURCE_JSON = REPO_ROOT / "data" / "srsod-inspection.json"
DEFAULT_REFERENCE_HTML = REPO_ROOT / "data" / "export" / "srsod-reconstructed.html"
DEFAULT_OUTPUT_HTML = REPO_ROOT / "data" / "export" / "SRCities_terminology_review.html"
REPORT_HEADER_TITLE = "SRCities terminology review (version Sep 3, 2026 based on SOD)"
REFERENCE_KICKER = '<p class="kicker">Reconstructed report</p>'
OUTPUT_KICKER = '<p class="kicker">Reconstructed report in HTML</p>'
GLOSSARY_TAB_ID = "glossary-overview-tab"
GLOSSARY_PANEL_ID = "glossary-overview-panel"
CAE_TAB_ID = "cae-check-tab"
CAE_PANEL_ID = "cae-check-panel"
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
GlossaryOccurrence = tuple[str, str, str, str, str]
GlossaryOccurrences = dict[str, list[GlossaryOccurrence]]
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
            .glossary-detail[hidden] { display: none; }
            .glossary-detail h3 { font-size: 1.45rem; font-weight: 400; margin-top: .8rem; }
            .glossary-detail-heading {
                align-items: baseline;
                display: flex;
                flex-wrap: wrap;
                gap: .3rem .45rem;
            }
            .glossary-parent { color: var(--muted); margin: .4rem 0 0; }
            .glossary-definition {
                border-left: 3px solid #9fcdd8;
                margin-top: 1rem;
                padding-left: .8rem;
            }
            .glossary-definition h4 { color: var(--ipcc-blue); font-size: 1rem; }
            .glossary-definition p { margin: .35rem 0 0; overflow-wrap: anywhere; }
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
                    const parent = sourceDetail.querySelector(".glossary-parent");
                    if (parent) dialogContent.append(parent.cloneNode(true));
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


class GlossaryMarkupLinker(HTMLParser):
    """Link glossary terms in narrative report text while preserving source markup."""

    VOID_TAGS = {"area", "base", "br", "col", "embed", "hr", "img", "input", "link", "meta", "source", "track", "wbr"}
    PROTECTED_TAGS = {"a", "button", "code", "pre", "script", "style"}
    PROTECTED_CLASSES = {"node-code", "source-reference"}
    LINKABLE_CLASSES = {"paragraph", "figure-explanation"}

    def __init__(self, markup: str, glossary: Glossary, excluded_root_ids: set[str]) -> None:
        super().__init__(convert_charrefs=False)
        self.markup = markup
        self.glossary = glossary
        self.excluded_root_ids = excluded_root_ids
        self.term_pattern = build_glossary_pattern(glossary)
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
        linked_text = linkify_glossary_terms(data, self.glossary, self.term_pattern)
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


def linkify_report_markup(markup: str, glossary: Glossary, excluded_root_ids: set[str]) -> str:
    """Make glossary terms clickable in eligible paragraphs and figure explanations."""
    parser = GlossaryMarkupLinker(markup, glossary, excluded_root_ids)
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


def glossary_keys_in_text(
    text: str,
    glossary: Glossary,
    term_pattern: re.Pattern[str] | None,
) -> set[str]:
    """Return boundary-safe glossary keys found in one sentence."""
    if term_pattern is None:
        return set()

    matched_keys: set[str] = set()
    for match in term_pattern.finditer(text):
        start, end = match.span()
        matched_text = match.group(0)
        term_key = matched_text.casefold()
        if term_key not in glossary:
            continue
        if matched_text[0].isalnum() and start > 0 and (text[start - 1].isalnum() or text[start - 1] == "_"):
            continue
        if matched_text[-1].isalnum() and end < len(text) and (text[end].isalnum() or text[end] == "_"):
            continue
        matched_keys.add(term_key)
    return matched_keys


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
    glossary: Glossary,
    node_codes: dict[str, str],
) -> GlossaryOccurrences:
    """Find glossary terms in report content outside references and supplements."""
    reports = payload.get("reports")
    if not isinstance(reports, list):
        raise ValueError("Inspection JSON must contain a reports list.")

    occurrences: GlossaryOccurrences = {term_key: [] for term_key in glossary}
    term_pattern = build_glossary_pattern(glossary)

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
                for term_key in glossary_keys_in_text(sentence, glossary, term_pattern):
                    occurrences[term_key].append((source_label, node_codes[node_id], node_id, sentence, text))
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
        '<h1>Confidence, Agreement, and Evidence check'
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


def render_glossary_evidence(term: str, matches: list[GlossaryOccurrence]) -> str:
    """Render section and sentence evidence without LLM content."""
    if not matches:
        return '<p class="glossary-result-count">No usage found in the SPM or Executive Summaries.</p>'

    rows = [
        '<details class="glossary-evidence">',
        f"<summary>Section / Sentence table ({len(matches)})</summary>",
        "<table>",
        "<thead><tr><th>Section</th><th>Sentence</th></tr></thead>",
        "<tbody>",
    ]
    for source_label, node_code, node_id, sentence, _ in matches:
        rows.append(
            "<tr>"
            '<td><span class="glossary-evidence-location">'
            f"{html.escape(source_label)}</span>"
            '<button class="glossary-evidence-code" type="button" '
            f'data-source-node-id="{html.escape(node_id, quote=True)}">{html.escape(node_code)}</button></td>'
            f"<td>{highlight_term(sentence, term)}</td>"
            "</tr>"
        )
    rows.extend(["</tbody>", "</table>", "</details>"])
    return "".join(rows)


def render_glossary_panel(
    glossary: Glossary,
    occurrences: GlossaryOccurrences,
) -> tuple[str, int]:
    """Render the Streamlit Glossary Overview workflow as a static tab panel."""
    usage_counts = {term_key: len(matches) for term_key, matches in occurrences.items()}
    used_term_count = sum(count > 0 for count in usage_counts.values())
    ordered_definitions = sorted(glossary.values(), key=lambda entries: entries[0][0].casefold())

    term_rows: list[str] = []
    detail_panels: list[str] = []
    for index, definitions in enumerate(ordered_definitions, start=1):
        term = definitions[0][0]
        term_key = term.casefold()
        frequency = usage_counts[term_key]
        sources = glossary_source_label(definitions)
        detail_id = f"glossary-term-detail-{index}"
        disabled = " disabled" if frequency == 0 else ""
        detail_attribute = f' data-detail-id="{detail_id}"' if frequency else ""
        term_rows.append(
            f'<li class="glossary-term-row" data-search="{html.escape(term_key, quote=True)}" '
            f'data-usage-count="{frequency}">'
            f'<button class="glossary-term-button" type="button" aria-pressed="false"{detail_attribute}{disabled}>'
            f"&#8226; {html.escape(term)} "
            f'<span class="glossary-term-count">[{frequency}]</span> '
            f'<span class="glossary-term-source">[{html.escape(sources)}]</span>'
            "</button>"
            "</li>"
        )
        if not frequency:
            continue

        parent = glossary_parent_label(definitions)
        parent_markup = f'<p class="glossary-parent">Parent: {html.escape(parent)}</p>' if parent else ""
        definition_markup = []
        for _, definition, _, source in definitions:
            definition_markup.append(
                '<section class="glossary-definition">'
                f"<h4>Explanation in {html.escape(source)}</h4>"
                f"<p>{html.escape(definition)}</p>"
                "</section>"
            )
        matches = occurrences[term_key]
        detail_panels.append(
            f'<article class="glossary-detail" id="{detail_id}" '
            f'data-term="{html.escape(term, quote=True)}" hidden>'
            '<div class="glossary-detail-heading">'
            f"<h3>{html.escape(term)}</h3>"
            f'<span class="glossary-term-count">[{len(matches)}]</span>'
            f'<span class="glossary-term-source">[{html.escape(sources)}]</span>'
            "</div>"
            f"{parent_markup}"
            f'{"".join(definition_markup)}'
            f"{render_glossary_evidence(term, matches)}"
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
        '<p class="source-reference">SRCities-SOD and AR6</p>'
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


def validate_cae_output(markup: str, result: CaeCheckResult) -> None:
    """Verify CAE tab order, aggregate counts, and malformed-case rows."""
    navigation, panels = parse_report_markup(markup)
    if len(panels) != len(REPORT_ORDER) + 2:
        raise ValueError("Generated output must contain seven reports, CAE check, and Glossary Overview.")
    cae_panel = panels[-2]
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
    if len(tabs) != len(REPORT_ORDER) + 2 or tabs[-2].attributes.get("aria-controls") != CAE_PANEL_ID:
        raise ValueError("CAE check must be the penultimate chapter-navigation tab.")

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


def validate_glossary_output(markup: str, glossary: Glossary, used_term_count: int) -> None:
    """Verify the glossary tab mirrors app data without any LLM section."""
    _, panels = parse_report_markup(markup)
    if len(panels) != len(REPORT_ORDER) + 2:
        raise ValueError("Generated output must contain seven reports, CAE check, and Glossary Overview.")
    if panels[-1].attributes.get("id") != GLOSSARY_PANEL_ID or "hidden" not in panels[-1].attributes:
        raise ValueError("Glossary Overview must be the final, initially hidden panel.")

    navigation, _ = parse_report_markup(markup)
    tabs = parse_chapter_tabs(markup[navigation.start : navigation.end])
    if len(tabs) != len(REPORT_ORDER) + 2 or tabs[-1].attributes.get("aria-controls") != GLOSSARY_PANEL_ID:
        raise ValueError("Glossary Overview must be the final chapter-navigation tab.")

    glossary_panel = markup[panels[-1].start : panels[-1].end]
    if glossary_panel.count('class="glossary-term-row"') != len(glossary):
        raise ValueError("Generated glossary term count differs from the encrypted bundle.")
    if glossary_panel.count('class="glossary-detail"') != used_term_count:
        raise ValueError("Generated selectable glossary count differs from app usage counts.")
    forbidden_llm_features = ('class="term-llm-summary"', "LLM summary", "llm_summary_")
    if any(feature in glossary_panel for feature in forbidden_llm_features):
        raise ValueError("Glossary Overview must not contain LLM summary content or controls.")


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
    parser.add_argument("--archive", type=Path, default=ENCRYPTED_REPORT_PATH)
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT_HTML)
    return parser.parse_args()


def main() -> None:
    """Generate and validate the report viewer."""
    args = parse_args()
    source_json = args.source_json.expanduser()
    reference_html = args.reference_html.expanduser()
    archive_path = args.archive.expanduser()
    output_path = args.output.expanduser()

    payload = json.loads(source_json.read_text(encoding="utf-8"))
    if not isinstance(payload, dict):
        raise ValueError("Inspection JSON must have an object at its root.")
    root_ids, node_ids, figure_count = canonical_report_data(payload)

    markup = reference_html.read_text(encoding="utf-8")
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

    if not archive_path.is_file():
        raise FileNotFoundError("Encrypted glossary archive is unavailable.")
    _, glossary, _, _, _, _ = load_encrypted_assets(
        str(archive_path),
        archive_path.stat().st_mtime_ns,
    )
    node_codes = report_node_codes(markup)
    occurrences = full_report_term_occurrences(payload, glossary, node_codes)
    cae_result = full_report_cae_check(payload, node_codes)
    markup = linkify_report_markup(markup, glossary, excluded_glossary_root_ids(payload))
    markup = add_cae_check(markup, render_cae_panel(cae_result))
    glossary_panel, used_term_count = render_glossary_panel(glossary, occurrences)
    markup = add_glossary_overview(markup, glossary_panel)
    validate_cae_output(markup, cae_result)
    validate_glossary_output(markup, glossary, used_term_count)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(markup, encoding="utf-8")
    print(
        f"Wrote {output_path} with {len(root_ids)} reports, {sum(node_ids.values())} nodes, "
        f"{figure_count} figures, {len(glossary)} glossary terms, and "
        f"{cae_result.candidate_count} CAE candidates."
    )


if __name__ == "__main__":
    main()
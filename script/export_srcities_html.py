#!/usr/bin/env python3
"""Export SRCities review content as a standalone local HTML document.

The output contains plaintext report material. Keep the generated file in an
approved internal location and do not publish it to a public web host.
"""

from __future__ import annotations

import argparse
import html
from pathlib import Path
import re

from srcities_streamlit_app import (
    ENCRYPTED_REPORT_PATH,
    EXECUTIVE_SUMMARY_ITEM_RE,
    FIGURE_CAPTION_RE,
    FIRST_LEVEL_STATEMENT_RE,
    LLM_SUMMARY_EXCLUDED_TERM_KEYS,
    MARKDOWN_RENDERER,
    REPO_ROOT,
    SUMMARY_MARKDOWN_RENDERER,
    SPM_SECTION_LABELS,
    STATEMENT_NUMBER_RE,
    ExecutiveSummaries,
    Glossary,
    StatementOccurrence,
    TermUsageSummaries,
    build_glossary_pattern,
    glossary_parent_label,
    glossary_source_label,
    glossary_usage_counts,
    highlight_term,
    linkify_glossary_html,
    linkify_glossary_terms,
    llm_summary_not_applied_message,
    load_encrypted_assets,
    normalize_text,
    parse_markdown_sections,
    summary_has_potential_issue,
    term_anchor_id,
    term_occurrences,
)


DEFAULT_OUTPUT_PATH = REPO_ROOT / "data" / "export" / "SRCities_terminology_review.html"
EXPORT_TITLE = "SRCities terminology review (version August 21, 2026 based on SOD)"
TERM_LINK_RE = re.compile(r'<a href="#" data-term="(?P<term>[^"]+)">')
SourceStatementKey = tuple[str, str]
SourceStatementAnchors = dict[SourceStatementKey, str]


DOCUMENT_CSS = """
:root {
  color-scheme: light;
  --canvas: #edf5f5;
  --surface: #ffffff;
  --ink: #1e2a30;
  --muted: #5b696e;
  --line: #c7d8d9;
  --accent: #006b8f;
  --accent-deep: #004f6a;
  --highlight: #e1f5fa;
  --target: #edf9fb;
  --notice: #fff6d5;
  --issue: #b42318;
}

* {
  box-sizing: border-box;
}

html {
  scroll-behavior: smooth;
}

body {
  background: var(--canvas);
  color: var(--ink);
  font-family: "Avenir Next", "Helvetica Neue", sans-serif;
  line-height: 1.58;
  margin: 0;
}

a {
  color: var(--accent-deep);
}

a:focus-visible,
input:focus-visible,
summary:focus-visible {
  outline: 3px solid #d28a26;
  outline-offset: 2px;
}

.site-header {
  background: var(--surface);
  border-bottom: 4px solid var(--accent);
  padding: 1.35rem 1.5rem 1rem;
}

.site-header__inner,
.page-content,
.site-footer {
  margin: 0 auto;
  max-width: 1120px;
}

.site-header h1,
.section-header h2,
.glossary-term h3 {
  font-family: "Iowan Old Style", "Palatino Linotype", Georgia, serif;
  letter-spacing: 0;
}

.site-header h1 {
  font-size: 2rem;
  line-height: 1.15;
  margin: 0;
}

.primary-nav,
.section-index {
  display: flex;
  flex-wrap: wrap;
  gap: 0.35rem 0.9rem;
}

.primary-nav {
  margin-top: 0.9rem;
}

.primary-nav a {
  font-weight: 800;
}

.page-content {
  padding: 1.25rem 1.5rem 2rem;
}

.document-section {
  background: var(--surface);
  border-top: 3px solid var(--accent);
  margin-top: 1.25rem;
  padding: 1.2rem;
}

.document-section:first-child {
  margin-top: 0;
}

.section-header {
  border-bottom: 1px solid var(--line);
  margin-bottom: 0.9rem;
  padding-bottom: 0.7rem;
}

.section-header h2 {
  color: var(--accent-deep);
  font-size: 1.55rem;
  line-height: 1.2;
  margin: 0;
}

.section-header p {
  color: var(--muted);
  margin: 0.35rem 0 0;
}

.section-index {
  margin: 0 0 1rem;
}

.section-index a {
  font-size: 0.92rem;
}

.source-section,
.executive-summary {
  border-top: 1px solid var(--line);
  margin-top: 1rem;
  padding-top: 1rem;
}

.source-section h3,
.executive-summary > h3 {
  color: var(--accent-deep);
  font-size: 1.2rem;
  letter-spacing: 0;
  margin: 0 0 0.65rem;
}

.source-section h4,
.executive-summary h4,
.executive-summary h5,
.executive-summary h6 {
  color: var(--accent-deep);
  font-size: 1rem;
  letter-spacing: 0;
  margin: 0.95rem 0 0.4rem;
}

.source-section p,
.executive-summary p {
  margin: 0 0 0.7rem;
}

.executive-summary ul,
.executive-summary ol {
  margin: 0.35rem 0 0.75rem;
  padding-left: 1.3rem;
}

.glossary-link {
  background: var(--highlight);
  border-radius: 2px;
  color: var(--accent-deep);
  padding: 0 0.08em;
  text-decoration-thickness: 1px;
  text-underline-offset: 0.13em;
}

.statement-link {
  color: var(--accent-deep);
  font-weight: 800;
  text-decoration-thickness: 1px;
  text-underline-offset: 0.13em;
}

.statement-link:hover {
  color: var(--accent);
}

.statement-target {
  scroll-margin-top: 1rem;
}

.statement-target:target {
  background: var(--target);
  box-shadow: inset 4px 0 0 var(--accent);
  padding-left: 0.65rem;
}

.glossary-summary-note {
  background: var(--notice);
  border-left: 3px solid #d28a26;
  margin: 0 0 1rem;
  padding: 0.7rem 0.85rem;
}

.glossary-controls {
  align-items: end;
  display: flex;
  flex-wrap: wrap;
  gap: 0.6rem 1rem;
  margin-bottom: 0.75rem;
}

.glossary-controls label {
  display: grid;
  font-weight: 800;
  gap: 0.25rem;
}

.glossary-controls input {
  border: 1px solid #8aaeb9;
  border-radius: 3px;
  color: var(--ink);
  font: inherit;
  min-width: min(22rem, 70vw);
  padding: 0.45rem 0.55rem;
}

.glossary-result-count {
  color: var(--muted);
  margin: 0 0 0.15rem;
}

.glossary-term {
  border-top: 1px solid var(--line);
  padding: 1rem 0;
  scroll-margin-top: 1rem;
}

.glossary-term[hidden] {
  display: none;
}

.glossary-term.is-focused {
  background: var(--target);
  box-shadow: inset 4px 0 0 var(--accent);
  padding-left: 0.75rem;
}

.glossary-term__heading {
  align-items: baseline;
  display: flex;
  flex-wrap: wrap;
  gap: 0.25rem 0.45rem;
}

.glossary-term h3 {
  font-size: 1.3rem;
  margin: 0;
}

.term-count {
  color: #6f4a20;
  font-weight: 800;
}

.term-source {
  color: var(--accent);
  font-weight: 800;
}

.term-issue {
  color: var(--issue);
  font-weight: 800;
}

.term-parent {
  color: var(--muted);
  margin: 0.35rem 0 0;
}

.definition-block {
  border-left: 2px solid #9fcdd8;
  margin-top: 0.7rem;
  min-width: 0;
  padding-left: 0.75rem;
}

.definition-block h4 {
  font-size: 0.96rem;
  letter-spacing: 0;
  margin: 0;
}

.definition-block p {
  margin: 0.25rem 0 0;
  overflow-wrap: anywhere;
}

.term-llm-summary,
.term-evidence {
  border-top: 1px dashed #a9bdc2;
  margin-top: 0.8rem;
  padding-top: 0.65rem;
}

.term-llm-summary summary,
.term-evidence summary {
  color: var(--accent-deep);
  cursor: pointer;
  font-weight: 800;
}

.term-llm-summary__content,
.term-evidence__content {
  margin-top: 0.7rem;
  min-width: 0;
  overflow-wrap: anywhere;
}

.term-llm-summary__content p,
.term-llm-summary__content ul,
.term-llm-summary__content ol,
.term-evidence__content p {
  margin: 0.35rem 0 0.7rem;
  max-width: 100%;
}

.term-llm-summary__content li,
.term-llm-summary__content td {
  overflow-wrap: anywhere;
}

.term-llm-summary__content table,
.term-evidence__table {
  border-collapse: collapse;
  display: block;
  margin: 0.6rem 0;
  max-width: 100%;
  overflow-x: auto;
}

.term-llm-summary__content th,
.term-llm-summary__content td,
.term-evidence__table th,
.term-evidence__table td {
  border: 1px solid var(--line);
  padding: 0.4rem 0.5rem;
  text-align: left;
  vertical-align: top;
}

.term-evidence__table th {
  background: #edf7f8;
  color: var(--accent-deep);
}

.term-evidence__table td:first-child {
  min-width: 6rem;
  overflow-wrap: anywhere;
}

mark {
  background: #fff0b8;
  color: inherit;
  padding: 0 0.08em;
}

.site-footer {
  color: var(--muted);
  font-size: 0.9rem;
  padding: 0 1.5rem 1.5rem;
}

@media (max-width: 720px) {
  .site-header,
  .page-content {
    padding-left: 1rem;
    padding-right: 1rem;
  }

  .document-section {
    padding: 1rem;
  }

  .site-header h1 {
    font-size: 1.65rem;
  }
}

@media print {
  body {
    background: #ffffff;
  }

  .primary-nav,
  .glossary-controls,
  .glossary-result-count {
    display: none;
  }
}
"""


DOCUMENT_JAVASCRIPT = """
(() => {
  const searchInput = document.getElementById("glossary-search");
  const resultCount = document.getElementById("glossary-result-count");
  const entries = [...document.querySelectorAll(".glossary-term")];

  const filterGlossary = () => {
    const query = searchInput ? searchInput.value.trim().toLocaleLowerCase() : "";
    let visibleCount = 0;
    for (const entry of entries) {
      const visible = !query || entry.dataset.search.includes(query);
      entry.hidden = !visible;
      if (visible) visibleCount += 1;
    }
    if (resultCount) {
      resultCount.textContent = query
        ? `${visibleCount} matching glossary term${visibleCount === 1 ? "" : "s"}`
        : `${entries.length} glossary terms`;
    }
  };

  searchInput?.addEventListener("input", filterGlossary);
  filterGlossary();

  document.addEventListener("click", event => {
    const link = event.target.closest("a.glossary-link");
    if (!link) return;

    const targetId = link.getAttribute("href")?.slice(1);
    const target = targetId ? document.getElementById(targetId) : null;
    if (!target) return;

    event.preventDefault();
    if (searchInput) searchInput.value = "";
    filterGlossary();
    target.classList.add("is-focused");
    target.scrollIntoView({ behavior: "smooth", block: "start" });
    window.location.hash = target.id;
    window.setTimeout(() => target.classList.remove("is-focused"), 1800);
  });
})();
"""


def heading_anchor_id(prefix: str, heading: str) -> str:
    """Build a stable anchor identifier for a document heading."""
    slug = re.sub(r"[^a-z0-9]+", "-", heading.casefold()).strip("-")
    return f"{prefix}-{slug or 'section'}"


def source_statement_key(statement_number: str, statement_text: str) -> SourceStatementKey:
    """Build the identity shared by a rendered source item and an occurrence match."""
    return (
        normalize_text(statement_number).casefold(),
        normalize_text(statement_text).casefold(),
    )


def spm_statement_values(section_name: str, line: str) -> tuple[str, str]:
    """Resolve an SPM line using the app's statement-occurrence rules."""
    number_match = STATEMENT_NUMBER_RE.match(line)
    first_level_match = FIRST_LEVEL_STATEMENT_RE.match(line)
    figure_match = FIGURE_CAPTION_RE.match(line)
    if figure_match and section_name.startswith("Section "):
        section_identifier = section_name.removeprefix("Section ")
        return f"{section_identifier}-Figure {figure_match.group(1)}", figure_match.group(2)
    if number_match:
        return number_match.groups()
    if first_level_match:
        return first_level_match.groups()
    return section_name, line


def executive_summary_key(markdown_block: str) -> SourceStatementKey | None:
    """Return the occurrence key represented by one Executive Summary block."""
    item_match = EXECUTIVE_SUMMARY_ITEM_RE.match(markdown_block.strip())
    if not item_match:
        return None
    statement_number, markdown_text = item_match.groups()
    return source_statement_key(statement_number, markdown_text)


def build_source_statement_anchors(
    sections: dict[str, list[str]],
    executive_summaries: ExecutiveSummaries,
) -> SourceStatementAnchors:
    """Assign one source anchor to each unique occurrence source unit."""
    anchors: SourceStatementAnchors = {}

    def add_anchor(key: SourceStatementKey) -> None:
        if key not in anchors:
            anchors[key] = f"source-statement-{len(anchors) + 1}"

    for section_name, lines in sections.items():
        for line in lines:
            add_anchor(source_statement_key(*spm_statement_values(section_name, line)))

    for summary_text in executive_summaries.values():
        for raw_line in summary_text.splitlines():
            key = executive_summary_key(raw_line)
            if key is not None:
                add_anchor(key)

    return anchors


def source_attributes(
    key: SourceStatementKey | None,
    source_anchors: SourceStatementAnchors,
    rendered_keys: set[SourceStatementKey],
) -> str:
    """Return id and class attributes for the first source rendering of a key."""
    if key is None or key in rendered_keys:
        return ""
    anchor = source_anchors.get(key)
    if anchor is None:
        return ""
    rendered_keys.add(key)
    return f' id="{html.escape(anchor, quote=True)}" class="statement-target"'


def glossary_anchor_ids(glossary: Glossary) -> dict[str, str]:
    """Build unique in-page anchors for normalized glossary terms."""
    anchors: dict[str, str] = {}
    claimed_ids: set[str] = set()
    for term_key in sorted(glossary):
        base_anchor = term_anchor_id(term_key)
        anchor = base_anchor
        suffix = 2
        while anchor in claimed_ids:
            anchor = f"{base_anchor}-{suffix}"
            suffix += 1
        anchors[term_key] = anchor
        claimed_ids.add(anchor)
    return anchors


def glossary_links_to_anchors(rendered_html: str, glossary_anchors: dict[str, str]) -> str:
    """Convert the app component's term triggers to document-local glossary links."""

    def replace_link(match: re.Match[str]) -> str:
        term_key = normalize_text(html.unescape(match.group("term"))).casefold()
        anchor = glossary_anchors.get(term_key)
        if anchor is None:
            return match.group(0)
        return f'<a class="glossary-link" href="#{html.escape(anchor, quote=True)}">'

    return TERM_LINK_RE.sub(replace_link, rendered_html)


def statement_link_html(
    statement_number: str,
    statement_text: str,
    source_anchors: SourceStatementAnchors,
) -> str:
    """Render an occurrence statement label as a link back to its source text."""
    label = html.escape(statement_number)
    source_anchor = source_anchors.get(source_statement_key(statement_number, statement_text))
    if source_anchor is None:
        return label
    return f'<a class="statement-link" href="#{html.escape(source_anchor, quote=True)}">{label}</a>'


def linkify_summary_statement_references(
    rendered_html: str,
    matches: list[StatementOccurrence],
    source_anchors: SourceStatementAnchors,
) -> str:
    """Link LLM-summary statement references to their rendered source text."""
    statement_anchors: dict[str, str] = {}
    for statement_number, _, statement_text in matches:
        source_anchor = source_anchors.get(source_statement_key(statement_number, statement_text))
        if source_anchor:
            statement_anchors[normalize_text(statement_number).casefold()] = source_anchor
    if not statement_anchors:
        return rendered_html

    identifiers = sorted(statement_anchors, key=len, reverse=True)
    identifier_pattern = re.compile(
        rf"(?<![A-Za-z0-9.-])(?P<identifier>{'|'.join(re.escape(identifier) for identifier in identifiers)})"
        r"(?![A-Za-z0-9.-])",
        re.IGNORECASE,
    )

    def linkify_text(text: str) -> str:
        fragments: list[str] = []
        last_end = 0
        for match in identifier_pattern.finditer(text):
            identifier = match.group("identifier")
            source_anchor = statement_anchors[identifier.casefold()]
            fragments.append(html.escape(text[last_end:match.start()]))
            fragments.append(
                f'<a class="statement-link" href="#{html.escape(source_anchor, quote=True)}">'
                f"{html.escape(identifier)}</a>"
            )
            last_end = match.end()
        fragments.append(html.escape(text[last_end:]))
        return "".join(fragments)

    fragments: list[str] = []
    protected_depth = 0
    protected_tags = {"a", "code", "pre"}
    for fragment in re.split(r"(<[^>]+>)", rendered_html):
        tag_match = re.match(r"</?([a-zA-Z][a-zA-Z0-9]*)", fragment)
        if not tag_match:
            fragments.append(fragment if protected_depth else linkify_text(html.unescape(fragment)))
            continue

        tag_name = tag_match.group(1).lower()
        if tag_name in protected_tags and fragment.startswith("</"):
            protected_depth -= 1
        fragments.append(fragment)
        if tag_name in protected_tags and not fragment.startswith("</") and not fragment.endswith("/>"):
            protected_depth += 1

    return "".join(fragments)


def render_spm_sections(
    sections: dict[str, list[str]],
    glossary: Glossary,
    glossary_anchors: dict[str, str],
    source_anchors: SourceStatementAnchors,
    rendered_source_keys: set[SourceStatementKey],
) -> str:
    """Render the six SPM sections with source and glossary anchors."""
    term_pattern = build_glossary_pattern(glossary)
    spm_sections = {section_name: sections.get(section_name, []) for section_name in SPM_SECTION_LABELS}
    section_links = [(name, heading_anchor_id("spm", name)) for name in spm_sections]
    fragments = [
        '<section class="document-section" id="spm">',
        '<header class="section-header"><h2>Summary for Policymakers</h2>'
      '<p>Glossary terms are highlighted and linked to the glossary overview.</p></header>',
        '<nav class="section-index" aria-label="SPM sections">',
    ]
    fragments.extend(
        f'<a href="#{html.escape(anchor, quote=True)}">{html.escape(section_name)}</a>'
        for section_name, anchor in section_links
    )
    fragments.append("</nav>")

    for section_name, section_anchor in section_links:
        fragments.extend(
            [
                f'<section class="source-section" id="{html.escape(section_anchor, quote=True)}">',
                f"<h3>{html.escape(section_name)}</h3>",
            ]
        )
        lines = spm_sections[section_name]
        if not lines:
            fragments.append('<p class="term-parent">No content parsed for this section.</p>')
        for line in lines:
            source_key = source_statement_key(*spm_statement_values(section_name, line))
            attributes = source_attributes(source_key, source_anchors, rendered_source_keys)
            rendered_line = glossary_links_to_anchors(
                linkify_glossary_terms(line, glossary, term_pattern),
                glossary_anchors,
            )
            if FIRST_LEVEL_STATEMENT_RE.match(line):
                fragments.append(f"<h4{attributes}>{rendered_line}</h4>")
            else:
                fragments.append(f"<p{attributes}>{rendered_line}</p>")
        fragments.append("</section>")

    fragments.append("</section>")
    return "\n".join(fragments)


def render_executive_summaries(
    executive_summaries: ExecutiveSummaries,
    glossary: Glossary,
    glossary_anchors: dict[str, str],
    source_anchors: SourceStatementAnchors,
    rendered_source_keys: set[SourceStatementKey],
) -> str:
    """Render Executive Summary Markdown with source and glossary anchors."""
    term_pattern = build_glossary_pattern(glossary)
    chapter_anchors = {
        chapter: heading_anchor_id("executive-summary", chapter)
        for chapter in executive_summaries
    }
    fragments = [
        '<section class="document-section" id="executive-summaries">',
        '<header class="section-header"><h2>Chapter Executive Summaries</h2>'
      '<p>Glossary terms are highlighted and linked to the glossary overview.</p></header>',
        '<nav class="section-index" aria-label="Executive Summary chapters">',
    ]
    fragments.extend(
        f'<a href="#{html.escape(anchor, quote=True)}">{html.escape(chapter)}</a>'
        for chapter, anchor in chapter_anchors.items()
    )
    fragments.append("</nav>")

    for chapter, markdown_text in executive_summaries.items():
        rendered_markdown = re.sub(
            r"^(#{1,4})\s",
            lambda match: f"{'#' * min(len(match.group(1)) + 2, 6)} ",
            markdown_text,
            flags=re.MULTILINE,
        )
        fragments.extend(
            [
                f'<section class="executive-summary" id="{html.escape(chapter_anchors[chapter], quote=True)}">',
                f"<h3>{html.escape(chapter)}</h3>",
            ]
        )
        for block in re.split(r"\n\s*\n", rendered_markdown):
            if not block.strip():
                continue
            rendered_html = MARKDOWN_RENDERER.render(block)
            linked_html = glossary_links_to_anchors(
                linkify_glossary_html(rendered_html, glossary, term_pattern),
                glossary_anchors,
            )
            attributes = source_attributes(
                executive_summary_key(block),
                source_anchors,
                rendered_source_keys,
            )
            if attributes:
                fragments.append(f"<div{attributes}>{linked_html}</div>")
            else:
                fragments.append(linked_html)
        fragments.append("</section>")

    fragments.append("</section>")
    return "\n".join(fragments)


def render_llm_summary(
    term_key: str,
    term: str,
    summaries: TermUsageSummaries,
    matches: list[StatementOccurrence],
    source_anchors: SourceStatementAnchors,
) -> str:
    """Render a precomputed LLM summary with source-text statement links."""
    if term_key in LLM_SUMMARY_EXCLUDED_TERM_KEYS:
        content = f'<p>{html.escape(llm_summary_not_applied_message(term))}</p>'
    else:
        summary = summaries.get(term_key, "")
        if summary:
            content = linkify_summary_statement_references(
                SUMMARY_MARKDOWN_RENDERER.render(summary),
                matches,
                source_anchors,
            )
        else:
            content = "<p>No precomputed LLM summary is available for this term.</p>"
    return (
        '<details class="term-llm-summary">'
        "<summary>LLM summary</summary>"
        f'<div class="term-llm-summary__content">{content}</div>'
        "</details>"
    )


def render_term_evidence(
    term: str,
    matches: list[StatementOccurrence],
    source_anchors: SourceStatementAnchors,
) -> str:
    """Render the app's Statement / Sentence evidence table as a disclosure."""
    if not matches:
        content = "<p>No usage found in the SPM or Executive Summaries.</p>"
    else:
        rows = [
            '<table class="term-evidence__table">',
            "<thead><tr><th>Statement</th><th>Sentence</th></tr></thead>",
            "<tbody>",
        ]
        for statement_number, sentence, statement_text in matches:
            rows.append(
                "<tr>"
                f"<td>{statement_link_html(statement_number, statement_text, source_anchors)}</td>"
                f"<td>{highlight_term(sentence, term)}</td>"
                "</tr>"
            )
        rows.append("</tbody></table>")
        content = "".join(rows)
    return (
        '<details class="term-evidence">'
        f"<summary>Statement / Sentence table ({len(matches)})</summary>"
        f'<div class="term-evidence__content">{content}</div>'
        "</details>"
    )


def render_glossary_overview(
    glossary: Glossary,
    sections: dict[str, list[str]],
    executive_summaries: ExecutiveSummaries,
    summaries: TermUsageSummaries,
    glossary_anchors: dict[str, str],
    source_anchors: SourceStatementAnchors,
) -> str:
    """Render every glossary term, definition, summary, and evidence table."""
    usage_counts = glossary_usage_counts(glossary, sections, executive_summaries)
    used_term_count = sum(count > 0 for count in usage_counts.values())
    fragments = [
        '<section class="document-section" id="glossary-overview">',
        '<header class="section-header"><h2>Glossary Overview</h2>',
        f"<p>{used_term_count} of {len(glossary)} glossary terms occur in the SPM or Executive Summaries.</p>",
        "</header>",
        '<p class="glossary-summary-note">'
        "Each entry contains the same precomputed LLM summary used by the app. "
        "It is reference material only and does not replace human judgment for consistency checks."
        "</p>",
        '<div class="glossary-controls">',
        '<label for="glossary-search">Find a glossary term'
        '<input id="glossary-search" type="search" autocomplete="off" placeholder="Type a term name">'
        "</label>",
        '<p class="glossary-result-count" id="glossary-result-count" aria-live="polite"></p>',
        "</div>",
    ]

    ordered_definitions = sorted(glossary.values(), key=lambda entries: entries[0][0].casefold())
    for definitions in ordered_definitions:
        term = definitions[0][0]
        term_key = term.casefold()
        matches = term_occurrences(term, "", sections, executive_summaries)
        issue_indicator = ""
        if term_key not in LLM_SUMMARY_EXCLUDED_TERM_KEYS and summary_has_potential_issue(summaries.get(term_key, "")):
            issue_indicator = (
                '<span class="term-issue" role="img" '
                'aria-label="Potential issue needing substantive review" '
                'title="Potential issue needing substantive review">!</span>'
            )
        parent = glossary_parent_label(definitions)
        parent_html = f'<p class="term-parent">Parent: {html.escape(parent)}</p>' if parent else ""
        definition_fragments = []
        for _, definition, _, source in definitions:
            definition_fragments.append(
                '<section class="definition-block">'
                f"<h4>Explanation in {html.escape(source)}</h4>"
                f"<p>{html.escape(definition)}</p>"
                "</section>"
            )
        frequency = usage_counts[term_key]
        sources = glossary_source_label(definitions)
        fragments.extend(
            [
                f'<article class="glossary-term" id="{html.escape(glossary_anchors[term_key], quote=True)}" '
                f'data-search="{html.escape(term.casefold(), quote=True)}">',
                '<div class="glossary-term__heading">',
                f"<h3>{html.escape(term)}</h3>",
                f'<span class="term-count">[{frequency}]</span>',
                f'<span class="term-source">[{html.escape(sources)}]</span>',
                issue_indicator,
                "</div>",
                parent_html,
                *definition_fragments,
                render_llm_summary(term_key, term, summaries, matches, source_anchors),
                render_term_evidence(term, matches, source_anchors),
                "</article>",
            ]
        )

    fragments.append("</section>")
    return "\n".join(fragments)


def build_document(
    sections: dict[str, list[str]],
    executive_summaries: ExecutiveSummaries,
    glossary: Glossary,
    summaries: TermUsageSummaries,
) -> str:
    """Build the complete standalone document from encrypted application assets."""
    glossary_anchors = glossary_anchor_ids(glossary)
    source_anchors = build_source_statement_anchors(sections, executive_summaries)
    rendered_source_keys: set[SourceStatementKey] = set()
    body = "\n".join(
        [
            '<header class="site-header">',
            '<div class="site-header__inner">',
            f"<h1>{html.escape(EXPORT_TITLE)}</h1>",
            '<nav class="primary-nav" aria-label="Document sections">',
            '<a href="#spm">SPM</a>',
            '<a href="#executive-summaries">Executive Summaries</a>',
            '<a href="#glossary-overview">Glossary Overview</a>',
            "</nav>",
            "</div>",
            "</header>",
            '<main class="page-content">',
            render_spm_sections(
                sections,
                glossary,
                glossary_anchors,
                source_anchors,
                rendered_source_keys,
            ),
            render_executive_summaries(
                executive_summaries,
                glossary,
                glossary_anchors,
                source_anchors,
                rendered_source_keys,
            ),
            render_glossary_overview(
                glossary,
                sections,
                executive_summaries,
                summaries,
                glossary_anchors,
                source_anchors,
            ),
            "</main>",
            '<footer class="site-footer">Generated locally from the encrypted application bundle.</footer>',
        ]
    )
    return "\n".join(
        [
            "<!doctype html>",
            '<html lang="en">',
            "<head>",
            '  <meta charset="utf-8">',
            '  <meta name="viewport" content="width=device-width, initial-scale=1">',
            f"  <title>{html.escape(EXPORT_TITLE)}</title>",
            "  <style>",
            DOCUMENT_CSS,
            "  </style>",
            "</head>",
            "<body>",
            body,
            "<script>",
            DOCUMENT_JAVASCRIPT,
            "</script>",
            "</body>",
            "</html>",
        ]
    )


def parse_args() -> argparse.Namespace:
    """Parse exporter paths while retaining secure defaults from the app."""
    parser = argparse.ArgumentParser(description="Export SRCities review content to standalone HTML.")
    parser.add_argument(
        "--archive",
        type=Path,
        default=ENCRYPTED_REPORT_PATH,
        help="Encrypted app archive to export from.",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=DEFAULT_OUTPUT_PATH,
        help="Destination HTML file. The default path is ignored by Git.",
    )
    return parser.parse_args()


def main() -> None:
    """Load the encrypted application data and write its HTML export."""
    args = parse_args()
    archive_path = args.archive.expanduser()
    if not archive_path.is_file():
        raise SystemExit("Encrypted report archive is unavailable.")

    try:
        (
            report_text,
            glossary,
            _,
            executive_summaries,
            _,
            summaries,
        ) = load_encrypted_assets(str(archive_path), archive_path.stat().st_mtime_ns)
    except (FileNotFoundError, RuntimeError, ValueError) as error:
        raise SystemExit(
            "Could not load the encrypted report archive. Ensure FERNET_KEY is configured before exporting."
        ) from error

    _, sections = parse_markdown_sections(report_text)
    output_path = args.output.expanduser()
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(
        build_document(sections, executive_summaries, glossary, summaries),
        encoding="utf-8",
    )
    print(
        f"Wrote {output_path} with {len(SPM_SECTION_LABELS)} SPM sections, "
        f"{len(executive_summaries)} Executive Summaries, and {len(glossary)} glossary terms."
    )


if __name__ == "__main__":
    main()
#!/usr/bin/env python3
"""
CSS Proposal Generator — standalone single-file module.

Does not modify app.py or any other project files. Run directly:

    python css_proposal_generator.py              # web UI on http://0.0.0.0:8010/
    python css_proposal_generator.py --build-template   # rebuild base .docx from source

Assets expected in the same folder as this script:
    CSS_Proposal_Source.docx       (original template; replace and run --build-template to update)
    CSS_Proposal_Base_Template.docx (prepared docxtpl template; auto-built if missing)
    default_logo.png               (optional fallback logo)
"""
from __future__ import annotations

import argparse
import copy
import json
import os
import re
import shutil
import tempfile
import traceback
import zipfile
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches
from docxtpl import DocxTemplate, InlineImage
from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, HTMLResponse

try:
    from PIL import Image

    PIL_AVAILABLE = True
except ImportError:
    PIL_AVAILABLE = False

# ---------------------------------------------------------------------------
# Paths (all relative to this file — no imports from other project modules)
# ---------------------------------------------------------------------------
HERE = Path(__file__).resolve().parent
SOURCE_TEMPLATE = HERE / "CSS_Proposal_Source.docx"
BASE_TEMPLATE = HERE / "CSS_Proposal_Base_Template.docx"
DEFAULT_LOGO = HERE / "default_logo.png"
UPLOAD_DIR = HERE / "uploads" / "css_proposal"
VERSION = "1.2.0"
DEFAULT_PORT = 8010

# Layout-based page estimation (lines per printed page; tuned for A4 proposal layout).
LINES_PER_PAGE = 38.0
CHARS_PER_LINE = 82
PARAS_PER_PAGE = LINES_PER_PAGE  # backwards compatibility for callers
CONTENTS_HEADING = "Contents"
CONCLUSION_TOC_PREFIX = "Conclusion & Pricing"

# Default Contents page numbers for the unchanged intro block (from base template).
DEFAULT_STATIC_TOC_PAGES: tuple[int, ...] = (3, 4, 4, 5, 7, 8)
DEFAULT_FIRST_ANNEXURE_PAGE = 13
TERMS_MARKER = "Terms & Conditions"
TERMS_TOC_LABEL = "Terms & Conditions"
# Compact Contents spacing (twips): keeps all TOC lines on one page.
TOC_LINE_BEFORE_TWIPS = "40"       # ~2pt between entries
TOC_FIRST_LINE_BEFORE_TWIPS = "100"  # ~5pt after "Contents" heading
TOC_TAB_POS = "9887"
STATIC_TOC_ANCHORS: tuple[tuple[str, int], ...] = (
    ("INTRODUCTION", 37),
    ("OUR UNDERSTANDING OF YOUR REQUIREMENT", 37),
    ("METHODOLOGY OVERVIEW", 37),
    ("VULNERABILITY ASSESSMENT METHODOLOGY", 37),
    ("PROPOSED TIMELINES & PROJECT MANAGEMENT", 47),
    ("TEAM OVERVIEW", 55),
)

TRAILER_MARKER = "Conclusion & Pricing"
CLIENT_NAME_LITERALS: tuple[str, ...] = ("XYZ LLC", "Company name")

# Default Conclusion pricing-table cells for known assessments (quantity, frequency, duration).
PRICING_ROW_DEFAULTS: dict[str, tuple[str, str, str]] = {
    "post_breach": ("1 Web Application", "One Time", ""),
    "active_directory": ("2 Web Application", "Three Time", ""),
    "web_app_pt": ("5 Web Applications", "One Time", ""),
    "mobile_app_pt": ("Android - 6, IoS - 3", "One Time", ""),
}
ANNEXURE_HEADING_RE = re.compile(r"^Annexure\s+(\d+)\s*:\s*(.+)$", re.IGNORECASE)

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
PIC_TAG = f"{{{W_NS}}}drawing"

TEXT_REPLACEMENTS: tuple[tuple[str, str], ...] = (
    ("XYZ LLC", "{{ client_name }}"),
    ("Company name", "{{ client_name }}"),
    ("EN-01-12-25", "{{ proposal_ref }}"),
)

XML_PARTS = tuple(
    f"word/{name}"
    for name in (
        "document.xml",
        "header1.xml",
        "header2.xml",
        "header3.xml",
        "header4.xml",
        "header5.xml",
        "footer1.xml",
        "footer2.xml",
        "footer3.xml",
        "footer4.xml",
        "footer5.xml",
        "footer6.xml",
    )
)

ASSESSMENT_CATALOG: tuple[dict[str, Any], ...] = (
    {"key": "post_breach", "title": "Post Breach Scenario Assessment", "toc_title": "Antivirus Bypass Analysis", "annexure_number": 1},
    {"key": "active_directory", "title": "Active Directory Assessment", "toc_title": "Active Directory Post Breach Security Attacks", "annexure_number": 2},
    {"key": "network_pt", "title": "Network Penetration Testing Assessment", "toc_title": "Network Penetration testing", "annexure_number": 3},
    {"key": "web_app_pt", "title": "Web Application Penetration Testing Assessment", "toc_title": "Web Application Penetration testing", "annexure_number": 4},
    {"key": "network_config_review", "title": "Network Configuration Review Assessment", "toc_title": "Network Penetration Review Assessment", "annexure_number": 5},
    {"key": "insider_threat", "title": "Insider Threat Assessment for Data Infiltration & Exfiltration", "toc_title": "Insider Threat Assessment for Data Infiltration and Exfiltration", "annexure_number": 6},
    {"key": "mobile_app_pt", "title": "Mobile Application Penetration Testing", "toc_title": "Mobile Application Penetration testing", "annexure_number": 7},
    {"key": "ad_security", "title": "Active Directory Security Assessment", "toc_title": "Active Directory Security Assessment", "annexure_number": 8},
    {"key": "soc_simulation", "title": "SOC Simulation Exercise", "toc_title": "SOC Simulation Exercise", "annexure_number": 9},
    {"key": "cloud_pt", "title": "Cloud Penetration Testing", "toc_title": "Cloud Penetration testing", "annexure_number": 10},
    {"key": "api_pt", "title": "API Penetration Testing Assessment", "toc_title": "API Penetration testing Assessment", "annexure_number": 11},
    {"key": "offensive_defensive_training", "title": "Offensive & Defensive Security Training", "toc_title": "Offensive & Defensive security Training", "annexure_number": 12},
    {"key": "source_code_review", "title": "Source Code Review", "toc_title": "Source Code Review", "annexure_number": 13},
)

CATALOG_BY_KEY: dict[str, dict[str, Any]] = {item["key"]: dict(item) for item in ASSESSMENT_CATALOG}
CATALOG_BY_NUMBER: dict[int, dict[str, Any]] = {
    int(item["annexure_number"]): dict(item) for item in ASSESSMENT_CATALOG
}


# ---------------------------------------------------------------------------
# Phase 2 extension point
# ---------------------------------------------------------------------------
@dataclass
class AssessmentItem:
    key: str
    title: str
    annexure_number: int | None = None
    page_number: int | None = None
    enabled: bool = True
    extra: dict[str, Any] = field(default_factory=dict)


def parse_annexure_order_json(raw: str | None) -> list[str] | None:
    """Parse optional annexure order from the form.

    Accepts either a JSON array of catalog keys or
    ``{"enabled": true, "order": ["post_breach", "web_app_pt", ...]}``.
    Returns ``None`` when customization is off or the order is empty (keep template default).
    """
    if not raw or not raw.strip():
        return None
    data = json.loads(raw)
    if isinstance(data, dict):
        if not data.get("enabled"):
            return None
        keys = data.get("order") or []
    elif isinstance(data, list):
        keys = data
    else:
        raise ValueError("annexure_order_json must be a JSON array or object")

    if not keys:
        return None

    seen: set[str] = set()
    ordered: list[str] = []
    for key in keys:
        k = str(key).strip()
        if not k:
            continue
        if k not in CATALOG_BY_KEY:
            raise ValueError(f"Unknown annexure key: {k}")
        if k in seen:
            raise ValueError(f"Duplicate annexure key: {k}")
        seen.add(k)
        ordered.append(k)
    return ordered or None


def parse_assessments_json(raw: str | None) -> list[AssessmentItem] | None:
    if not raw or not raw.strip():
        return None
    data = json.loads(raw)
    if not isinstance(data, list):
        raise ValueError("assessments_json must be a JSON array")
    items: list[AssessmentItem] = []
    for row in data:
        if not isinstance(row, dict):
            raise ValueError("Each assessment entry must be an object")
        items.append(
            AssessmentItem(
                key=str(row.get("key", "")),
                title=str(row.get("title", "")),
                annexure_number=row.get("annexure_number"),
                page_number=row.get("page_number"),
                enabled=bool(row.get("enabled", True)),
                extra={
                    k: v
                    for k, v in row.items()
                    if k not in {"key", "title", "annexure_number", "page_number", "enabled"}
                },
            )
        )
    return items


def _find_annexure_boundaries(doc: Document) -> list[tuple[int, int, str]]:
    """Return [(para_index, default_number, full_heading_text), ...] in document order."""
    hits: list[tuple[int, int, str]] = []
    for i, para in enumerate(doc.paragraphs):
        text = (para.text or "").strip()
        match = ANNEXURE_HEADING_RE.match(text)
        if match:
            hits.append((i, int(match.group(1)), text))
    return hits


def _find_trailer_start(doc: Document, section_start: int, section_end: int) -> int:
    for i in range(section_start, section_end):
        text = (doc.paragraphs[i].text or "").strip()
        if text == TRAILER_MARKER:
            return i
    return section_end


def _find_terms_heading_indices(doc: Document, section_start: int, section_end: int) -> list[int]:
    hits: list[int] = []
    for i in range(section_start, min(section_end, len(doc.paragraphs))):
        if (doc.paragraphs[i].text or "").strip() == TERMS_MARKER:
            hits.append(i)
    return hits


def _split_annexure13_tail(doc: Document, section_start: int, section_end: int) -> tuple[int, int, int]:
    """Return (annexure_content_end, new_terms_start, trailer_start) within annexure 13."""
    terms_hits = _find_terms_heading_indices(doc, section_start, section_end)
    trailer_start = _find_trailer_start(doc, section_start, section_end)
    first_terms = terms_hits[0] if terms_hits else trailer_start
    new_terms = terms_hits[1] if len(terms_hits) >= 2 else first_terms
    return first_terms, new_terms, trailer_start


def _replace_client_literals_in_elements(elements, company_name: str) -> None:
    replacements = {old: company_name for old in CLIENT_NAME_LITERALS}
    for el in elements:
        for node in el.iter(qn("w:t")):
            if not node.text:
                continue
            text = node.text
            for old, new in replacements.items():
                text = text.replace(old, new)
            node.text = text


def _set_table_cell_text(tc, text: str) -> None:
    paras = tc.findall(qn("w:p"))
    if not paras:
        paras = [OxmlElement("w:p")]
        tc.append(paras[0])
    _set_paragraph_text(paras[0], text)
    for extra in paras[1:]:
        _set_paragraph_text(extra, "")


def _find_assessment_pricing_table(root) -> Any:
    for tbl in root.iter(qn("w:tbl")):
        rows = tbl.findall(qn("w:tr"))
        if not rows:
            continue
        header_text = " ".join(_paragraph_text(tc) for tc in rows[0].findall(qn("w:tc")))
        if "Assessment Type" in header_text:
            return tbl
    return None


def _update_conclusion_pricing_table(elements, ordered_keys: list[str]) -> None:
    """Rebuild the Conclusion pricing table Assessment Type rows from selected annexures."""
    tbl = None
    for el in elements:
        tbl = _find_assessment_pricing_table(el)
        if tbl is not None:
            break
    if tbl is None:
        return

    rows = tbl.findall(qn("w:tr"))
    if len(rows) < 2:
        return

    template_row = copy.deepcopy(rows[1])
    for row in rows[1:-1]:
        tbl.remove(row)
    total_row = tbl.findall(qn("w:tr"))[-1]

    for slot, key in enumerate(ordered_keys, start=1):
        item = CATALOG_BY_KEY[key]
        row_copy = copy.deepcopy(template_row)
        cells = row_copy.findall(qn("w:tc"))
        qty, freq, dur = PRICING_ROW_DEFAULTS.get(key, ("", "One Time", ""))
        if len(cells) >= 1:
            _set_table_cell_text(cells[0], str(slot))
        if len(cells) >= 2:
            _set_table_cell_text(cells[1], str(item["title"]))
        if len(cells) >= 3:
            _set_table_cell_text(cells[2], qty)
        if len(cells) >= 4:
            _set_table_cell_text(cells[3], freq)
        if len(cells) >= 5:
            _set_table_cell_text(cells[4], dur)
        _strip_yellow_from_element(row_copy)
        total_row.addprevious(row_copy)


def _paragraph_text(p_element) -> str:
    parts: list[str] = []
    for node in p_element.iter(qn("w:t")):
        if node.text:
            parts.append(node.text)
    return "".join(parts).strip()


def _set_paragraph_text(p_element, text: str) -> None:
    nodes = list(p_element.iter(qn("w:t")))
    if not nodes:
        return
    nodes[0].text = text
    for node in nodes[1:]:
        node.text = ""


def _has_page_break(p_element) -> bool:
    for br in p_element.iter(qn("w:br")):
        if br.get(qn("w:type")) == "page":
            return True
    return False


def _is_page_break_only_paragraph(p_element) -> bool:
    return not _paragraph_text(p_element) and _has_page_break(p_element)


def _make_page_break_paragraph():
    p = OxmlElement("w:p")
    r = OxmlElement("w:r")
    br = OxmlElement("w:br")
    br.set(qn("w:type"), "page")
    r.append(br)
    p.append(r)
    return p


def _get_body_element_index(doc: Document, para_idx: int) -> int | None:
    """Get the body child index for a paragraph at given para_idx."""
    if para_idx < 0 or para_idx >= len(doc.paragraphs):
        return None
    target_element = doc.paragraphs[para_idx]._element
    body = doc.element.body
    for i, child in enumerate(body):
        if child is target_element:
            return i
    return None


def _extract_block_elements(doc: Document, start: int, end: int) -> list:
    """Extract all body elements (paragraphs + tables) between paragraph indices start and end."""
    if start < 0 or end <= start or start >= len(doc.paragraphs):
        return []
    
    body = doc.element.body
    start_idx = _get_body_element_index(doc, start)
    end_idx = _get_body_element_index(doc, min(end - 1, len(doc.paragraphs) - 1))
    
    if start_idx is None or end_idx is None:
        # Fallback to paragraph-only extraction
        return [copy.deepcopy(doc.paragraphs[i]._element) for i in range(start, end)]
    
    # Extract all body elements from start to end (inclusive of end)
    result = []
    for i in range(start_idx, end_idx + 1):
        result.append(copy.deepcopy(body[i]))
    return result


def _block_range_with_leading_break(doc: Document, start: int, end: int) -> tuple[int, int]:
    """Include any page-break-only paragraph before start, plus preserve tables after last para."""
    if start > 0 and _is_page_break_only_paragraph(doc.paragraphs[start - 1]._element):
        return start - 1, end
    # Check if end should include trailing tables
    if end < len(doc.paragraphs):
        body = doc.element.body
        end_body_idx = _get_body_element_index(doc, end - 1)
        if end_body_idx is not None and end_body_idx + 1 < len(body):
            # Look ahead for tables immediately after the last paragraph
            next_elem = body[end_body_idx + 1]
            if next_elem.tag == qn("w:tbl"):
                # Don't auto-extend; caller must handle this
                pass
    return start, end


def _insert_elements_before_sectpr(body, elements: list) -> None:
    sect_pr = body.find(qn("w:sectPr"))
    for el in elements:
        if sect_pr is not None:
            sect_pr.addprevious(el)
        else:
            body.append(el)


def _remove_paragraph_at(doc: Document, index: int) -> None:
    if index < 0 or index >= len(doc.paragraphs):
        return
    el = doc.paragraphs[index]._element
    parent = el.getparent()
    if parent is not None:
        parent.remove(el)


def _remove_paragraphs_from(doc: Document, start_index: int) -> None:
    for i in range(len(doc.paragraphs) - 1, start_index - 1, -1):
        _remove_paragraph_at(doc, i)


def _find_contents_region(doc: Document) -> tuple[int | None, int | None, int | None, int | None]:
    """Return (contents_heading_idx, first_assessment_toc_idx, terms_toc_idx, conclusion_toc_idx)."""
    boundaries = _find_annexure_boundaries(doc)
    first_annexure_idx = boundaries[0][0] if boundaries else len(doc.paragraphs)
    contents_idx: int | None = None
    first_assessment_idx: int | None = None
    terms_idx: int | None = None
    conclusion_idx: int | None = None

    for i in range(first_annexure_idx):
        text = (doc.paragraphs[i].text or "").strip()
        if text == CONTENTS_HEADING:
            contents_idx = i
        elif text.startswith("TEAM OVERVIEW") and "\t" in text:
            first_assessment_idx = i + 1
        elif text.startswith(TERMS_TOC_LABEL) and "\t" in text:
            terms_idx = i
        elif text.startswith(CONCLUSION_TOC_PREFIX) and "\t" in text:
            conclusion_idx = i

    if first_assessment_idx is None and contents_idx is not None:
        first_assessment_idx = contents_idx + 1 + len(STATIC_TOC_ANCHORS)

    if terms_idx is None and conclusion_idx is not None and conclusion_idx > 0:
        prev = (doc.paragraphs[conclusion_idx - 1].text or "").strip()
        if prev.startswith(TERMS_TOC_LABEL) and "\t" in prev:
            terms_idx = conclusion_idx - 1

    return contents_idx, first_assessment_idx, terms_idx, conclusion_idx


_TOC_PPR_CACHE: Any = None


def _reference_toc_ppr() -> Any:
    """Cached paragraph properties from a template Contents line (incl. dot leaders)."""
    global _TOC_PPR_CACHE
    if _TOC_PPR_CACHE is not None:
        return _TOC_PPR_CACHE
    if not BASE_TEMPLATE.is_file():
        return None
    ref = Document(str(BASE_TEMPLATE))
    _, first_assessment_idx, _, _ = _find_contents_region(ref)
    sample_idx = min(first_assessment_idx or 25, len(ref.paragraphs) - 1)
    ppr = ref.paragraphs[sample_idx]._element.find(qn("w:pPr"))
    _TOC_PPR_CACHE = copy.deepcopy(ppr) if ppr is not None else None
    return _TOC_PPR_CACHE


def _sanitize_toc_ppr(ppr) -> None:
    """Remove layout/highlight artifacts copied from the template TOC lines."""
    for tag in (qn("w:pageBreakBefore"), qn("w:keepNext"), qn("w:pageBreakAfter")):
        node = ppr.find(tag)
        if node is not None:
            ppr.remove(node)
    rpr = ppr.find(qn("w:rPr"))
    if rpr is not None:
        highlight = rpr.find(qn("w:highlight"))
        if highlight is not None:
            rpr.remove(highlight)
        if not len(rpr):
            ppr.remove(rpr)


def _apply_compact_toc_ppr(p_element, *, first_line: bool = False) -> None:
    """Apply compact Contents paragraph formatting with right-aligned dot leader tab."""
    existing = p_element.find(qn("w:pPr"))
    if existing is not None:
        p_element.remove(existing)

    ref = _reference_toc_ppr()
    if ref is not None:
        ppr = copy.deepcopy(ref)
        _sanitize_toc_ppr(ppr)
    else:
        ppr = OxmlElement("w:pPr")
        tabs = OxmlElement("w:tabs")
        tab = OxmlElement("w:tab")
        tab.set(qn("w:val"), "right")
        tab.set(qn("w:leader"), "dot")
        tab.set(qn("w:pos"), TOC_TAB_POS)
        tabs.append(tab)
        ppr.append(tabs)

    spacing = ppr.find(qn("w:spacing"))
    if spacing is None:
        spacing = OxmlElement("w:spacing")
        ppr.append(spacing)
    spacing.set(qn("w:before"), TOC_FIRST_LINE_BEFORE_TWIPS if first_line else TOC_LINE_BEFORE_TWIPS)
    spacing.set(qn("w:after"), "0")
    spacing.set(qn("w:line"), "240")
    spacing.set(qn("w:lineRule"), "auto")

    _sanitize_toc_ppr(ppr)

    # Ensure dot-leader tab stop is present.
    tabs_el = ppr.find(qn("w:tabs"))
    if tabs_el is None:
        tabs_el = OxmlElement("w:tabs")
        ppr.append(tabs_el)
    if not tabs_el.findall(qn("w:tab")):
        tab = OxmlElement("w:tab")
        tab.set(qn("w:val"), "right")
        tab.set(qn("w:leader"), "dot")
        tab.set(qn("w:pos"), TOC_TAB_POS)
        tabs_el.append(tab)

    p_element.insert(0, ppr)


def _set_toc_line(p_element, title: str, page: int, *, first_line: bool = False) -> None:
    """Rebuild a Contents line as ``Title<TAB>Page`` with compact layout."""
    _apply_compact_toc_ppr(p_element, first_line=first_line)
    for child in list(p_element):
        if child.tag != qn("w:pPr"):
            p_element.remove(child)

    r_title = OxmlElement("w:r")
    t_title = OxmlElement("w:t")
    t_title.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    t_title.text = title
    r_title.append(t_title)
    p_element.append(r_title)

    r_tab = OxmlElement("w:r")
    r_tab.append(OxmlElement("w:tab"))
    p_element.append(r_tab)

    r_page = OxmlElement("w:r")
    t_page = OxmlElement("w:t")
    t_page.text = str(page)
    r_page.append(t_page)
    p_element.append(r_page)


def _parse_toc_line(text: str) -> tuple[str, int]:
    parts = (text or "").split("\t")
    title = parts[0].strip() if parts else ""
    page_raw = parts[1].strip() if len(parts) > 1 else "1"
    page_digits = re.sub(r"[^\d]", "", page_raw)
    return title, int(page_digits) if page_digits else 1


def _find_introduction_body_idx(doc: Document) -> int | None:
    """Return paragraph index of the body INTRODUCTION heading (not the Contents line)."""
    _, _, _, conclusion_idx = _find_contents_region(doc)
    if conclusion_idx is None:
        return None
    for i in range(conclusion_idx + 1, len(doc.paragraphs)):
        text = (doc.paragraphs[i].text or "").strip()
        if text == "INTRODUCTION" and "\t" not in text:
            return i
    return None


def _ensure_page_break_before_introduction(doc: Document) -> None:
    """Start the Introduction section on the page after Contents."""
    intro_idx = _find_introduction_body_idx(doc)
    if intro_idx is None:
        return

    if intro_idx > 0:
        prev_el = doc.paragraphs[intro_idx - 1]._element
        if _is_page_break_only_paragraph(prev_el) or _has_page_break(prev_el):
            return

    intro_el = doc.paragraphs[intro_idx]._element
    ppr = intro_el.find(qn("w:pPr"))
    if ppr is not None and ppr.find(qn("w:pageBreakBefore")) is not None:
        return

    _insert_paragraph_before(doc, intro_idx, _make_page_break_paragraph())


def _normalize_contents_section(doc: Document) -> None:
    """Compact Contents layout: tight spacing, no page breaks, single-page TOC block."""
    contents_idx, first_assessment_idx, _, conclusion_idx = _find_contents_region(doc)
    if contents_idx is None or conclusion_idx is None:
        return

    # Drop blank lines and page-break-only paragraphs inside the Contents block.
    for i in range(conclusion_idx - 1, contents_idx, -1):
        if i >= len(doc.paragraphs):
            continue
        para = doc.paragraphs[i]
        if _is_page_break_only_paragraph(para._element):
            _remove_paragraph_at(doc, i)
            continue
        txt = (para.text or "").strip()
        if txt:
            continue
        if para._element.findall(f".//{PIC_TAG}"):
            continue
        _remove_paragraph_at(doc, i)

    _, first_assessment_idx, _, conclusion_idx = _find_contents_region(doc)
    if conclusion_idx is None:
        return

    for i in range(contents_idx + 1, conclusion_idx + 1):
        if i >= len(doc.paragraphs):
            break
        p_element = doc.paragraphs[i]._element
        for br in list(p_element.iter(qn("w:br"))):
            if br.get(qn("w:type")) == "page":
                parent = br.getparent()
                if parent is not None:
                    parent.remove(br)
        title, page_num = _parse_toc_line(doc.paragraphs[i].text or "")
        if title:
            _set_toc_line(
                p_element,
                title,
                page_num,
                first_line=(i == contents_idx + 1),
            )


def _strip_yellow_from_element(element) -> int:
    removed = 0
    for rpr in element.iter(qn("w:rPr")):
        highlight = rpr.find(qn("w:highlight"))
        if highlight is None:
            continue
        val = (highlight.get(qn("w:val")) or "").lower()
        if val not in ("", "none", "white"):
            rpr.remove(highlight)
            removed += 1
    return removed


def _strip_all_yellow_highlights(doc: Document) -> int:
    """Remove all visible highlights from the document (incl. tracked-change runs)."""
    removed = _strip_yellow_from_element(doc.element.body)
    for section in doc.sections:
        for hf in (
            section.header,
            section.footer,
            section.first_page_header,
            section.first_page_footer,
        ):
            if hf is not None and hasattr(hf, "_element"):
                removed += _strip_yellow_from_element(hf._element)
    return removed


def _finalize_document(docx_path: Path, ordered_keys: list[str] | None = None) -> None:
    """Apply final layout fixes and refresh all Contents page numbers."""
    if not ordered_keys:
        ordered_keys = [item["key"] for item in ASSESSMENT_CATALOG]

    doc = Document(str(docx_path))
    _normalize_contents_section(doc)
    _ensure_page_break_before_introduction(doc)

    pages = _compute_layout_pages(doc)
    _refresh_all_toc_pages(doc, ordered_keys, pages)
    _normalize_contents_section(doc)

    _strip_all_yellow_highlights(doc)
    doc.save(str(docx_path))
    _set_update_fields_on_open(docx_path)


def _clone_toc_paragraph(doc: Document) -> OxmlElement:
    """Create a new compact Contents line paragraph."""
    new_p = OxmlElement("w:p")
    _set_toc_line(new_p, "", 1)
    return new_p


def _insert_paragraph_before(doc: Document, index: int, p_element) -> None:
    anchor = doc.paragraphs[index]._element
    anchor.addprevious(p_element)


def _paragraph_lines(p_element) -> float:
    """Estimate how many printed lines a paragraph consumes."""
    text = _paragraph_text(p_element)
    lines = max(0.35, len(text) / CHARS_PER_LINE) if text else 0.35
    ppr = p_element.find(qn("w:pPr"))
    if ppr is not None:
        sp = ppr.find(qn("w:spacing"))
        if sp is not None:
            before = int(sp.get(qn("w:before")) or 0)
            after = int(sp.get(qn("w:after")) or 0)
            lines += (before + after) / 240.0
    return lines


def _table_lines(tbl_element) -> float:
    """Estimate printed lines consumed by a table."""
    total = 0.75
    for tr in tbl_element.findall(qn("w:tr")):
        row_lines = 1.0
        for tc in tr.findall(qn("w:tc")):
            cell_text = "".join(t.text or "" for t in tc.iter(qn("w:t")))
            row_lines = max(row_lines, max(1.0, len(cell_text) / CHARS_PER_LINE))
        total += row_lines
    return total


def _compute_layout_pages(doc: Document) -> dict[int, int]:
    """Estimate page numbers for each ``doc.paragraphs`` index from final layout."""
    pages: dict[int, int] = {}
    current_page = 1
    lines_used = 0.0

    for i, para in enumerate(doc.paragraphs):
        el = para._element
        pages[i] = current_page

        ppr = el.find(qn("w:pPr"))
        if ppr is not None and ppr.find(qn("w:pageBreakBefore")) is not None:
            current_page += 1
            lines_used = 0.0
            pages[i] = current_page

        if _is_page_break_only_paragraph(el):
            current_page += 1
            lines_used = 0.0
            continue

        lines = _paragraph_lines(el)
        if lines_used + lines > LINES_PER_PAGE:
            current_page += 1
            lines_used = lines
        else:
            lines_used += lines

        if _has_page_break(el):
            current_page += 1
            lines_used = 0.0

    return pages


def _compute_paragraph_pages(doc: Document, paras_per_page: float = PARAS_PER_PAGE) -> dict[int, int]:
    """Estimate page number for each paragraph index."""
    del paras_per_page  # layout model supersedes the old paragraph counter
    return _compute_layout_pages(doc)


def _find_body_heading_idx(
    doc: Document,
    label_prefix: str,
    *,
    min_index: int = 0,
    exact: bool = False,
) -> int | None:
    prefix = label_prefix.strip().lower()
    for i, para in enumerate(doc.paragraphs):
        if i < min_index:
            continue
        text = (para.text or "").strip()
        if not text or "\t" in text:
            continue
        lower = text.lower()
        if exact:
            if lower == prefix:
                return i
        elif lower == prefix or lower.startswith(prefix):
            return i
    return None


def _find_body_anchor_page(
    doc: Document,
    label_prefix: str,
    pages: dict[int, int],
    min_index: int = 30,
    *,
    exact: bool = False,
) -> int:
    idx = _find_body_heading_idx(doc, label_prefix, min_index=min_index, exact=exact)
    if idx is not None:
        return pages.get(idx, 1)
    return 1


def _find_terms_body_page(doc: Document, pages: dict[int, int]) -> int:
    boundaries = _find_annexure_boundaries(doc)
    search_from = boundaries[-1][0] if boundaries else 0
    for i, para in enumerate(doc.paragraphs):
        if i < search_from:
            continue
        text = (para.text or "").strip()
        if text == TERMS_MARKER:
            return pages.get(i, 1)
        if text == TRAILER_MARKER:
            break
    return _find_body_anchor_page(doc, TERMS_TOC_LABEL, pages)


def _find_trailer_page(doc: Document, pages: dict[int, int]) -> int:
    for i, para in enumerate(doc.paragraphs):
        if (para.text or "").strip() == TRAILER_MARKER:
            return pages.get(i, 1)
    return _find_body_anchor_page(doc, CONCLUSION_TOC_PREFIX, pages)


def _find_annexure_heading_pages(doc: Document, pages: dict[int, int]) -> dict[int, int]:
    """Map annexure slot number -> page (Annexure 1, Annexure 2, ...)."""
    slot_pages: dict[int, int] = {}
    for i, para in enumerate(doc.paragraphs):
        text = (para.text or "").strip()
        match = re.match(r"^Annexure\s+(\d+)\s*:", text, re.IGNORECASE)
        if match:
            slot_pages[int(match.group(1))] = pages.get(i, 1)
    return slot_pages


def _update_contents_toc_structure(doc: Document, ordered_keys: list[str]) -> None:
    """Rebuild assessment lines in Contents (titles only; pages filled later)."""
    contents_idx, first_assessment_idx, _, conclusion_idx = _find_contents_region(doc)
    if contents_idx is None or conclusion_idx is None or first_assessment_idx is None:
        return

    for i in range(conclusion_idx - 1, first_assessment_idx - 1, -1):
        _remove_paragraph_at(doc, i)

    _, first_assessment_idx, _, conclusion_idx = _find_contents_region(doc)
    if conclusion_idx is None or first_assessment_idx is None:
        return

    insert_at = first_assessment_idx
    for key in ordered_keys:
        item = CATALOG_BY_KEY[key]
        toc_line = _clone_toc_paragraph(doc)
        _set_toc_line(toc_line, str(item["toc_title"]), 1)
        _insert_paragraph_before(doc, insert_at, toc_line)
        insert_at += 1

    terms_line = _clone_toc_paragraph(doc)
    _set_toc_line(terms_line, TERMS_TOC_LABEL, 1)
    _insert_paragraph_before(doc, insert_at, terms_line)


def _refresh_all_toc_pages(doc: Document, ordered_keys: list[str], pages: dict[int, int]) -> None:
    """Update every Contents line page number from the final document layout."""
    contents_idx, first_assessment_idx, terms_idx, conclusion_idx = _find_contents_region(doc)
    if contents_idx is None:
        return

    body_min = (conclusion_idx if conclusion_idx is not None else contents_idx or 0) + 1
    intro_idx = _find_introduction_body_idx(doc)
    intro_page = int(pages.get(intro_idx, 1)) if intro_idx is not None else 1

    for offset, (label, _anchor) in enumerate(STATIC_TOC_ANCHORS):
        toc_idx = contents_idx + 1 + offset
        if toc_idx >= len(doc.paragraphs):
            break
        page = int(_find_body_anchor_page(doc, label, pages, min_index=body_min, exact=True))
        if page <= 1 and intro_page > 1:
            page = intro_page
        _set_toc_line(
            doc.paragraphs[toc_idx]._element,
            label,
            page,
            first_line=(offset == 0),
        )

    if first_assessment_idx is None or conclusion_idx is None:
        return

    annexure_pages = _find_annexure_heading_pages(doc, pages)
    for slot, key in enumerate(ordered_keys, start=1):
        toc_idx = first_assessment_idx + slot - 1
        if terms_idx is not None and toc_idx >= terms_idx:
            break
        if toc_idx >= conclusion_idx:
            break
        item = CATALOG_BY_KEY[key]
        page = int(annexure_pages.get(slot, DEFAULT_FIRST_ANNEXURE_PAGE))
        _set_toc_line(doc.paragraphs[toc_idx]._element, str(item["toc_title"]), page)

    _, _, terms_idx, conclusion_idx = _find_contents_region(doc)
    if terms_idx is not None:
        terms_page = int(_find_terms_body_page(doc, pages))
        _set_toc_line(doc.paragraphs[terms_idx]._element, TERMS_TOC_LABEL, terms_page)
    if conclusion_idx is not None:
        conclusion_page = int(_find_trailer_page(doc, pages))
        _set_toc_line(
            doc.paragraphs[conclusion_idx]._element,
            CONCLUSION_TOC_PREFIX,
            conclusion_page,
        )


def _update_contents_toc(doc: Document, ordered_keys: list[str], pages: dict[int, int] | None = None) -> None:
    """Rebuild Contents assessment entries (page numbers refreshed in finalize)."""
    del pages
    _update_contents_toc_structure(doc, ordered_keys)


def _set_update_fields_on_open(docx_path: Path) -> None:
    """Ask Word to refresh fields/page numbers when the document opens."""
    settings_name = "word/settings.xml"
    try:
        with zipfile.ZipFile(docx_path, "r") as zin:
            if settings_name not in zin.namelist():
                return
            entries = {n: zin.read(n) for n in zin.namelist()}
    except (zipfile.BadZipFile, OSError):
        return

    settings = entries[settings_name].decode("utf-8")
    if "updateFields" in settings:
        if 'w:val="true"' not in settings:
            settings = re.sub(
                r"<w:updateFields[^/]*/>",
                '<w:updateFields w:val="true"/>',
                settings,
                count=1,
            )
    else:
        settings = settings.replace(
            "</w:settings>",
            '<w:updateFields w:val="true"/></w:settings>',
        )
    entries[settings_name] = settings.encode("utf-8")
    tmp = docx_path.with_suffix(".tmp.docx")
    with zipfile.ZipFile(tmp, "w", compression=zipfile.ZIP_DEFLATED) as zout:
        for name, data in entries.items():
            zout.writestr(name, data)
    shutil.move(str(tmp), str(docx_path))


def apply_annexure_order(
    docx_path: Path,
    ordered_keys: list[str] | None,
    *,
    company_name: str = "Client",
) -> None:
    """Reorder / subset annexure sections, append Terms & Conclusion, rebuild Contents."""
    if not ordered_keys:
        return

    doc = Document(str(docx_path))
    boundaries = _find_annexure_boundaries(doc)
    if not boundaries:
        raise ValueError("No annexure headings found in the generated document.")

    first_idx = boundaries[0][0]
    section_ranges: dict[int, tuple[int, int]] = {}
    for j, (start, default_num, _heading) in enumerate(boundaries):
        end = boundaries[j + 1][0] if j + 1 < len(boundaries) else len(doc.paragraphs)
        section_ranges[default_num] = (start, end)

    if 13 not in section_ranges:
        raise ValueError("Template is missing Annexure 13 block required for document trailer.")

    sec13_start, sec13_end = section_ranges[13]
    annexure13_end, new_terms_start, trailer_start = _split_annexure13_tail(doc, sec13_start, sec13_end)

    blocks: dict[int, list] = {}
    for num, (start, end) in section_ranges.items():
        block_end = annexure13_end if num == 13 else end
        bs, be = _block_range_with_leading_break(doc, start, block_end)
        blocks[num] = _extract_block_elements(doc, bs, be)
        if not blocks[num] and num != 13:
            print(f"Warning: Annexure {num} extracted no elements from paragraph {bs} to {be}")

    terms_bs, _ = _block_range_with_leading_break(doc, new_terms_start, trailer_start)
    terms_elements = _extract_block_elements(doc, terms_bs, trailer_start)

    trailer_start_full = trailer_start
    if trailer_start_full > 0 and _is_page_break_only_paragraph(doc.paragraphs[trailer_start_full - 1]._element):
        trailer_start_full -= 1
    trailer_elements = _extract_block_elements(doc, trailer_start_full, sec13_end)

    body = doc.element.body
    body_children = list(body)
    start_body_idx = _get_body_element_index(doc, first_idx)
    if start_body_idx is not None:
        for i in range(len(body_children) - 1, start_body_idx - 1, -1):
            body.remove(body_children[i])
    else:
        _remove_paragraphs_from(doc, first_idx)
    body = doc.element.body

    for slot, key in enumerate(ordered_keys, start=1):
        item = CATALOG_BY_KEY[key]
        default_num = int(item["annexure_number"])
        block = blocks.get(default_num)
        if not block:
            raise ValueError(f"Missing template section for annexure {default_num} ({key}).")

        cloned: list = []
        cloned.append(_make_page_break_paragraph())

        for el in block:
            el_copy = copy.deepcopy(el)
            if el_copy.tag == qn("w:p") and _is_page_break_only_paragraph(el_copy):
                continue
            if el_copy.tag == qn("w:p"):
                text = _paragraph_text(el_copy)
                if text.startswith("Annexure "):
                    _set_paragraph_text(el_copy, f"Annexure {slot}: {item['title']}")
                    _strip_yellow_from_element(el_copy)
            cloned.append(el_copy)

        _insert_elements_before_sectpr(body, cloned)

    if terms_elements:
        terms_clone = [_make_page_break_paragraph()]
        for el in terms_elements:
            el_copy = copy.deepcopy(el)
            if el_copy.tag == qn("w:p") and _is_page_break_only_paragraph(el_copy):
                continue
            terms_clone.append(el_copy)
        _replace_client_literals_in_elements(terms_clone, company_name)
        for el in terms_clone:
            _strip_yellow_from_element(el)
        _insert_elements_before_sectpr(body, terms_clone)

    trailer_clone = [_make_page_break_paragraph()]
    for el in trailer_elements:
        el_copy = copy.deepcopy(el)
        if el_copy.tag == qn("w:p") and _is_page_break_only_paragraph(el_copy):
            continue
        trailer_clone.append(el_copy)
    _update_conclusion_pricing_table(trailer_clone, ordered_keys)
    _replace_client_literals_in_elements(trailer_clone, company_name)
    for el in trailer_clone:
        _strip_yellow_from_element(el)
    _insert_elements_before_sectpr(body, trailer_clone)

    _update_contents_toc(doc, ordered_keys, {})
    doc.save(str(docx_path))


def apply_assessments(
    docx_path: Path,
    assessments: list[AssessmentItem] | None,
    *,
    company_name: str = "Client",
) -> None:
    """Legacy hook — prefer ``apply_annexure_order`` via ``annexure_order`` keys."""
    if not assessments:
        return
    keys = [a.key for a in assessments if a.enabled and a.key in CATALOG_BY_KEY]
    if keys:
        apply_annexure_order(docx_path, keys, company_name=company_name)


# ---------------------------------------------------------------------------
# Template builder (embedded — no separate _build_proposal_template.py)
# ---------------------------------------------------------------------------
def _replace_literals_in_xml(xml_text: str) -> str:
    def repl_wt(match: re.Match[str]) -> str:
        opening, content, closing = match.groups()
        for old, new in TEXT_REPLACEMENTS:
            content = content.replace(old, new)
        return opening + content + closing

    return re.sub(
        r'(<w:t(?:\s+xml:space="preserve")?>)(.*?)(</w:t>)',
        repl_wt,
        xml_text,
        flags=re.DOTALL,
    )


def _process_xml_part(data: bytes) -> bytes:
    text = _replace_literals_in_xml(data.decode("utf-8"))
    # Keep yellow highlights to identify placeholders for post-processing
    return text.encode("utf-8")


def _find_cover_logo_paragraph(doc: Document):
    """Return the cover paragraph that holds the client logo slot."""
    seen_heading = False
    for para in doc.paragraphs:
        txt = (para.text or "").strip()
        if "Statement of Work" in txt:
            seen_heading = True
            continue
        if not seen_heading:
            continue
        if "{{ client_logo }}" in (para.text or ""):
            return para
        if para._element.findall(f".//{PIC_TAG}") and not txt:
            return para
    return None


def _insert_logo_on_cover(doc: Document, logo_path: Path) -> bool:
    """Replace the cover client-logo slot with the uploaded image."""
    if not logo_path.is_file():
        return False
    logo_para = _find_cover_logo_paragraph(doc)
    if logo_para is None:
        return False
    el = logo_para._element
    for child in list(el):
        if child.tag != qn("w:pPr"):
            el.remove(child)
    logo_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    try:
        width = calculate_optimal_logo_size(logo_path)
        logo_para.add_run().add_picture(str(logo_path), width=Inches(width))
        return True
    except Exception:
        traceback.print_exc()
        return False


def _insert_client_logo_placeholder(docx_path: Path) -> None:
    """Insert docxtpl tag on the cover logo paragraph (used when building base template)."""
    doc = Document(str(docx_path))
    logo_para = _find_cover_logo_paragraph(doc)
    if logo_para is None:
        raise RuntimeError("Cover client-logo paragraph not found.")
    el = logo_para._element
    for child in list(el):
        if child.tag != qn("w:pPr"):
            el.remove(child)
    logo_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    logo_para.add_run("{{ client_logo }}")
    doc.save(str(docx_path))


def _ensure_client_logo_placeholder_in_template() -> None:
    """Make sure the base template contains a docxtpl logo placeholder."""
    if not BASE_TEMPLATE.is_file():
        return
    try:
        with zipfile.ZipFile(BASE_TEMPLATE, "r") as zf:
            xml = zf.read("word/document.xml")
        if b"client_logo" in xml:
            return
        _insert_client_logo_placeholder(BASE_TEMPLATE)
    except Exception:
        traceback.print_exc()


def _replace_highlighted_placeholder_text(doc: Document, old_text: str, new_text: str) -> None:
    """Replace yellow-highlighted placeholder text and remove the highlight."""
    for paragraph in doc.paragraphs:
        for run in paragraph.runs:
            if not run.text:
                continue
            # Check if run has yellow highlight
            has_yellow = False
            if run._element.get_or_add_rPr() is not None:
                highlight = run._element.rPr.find(qn("w:highlight"))
                if highlight is not None and highlight.get(qn("w:val")) == "yellow":
                    has_yellow = True
            if has_yellow and old_text in run.text:
                run.text = run.text.replace(old_text, new_text)
                # Remove the yellow highlight after replacement
                rPr = run._element.rPr
                highlight = rPr.find(qn("w:highlight"))
                if highlight is not None:
                    rPr.remove(highlight)
    
    # Also check in tables
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    for run in para.runs:
                        if not run.text:
                            continue
                        has_yellow = False
                        if run._element.get_or_add_rPr() is not None:
                            highlight = run._element.rPr.find(qn("w:highlight"))
                            if highlight is not None and highlight.get(qn("w:val")) == "yellow":
                                has_yellow = True
                        if has_yellow and old_text in run.text:
                            run.text = run.text.replace(old_text, new_text)
                            # Remove the yellow highlight after replacement
                            rPr = run._element.rPr
                            highlight = rPr.find(qn("w:highlight"))
                            if highlight is not None:
                                rPr.remove(highlight)


def _replace_highlighted_placeholder_image(doc: Document, logo_path: Path) -> bool:
    """Replace yellow-highlighted image placeholder with new logo. Returns True if replaced."""
    replaced = False
    for paragraph in doc.paragraphs:
        # Check if this paragraph is inside a yellow-highlighted run
        has_highlight = False
        for run in paragraph.runs:
            if run._element.get_or_add_rPr() is not None:
                highlight = run._element.rPr.find(qn("w:highlight"))
                if highlight is not None and highlight.get(qn("w:val")) == "yellow":
                    has_highlight = True
                    break
        
        # Check if paragraph has any images
        p_elem = paragraph._element
        has_drawing = False
        drawing_elements = []
        
        # Look for all drawing elements (both inline and anchored)
        for drawing in p_elem.iter(qn("wp:anchor")):  # anchored images
            drawing_elements.append(drawing)
            has_drawing = True
        for drawing in p_elem.iter(qn("wp:inline")):  # inline images
            drawing_elements.append(drawing)
            has_drawing = True
        
        # If paragraph has yellow highlight and images, replace the image
        if has_highlight and has_drawing and logo_path.is_file():
            # Remove all existing drawings
            for drawing in drawing_elements:
                parent = drawing.getparent()
                if parent is not None:
                    parent.remove(drawing)
            
            # Remove all runs and add new picture
            paragraph.clear()
            try:
                width_inches = calculate_optimal_logo_size(logo_path)
                run = paragraph.add_run()
                picture = run.add_picture(str(logo_path), width=Inches(width_inches))
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                
                # Remove yellow highlight from the new run
                if run._element.get_or_add_rPr() is not None:
                    rPr = run._element.rPr
                    highlight = rPr.find(qn("w:highlight"))
                    if highlight is not None:
                        rPr.remove(highlight)
                
                replaced = True
            except Exception as e:
                traceback.print_exc()
    
    return replaced


def build_base_template(
    source: Path = SOURCE_TEMPLATE,
    output: Path = BASE_TEMPLATE,
) -> Path:
    """Convert source DOCX placeholders to docxtpl tags."""
    if not source.is_file():
        raise FileNotFoundError(f"Source template not found: {source}")
    shutil.copy2(source, output)
    with zipfile.ZipFile(output, "r") as zin:
        entries = {name: zin.read(name) for name in zin.namelist()}
    for part in XML_PARTS:
        if part in entries:
            entries[part] = _process_xml_part(entries[part])
    with zipfile.ZipFile(output, "w", compression=zipfile.ZIP_DEFLATED) as zout:
        for name, data in entries.items():
            zout.writestr(name, data)
    _insert_client_logo_placeholder(output)
    return output


def ensure_base_template() -> Path:
    if not BASE_TEMPLATE.is_file():
        if not SOURCE_TEMPLATE.is_file():
            raise FileNotFoundError(
                f"Missing {BASE_TEMPLATE.name} and {SOURCE_TEMPLATE.name}. "
                "Place CSS_Proposal_Source.docx next to this script and run "
                "`python css_proposal_generator.py --build-template`."
            )
        build_base_template()
    else:
        _ensure_client_logo_placeholder_in_template()
    return BASE_TEMPLATE


# ---------------------------------------------------------------------------
# Proposal generation
# ---------------------------------------------------------------------------
def default_proposal_ref(dt: datetime | None = None) -> str:
    dt = dt or datetime.now()
    return dt.strftime("%d/%m/%Y")


def calculate_optimal_logo_size(
    image_path: str | Path,
    max_width: float = 3.0,
    max_height: float = 1.5,
) -> float:
    if not PIL_AVAILABLE:
        return max_width
    try:
        with Image.open(image_path) as img:
            w, h = img.size
        w_in, h_in = w / 96, h / 96
        if w_in <= max_width and h_in <= max_height:
            return max(1.5, w_in)
        aspect = w / h
        optimal = min(max_height * aspect, max_width)
        return max(1.5, optimal)
    except Exception:
        return max_width


def _resolve_logo_inline(tpl: DocxTemplate, logo_path: str | Path | None) -> InlineImage | None:
    try:
        path = Path(logo_path) if logo_path else None
        if path and path.is_file():
            return InlineImage(tpl, str(path), width=Inches(calculate_optimal_logo_size(path)))
        if DEFAULT_LOGO.is_file():
            return InlineImage(tpl, str(DEFAULT_LOGO), width=Inches(calculate_optimal_logo_size(DEFAULT_LOGO)))
    except Exception:
        traceback.print_exc()
    return None


def generate_proposal(
    *,
    company_name: str,
    proposal_ref: str,
    logo_path: str | Path | None,
    output_path: str | Path,
    annexure_order: list[str] | None = None,
    assessments: list[AssessmentItem] | None = None,
    template_path: str | Path | None = None,
) -> Path:
    company_name = (company_name or "").strip() or "Client"
    proposal_ref = (proposal_ref or "").strip() or default_proposal_ref()
    template = Path(template_path) if template_path else ensure_base_template()
    output = Path(output_path)

    tpl = DocxTemplate(str(template))
    context: dict[str, Any] = {
        "client_name": company_name,
        "proposal_ref": proposal_ref,
        "client_logo": _resolve_logo_inline(tpl, logo_path),
        "report_date": datetime.now().strftime("%d %B %Y"),
    }
    try:
        tpl.render(context)
    except Exception as exc:
        raise RuntimeError(f"Template render failed: {exc}") from exc
    output.parent.mkdir(parents=True, exist_ok=True)
    tpl.save(str(output))

    logo_file: Path | None = None
    if logo_path:
        candidate = Path(logo_path)
        if candidate.is_file():
            logo_file = candidate

    # Post-processing: replace placeholders and insert cover logo.
    try:
        doc = Document(str(output))
        _replace_highlighted_placeholder_text(doc, "XYZ LLC", company_name)
        _replace_highlighted_placeholder_text(doc, "Company name", company_name)
        _replace_highlighted_placeholder_text(doc, "EN-01-12-25", proposal_ref)

        if logo_file:
            if not _insert_logo_on_cover(doc, logo_file):
                _replace_highlighted_placeholder_image(doc, logo_file)

        doc.save(str(output))
    except Exception:
        traceback.print_exc()
    
    try:
        ordered_keys = annexure_order or [item["key"] for item in ASSESSMENT_CATALOG]
        apply_annexure_order(output, ordered_keys, company_name=company_name)
    except Exception as exc:
        raise RuntimeError(f"Annexure layout failed: {exc}") from exc

    _finalize_document(output, ordered_keys)
    return output


def _safe_filename(name: str, ext: str = ".docx") -> str:
    stem = re.sub(r"[^\w\-]+", "_", name.strip())[:80] or "Proposal"
    return f"{stem}{ext}"


# ---------------------------------------------------------------------------
# FastAPI app (self-hosted — not mounted in app.py)
# ---------------------------------------------------------------------------
app = FastAPI(title="CSS Proposal Generator", version=VERSION)
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)
UPLOAD_DIR.mkdir(parents=True, exist_ok=True)


@app.get("/catalog/")
async def annexure_catalog():
    return list(ASSESSMENT_CATALOG)


@app.get("/", response_class=HTMLResponse)
async def proposal_form() -> HTMLResponse:
    today_ref = default_proposal_ref()
    catalog_json = json.dumps(list(ASSESSMENT_CATALOG))
    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>CSS Proposal Generator</title>
  <style>
    body {{ font-family: Arial, sans-serif; background: #1a1a2e; color: #fff; margin: 0; padding: 20px; }}
    .container {{ max-width: 820px; margin: 0 auto; background: #16213e; border-radius: 8px; padding: 24px; }}
    h1 {{ color: #8a2be2; text-align: center; margin-top: 0; }}
    h2 {{ color: #b388ff; font-size: 1.05rem; margin: 24px 0 12px; }}
    .form-group {{ margin-bottom: 16px; }}
    label {{ display: block; margin-bottom: 6px; font-weight: bold; }}
    input[type="text"], input[type="file"], select {{ width: 100%; box-sizing: border-box; }}
    input[type="text"], select {{
      padding: 10px; border: 1px solid #444; border-radius: 4px; background: #0f3460; color: #fff;
    }}
    select option {{ background: #0f3460; color: #fff; }}
    button {{
      background: #8a2be2; color: #fff; border: none; padding: 10px 16px; border-radius: 4px;
      cursor: pointer; font-size: 15px;
    }}
    button:hover {{ background: #7a1fd2; }}
    button.secondary {{ background: #394867; }}
    button.secondary:hover {{ background: #4b5875; }}
    button.danger {{ background: #7a2230; padding: 6px 10px; font-size: 13px; }}
    button.danger:hover {{ background: #922635; }}
    .submit-btn {{ width: 100%; padding: 12px 24px; font-size: 16px; margin-top: 8px; }}
    .hint {{ color: #aaa; font-size: 0.9em; margin-top: 4px; }}
    #status {{ margin-top: 16px; min-height: 1.2em; }}
    .error {{ color: #ff6b6b; }}
    .success {{ color: #69db7c; }}
    .annexure-panel {{
      border: 1px solid #394867; border-radius: 6px; padding: 16px; background: #0f2744;
    }}
    .annexure-row {{
      display: grid; grid-template-columns: 120px 1fr auto; gap: 10px; align-items: center;
      margin-bottom: 10px;
    }}
    .annexure-row label {{ margin: 0; font-weight: 600; }}
    .annexure-actions {{ display: flex; gap: 8px; flex-wrap: wrap; margin-top: 10px; }}
    .checkbox-row {{ display: flex; align-items: center; gap: 10px; }}
    .checkbox-row input {{ width: auto; }}
    .checkbox-row label {{ margin: 0; font-weight: normal; }}
    .hidden {{ display: none; }}
  </style>
</head>
<body>
  <div class="container">
    <h1>CSS Proposal Generator</h1>
    <p class="hint">Company name, proposal reference, cover logo, and optional annexure selection/order.</p>
    <form id="proposalForm">
      <div class="form-group">
        <label for="company_name">Company name</label>
        <input type="text" id="company_name" name="company_name" required placeholder="Acme Corp" />
      </div>
      <div class="form-group">
        <label for="proposal_ref">Proposal reference (date: DD/MM/YYYY)</label>
        <input type="text" id="proposal_ref" name="proposal_ref" value="{today_ref}" placeholder="23/05/2026" />
        <p class="hint">This date will replace any placeholders in the template.</p>
      </div>
      <div class="form-group">
        <label for="client_logo">Client logo (cover page)</label>
        <input type="file" id="client_logo" name="client_logo" accept="image/png,image/jpeg,image/jpg,image/webp" />
        <p class="hint">Upload your company logo. It will appear on the cover page.</p>
      </div>

      <h2>Annexures (optional)</h2>
      <div class="annexure-panel">
        <div class="checkbox-row form-group">
          <input type="checkbox" id="customize_annexures" />
          <label for="customize_annexures">Customize which annexures to include and their order</label>
        </div>
        <p class="hint">Leave unchecked to keep all 13 annexures in the template default order.</p>
        <div id="annexureBuilder" class="hidden">
          <p class="hint">Pick Annexure 1, then Annexure 2 from the remaining list, and so on. Add only as many as you need.</p>
          <div id="annexureSlots"></div>
          <div class="annexure-actions">
            <button type="button" class="secondary" id="addAnnexureBtn">+ Add annexure slot</button>
            <button type="button" class="secondary" id="clearAnnexuresBtn">Clear all</button>
          </div>
        </div>
      </div>

      <button type="submit" class="submit-btn">Generate Proposal</button>
    </form>
    <div id="status"></div>
  </div>
  <script>
    const CATALOG = {catalog_json};
    const slotsEl = document.getElementById('annexureSlots');
    const customizeEl = document.getElementById('customize_annexures');
    const builderEl = document.getElementById('annexureBuilder');
    let slotSelections = [];

    function selectedKeys() {{
      return slotSelections.filter(Boolean);
    }}

    function availableForSlot(slotIndex) {{
      const used = new Set(selectedKeys());
      const current = slotSelections[slotIndex];
      if (current) used.delete(current);
      return CATALOG.filter(item => !used.has(item.key));
    }}

    function renderSlots() {{
      slotsEl.innerHTML = '';
      slotSelections.forEach((selected, index) => {{
        const row = document.createElement('div');
        row.className = 'annexure-row';

        const label = document.createElement('label');
        label.textContent = `Annexure ${{index + 1}}`;

        const select = document.createElement('select');
        select.dataset.index = String(index);
        const empty = document.createElement('option');
        empty.value = '';
        empty.textContent = '-- Select assessment --';
        select.appendChild(empty);

        availableForSlot(index).forEach(item => {{
          const opt = document.createElement('option');
          opt.value = item.key;
          opt.textContent = item.title;
          if (item.key === selected) opt.selected = true;
          select.appendChild(opt);
        }});

        select.addEventListener('change', (e) => {{
          slotSelections[index] = e.target.value || '';
          renderSlots();
        }});

        const removeBtn = document.createElement('button');
        removeBtn.type = 'button';
        removeBtn.className = 'danger';
        removeBtn.textContent = 'Remove';
        removeBtn.addEventListener('click', () => {{
          slotSelections.splice(index, 1);
          renderSlots();
        }});

        row.appendChild(label);
        row.appendChild(select);
        row.appendChild(removeBtn);
        slotsEl.appendChild(row);
      }});
    }}

    function addSlot() {{
      if (slotSelections.length >= CATALOG.length) return;
      if (selectedKeys().length >= CATALOG.length) return;
      slotSelections.push('');
      renderSlots();
    }}

    customizeEl.addEventListener('change', () => {{
      builderEl.classList.toggle('hidden', !customizeEl.checked);
      if (customizeEl.checked && slotSelections.length === 0) addSlot();
    }});

    document.getElementById('addAnnexureBtn').addEventListener('click', addSlot);
    document.getElementById('clearAnnexuresBtn').addEventListener('click', () => {{
      slotSelections = [];
      renderSlots();
    }});

    function buildAnnexurePayload() {{
      if (!customizeEl.checked) return null;
      const order = selectedKeys();
      if (!order.length) return null;
      return {{ enabled: true, order }};
    }}

    document.getElementById('proposalForm').addEventListener('submit', async (e) => {{
      e.preventDefault();
      const status = document.getElementById('status');
      status.textContent = 'Generating...';
      status.className = '';

      if (customizeEl.checked) {{
        const order = selectedKeys();
        if (order.length !== slotSelections.filter(s => s !== undefined).length) {{
          status.textContent = 'Please select an assessment for each annexure slot, or remove empty slots.';
          status.className = 'error';
          return;
        }}
        if (!order.length) {{
          status.textContent = 'Add at least one annexure or uncheck customization to use the default 13.';
          status.className = 'error';
          return;
        }}
      }}

      const formData = new FormData();
      formData.append('company_name', document.getElementById('company_name').value);
      formData.append('proposal_ref', document.getElementById('proposal_ref').value);
      const logo = document.getElementById('client_logo').files[0];
      if (logo) formData.append('client_logo', logo);

      const annexurePayload = buildAnnexurePayload();
      if (annexurePayload) {{
        formData.append('annexure_order_json', JSON.stringify(annexurePayload));
      }}

      try {{
        const resp = await fetch('generate/', {{ method: 'POST', body: formData }});
        if (!resp.ok) {{
          const err = await resp.json().catch(() => ({{ detail: resp.statusText }}));
          throw new Error(err.detail || 'Generation failed');
        }}
        const blob = await resp.blob();
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = 'CSS_Proposal.docx';
        document.body.appendChild(a);
        a.click();
        a.remove();
        URL.revokeObjectURL(url);
        status.textContent = 'Proposal downloaded successfully.';
        status.className = 'success';
      }} catch (err) {{
        status.textContent = err.message || String(err);
        status.className = 'error';
      }}
    }});
  </script>
</body>
</html>"""
    return HTMLResponse(content=html)


@app.post("/generate/")
async def generate_proposal_endpoint(
    company_name: str = Form(...),
    proposal_ref: str = Form(""),
    client_logo: UploadFile | None = File(None),
    annexure_order_json: str = Form(""),
    assessments_json: str = Form(""),
) -> FileResponse:
    tmp_dir: str | None = None
    logo_path: Path | None = None
    try:
        ref = (proposal_ref or "").strip() or default_proposal_ref()
        annexure_order = parse_annexure_order_json(annexure_order_json)
        assessments = parse_assessments_json(assessments_json) if not annexure_order else None
        tmp_dir = tempfile.mkdtemp(prefix="css_proposal_")
        if client_logo and client_logo.filename:
            ext = Path(client_logo.filename).suffix or ".png"
            logo_path = Path(tmp_dir) / f"client_logo{ext}"
            logo_path.write_bytes(await client_logo.read())
        out_name = _safe_filename(company_name)
        output_path = Path(tmp_dir) / out_name
        generate_proposal(
            company_name=company_name,
            proposal_ref=ref,
            logo_path=str(logo_path) if logo_path else None,
            output_path=output_path,
            annexure_order=annexure_order,
            assessments=assessments,
        )
        final_path = UPLOAD_DIR / out_name
        shutil.copy2(output_path, final_path)
        return FileResponse(
            path=str(final_path),
            filename=out_name,
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    except FileNotFoundError as exc:
        raise HTTPException(status_code=500, detail=str(exc)) from exc
    except Exception as exc:
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=str(exc)) from exc
    finally:
        if tmp_dir and os.path.isdir(tmp_dir):
            shutil.rmtree(tmp_dir, ignore_errors=True)


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------
def main() -> None:
    parser = argparse.ArgumentParser(description="CSS Proposal Generator (standalone)")
    parser.add_argument(
        "--build-template",
        action="store_true",
        help="Rebuild CSS_Proposal_Base_Template.docx from CSS_Proposal_Source.docx",
    )
    parser.add_argument("--host", default="localhost")
    parser.add_argument("--port", type=int, default=DEFAULT_PORT)
    args = parser.parse_args()

    if args.build_template:
        out = build_base_template()
        print(f"Built base template: {out}")
        return

    import uvicorn

    print(f"CSS Proposal Generator v{VERSION}")
    print(f"Open http://localhost:{args.port}/")
    uvicorn.run(app, host=args.host, port=args.port)


if __name__ == "__main__":
    main()

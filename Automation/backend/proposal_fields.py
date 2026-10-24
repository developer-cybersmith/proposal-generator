"""
Document edits for the Excel-driven flow: the `Count:` and `Frequency:` lines inside annexures,
and annexure cross-reference renumbering.

All formatting comes from proposal_tokens (plan section 0). Operates on lists of OOXML elements
(the cloned annexure block) or on a python-docx Document.

Decisions (docs/ADVANCED_FEATURE_PLAN.md):
  D30/D31  Count line:  "Count: N <label>" (label per annexure; two labels joined for Annexure 3)
  D23/D34  Frequency is one shared field. While it still holds the default ("One Time") the
           annexure keeps its own template wording (e.g. "One time with Revalidation"), so no
           template commitment is silently dropped; once the user edits it, that text is used.
"""
from __future__ import annotations

import copy
import re
from typing import Iterable

from docx.oxml import OxmlElement
from docx.oxml.ns import qn

import proposal_tokens as tokens

FREQUENCY_RE = re.compile(r"^\s*Frequency\s*:", re.IGNORECASE)
COUNT_RE = re.compile(r"^\s*Count\s*:", re.IGNORECASE)
DELIVERABLES_RE = re.compile(r"^\s*Deliverables\b", re.IGNORECASE)
ANNEXURE_REF_RE = re.compile(r"\bAnnexure\s+(\d+)\b", re.IGNORECASE)
ANNEXURE_HEADING_RE = re.compile(r"^\s*Annexure\s+\d+\s*:", re.IGNORECASE)


# --------------------------------------------------------------------------- OOXML helpers
def paragraph_text(p) -> str:
    return "".join(t.text or "" for t in p.iter(qn("w:t")))


def make_run(text: str, *, bold: bool = False, italic: bool = False, color: str,
             size_pt: float, font: str = tokens.FONT):
    """A run formatted only from tokens."""
    r = OxmlElement("w:r")
    rpr = OxmlElement("w:rPr")
    rfonts = OxmlElement("w:rFonts")
    for attr in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        rfonts.set(qn(attr), font)
    rpr.append(rfonts)
    if bold:
        rpr.append(OxmlElement("w:b"))
        rpr.append(OxmlElement("w:bCs"))
    if italic:
        rpr.append(OxmlElement("w:i"))
        rpr.append(OxmlElement("w:iCs"))
    c = OxmlElement("w:color")
    c.set(qn("w:val"), color)
    rpr.append(c)
    half_points = str(int(round(size_pt * 2)))
    for tag in ("w:sz", "w:szCs"):
        sz = OxmlElement(tag)
        sz.set(qn("w:val"), half_points)
        rpr.append(sz)
    r.append(rpr)
    t = OxmlElement("w:t")
    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    t.text = text
    r.append(t)
    return r


def _spacing_ppr(after_pt: float = tokens.BODY_SPACE_AFTER_PT):
    ppr = OxmlElement("w:pPr")
    sp = OxmlElement("w:spacing")
    sp.set(qn("w:before"), str(int(tokens.BODY_SPACE_BEFORE_PT * 20)))
    sp.set(qn("w:after"), str(int(after_pt * 20)))
    ppr.append(sp)
    return ppr


def label_value_paragraph(label: str, value: str):
    """New paragraph: bold purple label, regular body-colour value (spec section 8)."""
    p = OxmlElement("w:p")
    p.append(_spacing_ppr())
    p.append(make_run(f"{label}: ", bold=True, color=tokens.BRAND_PURPLE, size_pt=tokens.SIZE_BODY))
    if value:
        p.append(make_run(value, color=tokens.BODY_TEXT, size_pt=tokens.SIZE_BODY))
    return p


def rewrite_label_value(p, label: str, value: str) -> None:
    """Replace every run in *p* with a bold label run and a regular value run.

    Paragraph properties (indent, spacing) are kept, the paragraph-mark run properties are
    dropped (they carried highlights/legacy fonts), bookmarks stay.
    """
    for child in list(p):
        if child.tag == qn("w:pPr"):
            mark = child.find(qn("w:rPr"))
            if mark is not None:
                child.remove(mark)
        elif child.tag in (qn("w:r"), qn("w:hyperlink"), qn("w:smartTag"), qn("w:ins")):
            p.remove(child)
    p.append(make_run(f"{label}: ", bold=True, color=tokens.BRAND_PURPLE, size_pt=tokens.SIZE_BODY))
    if value:
        p.append(make_run(value, color=tokens.BODY_TEXT, size_pt=tokens.SIZE_BODY))


def _value_after_colon(p) -> str:
    return paragraph_text(p).split(":", 1)[1].strip() if ":" in paragraph_text(p) else ""


# --------------------------------------------------------------------------- Count / Frequency
def count_line_text(entries: Iterable[tuple[int, str]]) -> str:
    """[(6, 'Web Application categories')] -> '6 Web Application categories'.

    Several entries (External + Internal share Annexure 3) are joined with ', '.
    """
    return ", ".join(f"{n} {label}".strip() for n, label in entries)


def apply_count_and_frequency(
    elements: list,
    *,
    count_text: str | None,
    frequency: str | None,
    default_frequency: str = "One Time",
) -> dict[str, str]:
    """Rewrite or insert the Count / Frequency lines in a cloned annexure block.

    count_text  - text after "Count: " (None leaves the line untouched, e.g. manual flow)
    frequency   - the shared Frequency value. Equal to default_frequency (or empty/None) keeps
                  the annexure's own wording; any other text replaces it.
    Returns {"count": "rewritten|inserted|skipped", "frequency": "..."} for tests and logging.
    """
    outcome = {"count": "skipped", "frequency": "skipped"}
    paragraphs = [el for el in elements if el.tag == qn("w:p")]
    freq_p = next((p for p in paragraphs if FREQUENCY_RE.match(paragraph_text(p))), None)
    count_p = next((p for p in paragraphs if COUNT_RE.match(paragraph_text(p))), None)

    use_custom_freq = bool(frequency) and frequency.strip() != default_frequency
    if use_custom_freq:
        if freq_p is not None:
            rewrite_label_value(freq_p, "Frequency", frequency.strip())
            outcome["frequency"] = "rewritten"
        else:
            freq_p = label_value_paragraph("Frequency", frequency.strip())
            _insert_field_paragraph(elements, freq_p, anchor=count_p)
            outcome["frequency"] = "inserted"
    elif freq_p is not None:
        # keep the annexure's own wording, but give the line the specification formatting
        rewrite_label_value(freq_p, "Frequency", _value_after_colon(freq_p))
        outcome["frequency"] = "kept"
    elif count_text is not None:
        # Excel flow, annexure has no Frequency line: add the default so every annexure has one
        freq_p = label_value_paragraph("Frequency", default_frequency)
        _insert_field_paragraph(elements, freq_p, anchor=count_p)
        outcome["frequency"] = "inserted"

    if count_text is not None:
        if count_p is not None:
            rewrite_label_value(count_p, "Count", count_text)
            outcome["count"] = "rewritten"
        else:
            count_p = label_value_paragraph("Count", count_text)
            if freq_p is not None and freq_p in elements:
                elements.insert(elements.index(freq_p) + 1, count_p)
            else:
                _insert_field_paragraph(elements, count_p, anchor=None)
            outcome["count"] = "inserted"
    return outcome


def _insert_field_paragraph(elements: list, new_p, *, anchor) -> None:
    """Insert before the Deliverables heading when present, otherwise at the end."""
    if anchor is not None and anchor in elements:
        elements.insert(elements.index(anchor), new_p)
        return
    for i, el in enumerate(elements):
        if el.tag == qn("w:p") and DELIVERABLES_RE.match(paragraph_text(el)):
            elements.insert(i, new_p)
            return
    elements.append(new_p)


# --------------------------------------------------------------------------- cross references
def renumber_cross_references(elements_or_doc, mapping: dict[int, int]) -> int:
    """Rewrite 'Annexure N' mentions in body text (not headings) using old->new numbers.

    Only references contained in a single text node are rewritten; returns how many were.
    Headings ("Annexure N: Title") are managed elsewhere and left alone.
    """
    if hasattr(elements_or_doc, "element"):
        roots = [elements_or_doc.element.body]
    else:
        roots = list(elements_or_doc)
    changed = 0
    for root in roots:
        for p in root.iter(qn("w:p")):
            if ANNEXURE_HEADING_RE.match(paragraph_text(p)) or "\t" in paragraph_text(p):
                continue  # headings and Contents lines
            for t in p.iter(qn("w:t")):
                if not t.text:
                    continue

                def repl(m: re.Match) -> str:
                    nonlocal changed
                    old = int(m.group(1))
                    new = mapping.get(old)
                    if new is None or new == old:
                        return m.group(0)
                    changed += 1
                    return m.group(0).replace(str(old), str(new))

                t.text = ANNEXURE_REF_RE.sub(repl, t.text)
    return changed

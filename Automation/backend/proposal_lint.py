"""
Structural and formatting lint for generated proposals (.docx).

Hard acceptance rules from docs/ADVANCED_FEATURE_PLAN.md section 0:
  * annexure numbering is sequential (1..N) with no duplicates,
  * the Contents page lists every annexure in the body,
  * no empty or near-empty pages (structural proxy: no stacked page breaks, no runs of blank
    paragraphs, no trailing break),
  * tables stay within the text width.
Formatting findings (fonts/colours outside the tokens) are reported as warnings: the current
template predates the specification and is brought to it in Track B.

Usage:  report = lint_docx("proposal.docx");  report.errors / report.warnings / report.ok
"""
from __future__ import annotations

import re
from collections import Counter
from dataclasses import dataclass, field
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn

import proposal_contents
import proposal_tokens as tokens

ANNEXURE_RE = re.compile(r"^Annexure\s+(\d+)\s*:\s*(.+)$", re.IGNORECASE)
BLANK_RUN_LIMIT = 3          # this many consecutive empty paragraphs is "spacing by blank lines"
WIDTH_TOLERANCE = 1.02       # allow 2% over the text width
TWIPS_PER_INCH = 1440


@dataclass
class Finding:
    severity: str            # "error" | "warning"
    code: str
    message: str

    def __str__(self) -> str:
        return f"[{self.severity}] {self.code}: {self.message}"


@dataclass
class LintReport:
    findings: list[Finding] = field(default_factory=list)

    @property
    def errors(self) -> list[Finding]:
        return [f for f in self.findings if f.severity == "error"]

    @property
    def warnings(self) -> list[Finding]:
        return [f for f in self.findings if f.severity == "warning"]

    @property
    def ok(self) -> bool:
        return not self.errors

    def codes(self) -> set[str]:
        return {f.code for f in self.findings}


# --------------------------------------------------------------------------- paragraph helpers
def body_elements_after_cover(doc) -> list:
    """Top-level body elements from the INTRODUCTION heading on.

    The cover page, the Contents block and the fixed front pages (Introduction ... tools, pages 3-8
    of the organisation's master) keep the organisation's own design: blank-line spacing places
    their diagrams. Formatting checks therefore start at the first annexure. A document without
    that heading is checked from INTRODUCTION, and one without either in full.
    """
    children = list(doc.element.body)
    seen_contents = False
    intro = None
    for i, el in enumerate(children):
        if el.tag != qn("w:p"):
            continue
        t = " ".join(_text(el).split())
        if t == "Contents":
            seen_contents = True
        elif seen_contents and t == "INTRODUCTION" and intro is None:
            intro = i
        elif intro is not None and ANNEXURE_RE.match(t):
            return children[i:]
    return children[intro:] if intro is not None else children


def _text(p) -> str:
    return "".join(t.text or "" for t in p.iter(qn("w:t"))).strip()


def _has_page_break(p) -> bool:
    return any(br.get(qn("w:type")) == "page" for br in p.iter(qn("w:br")))


def _has_drawing(p) -> bool:
    return any(True for _ in p.iter(qn("w:drawing"))) or any(True for _ in p.iter(qn("w:pict")))


def _is_blank(p) -> bool:
    return not _text(p) and not _has_page_break(p) and not _has_drawing(p) and not any(
        True for _ in p.iter(qn("w:sectPr")))


def _body_paragraph_elements(doc) -> list:
    """Top-level paragraphs in document order (tables and content controls are skipped)."""
    return [el for el in doc.element.body if el.tag == qn("w:p")]


# --------------------------------------------------------------------------- checks
def _body_annexure_headings(doc) -> list[str]:
    """Annexure heading texts in the body (the Contents lines that repeat them are not headings)."""
    block = proposal_contents.contents_block(doc.element.body)      # kept alive: lxml proxies must not be freed
    contents = {id(el) for el in block}
    return [_text(el) for el in doc.element.body
            if el.tag == qn("w:p") and id(el) not in contents and ANNEXURE_RE.match(_text(el))]


def check_annexure_numbering(doc) -> list[Finding]:
    numbers = []
    for text in _body_annexure_headings(doc):
        m = ANNEXURE_RE.match(text)
        if m:
            numbers.append(int(m.group(1)))
    findings: list[Finding] = []
    if not numbers:
        return findings
    expected = list(range(1, len(numbers) + 1))
    if numbers != expected:
        dupes = sorted(n for n, c in Counter(numbers).items() if c > 1)
        detail = f" duplicates: {dupes}" if dupes else ""
        findings.append(Finding("error", "annexure_numbering",
                                f"annexure headings are {numbers}, expected {expected}.{detail}"))
    return findings


def _contents_entries(doc) -> list[str]:
    """The label of every entry in the Contents block (hyperlink text, or text before the tab)."""
    labels: list[str] = []
    for p in proposal_contents.contents_block(doc.element.body):
        link = p.find(qn("w:hyperlink"))
        if link is not None:
            texts = [t.text or "" for t in link.iter(qn("w:t"))]
            if texts and texts[0].strip():
                labels.append(texts[0].strip())
        elif "\t" in "".join(t.text or "" for t in p.iter(qn("w:t"))):
            labels.append(_text(p).split("\t")[0].strip())
    return labels


def check_toc_lists_all_annexures(doc) -> list[Finding]:
    """The Contents lists exactly the headings the body has (fixed sections, every annexure, closing)."""
    entries = _contents_entries(doc)
    if not entries:
        return [Finding("warning", "toc_not_found", "could not locate the Contents entries to compare")]
    headings = [h.label for h in proposal_contents.find_headings(doc.element.body)]
    toc_annexures = sum(1 for e in entries if ANNEXURE_RE.match(e))
    body_annexures = sum(1 for h in headings if ANNEXURE_RE.match(h))
    if toc_annexures != body_annexures:
        return [Finding("error", "toc_incomplete",
                        f"Contents lists {toc_annexures} annexure(s) but the body has {body_annexures}")]
    if entries != headings:
        missing = [h for h in headings if h not in entries]
        extra = [e for e in entries if e not in headings]
        return [Finding("error", "toc_mismatch",
                        f"Contents does not match the body headings (missing {missing}, extra {extra})")]
    return []


def check_blank_spacing(doc) -> list[Finding]:
    findings: list[Finding] = []
    paragraphs = _body_paragraph_elements(doc)
    body_ids = {id(el) for el in body_elements_after_cover(doc)}
    blank_run = 0
    break_run = 0
    for i, p in enumerate(paragraphs):
        if _is_blank(p):
            blank_run += 1
            if blank_run == BLANK_RUN_LIMIT and id(p) in body_ids:      # the cover keeps its own layout
                findings.append(Finding("warning", "blank_paragraph_run",
                                        f"{BLANK_RUN_LIMIT}+ empty paragraphs in a row near body paragraph {i}"))
        else:
            blank_run = 0
        if _has_page_break(p) and not _text(p):
            break_run += 1
            if break_run == 2:
                findings.append(Finding("error", "stacked_page_breaks",
                                        f"consecutive page breaks near body paragraph {i} (empty page)"))
        elif not _is_blank(p):
            break_run = 0
    # a trailing page break leaves an empty last page
    for p in reversed(paragraphs):
        if _is_blank(p):
            continue
        if _has_page_break(p) and not _text(p):
            findings.append(Finding("error", "trailing_page_break", "document ends with a page break"))
        break
    return findings


def _text_width_twips(doc) -> int:
    s = doc.sections[0]
    return int((s.page_width - s.left_margin - s.right_margin) / 635)  # EMU -> twips


def check_tables_within_width(doc) -> list[Finding]:
    limit = _text_width_twips(doc) * WIDTH_TOLERANCE
    findings: list[Finding] = []
    for n, tbl in enumerate(doc.element.body.iter(qn("w:tbl")), start=1):
        grid = tbl.find(qn("w:tblGrid"))
        if grid is None:
            continue
        widths = [int(g.get(qn("w:w")) or 0) for g in grid.findall(qn("w:gridCol"))]
        total = sum(widths)
        if total > limit:
            findings.append(Finding("error", "table_too_wide",
                                    f"table {n} is {total / TWIPS_PER_INCH:.2f} in wide, text width is "
                                    f"{limit / WIDTH_TOLERANCE / TWIPS_PER_INCH:.2f} in"))
    return findings


def check_run_formatting(doc) -> list[Finding]:
    """Warn about explicit fonts/colours outside the tokens (aggregated, one finding each)."""
    fonts: Counter = Counter()
    colors: Counter = Counter()
    runs = [r for el in body_elements_after_cover(doc) for r in el.iter(qn("w:r"))]
    for r in runs:
        if not "".join(t.text or "" for t in r.iter(qn("w:t"))).strip():
            continue
        rpr = r.find(qn("w:rPr"))
        if rpr is None:
            continue
        rf = rpr.find(qn("w:rFonts"))
        if rf is not None:
            name = rf.get(qn("w:ascii")) or rf.get(qn("w:hAnsi"))
            if name and name != tokens.FONT:
                fonts[name] += 1
        col = rpr.find(qn("w:color"))
        if col is not None:
            val = (col.get(qn("w:val")) or "").upper()
            if val and val not in ("AUTO",) and val not in tokens.ALLOWED_COLORS:
                colors[val] += 1
    findings = [Finding("warning", "font_outside_tokens", f"{n} run(s) use font {name!r}, token font is {tokens.FONT!r}")
                for name, n in fonts.most_common()]
    findings += [Finding("warning", "color_outside_tokens", f"{n} run(s) use colour #{val}, not in the token set")
                 for val, n in colors.most_common()]
    return findings


def check_final_section(doc) -> list[Finding]:
    """The body must end with section properties, or Word falls back to application defaults."""
    if doc.element.body.find(qn("w:sectPr")) is None:
        return [Finding("error", "missing_final_section",
                        "the document has no closing section properties (page size/margins undefined)")]
    return []


def check_contents_has_no_phantom_entries(doc) -> list[Finding]:
    """Every fixed Contents entry must have a matching heading in the body."""
    paragraphs = doc.paragraphs
    start = next((i for i, p in enumerate(paragraphs) if p.text.strip() == "Contents"), None)
    if start is None:
        return []
    entries, end = [], start
    for i in range(start + 1, len(paragraphs)):
        t = paragraphs[i].text
        if "\t" in t and t.strip():
            entries.append(t.split("\t")[0].strip())
            end = i
        elif paragraphs[i].style is not None and paragraphs[i].style.name.lower().startswith("heading") and t.strip():
            break
    body = {p.text.strip().lower() for p in paragraphs[end + 1:] if "\t" not in p.text}
    fixed = {"introduction", "our understanding of your requirement", "methodology overview",
             "vulnerability assessment methodology", "proposed timelines & project management", "team overview"}
    return [Finding("error", "contents_entry_without_heading", f"Contents lists {e!r} but the body has no such heading")
            for e in entries if e.lower() in fixed and e.lower() not in body]


def check_package(path: str | Path) -> list[Finding]:
    """OOXML package checks that Microsoft Word is strict about."""
    import zipfile
    from lxml import etree
    findings: list[Finding] = []
    try:
        with zipfile.ZipFile(path) as z:
            if "docProps/core.xml" in z.namelist():
                root = etree.fromstring(z.read("docProps/core.xml"))
                core_ns = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
                for tag in ("lastModifiedBy", "revision", "keywords", "category", "contentStatus", "lastPrinted"):
                    for el in root.iter():
                        local = etree.QName(el).localname
                        if local == tag and etree.QName(el).namespace != core_ns:
                            findings.append(Finding("error", "core_properties_invalid",
                                                    f"docProps/core.xml: <{tag}> is in namespace "
                                                    f"{etree.QName(el).namespace!r}; Word rejects the file as corrupt"))
    except (zipfile.BadZipFile, etree.XMLSyntaxError, KeyError) as exc:
        findings.append(Finding("error", "package_unreadable", f"{type(exc).__name__}: {exc}"))
    return findings


CHECKS = (
    check_annexure_numbering,
    check_toc_lists_all_annexures,
    check_contents_has_no_phantom_entries,
    check_blank_spacing,
    check_final_section,
    check_tables_within_width,
    check_run_formatting,
)


def lint_docx(path_or_doc) -> LintReport:
    is_path = isinstance(path_or_doc, (str, Path))
    doc = Document(str(path_or_doc)) if is_path else path_or_doc
    report = LintReport()
    if is_path:
        report.findings.extend(check_package(path_or_doc))
    for check in CHECKS:
        report.findings.extend(check(doc))
    return report


if __name__ == "__main__":  # python proposal_lint.py file.docx
    import sys
    rep = lint_docx(sys.argv[1])
    for f in rep.findings:
        print(f)
    print("OK" if rep.ok else f"{len(rep.errors)} error(s)")
    sys.exit(0 if rep.ok else 1)

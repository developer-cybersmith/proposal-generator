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
def check_annexure_numbering(doc) -> list[Finding]:
    numbers = []
    for p in doc.paragraphs:
        m = ANNEXURE_RE.match(p.text.strip())
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


def _toc_annexure_entries(doc) -> tuple[int, int | None]:
    """Number of annexure lines in Contents.

    Contents order is: fixed intro entries ... TEAM OVERVIEW, one line per annexure,
    Terms & Conditions, Conclusion & Pricing. Returns (-1, None) when that frame is not found.
    """
    paragraphs = doc.paragraphs
    start = next((i for i, p in enumerate(paragraphs) if p.text.strip() == "Contents"), None)
    if start is None:
        return -1, None
    entries: list[str] = []
    for p in paragraphs[start + 1:]:
        t = p.text.strip()
        if p.style is not None and p.style.name.lower().startswith("heading") and t:
            break
        if "\t" in p.text and t:
            entries.append(t.split("\t")[0].strip())
    try:
        first = next(i for i, e in enumerate(entries) if e.lower().startswith("team overview"))
        last = next(i for i, e in enumerate(entries) if e.lower().startswith("terms"))
    except StopIteration:
        return -1, None
    return max(0, last - first - 1), None


def check_toc_lists_all_annexures(doc) -> list[Finding]:
    body_count = sum(1 for p in doc.paragraphs if ANNEXURE_RE.match(p.text.strip()))
    toc_count, _ = _toc_annexure_entries(doc)
    if toc_count < 0:
        return [Finding("warning", "toc_not_found", "could not locate the Contents entries to compare")]
    if toc_count != body_count:
        return [Finding("error", "toc_incomplete",
                        f"Contents lists {toc_count} annexure(s) but the body has {body_count}")]
    return []


def check_blank_spacing(doc) -> list[Finding]:
    findings: list[Finding] = []
    paragraphs = _body_paragraph_elements(doc)
    blank_run = 0
    break_run = 0
    for i, p in enumerate(paragraphs):
        if _is_blank(p):
            blank_run += 1
            if blank_run == BLANK_RUN_LIMIT:
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
    for r in doc.element.body.iter(qn("w:r")):
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


CHECKS = (
    check_annexure_numbering,
    check_toc_lists_all_annexures,
    check_blank_spacing,
    check_tables_within_width,
    check_run_formatting,
)


def lint_docx(path_or_doc) -> LintReport:
    doc = Document(str(path_or_doc)) if isinstance(path_or_doc, (str, Path)) else path_or_doc
    report = LintReport()
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

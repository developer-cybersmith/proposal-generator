"""The proposal's Contents page.

The organisation's master document has a real Word TOC field: one hyperlinked line per heading,
each with a PAGEREF field, linked to a ``_Toc`` bookmark on the heading. This module writes that
same structure for whatever the generated proposal contains, so

  * the Contents always lists exactly the headings that are in the body (Introduction ... Methodology,
    "Annexure N: Title" for every annexure, Terms & Conditions, Conclusion & Pricing),
  * the cached page numbers are the measured ones (or an estimate), and
  * when Word updates fields it rebuilds the same list from the heading styles.

Entry key (used for page numbers): the lower-case heading text, or ``annexure N``.
"""
from __future__ import annotations

import re
from dataclasses import dataclass
from xml.sax.saxutils import escape

from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

import proposal_tokens as tokens

MARKER = "[[CONTENTS]]"
CONTENTS_TEXT = "Contents"
TOC_INSTRUCTION = ' TOC \\t "Heading 1,1,Heading 2,1,Annexure Title,2" \\h \\z '

STATIC_HEADINGS = (
    "INTRODUCTION",
    "OUR UNDERSTANDING OF YOUR REQUIREMENT",
    "TEAM OVERVIEW",
    "METHODOLOGY OVERVIEW",
    "VULNERABILITY ASSESSMENT METHODOLOGY",
)
TERMS_TEXT = "Terms & Conditions"
CONCLUSION_TEXT = "Conclusion & Pricing"
ANNEXURE_RE = re.compile(r"^\s*Annexure\s+(\d+)\s*:\s*(.+?)\s*$", re.IGNORECASE)
_TOC_STYLES = ("TOC1", "TOC2")
_XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"


@dataclass
class Heading:
    label: str                  # text shown in the Contents
    key: str                    # page-number key: lower-case text, or "annexure N"
    level: int                  # 1 = TOC1 (bold), 2 = TOC2 (annexures)
    paragraph: object           # the heading's w:p element


def _text(p) -> str:
    return "".join(t.text or "" for t in p.iter(qn("w:t")))


def _pstyle(p) -> str | None:
    ppr = p.find(qn("w:pPr"))
    node = ppr.find(qn("w:pStyle")) if ppr is not None else None
    return node.get(qn("w:val")) if node is not None else None


def key_for(label: str) -> str:
    m = ANNEXURE_RE.match(label)
    return f"annexure {int(m.group(1))}" if m else re.sub(r"\s+", " ", label).strip().lower()


# ------------------------------------------------------------------------------------ locating
def _contents_heading(body):
    for el in body:
        if el.tag == qn("w:p") and re.sub(r"\s+", " ", _text(el)).strip() == CONTENTS_TEXT:
            return el
    return None


def _is_contents_paragraph(el) -> bool:
    """The marker, a Contents entry, or the paragraph that closes the TOC field."""
    if el.tag != qn("w:p"):
        return False
    if _text(el).strip() == MARKER or _pstyle(el) in _TOC_STYLES:
        return True
    fld = [c.get(qn("w:fldCharType")) for c in el.iter(qn("w:fldChar"))]
    return fld == ["end"] and not _text(el).strip()


def contents_block(body) -> list:
    """The paragraphs between the Contents heading and the next real content."""
    heading = _contents_heading(body)
    if heading is None:
        return []
    block = []
    el = heading.getnext()
    while el is not None and _is_contents_paragraph(el):
        block.append(el)
        el = el.getnext()
    return block


def find_headings(body) -> list[Heading]:
    """Every heading that belongs in the Contents, in document order, after the Contents block."""
    heading = _contents_heading(body)
    block = contents_block(body)
    start = block[-1] if block else heading
    found: list[Heading] = []
    el = start.getnext() if start is not None else None
    seen_static: set[str] = set()
    while el is not None:
        if el.tag == qn("w:p"):
            text = re.sub(r"\s+", " ", _text(el)).strip()
            m = ANNEXURE_RE.match(text)
            if m:
                found.append(Heading(text, key_for(text), 2, el))
            elif text in STATIC_HEADINGS and text not in seen_static:
                seen_static.add(text)
                found.append(Heading(text, key_for(text), 1, el))
            elif text == TERMS_TEXT and not any(h.key == key_for(TERMS_TEXT) for h in found):
                found.append(Heading(text, key_for(text), 1, el))
            elif text == CONCLUSION_TEXT and not any(h.key == key_for(CONCLUSION_TEXT) for h in found):
                found.append(Heading(text, key_for(text), 1, el))
        el = el.getnext()
    return found


# ------------------------------------------------------------------------------------ bookmarks
def _next_bookmark_id(body) -> int:
    ids = [int(b.get(qn("w:id"))) for b in body.iter(qn("w:bookmarkStart")) if (b.get(qn("w:id")) or "").isdigit()]
    return max(ids + [100]) + 1


def ensure_bookmark(body, p) -> str:
    """Name of a ``_Toc`` bookmark that spans the heading paragraph (created when missing)."""
    for b in p.findall(qn("w:bookmarkStart")):
        name = b.get(qn("w:name")) or ""
        if name.startswith("_Toc"):
            return name
    bid = _next_bookmark_id(body)
    name = f"_Toc{700000000 + bid}"
    start = OxmlElement("w:bookmarkStart")
    start.set(qn("w:id"), str(bid))
    start.set(qn("w:name"), name)
    end = OxmlElement("w:bookmarkEnd")
    end.set(qn("w:id"), str(bid))
    ppr = p.find(qn("w:pPr"))
    if ppr is not None:
        ppr.addnext(start)
    else:
        p.insert(0, start)
    p.append(end)
    return name


# ------------------------------------------------------------------------------------ building
def _run(inner: str, *, bold=False, italic=False, hidden=False) -> str:
    props = f'<w:rFonts w:ascii="{tokens.FONT}" w:hAnsi="{tokens.FONT}" w:cs="{tokens.FONT}"/>'
    props += "<w:b/>" if bold else ""
    props += "<w:i/>" if italic else ""
    props += "<w:noProof/>"
    props += f'<w:color w:val="{tokens.BODY_TEXT}"/>'
    props += "<w:webHidden/>" if hidden else ""
    return f"<w:r><w:rPr>{props}</w:rPr>{inner}</w:r>"


def _fld(kind: str, **kw) -> str:
    return _run(f'<w:fldChar w:fldCharType="{kind}"/>', **kw)


def _entry_paragraph(heading: Heading, bookmark: str, page: int, *, first: bool):
    style = "TOC2" if heading.level == 2 else "TOC1"
    bold = True
    italic = heading.level == 2
    lead = ""
    if first:
        lead = (_fld("begin")
                + _run(f'<w:instrText xml:space="preserve">{escape(TOC_INSTRUCTION)}</w:instrText>')
                + _fld("separate"))
    link = (f'<w:hyperlink w:anchor="{bookmark}" w:history="1">'
            + _run(f'<w:t xml:space="preserve">{escape(heading.label)}</w:t>', bold=bold, italic=italic)
            + _run("<w:tab/>", hidden=True)
            + _fld("begin", hidden=True)
            + _run(f'<w:instrText xml:space="preserve"> PAGEREF {bookmark} \\h </w:instrText>', hidden=True)
            + _fld("separate", hidden=True)
            + _run(f"<w:t>{int(page)}</w:t>", hidden=True)
            + _fld("end", hidden=True)
            + "</w:hyperlink>")
    return parse_xml(f'<w:p {nsdecls("w")}><w:pPr><w:pStyle w:val="{style}"/></w:pPr>{lead}{link}</w:p>')


def _closing_paragraph():
    return parse_xml(f'<w:p {nsdecls("w")}><w:pPr><w:spacing w:before="0" w:after="0"/></w:pPr>{_fld("end")}</w:p>')


def build(doc, pages: dict[str, int] | None = None, default_page: int = 1) -> list[Heading]:
    """(Re)write the Contents entries from the headings now in the body. Returns the headings used.

    *pages* maps an entry key to its page; a missing key keeps the page already cached for that
    entry (when rebuilding) or *default_page*.
    """
    body = doc.element.body
    heading = _contents_heading(body)
    if heading is None:
        return []
    cached = read_pages(doc)
    entries = find_headings(body)
    old = contents_block(body)
    anchor = heading
    for el in old:
        body.remove(el)
    paras = []
    for i, h in enumerate(entries):
        bookmark = ensure_bookmark(body, h.paragraph)
        page = (pages or {}).get(h.key) or cached.get(h.key) or default_page
        paras.append(_entry_paragraph(h, bookmark, int(page), first=(i == 0)))
    if not paras:                                              # nothing to list: keep a marker for next time
        marker = OxmlElement("w:p")
        run = OxmlElement("w:r")
        t = OxmlElement("w:t")
        t.text = MARKER
        run.append(t)
        marker.append(run)
        paras = [marker]
    else:
        paras.append(_closing_paragraph())
    for p in paras:
        anchor.addnext(p)
        anchor = p
    return entries


def read_pages(doc) -> dict[str, int]:
    """The page numbers currently cached in the Contents entries, by entry key."""
    pages: dict[str, int] = {}
    for p in contents_block(doc.element.body):
        link = p.find(qn("w:hyperlink"))
        if link is None:
            continue
        texts = [t.text or "" for t in link.iter(qn("w:t"))]
        if len(texts) >= 2 and texts[-1].strip().isdigit():
            pages[key_for(texts[0].strip())] = int(texts[-1])
    return pages


def refresh_pages(doc, measured: dict[str, int]) -> int:
    """Write measured page numbers into the cached results. Returns how many entries changed."""
    changed = 0
    for p in contents_block(doc.element.body):
        link = p.find(qn("w:hyperlink"))
        if link is None:
            continue
        ts = list(link.iter(qn("w:t")))
        if len(ts) < 2:
            continue
        key = key_for((ts[0].text or "").strip())
        new = measured.get(key)
        if new and (ts[-1].text or "").strip() != str(int(new)):
            ts[-1].text = str(int(new))
            changed += 1
    return changed

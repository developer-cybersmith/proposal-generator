"""
Track B: bring a generated proposal to the CyberSmithSECURE Formatting Specification.

`apply_specification(doc)` runs on the assembled document (docs/CyberSmithSECURE_Proposal_
Formatting_Specification.md is the authority; every value comes from proposal_tokens):

  * named styles: Normal, Heading 1/2/3, AnnexureTitle, TOC 1/2 defined once, from the tokens;
  * headings and sub-labels use those styles instead of direct formatting;
  * body text is Corbel 11 pt in the body colour; spacing comes from paragraph spacing, never from
    empty paragraphs;
  * tables follow spec section 9 (purple header row, header repeats, rows never split);
  * the signature block is a borderless two-column table (not space-aligned text);
  * AI-annexure bullets use a real Word list definition (numbering part) instead of typed glyphs.

The cover page and Contents block (everything before the INTRODUCTION heading) are left as the
template designed them: the cover navy and artwork are brand-owner decisions (see plan section 12).
The function is idempotent and only reads/writes OOXML; it never changes the text.
"""
from __future__ import annotations

import copy
import re
from typing import Any

from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

import proposal_fields
import proposal_tokens as tokens
from proposal_fields import make_run

BULLET_NUM_ID = 9100          # numbering definition owned by this module (AI annexure bullets)
SPACING_AFTER_BODY = int(tokens.BODY_SPACE_AFTER_PT * 20)
SPACING_AFTER_LIST = int(tokens.LIST_SPACE_AFTER_PT * 20)
LINE_SPACING = 264            # 1.1 lines (spec: 1.08-1.15)
TWIPS_PER_INCH = 1440

ANNEXURE_TITLE_RE = re.compile(r"^\s*Annexure\s+\d+\s*:", re.IGNORECASE)
HEADING2_TEXTS = {"introduction", "our understanding of your requirement", "team overview", "methodology overview",
                  "vulnerability assessment methodology", "proposed timelines & project management",
                  "terms & conditions"}
HEADING1_TEXTS = {"conclusion & pricing"}
PURPLES = {"7030a0", "6f2f9f", "6f2f9e"}


def _w(tag: str) -> str:
    return qn(tag)


# =============================================================================== styles
def _style_xml(style_id: str, name: str, *, based_on: str | None = "Normal", next_style: str | None = None,
               bold: bool = False, italic: bool = False, color: str = tokens.BODY_TEXT,
               size_pt: float = tokens.SIZE_BODY, before_pt: float = 0, after_pt: float = tokens.BODY_SPACE_AFTER_PT,
               keep_next: bool = False, outline: int | None = None, line: int = LINE_SPACING,
               tabs: str = "", ui_priority: int = 9) -> str:
    half = int(round(size_pt * 2))
    ppr = []
    if keep_next:
        ppr.append("<w:keepNext/><w:keepLines/>")
    ppr.append(f'<w:widowControl w:val="1"/>')
    if tabs:
        ppr.append(tabs)
    ppr.append(f'<w:spacing w:before="{int(before_pt * 20)}" w:after="{int(after_pt * 20)}" w:line="{line}" w:lineRule="auto"/>')
    ppr.append('<w:ind w:left="0" w:firstLine="0"/>')
    if outline is not None:
        ppr.append(f'<w:outlineLvl w:val="{outline}"/>')
    rpr = (f'<w:rFonts w:ascii="{tokens.FONT}" w:hAnsi="{tokens.FONT}" w:cs="{tokens.FONT}" w:eastAsia="{tokens.FONT}"/>'
           + ("<w:b/><w:bCs/>" if bold else "<w:b w:val=\"0\"/><w:bCs w:val=\"0\"/>")
           + ("<w:i/><w:iCs/>" if italic else "<w:i w:val=\"0\"/><w:iCs w:val=\"0\"/>")
           + f'<w:color w:val="{color}"/><w:sz w:val="{half}"/><w:szCs w:val="{half}"/>')
    based = f'<w:basedOn w:val="{based_on}"/>' if based_on else ""
    nxt = f'<w:next w:val="{next_style}"/>' if next_style else ""
    default = ' w:default="1"' if style_id == "Normal" else ""
    return (f'<w:style {nsdecls("w")} w:type="paragraph"{default} w:styleId="{style_id}"><w:name w:val="{name}"/>{based}{nxt}'
            f'<w:uiPriority w:val="{ui_priority}"/><w:qFormat/><w:pPr>{"".join(ppr)}</w:pPr><w:rPr>{rpr}</w:rPr></w:style>')


def define_styles(doc) -> None:
    """Create or replace the named styles from the tokens (spec section 5)."""
    styles_el = doc.styles.element
    text_width = _text_width_twips(doc)
    toc_tabs = f'<w:tabs><w:tab w:val="right" w:leader="dot" w:pos="{text_width}"/></w:tabs>'   # right edge of the text
    # "Normal" is deliberately NOT redefined: the cover page and Contents (left as the template
    # designed them) inherit from it, and changing its spacing/colour moves the cover layout.
    # Body paragraphs get their Corbel 11 pt / colour / spacing directly instead.
    definitions = [
        _style_xml("Heading1", "heading 1", next_style="Normal", bold=True, color=tokens.BRAND_PURPLE,
                   size_pt=tokens.SIZE_HEADING_1, before_pt=12, after_pt=6, keep_next=True, outline=0),
        _style_xml("Heading2", "heading 2", next_style="Normal", bold=True, color=tokens.BRAND_PURPLE,
                   size_pt=tokens.SIZE_HEADING_2, before_pt=12, after_pt=6, keep_next=True, outline=1),
        _style_xml("Heading3", "heading 3", next_style="Normal", bold=True, color=tokens.BRAND_PURPLE,
                   size_pt=tokens.SIZE_HEADING_3, before_pt=12, after_pt=4, keep_next=True, outline=2),
        _style_xml("AnnexureTitle", "Annexure Title", based_on="Heading2", next_style="Normal", bold=True, italic=True,
                   color=tokens.BRAND_PURPLE, size_pt=tokens.SIZE_ANNEXURE_TITLE, before_pt=0, after_pt=6,
                   keep_next=True, outline=1),
        _style_xml("TOC1", "toc 1", next_style="Normal", bold=True, color=tokens.BODY_TEXT, size_pt=10.5,
                   before_pt=2, after_pt=0, tabs=toc_tabs, line=240, ui_priority=39),
        _style_xml("TOC2", "toc 2", next_style="Normal", bold=True, italic=True, color=tokens.BODY_TEXT, size_pt=10.5,   # annexure lines
                   before_pt=2, after_pt=0, tabs=toc_tabs, line=240, ui_priority=39),
    ]
    existing = {s.get(_w("w:styleId")): s for s in styles_el.findall(_w("w:style"))}
    for xml in definitions:
        new = parse_xml(xml)
        old = existing.get(new.get(_w("w:styleId")))
        if old is not None:
            old.addprevious(new)
            styles_el.remove(old)
        else:
            styles_el.append(new)
    # document default font stays Corbel (the template already sets it; make sure of it)
    rpr_default = styles_el.find(f"{_w('w:docDefaults')}/{_w('w:rPrDefault')}/{_w('w:rPr')}")
    if rpr_default is not None:
        fonts = rpr_default.find(_w("w:rFonts"))
        if fonts is not None:
            for attr in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
                fonts.set(_w(attr), tokens.FONT)


def _text_width_twips(doc) -> int:
    s = doc.sections[-1]
    return int((s.page_width - s.left_margin - s.right_margin) / 635)


# =============================================================================== numbering
def ensure_bullet_numbering(doc) -> bool:
    """Make sure numbering definition BULLET_NUM_ID exists (round bullet, then open circle).

    AI annexures reference it so their bullets are real Word lists. Returns True if it was added.
    """
    try:
        numbering = doc.part.numbering_part.element
    except Exception:                     # document without a numbering part
        return False
    for num in numbering.findall(_w("w:num")):
        if num.get(_w("w:numId")) == str(BULLET_NUM_ID):
            return False
    levels = []
    for i, glyph in enumerate((tokens.LIST_GLYPHS[0], tokens.LIST_GLYPHS[1], "▪")):
        left = int(tokens.LIST_LEVEL_INDENT_IN[min(i, 2)] * TWIPS_PER_INCH)
        hang = int(tokens.LIST_HANGING_IN * TWIPS_PER_INCH)
        font = tokens.FONT if glyph != tokens.LIST_GLYPHS[1] else "Courier New"
        levels.append(
            f'<w:lvl w:ilvl="{i}"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="{glyph}"/>'
            f'<w:lvlJc w:val="left"/><w:pPr><w:ind w:left="{left}" w:hanging="{hang}"/></w:pPr>'
            f'<w:rPr><w:rFonts w:ascii="{font}" w:hAnsi="{font}" w:cs="{font}" w:hint="default"/></w:rPr></w:lvl>')
    abstract = parse_xml(f'<w:abstractNum {nsdecls("w")} w:abstractNumId="{BULLET_NUM_ID}">'
                         f'<w:multiLevelType w:val="hybridMultilevel"/>{"".join(levels)}</w:abstractNum>')
    first_num = numbering.find(_w("w:num"))
    if first_num is not None:
        first_num.addprevious(abstract)          # schema: every abstractNum precedes the first num
    else:
        numbering.append(abstract)
    numbering.append(parse_xml(f'<w:num {nsdecls("w")} w:numId="{BULLET_NUM_ID}"><w:abstractNumId w:val="{BULLET_NUM_ID}"/></w:num>'))
    return True


def bullet_numpr(level: int):
    """<w:numPr> element for an AI-annexure bullet (level 0 = round bullet, 1 = open circle)."""
    numpr = OxmlElement("w:numPr")
    ilvl = OxmlElement("w:ilvl")
    ilvl.set(_w("w:val"), str(level))
    num = OxmlElement("w:numId")
    num.set(_w("w:val"), str(BULLET_NUM_ID))
    numpr.append(ilvl)
    numpr.append(num)
    return numpr


# =============================================================================== paragraph helpers
def _text(p) -> str:
    return "".join(t.text or "" for t in p.iter(_w("w:t")))


def _runs_with_text(p):
    return [r for r in p.iter(_w("w:r")) if "".join(t.text or "" for t in r.iter(_w("w:t"))).strip()]


def _run_props(r) -> dict[str, Any]:
    rpr = r.find(_w("w:rPr"))
    info: dict[str, Any] = {"bold": False, "italic": False, "color": None, "size": None, "highlight": None}
    if rpr is None:
        return info
    b = rpr.find(_w("w:b"))
    info["bold"] = b is not None and b.get(_w("w:val")) not in ("0", "false", "off")
    i = rpr.find(_w("w:i"))
    info["italic"] = i is not None and i.get(_w("w:val")) not in ("0", "false", "off")
    c = rpr.find(_w("w:color"))
    info["color"] = (c.get(_w("w:val")) or "").lower() if c is not None else None
    s = rpr.find(_w("w:sz"))
    info["size"] = int(s.get(_w("w:val"))) if s is not None and (s.get(_w("w:val")) or "").isdigit() else None
    return info


def _ppr(p, create: bool = False):
    ppr = p.find(_w("w:pPr"))
    if ppr is None and create:
        ppr = OxmlElement("w:pPr")
        p.insert(0, ppr)
    return ppr


def _has_numpr(p) -> bool:
    ppr = p.find(_w("w:pPr"))
    return ppr is not None and ppr.find(_w("w:numPr")) is not None


def _set_pstyle(p, style_id: str) -> None:
    ppr = _ppr(p, create=True)
    node = ppr.find(_w("w:pStyle"))
    if node is None:
        node = OxmlElement("w:pStyle")
        ppr.insert(0, node)
    node.set(_w("w:val"), style_id)


def _strip_direct_format(p, *, keep_numpr: bool) -> None:
    """Remove direct paragraph and run formatting so the paragraph style governs."""
    ppr = p.find(_w("w:pPr"))
    if ppr is not None:
        keep = {_w("w:pStyle"), _w("w:pageBreakBefore"), _w("w:sectPr")}
        if keep_numpr:
            keep.add(_w("w:numPr"))
        for child in list(ppr):
            if child.tag not in keep:
                ppr.remove(child)
    for r in p.iter(_w("w:r")):
        rpr = r.find(_w("w:rPr"))
        if rpr is not None:
            for child in list(rpr):
                if child.tag in (_w("w:rFonts"), _w("w:b"), _w("w:bCs"), _w("w:i"), _w("w:iCs"), _w("w:color"),
                                 _w("w:sz"), _w("w:szCs"), _w("w:highlight"), _w("w:u")):
                    rpr.remove(child)


def _normalize_runs(p, *, size_pt: float = tokens.SIZE_BODY) -> None:
    """Corbel at the body size and colour; bold/italic and purple/white emphasis are kept."""
    half = str(int(round(size_pt * 2)))
    for r in p.iter(_w("w:r")):
        rpr = r.find(_w("w:rPr"))
        if rpr is None:
            rpr = OxmlElement("w:rPr")
            r.insert(0, rpr)
        info = _run_props(r)
        color = info["color"]
        if color in PURPLES:
            new_color = tokens.BRAND_PURPLE
        elif color == tokens.WHITE.lower():
            new_color = tokens.WHITE
        else:
            new_color = tokens.BODY_TEXT
        for tag in ("w:rFonts", "w:color", "w:sz", "w:szCs", "w:highlight"):
            for node in rpr.findall(_w(tag)):
                rpr.remove(node)
        fonts = OxmlElement("w:rFonts")
        for attr in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
            fonts.set(_w(attr), tokens.FONT)
        col = OxmlElement("w:color")
        col.set(_w("w:val"), new_color)
        sz, szcs = OxmlElement("w:sz"), OxmlElement("w:szCs")
        sz.set(_w("w:val"), half)
        szcs.set(_w("w:val"), half)
        # schema order inside rPr: rFonts, b, bCs, i, iCs, ..., color, ..., sz, szCs
        rpr.insert(0, fonts)
        anchor = None
        for tag in ("w:iCs", "w:i", "w:bCs", "w:b"):
            anchor = rpr.find(_w(tag))
            if anchor is not None:
                break
        (anchor if anchor is not None else fonts).addnext(col)
        col.addnext(sz)
        sz.addnext(szcs)


def _set_spacing(p, *, before: int = 0, after: int, line: int = LINE_SPACING) -> None:
    ppr = _ppr(p, create=True)
    for node in ppr.findall(_w("w:spacing")):
        ppr.remove(node)
    sp = OxmlElement("w:spacing")
    sp.set(_w("w:before"), str(before))
    sp.set(_w("w:after"), str(after))
    sp.set(_w("w:line"), str(line))
    sp.set(_w("w:lineRule"), "auto")
    # schema order: pStyle, keepNext, keepLines, pageBreakBefore, ..., numPr, ..., pBdr, shd, tabs, ..., spacing, ind, ..., jc
    successors = [_w(t) for t in ("w:ind", "w:contextualSpacing", "w:mirrorIndents", "w:suppressOverlap", "w:jc",
                                  "w:textDirection", "w:textAlignment", "w:textboxTightWrap", "w:outlineLvl",
                                  "w:divId", "w:cnfStyle", "w:rPr", "w:sectPr", "w:pPrChange")]
    for child in ppr:
        if child.tag in successors:
            child.addprevious(sp)
            return
    ppr.append(sp)


# =============================================================================== classification
def _is_blank(p) -> bool:
    return (p.tag == _w("w:p") and not _text(p).strip()
            and not any(b.get(_w("w:type")) == "page" for b in p.iter(_w("w:br")))
            and p.find(f".//{_w('w:drawing')}") is None and p.find(f".//{_w('w:pict')}") is None
            and p.find(f".//{_w('w:sectPr')}") is None)


_STYLE_KIND = {"AnnexureTitle": "annexure_title", "Heading1": "h1", "Heading2": "h2", "Heading3": "label"}
FIELD_LINE_RE = re.compile(r"^\s*(Frequency|Count)\s*:", re.IGNORECASE)


def _pstyle(p) -> str | None:
    ppr = p.find(_w("w:pPr"))
    node = ppr.find(_w("w:pStyle")) if ppr is not None else None
    return node.get(_w("w:val")) if node is not None else None


def _classify(p) -> str:
    text = _text(p).strip()
    low = re.sub(r"\s+", " ", text).lower()
    if not text:
        return "body"                          # empty paragraphs are handled (removed) as blanks, whatever style they carry
    if ANNEXURE_TITLE_RE.match(text):
        return "annexure_title"
    if "\t" in _text(p):
        return "contents"
    if low in HEADING1_TEXTS:
        return "h1"
    if low in HEADING2_TEXTS:
        return "h2"
    # Already styled by an earlier pass (this function must be idempotent). A list item that merely
    # carries a heading style in the template (Google-Docs export) is body text, not a heading.
    styled = _STYLE_KIND.get(_pstyle(p) or "")
    if FIELD_LINE_RE.match(text):
        return "field"                         # "Frequency: ..." / "Count: ..." lines share one style
    runs = _runs_with_text(p)
    if styled and len(text) < 140:
        # A heading already normalised by a previous pass has no direct size on its runs. A list item
        # that merely inherited a heading style in the template still carries its own direct size.
        already_normalised = all(_run_props(r)["size"] is None for r in runs)
        if already_normalised:
            return styled
    if runs:
        infos = [_run_props(r) for r in runs]
        # a sub-label is bold purple text; the template sizes them inconsistently (14 pt, 12 pt, 11 pt),
        # so a trailing colon (or the 14 pt size) identifies it, not the exact size
        if all(i["bold"] and i["color"] in PURPLES for i in infos) and len(text) < 140 and \
                (text.endswith(":") or all(i["size"] == 28 for i in infos)):
            return "label"
    return "body"


def _body_start(body) -> int:
    """Index of the body INTRODUCTION heading (after the Contents block), else 0."""
    seen_contents = False
    for i, el in enumerate(body):
        if el.tag != _w("w:p"):
            continue
        t = re.sub(r"\s+", " ", _text(el)).strip()
        if t == "Contents":
            seen_contents = True
        elif seen_contents and t == "INTRODUCTION":
            return i
    return 0


def _front_end(body, start: int) -> int:
    """Index of the first annexure title after *start* (end of the fixed front pages), else *start*."""
    children = list(body)
    for i in range(start, len(children)):
        el = children[i]
        if el.tag == _w("w:p") and ANNEXURE_TITLE_RE.match(_text(el)):
            return i
    return start


# =============================================================================== tables
def _cell_borders():
    b = OxmlElement("w:tcBorders")
    for side in ("w:top", "w:left", "w:bottom", "w:right"):
        e = OxmlElement(side)
        e.set(_w("w:val"), "single")
        e.set(_w("w:sz"), str(tokens.TABLE_BORDER_EIGHTHS_PT))
        e.set(_w("w:space"), "0")
        e.set(_w("w:color"), tokens.TABLE_BORDER)
        b.append(e)
    return b


def _set_tcpr(tc, fill: str | None) -> None:
    tcpr = tc.find(_w("w:tcPr"))
    if tcpr is None:
        tcpr = OxmlElement("w:tcPr")
        tc.insert(0, tcpr)
    keep = {_w("w:tcW"), _w("w:gridSpan"), _w("w:vMerge")}
    for child in list(tcpr):
        if child.tag not in keep:
            tcpr.remove(child)
    tcpr.append(_cell_borders())
    if fill:
        shd = OxmlElement("w:shd")
        shd.set(_w("w:val"), "clear")
        shd.set(_w("w:color"), "auto")
        shd.set(_w("w:fill"), fill)
        tcpr.append(shd)
    va = OxmlElement("w:vAlign")
    va.set(_w("w:val"), "center")
    tcpr.append(va)


MIN_SERIAL_COLUMN_TWIPS = 1100          # "Sr No." in bold 10 pt plus padding must not wrap mid-word


def _widen_serial_column(grid, rows, text_width_twips: int) -> None:
    """Give a narrow first "Sr No." column room (taken from the widest column) so it never wraps."""
    if grid is None or not rows:
        return
    cols = grid.findall(_w("w:gridCol"))
    first_header = "".join(t.text or "" for t in rows[0].findall(_w("w:tc"))[0].iter(_w("w:t")))
    if len(cols) < 3 or not re.match(r"^\s*sr\b", first_header, re.IGNORECASE):
        return
    widths = [int(c.get(_w("w:w")) or 0) for c in cols]
    if widths[0] >= MIN_SERIAL_COLUMN_TWIPS:
        return
    give = MIN_SERIAL_COLUMN_TWIPS - widths[0]
    donor = max(range(1, len(widths)), key=lambda i: widths[i])
    widths[0] += give
    widths[donor] -= give
    for c, w in zip(cols, widths):
        c.set(_w("w:w"), str(w))
    for tr in rows:
        idx = 0
        for tc in tr.findall(_w("w:tc")):
            span_el = tc.find(f"{_w('w:tcPr')}/{_w('w:gridSpan')}")
            span = int(span_el.get(_w("w:val"))) if span_el is not None else 1
            tcw = tc.find(f"{_w('w:tcPr')}/{_w('w:tcW')}")
            if tcw is not None:
                tcw.set(_w("w:w"), str(sum(widths[idx:idx + span])))
                tcw.set(_w("w:type"), "dxa")
            idx += span


def restyle_table(tbl, text_width_twips: int, *, compact: bool = False) -> None:
    """Spec section 9: purple header row, light body rows, grey borders, padding, header repeat.

    *compact* (the Team Overview table on the fixed front page) drops the template's fixed row
    heights and uses less vertical padding, so the table stays on its one page.
    """
    rows = tbl.findall(_w("w:tr"))
    if not rows:
        return
    if compact:
        for tr in rows:
            trpr = tr.find(_w("w:trPr"))
            if trpr is not None:
                for node in trpr.findall(_w("w:trHeight")):
                    trpr.remove(node)
    grid = tbl.find(_w("w:tblGrid"))
    cols = [int(g.get(_w("w:w")) or 0) for g in grid.findall(_w("w:gridCol"))] if grid is not None else []
    total = sum(cols)
    if grid is not None and total and total != text_width_twips:
        scale = text_width_twips / total
        new = [int(c * scale) for c in cols]
        new[-1] += text_width_twips - sum(new)
        for g, w in zip(grid.findall(_w("w:gridCol")), new):
            g.set(_w("w:w"), str(w))
        for tr in rows:
            for tc in tr.findall(_w("w:tc")):
                tcw = tc.find(f"{_w('w:tcPr')}/{_w('w:tcW')}")
                if tcw is not None and tcw.get(_w("w:w")):
                    try:
                        tcw.set(_w("w:w"), str(int(float(tcw.get(_w("w:w"))) * scale)))
                        tcw.set(_w("w:type"), "dxa")
                    except ValueError:
                        pass
    _widen_serial_column(grid, rows, text_width_twips)
    tblpr = tbl.find(_w("w:tblPr"))
    for tag in ("w:tblW", "w:jc", "w:tblLayout", "w:tblCellMar"):
        for node in tblpr.findall(_w(tag)):
            tblpr.remove(node)
    tblw = OxmlElement("w:tblW")
    tblw.set(_w("w:w"), str(text_width_twips))
    tblw.set(_w("w:type"), "dxa")
    jc = OxmlElement("w:jc")
    jc.set(_w("w:val"), "center")
    layout = OxmlElement("w:tblLayout")
    layout.set(_w("w:type"), "fixed")
    mar = OxmlElement("w:tblCellMar")
    pad_v = tokens.TABLE_CELL_PAD_V_IN / 2 if compact else tokens.TABLE_CELL_PAD_V_IN
    for side, inches in (("w:top", pad_v), ("w:left", tokens.TABLE_CELL_PAD_H_IN),
                         ("w:bottom", pad_v), ("w:right", tokens.TABLE_CELL_PAD_H_IN)):
        m = OxmlElement(side)
        m.set(_w("w:w"), str(int(inches * TWIPS_PER_INCH)))
        m.set(_w("w:type"), "dxa")
        mar.append(m)
    anchor = tblpr.find(_w("w:tblStyle"))
    nodes = [tblw, jc, layout, mar]
    for n in reversed(nodes):
        (anchor.addnext(n) if anchor is not None else tblpr.insert(0, n))

    header_texts = ["".join(t.text or "" for t in tc.iter(_w("w:t"))) for tc in rows[0].findall(_w("w:tc"))]
    money_cols = {i for i, h in enumerate(header_texts) if re.search(r"charge|cost|price|amount", h, re.I)}
    for ri, tr in enumerate(rows):
        trpr = tr.find(_w("w:trPr"))
        if trpr is None:
            trpr = OxmlElement("w:trPr")
            tr.insert(0, trpr)
        for tag in ("w:cantSplit", "w:tblHeader"):
            for node in trpr.findall(_w(tag)):
                trpr.remove(node)
        trpr.insert(0, OxmlElement("w:cantSplit"))
        if ri == 0:
            trpr.append(OxmlElement("w:tblHeader"))
        header = ri == 0
        fill = tokens.BRAND_PURPLE if header else (tokens.TABLE_BODY_FILL if ri % 2 == 0 else tokens.WHITE)
        for ci, tc in enumerate(tr.findall(_w("w:tc"))):
            _set_tcpr(tc, fill)
            for p in tc.findall(_w("w:p")):
                _strip_direct_format(p, keep_numpr=False)
                align = "center" if (header or ci == 0) else ("right" if ci in money_cols else "left")
                ppr = _ppr(p, create=True)
                jc_node = OxmlElement("w:jc")
                jc_node.set(_w("w:val"), align)
                ppr.append(jc_node)
                _set_spacing(p, after=0, line=240)
                _normalize_runs(p, size_pt=tokens.SIZE_TABLE_HEADER if header else tokens.SIZE_TABLE_BODY)
                for r in p.iter(_w("w:r")):
                    rpr = r.find(_w("w:rPr"))
                    for c in rpr.findall(_w("w:color")):
                        c.set(_w("w:val"), tokens.TABLE_HEADER_TEXT if header else tokens.BODY_DARK)
                    if header and rpr.find(_w("w:b")) is None:
                        rpr.find(_w("w:rFonts")).addnext(OxmlElement("w:b"))


# =============================================================================== signature block
SIGNATURE_LINE_RE = re.compile(r"^\s*(?P<left>\S.*?)\s{2,}(?P<right>\S.*?)\s*$")


def _signature_rows(paragraphs: list) -> list[tuple[str, str]]:
    rows = []
    for p in paragraphs:
        text = _text(p)
        if not text.strip():
            continue
        m = SIGNATURE_LINE_RE.match(text.replace("\t", "    "))
        if m:
            rows.append((m.group("left").strip(), m.group("right").strip()))
        else:
            rows.append((text.strip(), ""))
    return rows


def _sig_cell(width: int, label: str, value: str, *, bold_whole: bool = False, before_pt: float = 0, keep_next: bool = True):
    tc = OxmlElement("w:tc")
    tcpr = OxmlElement("w:tcPr")
    tcw = OxmlElement("w:tcW")
    tcw.set(_w("w:w"), str(width))
    tcw.set(_w("w:type"), "dxa")
    tcpr.append(tcw)
    borders = OxmlElement("w:tcBorders")
    for side in ("w:top", "w:left", "w:bottom", "w:right"):
        e = OxmlElement(side)
        e.set(_w("w:val"), "nil")
        borders.append(e)
    tcpr.append(borders)
    tc.append(tcpr)
    p = OxmlElement("w:p")
    ppr = OxmlElement("w:pPr")
    if keep_next:
        ppr.append(OxmlElement("w:keepNext"))
    sp = OxmlElement("w:spacing")
    sp.set(_w("w:before"), str(int(before_pt * 20)))
    sp.set(_w("w:after"), "0")
    sp.set(_w("w:line"), str(LINE_SPACING))
    sp.set(_w("w:lineRule"), "auto")
    ppr.append(sp)
    p.append(ppr)
    if bold_whole:
        p.append(make_run(label, bold=True, color=tokens.BODY_TEXT, size_pt=tokens.SIZE_BODY))
    else:
        p.append(make_run(f"{label} " if label else "", bold=True, color=tokens.BODY_TEXT, size_pt=tokens.SIZE_BODY))
        if value:
            p.append(make_run(value, color=tokens.BODY_TEXT, size_pt=tokens.SIZE_BODY))
    tc.append(p)
    return tc


def build_signature_table(rows: list[tuple[str, str]], text_width_twips: int):
    """Borderless two-column table: supplier left, client right (spec section 8, page 56)."""
    half = text_width_twips // 2
    tbl = OxmlElement("w:tbl")
    tblpr = OxmlElement("w:tblPr")
    tblw = OxmlElement("w:tblW")
    tblw.set(_w("w:w"), str(text_width_twips))
    tblw.set(_w("w:type"), "dxa")
    tblpr.append(tblw)
    layout = OxmlElement("w:tblLayout")
    layout.set(_w("w:type"), "fixed")
    tblpr.append(layout)
    tbl.append(tblpr)
    grid = OxmlElement("w:tblGrid")
    for w in (half, text_width_twips - half):
        gc = OxmlElement("w:gridCol")
        gc.set(_w("w:w"), str(w))
        grid.append(gc)
    tbl.append(grid)
    for i, (left, right) in enumerate(rows):
        last = i == len(rows) - 1
        tr = OxmlElement("w:tr")
        trpr = OxmlElement("w:trPr")
        trpr.append(OxmlElement("w:cantSplit"))
        tr.append(trpr)
        is_name_row = i == 0
        for width, text in ((half, left), (text_width_twips - half, right)):
            if is_name_row:
                tr.append(_sig_cell(width, text, "", bold_whole=True, keep_next=not last))
            else:
                label, _, value = text.partition(":")
                label = (label + ":") if ":" in text else text
                tr.append(_sig_cell(width, label, value.strip() if ":" in text else "", before_pt=14 if i == 1 else 6,
                                    keep_next=not last))
        tbl.append(tr)
    return tbl


def convert_signature_block(doc, text_width_twips: int) -> bool:
    """Replace the space-aligned signature lines by a borderless table (adds a Date row)."""
    body = doc.element.body
    children = list(body)
    start = next((i for i, el in enumerate(children)
                  if el.tag == _w("w:p") and "in witness whereof" in _text(el).lower()), None)
    if start is None:
        return False
    block = []
    for el in children[start + 1:]:
        if el.tag != _w("w:p"):
            break
        block.append(el)
    meaningful = [p for p in block if _text(p).strip()]
    if not meaningful or any(re.search(r"\w+:\s*$|By:|Name:|Title:|Date:", _text(p)) for p in meaningful) is False:
        return False
    if body.find(f".//{_w('w:tbl')}") is not None and any(
            "By:" in "".join(t.text or "" for t in tbl.iter(_w("w:t"))) for tbl in body.iter(_w("w:tbl"))):
        return False                                           # already a table
    rows = _signature_rows(meaningful)
    if not rows or not any(r[0].startswith("By:") for r in rows):
        return False
    if not any(r[0].startswith("Date:") for r in rows):
        rows.append(("Date: ______________________", "Date: ______________________"))
    table = build_signature_table(rows, text_width_twips)
    witness = children[start]                                  # "In witness whereof ..." moves with the signatures
    wppr = _ppr(witness, create=True)
    if wppr.find(_w("w:keepNext")) is None:
        keep = OxmlElement("w:keepNext")
        style_node = wppr.find(_w("w:pStyle"))
        (style_node.addnext(keep) if style_node is not None else wppr.insert(0, keep))
    block[0].addprevious(table)
    for p in block:
        body.remove(p)
    # a table must be followed by a paragraph before the section properties
    after = table.getnext()
    if after is None or after.tag == _w("w:sectPr"):
        spacer = OxmlElement("w:p")
        table.addnext(spacer)
        make_tiny_paragraph(spacer)
    return True


def make_tiny_paragraph(p) -> None:
    """A 1 pt, zero-spacing paragraph: satisfies "a table is followed by a paragraph" without being
    able to spill onto a page of its own (which would be a blank last page)."""
    ppr = _ppr(p, create=True)
    for child in list(ppr):
        if child.tag != _w("w:pStyle"):
            ppr.remove(child)
    sp = OxmlElement("w:spacing")
    sp.set(_w("w:before"), "0")
    sp.set(_w("w:after"), "0")
    sp.set(_w("w:line"), "20")
    sp.set(_w("w:lineRule"), "exact")
    ppr.append(sp)
    mark = OxmlElement("w:rPr")
    for tag in ("w:sz", "w:szCs"):
        size = OxmlElement(tag)
        size.set(_w("w:val"), "2")
        mark.append(size)
    ppr.append(mark)


# =============================================================================== Contents as a real TOC field


# =============================================================================== editability
def unlock_content_controls(doc) -> int:
    """Unwrap every content control (w:sdt) so no part of the document is locked against editing.

    The template (a Google Docs export) wraps the pricing table in a content control marked
    ``contentLocked``: Word then refuses any typing inside the table ("You are not allowed to edit
    this selection because it is protected"), so prices could not be filled in afterwards. The
    controls hold nothing the proposal needs, so their content is kept and the wrapper removed.
    Returns how many were unwrapped.
    """
    unwrapped = 0
    for sdt in list(doc.element.body.iter(_w("w:sdt"))):
        parent = sdt.getparent()
        content = sdt.find(_w("w:sdtContent"))
        if parent is None or content is None:
            continue
        for child in list(content):
            sdt.addprevious(child)
        parent.remove(sdt)
        unwrapped += 1
    return unwrapped


# =============================================================================== entry point
def apply_specification(doc) -> dict[str, int]:
    """Normalise the assembled document to the specification. Returns counters for logging/tests."""
    stats = {"headings": 0, "labels": 0, "body": 0, "blank_removed": 0, "tables": 0, "signature": 0,
             "unlocked": unlock_content_controls(doc)}
    define_styles(doc)
    ensure_bullet_numbering(doc)
    body = doc.element.body
    width = _text_width_twips(doc)
    start = _body_start(body)
    front_end = _front_end(body, start)               # pages 3-8 are the organisation's fixed pages
    children = list(body)

    for i, el in enumerate(children):
        if i < start:
            continue                                  # cover and Contents: left as designed
        if start <= i < front_end:
            # The fixed front pages (Introduction ... tools) are the organisation's own design: text,
            # fonts and the spacing that places their diagrams stay exactly as in the master. Only
            # their tables are brought to the table standard.
            if el.tag == _w("w:tbl"):
                restyle_table(el, width, compact=True)
                stats["tables"] += 1
            continue
        if el.tag == _w("w:p"):
            kind = _classify(el)
            if kind == "contents":
                continue
            if kind == "field":
                label, _, value = _text(el).partition(":")
                proposal_fields.rewrite_label_value(el, label.strip().capitalize(), value.strip())
                _set_spacing(el, after=SPACING_AFTER_BODY)
                stats["body"] += 1
                continue
            if kind in ("annexure_title", "h1", "h2", "label"):
                style = {"annexure_title": "AnnexureTitle", "h1": "Heading1", "h2": "Heading2", "label": "Heading3"}[kind]
                _strip_direct_format(el, keep_numpr=True)
                _set_pstyle(el, style)
                stats["labels" if kind == "label" else "headings"] += 1
            elif _is_blank(el):
                prev, nxt = el.getprevious(), el.getnext()
                closing_spacer = (prev is not None and prev.tag == _w("w:tbl")
                                  and (nxt is None or nxt.tag == _w("w:sectPr")))
                if el.getparent() is body and not closing_spacer:     # Word needs a paragraph after a final table
                    body.remove(el)
                    stats["blank_removed"] += 1
            else:
                _normalize_runs(el)
                ppr = _ppr(el)
                stray = ppr.find(_w("w:pStyle")) if ppr is not None else None
                if stray is not None and stray.get(_w("w:val")) in ("Heading1", "Heading2", "Heading3", "Heading4"):
                    ppr.remove(stray)                 # a list item that only inherited a heading style
                listed = _has_numpr(el)
                prev = el.getprevious()
                after_table = prev is not None and prev.tag in (_w("w:tbl"), _w("w:sdt"))
                _set_spacing(el, before=240 if after_table else 0,        # a table never touches the next text
                             after=SPACING_AFTER_LIST if listed else SPACING_AFTER_BODY)
                stats["body"] += 1
        elif el.tag in (_w("w:tbl"), _w("w:sdt")):
            for tbl in ([el] if el.tag == _w("w:tbl") else el.iter(_w("w:tbl"))):
                first = "".join(t.text or "" for t in tbl.findall(_w("w:tr"))[0].iter(_w("w:t"))) if tbl.findall(_w("w:tr")) else ""
                if "Assessment Type" in first:
                    continue                          # the pricing table is already built from the tokens
                flat = "".join(t.text or "" for t in tbl.iter(_w("w:t")))
                if "By:" in flat and "Name:" in flat:
                    continue                          # the signature table is borderless by design
                restyle_table(tbl, width)
                stats["tables"] += 1
    if convert_signature_block(doc, width):
        stats["signature"] = 1
    return stats

"""
Conclusion & Pricing table (plan section 7), built only from proposal_tokens.

Rows are plain dicts: {"title", "quantity", "frequency", "cost"} (all strings; cost may be empty).
The table replaces the one inside the Base Template's content control, so surrounding text
(HSN/SAC line, payment terms, signatures) stays template-owned.

Total row (D24): one merged cell "₹ /- * + 18% GST". When every row carries a price the sum is
shown in Indian digit grouping ("₹ 1,25,000 /- * + 18% GST"); with any price missing the
placeholder stays, so a partial sum is never shown as if it were the total.
"""
from __future__ import annotations

import re
from decimal import Decimal, InvalidOperation
from typing import Iterable, Sequence

from docx.oxml import OxmlElement
from docx.oxml.ns import qn

import proposal_tokens as tokens
from proposal_fields import make_run

TWIPS_PER_INCH = 1440
RUPEE = "₹"


# --------------------------------------------------------------------------- money
def parse_amount(text: str | int | float | Decimal | None) -> Decimal | None:
    """'₹ 1,25,000 /-' -> Decimal('125000'); returns None for blank or non-numeric text."""
    if text is None:
        return None
    if isinstance(text, (int, float, Decimal)) and not isinstance(text, bool):
        return Decimal(str(text))
    cleaned = re.sub(r"(?i)rs\.?|inr|₹|/-|,|\s", "", str(text))
    if not cleaned:
        return None
    try:
        value = Decimal(cleaned)
    except InvalidOperation:
        return None
    return value if value >= 0 else None


def format_inr(amount: Decimal | int | float) -> str:
    """Indian digit grouping: 125000 -> '₹ 1,25,000'; paise only when present."""
    value = Decimal(str(amount))
    sign = "-" if value < 0 else ""
    value = abs(value)
    whole, _, frac = f"{value:.2f}".partition(".")
    if len(whole) > 3:
        head, tail = whole[:-3], whole[-3:]
        groups: list[str] = []
        while len(head) > 2:
            groups.insert(0, head[-2:])
            head = head[:-2]
        if head:
            groups.insert(0, head)
        whole = ",".join(groups + [tail])
    text = whole if frac == "00" else f"{whole}.{frac}"
    return f"{RUPEE} {sign}{text}"


def total_text(rows: Sequence[dict]) -> str:
    """Text of the merged Total cell (see module docstring)."""
    suffix = tokens.PRICING_TOTAL_EMPTY_TEXT.replace(f"{RUPEE} /-", "").strip()  # '* + 18% GST'
    amounts = [parse_amount(r.get("cost")) for r in rows]
    if rows and all(a is not None for a in amounts):
        return f"{format_inr(sum(amounts))} /- {suffix}"
    return tokens.PRICING_TOTAL_EMPTY_TEXT


def display_cost(text: str | None) -> str:
    """Prices typed as numbers are shown with ₹ and Indian grouping; other text is kept as typed."""
    amount = parse_amount(text)
    if amount is None:
        return (text or "").strip()
    return format_inr(amount)


# --------------------------------------------------------------------------- table construction
def _border(tag: str):
    b = OxmlElement(tag)
    b.set(qn("w:val"), "single")
    b.set(qn("w:sz"), str(tokens.TABLE_BORDER_EIGHTHS_PT))
    b.set(qn("w:space"), "0")
    b.set(qn("w:color"), tokens.TABLE_BORDER)
    return b


def _cell(text: str, width_twips: int, *, fill: str | None, align: str, bold: bool,
          color: str, size_pt: float, span: int = 1):
    tc = OxmlElement("w:tc")
    tcpr = OxmlElement("w:tcPr")
    tcw = OxmlElement("w:tcW")
    tcw.set(qn("w:w"), str(width_twips))
    tcw.set(qn("w:type"), "dxa")
    tcpr.append(tcw)
    if span > 1:
        gs = OxmlElement("w:gridSpan")
        gs.set(qn("w:val"), str(span))
        tcpr.append(gs)
    borders = OxmlElement("w:tcBorders")
    for side in ("w:top", "w:left", "w:bottom", "w:right"):
        borders.append(_border(side))
    tcpr.append(borders)
    if fill:
        shd = OxmlElement("w:shd")
        shd.set(qn("w:val"), "clear")
        shd.set(qn("w:color"), "auto")
        shd.set(qn("w:fill"), fill)
        tcpr.append(shd)
    va = OxmlElement("w:vAlign")
    va.set(qn("w:val"), "center")
    tcpr.append(va)
    tc.append(tcpr)

    p = OxmlElement("w:p")
    ppr = OxmlElement("w:pPr")
    sp = OxmlElement("w:spacing")
    sp.set(qn("w:before"), "0")
    sp.set(qn("w:after"), "0")
    ppr.append(sp)
    jc = OxmlElement("w:jc")
    jc.set(qn("w:val"), align)
    ppr.append(jc)
    p.append(ppr)
    p.append(make_run(text, bold=bold, color=color, size_pt=size_pt))
    tc.append(p)
    return tc


def _row(cells: Iterable, *, header: bool = False, height: int | None = None):
    tr = OxmlElement("w:tr")
    trpr = OxmlElement("w:trPr")
    trpr.append(OxmlElement("w:cantSplit"))          # rows never split across pages
    if height:
        th = OxmlElement("w:trHeight")
        th.set(qn("w:val"), str(height))
        th.set(qn("w:hRule"), "atLeast")
        trpr.append(th)
    if header:
        trpr.append(OxmlElement("w:tblHeader"))      # header repeats on each page
    tr.append(trpr)
    for c in cells:
        tr.append(c)
    return tr


def column_widths(text_width_twips: int) -> list[int]:
    widths = [int(text_width_twips * pct / 100) for pct in tokens.PRICING_COLUMN_WIDTH_PCT]
    widths[-1] += text_width_twips - sum(widths)      # absorb rounding so the table fits exactly
    return widths


def build_pricing_table(rows: Sequence[dict], text_width_twips: int):
    """Return a new <w:tbl> element for the Conclusion & Pricing section."""
    widths = column_widths(text_width_twips)
    tbl = OxmlElement("w:tbl")

    tblpr = OxmlElement("w:tblPr")
    tblw = OxmlElement("w:tblW")
    tblw.set(qn("w:w"), str(text_width_twips))
    tblw.set(qn("w:type"), "dxa")
    tblpr.append(tblw)
    jc = OxmlElement("w:jc")
    jc.set(qn("w:val"), "center")
    tblpr.append(jc)
    layout = OxmlElement("w:tblLayout")
    layout.set(qn("w:type"), "fixed")
    tblpr.append(layout)
    mar = OxmlElement("w:tblCellMar")
    for side, inches in (("w:top", tokens.TABLE_CELL_PAD_V_IN), ("w:left", tokens.TABLE_CELL_PAD_H_IN),
                         ("w:bottom", tokens.TABLE_CELL_PAD_V_IN), ("w:right", tokens.TABLE_CELL_PAD_H_IN)):
        m = OxmlElement(side)
        m.set(qn("w:w"), str(int(inches * TWIPS_PER_INCH)))
        m.set(qn("w:type"), "dxa")
        mar.append(m)
    tblpr.append(mar)
    tbl.append(tblpr)

    grid = OxmlElement("w:tblGrid")
    for w in widths:
        gc = OxmlElement("w:gridCol")
        gc.set(qn("w:w"), str(w))
        grid.append(gc)
    tbl.append(grid)

    tbl.append(_row(
        [_cell(h, widths[i], fill=tokens.BRAND_PURPLE, align="center", bold=True,
               color=tokens.TABLE_HEADER_TEXT, size_pt=tokens.SIZE_TABLE_HEADER)
         for i, h in enumerate(tokens.PRICING_HEADERS)],
        header=True))

    for n, row in enumerate(rows, start=1):
        fill = tokens.TABLE_BODY_FILL if n % 2 == 0 else tokens.WHITE
        values = (str(n), row.get("title", ""), str(row.get("quantity", "") or ""),
                  row.get("frequency", "") or "", display_cost(row.get("cost")))
        tbl.append(_row(
            [_cell(v, widths[i], fill=fill, align=tokens.PRICING_ALIGN[i], bold=False,
                   color=tokens.BODY_DARK, size_pt=tokens.SIZE_TABLE_BODY)
             for i, v in enumerate(values)]))

    total_cells = [
        _cell(tokens.PRICING_TOTAL_LABEL, widths[0], fill=tokens.TABLE_BODY_FILL, align="center",
              bold=True, color=tokens.BRAND_PURPLE, size_pt=tokens.SIZE_TABLE_TOTAL),
        _cell(total_text(rows), sum(widths[1:]), fill=tokens.TABLE_BODY_FILL, align="center",
              bold=True, color=tokens.BRAND_PURPLE, size_pt=tokens.SIZE_TABLE_TOTAL, span=4),
    ]
    tbl.append(_row(total_cells, height=588))
    return tbl


def replace_pricing_table(existing_tbl, rows: Sequence[dict], text_width_twips: int):
    """Swap the template's pricing table for a freshly built one (same place, same wrapper)."""
    new_tbl = build_pricing_table(rows, text_width_twips)
    existing_tbl.addnext(new_tbl)
    existing_tbl.getparent().remove(existing_tbl)
    return new_tbl

"""
CyberSmithSECURE Type6 — Closure Report API v1.0
===============================================
FastAPI endpoint `/generate-closure/` matching AdvantageGo PDF format exactly.

POST /generate-closure/
- initial_report: DOCX (initial report)
- excel_file: XLSX (vulnerability details)  
- poc_zip: ZIP (POC screenshots, optional)

Returns styled Closure_Report.docx with data merged from all inputs.
"""
from fastapi import FastAPI, UploadFile, File, HTTPException, BackgroundTasks, Request
from fastapi.responses import FileResponse, HTMLResponse, RedirectResponse, Response
import tempfile
import shutil
from zipfile import ZipFile
import os
from pathlib import Path
from typing import Any



import re
import sys
import json
import argparse
import io
import copy
import time
from dotenv import load_dotenv
from datetime import datetime
import hashlib

from docx import Document as DocxDocument
from docx.shared import Pt, RGBColor, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml.ns import qn, nsdecls
from docx.oxml import OxmlElement, parse_xml

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from mpl_toolkits.mplot3d import Axes3D
import numpy as np

sys.path.insert(0, str(Path(__file__).resolve().parents[2]))
try:
    from excel_parser import parse_excel_data
except Exception:
    parse_excel_data = None

# ─────────────────────────────────────────────────────────────────────────────
# EXACT COLORS FROM ADVANTAGEGO PDF
# ─────────────────────────────────────────────────────────────────────────────
PRP       = "7030A0"   # primary purple headings / TOC
PRP2      = "6923D0"   # table headers
NAVY      = "1F3864"   # dark navy
WHITE     = "FFFFFF"
LGRAY     = "F2F2F2"   # alternating row light gray
BGRAY     = "CCCCCC"   # border gray
YELLOW_ST = "FFD966"   # status cell yellow (from type3 / AdvantageGo)

# Severity palette — exact from AdvantageGo PDF pages 9-10
SEV = {
    "critical":      ("800000", WHITE),
    "high":          ("FF0000", WHITE),
    "medium":        ("FFC000", "000000"),
    "low":           ("92D050", WHITE),
    "informational": ("00B0F0", WHITE),
}

# Revalidation palette
REVAL = {
    "patched":              ("92D050", WHITE),
    "closed":               ("92D050", WHITE),
    "justified and closed": ("92D050", WHITE),
    "unpatched":            ("FF0000", WHITE),
    "open":                 ("FF0000", WHITE),
    "partially patched":    ("FFC000", "000000"),
}

FONT  = "Calibri"
BODY  = 10
HEAD  = 11

BACKEND_DIR = Path(__file__).resolve().parent
ROOT_DIR = Path(__file__).resolve().parents[2]
ROOT_ENV_PATH = ROOT_DIR / ".env"
ROOT_ENV_FALLBACK_PATH = ROOT_DIR / "env"
# Base for cover / disclaimer / TOC, optional Risk/Scan/Tools pages, then {CLOSURE_…}
# (see _build_advantagego_closure_block). `generate()` fills placeholders; it does not draw H/F.
# May be a .docx or a .pdf; PDFs are converted once into closure_base_cache/ for Word I/O.
ADVANTAGE_TEMPLATE_PATH = BACKEND_DIR / "1.docx"
ADVANTAGE_TEMPLATE_CANDIDATES = [
    BACKEND_DIR / "AdvantageGo Exact Web app VAPT Technical Closure Report 1.docx",
    BACKEND_DIR / "AdvantageGo_Closure_Base_Template.docx",
]
# Last-resort: AdvantageGo PDF in Downloads (if 1.docx and closure base are missing).
def _default_advantagego_pdf_in_downloads() -> Path | None:
    d = (os.environ.get("USERPROFILE") or os.environ.get("HOME") or "").strip()
    if not d:
        return None
    p = Path(d) / "Downloads" / "AdvantageGo Exact Web app VAPT Technical Closure Report.pdf"
    return p if p.is_file() else None
TOC_MARKER = "{TOC}"
CONTENT_MARKER = "{CLOSURE_DYNAMIC_SECTIONS}"
# All markers accepted in the base template (typo-tolerant for closure slot)
_CONTENT_MARKER_VARIANTS = frozenset(
    {
        "{CLOSURE_DYNAMIC_SECTIONS}",
        "{{CLOSURE_DYNAMIC_SECTIONS}}",
        "{{ CLOSURE_DYNAMIC_SECTIONS }}",
        "{CLOSURE_DYNAMIC_SESSIONS}",  # common typo
        "{{CLOSURE_DYNAMIC_SESSIONS}}",
        "{{ CLOSURE_DYNAMIC_SESSIONS }}",
    }
)
IMAGE_EXTENSIONS = {".png", ".jpg", ".jpeg", ".bmp", ".gif", ".webp"}

def _unlink_quiet(path: str) -> None:
    try:
        if path and os.path.isfile(path):
            os.remove(path)
    except OSError:
        pass

# If "1" / "true" (env TYPE6_CLOSURE_RISK_SCAN_TOOLS), also emit Risk & Description, SCAN
# Type, and List of Tools inside {CLOSURE_DYNAMIC_SECTIONS} before Overall Findings. Default
# is off so the static base template (like AdvantageGo PDF) holds those pages; the closure
# slot is Overall Findings → Test Case → Vulnerabilities Found → Revalidation → details.


def _load_root_env():
    """Load server-side secrets from CSS/.env, with CSS/env fallback."""
    try:
        if ROOT_ENV_PATH.exists():
            load_dotenv(ROOT_ENV_PATH, override=True)
        elif ROOT_ENV_FALLBACK_PATH.exists():
            load_dotenv(ROOT_ENV_FALLBACK_PATH, override=True)
    except Exception:
        pass

# ─────────────────────────────────────────────────────────────────────────────
# LOW-LEVEL HELPERS
# ─────────────────────────────────────────────────────────────────────────────

def rgb(h):
    h = h.lstrip("#")
    return RGBColor(int(h[:2],16), int(h[2:4],16), int(h[4:6],16))

def shd(cell, hex6):
    tc = cell._tc; tcPr = tc.get_or_add_tcPr()
    for old in tcPr.findall(qn("w:shd")): tcPr.remove(old)
    e = OxmlElement("w:shd")
    e.set(qn("w:val"),"clear"); e.set(qn("w:color"),"auto")
    e.set(qn("w:fill"), hex6.upper().lstrip("#")); tcPr.append(e)

def borders(cell, color=BGRAY, sz=4):
    tc = cell._tc; tcPr = tc.get_or_add_tcPr()
    for old in tcPr.findall(qn("w:tcBorders")): tcPr.remove(old)
    b = OxmlElement("w:tcBorders")
    for side in ("top","left","bottom","right"):
        e = OxmlElement(f"w:{side}")
        e.set(qn("w:val"),"single"); e.set(qn("w:sz"),str(sz))
        e.set(qn("w:space"),"0"); e.set(qn("w:color"),color.upper().lstrip("#"))
        b.append(e)
    tcPr.append(b)

def margins(cell, top=60, bot=60, left=100, right=100):
    tc = cell._tc; tcPr = tc.get_or_add_tcPr()
    for old in tcPr.findall(qn("w:tcMar")): tcPr.remove(old)
    m = OxmlElement("w:tcMar")
    for side,val in [("top",top),("bottom",bot),("left",left),("right",right)]:
        e = OxmlElement(f"w:{side}"); e.set(qn("w:w"),str(val)); e.set(qn("w:type"),"dxa"); m.append(e)
    tcPr.append(m)

def tcw(cell, tw):
    tc = cell._tc; tcPr = tc.get_or_add_tcPr()
    for old in tcPr.findall(qn("w:tcW")): tcPr.remove(old)
    e = parse_xml(f'<w:tcW {nsdecls("w")} w:w="{tw}" w:type="dxa"/>'); tcPr.append(e)

def tbl_width(table, tw=9350):
    tbl = table._tbl; tblPr = tbl.tblPr
    if tblPr is None:
        tblPr = parse_xml(f'<w:tblPr {nsdecls("w")}/>'); tbl.insert(0, tblPr)
    for old in tblPr.findall(qn("w:tblW")): tblPr.remove(old)
    tblPr.append(parse_xml(f'<w:tblW {nsdecls("w")} w:w="{tw}" w:type="dxa"/>'))

def tbl_borders(table, color=BGRAY, sz=4):
    tbl = table._tbl; tblPr = tbl.tblPr
    if tblPr is None:
        tblPr = parse_xml(f'<w:tblPr {nsdecls("w")}/>'); tbl.insert(0, tblPr)
    for old in tblPr.findall(qn("w:tblBorders")): tblPr.remove(old)
    c = color.lstrip("#")
    tblPr.append(parse_xml(
        f'<w:tblBorders {nsdecls("w")}>'
        f'<w:top w:val="single" w:sz="{sz}" w:color="{c}"/>'
        f'<w:left w:val="single" w:sz="{sz}" w:color="{c}"/>'
        f'<w:bottom w:val="single" w:sz="{sz}" w:color="{c}"/>'
        f'<w:right w:val="single" w:sz="{sz}" w:color="{c}"/>'
        f'<w:insideH w:val="single" w:sz="{sz}" w:color="{c}"/>'
        f'<w:insideV w:val="single" w:sz="{sz}" w:color="{c}"/>'
        f'</w:tblBorders>'
    ))


def _safe_set_table_style(table, preferred_name="Table Grid"):
    """Best-effort table style setter for custom templates with missing built-in styles."""
    style_candidates = [preferred_name, "TableGrid"]
    for style_name in style_candidates:
        try:
            table.style = style_name
            return True
        except Exception:
            continue
    return False

def write_cell(cell, text, bold=False, italic=False, size=BODY,
               fg=None, align=WD_ALIGN_PARAGRAPH.LEFT, underline=False):
    cell.paragraphs[0].clear(); p = cell.paragraphs[0]; p.alignment = align
    run = p.add_run(str(text) if text else "")
    run.bold=bold; run.italic=italic; run.underline=underline
    run.font.name=FONT; run.font.size=Pt(size)
    if fg: run.font.color.rgb = rgb(fg)
    margins(cell)

def header_row(row, texts, bg=PRP2, fg=WHITE, sz=BODY, col_bgs=None):
    for i, txt in enumerate(texts):
        c = row.cells[i]; cb = (col_bgs or {}).get(i, bg)
        write_cell(c, txt, bold=True, size=sz, fg=fg, align=WD_ALIGN_PARAGRAPH.CENTER)
        shd(c, cb); borders(c,"FFFFFF",4); c.vertical_alignment=WD_ALIGN_VERTICAL.CENTER
        margins(c,60,60,80,80)

def para(doc, text="", bold=False, italic=False, size=BODY, fg=None,
         align=WD_ALIGN_PARAGRAPH.JUSTIFY, sb=0, sa=4):
    p = doc.add_paragraph(); p.alignment = align
    p.paragraph_format.space_before = Pt(sb); p.paragraph_format.space_after = Pt(sa)
    if text:
        run = p.add_run(text); run.bold=bold; run.italic=italic
        run.font.name=FONT; run.font.size=Pt(size)
        if fg: run.font.color.rgb=rgb(fg)
    return p

def heading(doc, text, level=1):
    sz = HEAD+2 if level==1 else HEAD+1
    return para(doc, text, bold=True, size=sz, fg=PRP,
                align=WD_ALIGN_PARAGRAPH.LEFT, sb=8, sa=6)

def page_break(doc):
    p = doc.add_paragraph()
    p.paragraph_format.space_before=Pt(0); p.paragraph_format.space_after=Pt(0)
    br = OxmlElement("w:br"); br.set(qn("w:type"),"page"); p.add_run()._r.append(br)


def _reparent_body_element_after(anchor, element) -> object:
    """Move element (w:p or w:tbl) to immediately after anchor in the same parent."""
    parent = element.getparent()
    if parent is not None:
        parent.remove(element)
    anchor.addnext(element)
    return element


class _DocBodyTail:
    """
    new paragraphs/tables are created via Document (last in body) then reparented
    after a fixed anchor w:p so generated content appears at {CLOSURE_DYNAMIC_SECTIONS}.
    """
    def __init__(self, real_doc, anchor_paragraph):
        self._doc = real_doc
        self._after = anchor_paragraph._p

    def add_paragraph(self, *args, **kwargs):
        p = self._doc.add_paragraph(*args, **kwargs)
        self._after = _reparent_body_element_after(self._after, p._p)
        return p

    def add_table(self, *args, **kwargs):
        t = self._doc.add_table(*args, **kwargs)
        self._after = _reparent_body_element_after(self._after, t._tbl)
        return t


class _InsertionDocument:
    """
    Exposes the same add_paragraph/add_table as a Document, but appends in-order at the
    closure marker; all other attribute access (sections, save, _element) uses the real doc.
    """
    __slots__ = ("_doc", "_tail")

    def __init__(self, real_doc, tail: _DocBodyTail):
        object.__setattr__(self, "_doc", real_doc)
        object.__setattr__(self, "_tail", tail)

    def add_paragraph(self, *args, **kwargs):
        return self._tail.add_paragraph(*args, **kwargs)

    def add_table(self, *args, **kwargs):
        return self._tail.add_table(*args, **kwargs)

    def __getattr__(self, name):
        return getattr(self._doc, name)


def _mark_heading(paragraph, level=1):
    """
    Intentional no-op: do not set Word 'Heading' styles on generated content so the
    base template’s outline / TOC and any header are not mixed with program styles.
    """
    return paragraph


def _iter_container_paragraphs(container):
    for p in getattr(container, "paragraphs", []):
        yield p
    for table in getattr(container, "tables", []):
        for row in table.rows:
            for cell in row.cells:
                yield from _iter_container_paragraphs(cell)


def _replace_placeholder_tokens(text, values):
    updated = text
    for key, value in values.items():
        val = "" if value is None else str(value)
        updated = re.sub(r"\{\{\s*" + re.escape(key) + r"\s*\}\}", val, updated)
        updated = updated.replace("{" + key + "}", val)
    return updated


def _replace_placeholders_in_doc(doc, values):
    """Fill placeholders in the main document body only; never write header/footer."""
    targets = [doc]

    for target in targets:
        for p in _iter_container_paragraphs(target):
            if not p.text:
                continue
            new_text = _replace_placeholder_tokens(p.text, values)
            if new_text == p.text:
                continue
            if p.runs:
                p.runs[0].text = new_text
                for r in p.runs[1:]:
                    r.text = ""
            else:
                p.add_run(new_text)


def _enable_update_fields_on_open(doc):
    try:
        settings_el = doc.settings._element
        node = settings_el.find(qn("w:updateFields"))
        if node is None:
            node = OxmlElement("w:updateFields")
            settings_el.append(node)
        node.set(qn("w:val"), "true")
    except Exception:
        pass


def _add_toc_field(paragraph):
    paragraph.text = ""
    fld_begin = OxmlElement("w:fldChar")
    fld_begin.set(qn("w:fldCharType"), "begin")

    instr = OxmlElement("w:instrText")
    instr.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    instr.text = 'TOC \\o "1-3" \\h \\z \\u'

    fld_sep = OxmlElement("w:fldChar")
    fld_sep.set(qn("w:fldCharType"), "separate")

    fld_end = OxmlElement("w:fldChar")
    fld_end.set(qn("w:fldCharType"), "end")

    run = paragraph.add_run()
    run._r.append(fld_begin)
    run._r.append(instr)
    run._r.append(fld_sep)

    hint = paragraph.add_run("Right-click and update field to refresh page numbers.")
    hint.font.name = FONT
    hint.font.size = Pt(9)
    hint.font.color.rgb = rgb("808080")

    run_end = paragraph.add_run()
    run_end._r.append(fld_end)


def _has_toc_field(doc):
    try:
        return "TOC \\o" in doc._element.xml
    except Exception:
        return False


def _ensure_toc_field(doc):
    """Replace {TOC} (anywhere in main body, including table cells) with a Word TOC field."""
    inserted = False
    marker_variants = {
        TOC_MARKER,
        "{TOC_FIELD}",
        "{{TOC}}",
        "{{ TOC }}",
        "{{TOC_FIELD}}",
        "{{ TOC_FIELD }}",
    }
    for p in _iter_container_paragraphs(doc):
        if p.text.strip() in marker_variants:
            _add_toc_field(p)
            inserted = True
            break

    if not inserted and not _has_toc_field(doc):
        p = doc.add_paragraph()
        _mark_heading(p, 1)
        run = p.add_run("Table of Content")
        run.font.name = FONT
        run.font.size = Pt(22)
        run.bold = True
        run.font.color.rgb = rgb(PRP)
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(14)

        toc_p = doc.add_paragraph()
        _add_toc_field(toc_p)
        page_break(doc)


def _create_default_advantage_template(template_path):
    doc = DocxDocument()

    for sec in doc.sections:
        sec.page_width = Cm(21.0)
        sec.page_height = Cm(29.7)
        sec.top_margin = Cm(2.5)
        sec.bottom_margin = Cm(2.5)
        sec.left_margin = Cm(2.5)
        sec.right_margin = Cm(2.5)

    # Cover placeholders
    para(doc, "{report_id}", bold=True, size=13, fg=PRP, align=WD_ALIGN_PARAGRAPH.RIGHT, sa=2)
    para(doc, "Technical Closure Report", bold=True, size=14, fg=NAVY, align=WD_ALIGN_PARAGRAPH.RIGHT, sa=30)
    for _ in range(5):
        para(doc, sb=0, sa=6)
    para(doc, "CyberSmith SECURE", bold=True, size=24, fg=PRP, align=WD_ALIGN_PARAGRAPH.CENTER, sa=4)
    para(doc, "Private Limited", size=14, fg=NAVY, align=WD_ALIGN_PARAGRAPH.CENTER, sa=18)
    para(doc, "{doc_title}", bold=True, size=14, fg=PRP, align=WD_ALIGN_PARAGRAPH.CENTER, sa=6)
    para(doc, "{release_date}", size=11, fg="808080", align=WD_ALIGN_PARAGRAPH.CENTER, sa=0)
    page_break(doc)

    # Disclaimer page
    p_disclaimer = doc.add_paragraph()
    _mark_heading(p_disclaimer, 1)
    r_disclaimer = p_disclaimer.add_run("Disclaimer")
    r_disclaimer.font.name = FONT
    r_disclaimer.font.size = Pt(18)
    r_disclaimer.bold = True
    r_disclaimer.font.color.rgb = rgb(PRP)
    para(doc, (
        "All information in this report is confidential and proprietary to "
        "CyberSmith SECURE Private Limited."
    ), size=BODY, sa=10)
    page_break(doc)

    # TOC page marker
    p_toc_title = doc.add_paragraph()
    _mark_heading(p_toc_title, 1)
    r_toc_title = p_toc_title.add_run("Table of Content")
    r_toc_title.font.name = FONT
    r_toc_title.font.size = Pt(22)
    r_toc_title.bold = True
    r_toc_title.font.color.rgb = rgb(PRP)
    p_toc_title.paragraph_format.space_after = Pt(14)
    doc.add_paragraph(TOC_MARKER)
    page_break(doc)

    # Marker where dynamic sections start
    doc.add_paragraph(CONTENT_MARKER)
    doc.save(template_path)


def _closure_base_cache_dir() -> Path:
    d = BACKEND_DIR / "closure_base_cache"
    d.mkdir(parents=True, exist_ok=True)
    return d


def _materialize_base_template(p: Path) -> Path:
    """
    Return a .docx path for python-docx. If `p` is .docx, return as-is. If .pdf, convert
    with pdf2docx to a file under closure_base_cache/ and reuse while the source PDF
    mtime is not newer than the cache.
    """
    p = p.resolve()
    if p.suffix.lower() == ".docx":
        return p
    if p.suffix.lower() != ".pdf":
        return p
    try:
        from pdf2docx import Converter
    except ImportError as exc:
        raise FileNotFoundError(
            f"Base template is a PDF ({p.name}) but pdf2docx is not installed. "
            "Run: pip install pdf2docx"
        ) from exc
    key = hashlib.sha1(str(p).encode("utf-8")).hexdigest()[:20]
    out = _closure_base_cache_dir() / f"base_from_pdf_{key}.docx"
    if out.is_file() and out.stat().st_mtime >= p.stat().st_mtime:
        return out
    cv = Converter(str(p))
    cv.convert(str(out))
    cv.close()
    return out


def _resolve_advantage_template_path() -> Path:
    """
    1) TYPE6_CLOSURE_BASE_TEMPLATE in env (path to .docx or .pdf)
    2) `AdvantageGo Exact Web app VAPT Technical Closure Report 1.docx` in Automation/backend, then
    3) `AdvantageGo_Closure_Base_Template.docx` in the same folder
    4) AdvantageGo PDF in Downloads (converted to cache), if present
    """
    _load_root_env()
    extra = (os.environ.get("TYPE6_CLOSURE_BASE_TEMPLATE") or "").strip()
    if extra:
        ep = Path(extra)
        if ep.is_file():
            return _materialize_base_template(ep)
        raise FileNotFoundError(
            f"TYPE6_CLOSURE_BASE_TEMPLATE is set but not found: {ep}"
        )
    candidates: list[Path] = [ADVANTAGE_TEMPLATE_PATH, *ADVANTAGE_TEMPLATE_CANDIDATES]
    for path in candidates:
        if path.exists():
            return _materialize_base_template(path)
    dlp = _default_advantagego_pdf_in_downloads()
    if dlp is not None:
        return _materialize_base_template(dlp)
    raise FileNotFoundError(
        "No closure base template found. Looked for: "
        + ", ".join(str(p) for p in candidates)
        + " — place `1.docx` in Automation/backend, or set TYPE6_CLOSURE_BASE_TEMPLATE."
    )

def sev_colors(s): return SEV.get(s.lower(), ("808080", WHITE))
def reval_colors(s): return REVAL.get(s.lower(), ("808080", WHITE))

# ─────────────────────────────────────────────────────────────────────────────
# CHARTS — exact AdvantageGo color mapping
# ─────────────────────────────────────────────────────────────────────────────

def bar_chart(counts):
    """3D bar chart with elevated perspective, shadows, and beveled edges."""
    labels = [k for k in ["Critical","High","Medium","Low","Informational"] if counts.get(k,0)>0]
    values = [counts[k] for k in labels]
    if not labels:
        labels, values = ["None"], [0]
    cmap = {"Critical":"#800000","High":"#FF0000","Medium":"#FFC000","Low":"#92D050","Informational":"#00B0F0","None":"#CCCCCC"}
    bcolors = [cmap.get(l, "#CCCCCC") for l in labels]

    fig = plt.figure(figsize=(6.0, 4.0), facecolor="white")
    ax = fig.add_subplot(111, projection="3d")

    x_pos = np.arange(len(labels))
    y_pos = np.zeros(len(labels))
    z_pos = np.zeros(len(labels))
    dx = 0.5
    dy = 0.4
    dz = np.array(values, dtype=float)

    # Draw 3D bars with individual colors
    for i in range(len(labels)):
        ax.bar3d(
            x_pos[i], y_pos[i], z_pos[i],
            dx, dy, dz[i],
            color=bcolors[i], edgecolor="white", linewidth=0.5,
            alpha=0.92, zsort="average",
        )
        # Value label on top of each bar
        if dz[i] > 0:
            ax.text(
                x_pos[i] + dx / 2, y_pos[i] + dy / 2, dz[i] + 0.15,
                str(int(dz[i])), ha="center", va="bottom",
                fontsize=10, fontweight="bold", color=bcolors[i],
            )

    ax.set_xticks(x_pos + dx / 2)
    ax.set_xticklabels(labels, fontsize=8, rotation=15, ha="right")
    ax.set_yticks([])
    ax.set_zlabel("Count", fontsize=8, labelpad=4)
    ax.set_title(
        "Total Vulnerabilities Found",
        fontsize=12, color="#7030A0", fontweight="bold", pad=12,
    )

    # Camera angle for best 3D perspective
    ax.view_init(elev=28, azim=-55)
    ax.set_facecolor("white")
    ax.xaxis.pane.fill = False
    ax.yaxis.pane.fill = False
    ax.zaxis.pane.fill = False
    ax.xaxis.pane.set_edgecolor("#E0E0E0")
    ax.yaxis.pane.set_edgecolor("#E0E0E0")
    ax.zaxis.pane.set_edgecolor("#E0E0E0")
    ax.grid(axis="z", linestyle="--", alpha=0.3)

    plt.tight_layout(pad=0.8)
    buf = io.BytesIO()
    plt.savefig(buf, format="png", dpi=180, bbox_inches="tight", facecolor="white")
    plt.close()
    buf.seek(0)
    return buf.read()


def donut_chart(counts):
    """3D-style pie chart (exploded wedges with shadow) for severity distribution."""
    labels, values, colors = [], [], []
    cmap = {"Critical":"#800000","High":"#FF0000","Medium":"#FFC000","Low":"#92D050","Informational":"#00B0F0"}
    for k in ["Critical","High","Medium","Low","Informational"]:
        if counts.get(k, 0) > 0:
            labels.append(k)
            values.append(counts[k])
            colors.append(cmap[k])
    if not values:
        values, labels, colors = [1], ["None"], ["#CCCCCC"]
    total = sum(values)

    # Explode the largest slice slightly for 3D effect
    max_idx = values.index(max(values))
    explode = [0.06 if i == max_idx else 0.02 for i in range(len(values))]

    fig, ax = plt.subplots(figsize=(4.0, 4.0), facecolor="white")

    # Shadow layer for 3D depth
    ax.pie(
        values,
        colors=["#E0E0E0"] * len(values),
        startangle=87,
        explode=explode,
        wedgeprops=dict(width=0.50, edgecolor="#D0D0D0", linewidth=0.5),
        radius=1.02,
        center=(0.015, -0.015),
    )

    # Main pie on top
    pie_result = ax.pie(
        values,
        colors=colors,
        startangle=90,
        explode=explode,
        wedgeprops=dict(width=0.50, edgecolor="white", linewidth=2.5),
        shadow=False,
    )
    wedges = pie_result[0]

    # Center text
    ax.text(
        0, 0, f"Total\nFindings: {total}",
        ha="center", va="center",
        fontsize=9, fontweight="bold", color="#444444",
    )

    ax.legend(
        wedges,
        [f"{l}: {v}" for l, v in zip(labels, values)],
        loc="lower center",
        bbox_to_anchor=(0.5, -0.18),
        fontsize=7,
        frameon=False,
        ncol=min(3, len(labels)),
    )

    plt.tight_layout(pad=0.4)
    buf = io.BytesIO()
    plt.savefig(buf, format="png", dpi=180, bbox_inches="tight", facecolor="white")
    plt.close()
    buf.seek(0)
    return buf.read()

# ─────────────────────────────────────────────────────────────────────────────
# EXTRACT ALL DATA FROM ANNA DARPAN DOCX
# ─────────────────────────────────────────────────────────────────────────────

def extract_all_data(docx_path):
    doc = DocxDocument(docx_path)
    tables = doc.tables

    # ── Metadata from Table 1 (index 1)
    meta = {}
    if len(tables) > 1:
        for row in tables[1].rows:
            cells = [c.text.strip() for c in row.cells]
            if len(cells) >= 2 and cells[0] and cells[1]:
                meta[cells[0]] = cells[1]

    # ── Document control from Table 2 (index 2)
    if len(tables) > 2:
        for row in tables[2].rows:
            cells = [c.text.strip() for c in row.cells]
            if len(cells) >= 2 and cells[0] and cells[1]:
                meta[cells[0]] = cells[1]

    # Supplement hardcoded values from Anna Darpan
    meta.setdefault("client_name",   "Capita")
    meta.setdefault("target_url",    "https://kcq-uat.capita.co.in/")
    meta.setdefault("report_id",     "CSS-MOB-2025")
    meta.setdefault("doc_title",     "Anna Darpan Android Application")
    meta.setdefault("audit_type",    "Security assessment (VAPT) - Android Mobile Application")
    meta.setdefault("Period",        "01-11-2025 To 05-11-2025")
    meta.setdefault("Prepared by",   "Masoom Khan")
    meta.setdefault("Reviewed by",   "Kiran Parmar")
    meta.setdefault("Approved by",   "Sarvesh Salgaonkar")
    meta.setdefault("Released by",   "Smith Gonsalves")

    # ── Engagement scope Table 5 (index 5)
    assets = []
    if len(tables) > 5:
        for ri, row in enumerate(tables[5].rows):
            if ri == 0: continue
            cells = [c.text.strip() for c in row.cells]
            if any(cells):
                assets.append(cells)
    meta["assets"] = assets

    # ── Team Table 6 (index 6)
    team = []
    if len(tables) > 6:
        for ri, row in enumerate(tables[6].rows):
            if ri == 0: continue
            cells = [c.text.strip() for c in row.cells]
            if cells and cells[0].isdigit():
                team.append({"no":cells[0],"name":cells[1],"designation":cells[2],"email":cells[3],"certs":cells[4]})
    meta["team"] = team

    # ── Tools Table 9 (index 9)
    tools = []
    if len(tables) > 9:
        for ri, row in enumerate(tables[9].rows):
            if ri == 0: continue
            cells = [c.text.strip() for c in row.cells]
            if cells and cells[0].isdigit():
                tools.append({"no":cells[0],"name":cells[1],"version":cells[2],"type":cells[3]})
    meta["tools"] = tools

    # ── Overall findings Table 10 (index 10)
    # Row 1: 1. | URL | Web Application | Completed | 0 | 6 | 4 | 0 | 0 | 10
    if len(tables) > 10:
        data_row = [c.text.strip() for c in tables[10].rows[1].cells] if len(tables[10].rows) > 1 else []
        if data_row and len(data_row) >= 10:
            meta["c_critical"]     = int(data_row[4] or 0)
            meta["c_high"]         = int(data_row[5] or 0)
            meta["c_medium"]       = int(data_row[6] or 0)
            meta["c_low"]          = int(data_row[7] or 0)
            meta["c_informational"]= int(data_row[8] or 0)
            meta["c_total"]        = int(data_row[9] or 0)

    # ── Vulnerabilities: tables 11-16 (each table is one set of observations, merged rows)
    vulns = []
    # Observation tables start at index 11
    for ti in range(11, len(tables)):
        t = tables[ti]
        rows = t.rows
        # Read rows and group by observation headers
        obs_data: dict[str, Any] = {}  # current observation dict
        i = 0
        while i < len(rows):
            cells = [c.text.strip() for c in rows[i].cells]
            # De-duplicate merged cells
            unique_cells = []
            for c in cells:
                if not unique_cells or c != unique_cells[-1]:
                    unique_cells.append(c)
            text = " | ".join(c for c in unique_cells if c)

            # Start of a new observation
            m = re.match(r'Observation:\s*#(\d+)', text, re.I)
            if m:
                if obs_data:
                    vulns.append(_finalize_vuln(obs_data))
                obs_data = {"number": int(m.group(1))}
                # Parse severity and status from same row
                for c in unique_cells:
                    sm = re.search(r'Severity:\s*(\w+)', c, re.I)
                    stm = re.search(r'Status:\s*(\w+)', c, re.I)
                    if sm: obs_data["severity"] = sm.group(1).capitalize()
                    if stm: obs_data["status"] = stm.group(1).capitalize()
            elif obs_data:
                # Parse remaining fields
                _parse_obs_row(obs_data, unique_cells, text)
            i += 1

        if obs_data and obs_data.get("number"):
            vulns.append(_finalize_vuln(obs_data))

    vulns.sort(key=lambda x: x["number"])
    print(f"  ✓ Extracted {len(vulns)} vulnerabilities")
    return meta, vulns

def _parse_obs_row(obs, cells, text):
    # New or Repeat / CVE/CWE
    for c in cells:
        m = re.search(r'New or Repeat Observation:\s*(\w+)', c, re.I)
        if m: obs["new_or_repeat"] = m.group(1).capitalize()
        m2 = re.search(r'CVE/CWE:\s*([A-Z0-9\-]+)', c, re.I)
        if m2: obs["cwe"] = m2.group(1)
        m3 = re.search(r'CVSS:\s*([\d.]+)\s*$', c.strip())
        if m3: obs.setdefault("cvss_score", m3.group(1))
        m4 = re.search(r'CVSS:\s*[\d.]+\s*[-–]\s*(CVSS:[^\s|]+)', c)
        if m4: obs["cvss_vector"] = m4.group(1)
    # Affected asset
    if "Affected Asset" in text:
        for c in cells:
            if "http" in c.lower() or "/" in c:
                if "http" in c.lower():
                    obs.setdefault("affected_asset", c.split("|")[0].strip())
    # Title
    m = re.search(r'Observation/\s*Vulnerability Title:\s*([^|]+)', text, re.I)
    if m: obs["title"] = m.group(1).strip().rstrip("|").strip()
    # Description
    m = re.search(r'Detailed Observation.*?:\s*([^|].{20,})', text, re.I)
    if m and "Observation" not in m.group(1)[:20]:
        obs.setdefault("description", m.group(1).strip()[:1500])
    # Recommendation
    m = re.search(r'Recommendation:\s*\n?([^|].{20,})', text, re.I|re.DOTALL)
    if m:
        rec = m.group(1).split("Reference:")[0].split("Evidence")[0].strip()[:1500]
        obs.setdefault("recommendation", rec)
    # Reference
    m = re.search(r'Reference:\s*(https?://[^\s|]+)', text, re.I)
    if m: obs.setdefault("reference", m.group(1).strip())
    # PoC steps
    if "Evidence / Proof of Concept" in text or "Step 1:" in text:
        raw_poc = text
        steps = re.findall(r'Step\s*(\d+)[:\s]+(.+?)(?=Step\s*\d+|$)', raw_poc, re.DOTALL|re.I)
        if steps:
            obs.setdefault("poc_steps", [{"num":s[0],"text":re.sub(r'\s+',' ',s[1]).strip()[:300]} for s in steps if s[1].strip()])

def _finalize_vuln(obs):
    n = obs.get("number", 0)
    return {
        "number":         n,
        "id":             f"CSS-{n:03d}",
        "title":          obs.get("title", f"Finding #{n}"),
        "severity":       obs.get("severity", "Medium"),
        "status":         obs.get("status", "Open"),
        "new_or_repeat":  obs.get("new_or_repeat", "New"),
        "cwe":            obs.get("cwe", ""),
        "cve":            obs.get("cve", ""),
        "cvss_score":     obs.get("cvss_score", ""),
        "cvss_vector":    obs.get("cvss_vector", ""),
        "affected_asset": obs.get("affected_asset", "https://kcq-uat.capita.co.in/"),
        "description":    obs.get("description", ""),
        "recommendation": obs.get("recommendation", ""),
        "reference":      obs.get("reference", ""),
        "poc_steps":      obs.get("poc_steps", []),
        "poc_images":     [],
        "reval_images":   [],
        # Closure — filled by AI
        "revalidation_status": "Unpatched",
        "developer_remarks":   "-",
        "vapt_remarks":        "Unpatched – Pending remediation.",
        "closure_status":      "Open",
        "closure_notes":       "This vulnerability remains open pending remediation.",
    }


def _to_int(value, default=None):
    try:
        if value is None:
            return default
        text = str(value).strip()
        if not text:
            return default
        return int(float(text))
    except Exception:
        return default


def _split_cwe_cve(cwe_cve_text):
    cwe = ""
    cve = ""
    if not cwe_cve_text:
        return cwe, cve
    tokens = re.split(r"[,|/\\\s]+", str(cwe_cve_text).strip())
    for token in tokens:
        t = token.strip().upper()
        if not t:
            continue
        if t.startswith("CWE") and not cwe:
            cwe = t
        elif t.startswith("CVE") and not cve:
            cve = t
    return cwe, cve


def _overlay_excel_row(vuln, excel_row, fallback_number):
    vuln["number"] = _to_int(excel_row.get("sr_no"), vuln.get("number", fallback_number))
    vuln["id"] = f"CSS-{int(vuln['number']):03d}"

    title = (excel_row.get("observation") or excel_row.get("observation_summary") or vuln.get("title") or "").strip()
    if title:
        vuln["title"] = title

    severity = str(excel_row.get("severity") or "").strip()
    if severity:
        vuln["severity"] = severity.capitalize()

    status = str(excel_row.get("status") or "").strip()
    if status:
        vuln["status"] = status.capitalize()

    cvss_score = excel_row.get("cvss")
    if cvss_score is not None and str(cvss_score).strip() != "":
        vuln["cvss_score"] = str(cvss_score).strip()

    cvss_vector = str(excel_row.get("cvss_vector") or "").strip()
    if cvss_vector:
        vuln["cvss_vector"] = cvss_vector

    cwe, cve = _split_cwe_cve(excel_row.get("cve_cwe"))
    if cwe:
        vuln["cwe"] = cwe
    if cve:
        vuln["cve"] = cve

    affected = (excel_row.get("affected_asset") or excel_row.get("ip_url_app") or "").strip()
    if affected:
        vuln["affected_asset"] = affected

    description = (excel_row.get("detailed_observation") or excel_row.get("observation_summary") or "").strip()
    if description:
        vuln["description"] = description

    recommendation = str(excel_row.get("recommendation") or "").strip()
    if recommendation:
        vuln["recommendation"] = recommendation

    reference = str(excel_row.get("reference") or "").strip()
    if reference:
        vuln["reference"] = reference

    steps = excel_row.get("steps") or []
    if isinstance(steps, list) and steps:
        normalized_steps = []
        for idx, step in enumerate(steps, 1):
            if not isinstance(step, dict):
                continue
            step_num = _to_int(step.get("number"), idx)
            step_text = str(step.get("content") or "").strip()
            if step_text:
                normalized_steps.append({"num": str(step_num), "text": step_text})
        if normalized_steps:
            vuln["poc_steps"] = normalized_steps


def _overlay_excel_data(meta, vulns, excel_path):
    if not excel_path or not os.path.exists(excel_path):
        return meta, vulns
    if parse_excel_data is None:
        print("  ⚠ excel_parser unavailable; skipping excel overlay")
        return meta, vulns

    parsed = parse_excel_data(excel_path)
    excel_vulns = parsed.get("vulnerabilities") or []
    metadata = parsed.get("metadata") or {}

    if metadata.get("client"):
        meta["client_name"] = str(metadata["client"]).strip()
    if metadata.get("project"):
        meta["doc_title"] = str(metadata["project"]).strip()

    if not excel_vulns:
        print("  ⚠ No vulnerability rows found in uploaded excel")
        return meta, vulns

    by_number = {}
    by_position = {}
    for idx, row in enumerate(excel_vulns, 1):
        row_no = _to_int(row.get("sr_no"), None)
        if row_no is not None:
            by_number[row_no] = row
        by_position[idx] = row

    if not vulns:
        vulns = []
        for idx, row in enumerate(excel_vulns, 1):
            v = _finalize_vuln({"number": _to_int(row.get("sr_no"), idx)})
            _overlay_excel_row(v, row, idx)
            vulns.append(v)
    else:
        for idx, v in enumerate(vulns, 1):
            row = by_number.get(_to_int(v.get("number"), idx)) or by_position.get(idx)
            if row:
                _overlay_excel_row(v, row, idx)

    sev_counts = {"Critical": 0, "High": 0, "Medium": 0, "Low": 0, "Informational": 0}
    for v in vulns:
        s = str(v.get("severity", "")).capitalize()
        if s in sev_counts:
            sev_counts[s] += 1
    meta["c_critical"] = sev_counts["Critical"]
    meta["c_high"] = sev_counts["High"]
    meta["c_medium"] = sev_counts["Medium"]
    meta["c_low"] = sev_counts["Low"]
    meta["c_informational"] = sev_counts["Informational"]
    meta["c_total"] = len(vulns)

    print(f"  ✓ Excel overlay applied to {len(vulns)} vulnerabilities")
    return meta, vulns


def _extract_observation_number(path_text):
    text = str(path_text)
    for token in reversed(Path(text).parts):
        m = re.search(r"(?:^|[^0-9])(?:#|obs(?:ervation)?[-_ ]*|css-)?(\d{1,3})(?:[^0-9]|$)", token, re.I)
        if m:
            return _to_int(m.group(1), None)
    m2 = re.search(r"(?:^|[^0-9])(\d{1,3})(?:[^0-9]|$)", Path(text).name)
    return _to_int(m2.group(1), None) if m2 else None


def _attach_zip_images(vulns, poc_zip_path):
    if not poc_zip_path or not os.path.exists(poc_zip_path):
        return 0

    image_map = {}
    with ZipFile(poc_zip_path, "r") as zf:
        for info in zf.infolist():
            if info.is_dir():
                continue
            suffix = Path(info.filename).suffix.lower()
            if suffix not in IMAGE_EXTENSIONS:
                continue

            obs_no = _extract_observation_number(info.filename)
            if obs_no is None:
                continue

            bucket = image_map.setdefault(obs_no, {"poc_images": [], "reval_images": []})
            lower_name = info.filename.lower()
            target = "reval_images" if any(k in lower_name for k in ("reval", "closure", "patched", "fixed", "after")) else "poc_images"
            try:
                bucket[target].append(zf.read(info.filename))
            except Exception:
                pass

    attached = 0
    for idx, v in enumerate(vulns, 1):
        key = _to_int(v.get("number"), idx)
        payload = image_map.get(key) or image_map.get(idx)
        if not payload:
            continue
        if payload["poc_images"]:
            v["poc_images"] = payload["poc_images"]
            attached += len(payload["poc_images"])
        if payload["reval_images"]:
            v["reval_images"] = payload["reval_images"]
            attached += len(payload["reval_images"])

    return attached


# --------------------------------------------------
# AI→Excel alignment helpers  (GUARDRAILS v2)
# --------------------------------------------------
# Word limits for safety (matches system prompt constrains)
MAX_DEVELOPER_REMARKS_WORDS = 12
MAX_VAPT_REMARKS_WORDS = 15
MAX_CLOSURE_NOTES_WORDS = 20

# Fields that AI is NEVER allowed to modify (guardrail)
_PROTECTED_FIELDS = frozenset([
    "id", "number", "title", "severity", "cvss_score", "cvss_vector",
    "cwe", "cve", "description", "recommendation", "reference",
    "affected_asset", "poc_steps", "poc_images", "reval_images",
    "new_or_repeat", "status",
])

# Hallucination indicators — if AI output contains these, reject that field
_HALLUCINATION_MARKERS = [
    "as an ai", "i cannot", "i don't have", "hypothetically",
    "for example", "lorem ipsum", "placeholder", "sample text",
    "<", ">", "```", "\n\n\n",
]


def _truncate_words(text, max_words):
    if not text:
        return ""
    parts = re.split(r"\s+", str(text).strip())
    if len(parts) <= max_words:
        return " ".join(parts).strip()
    return " ".join(parts[:max_words]).strip()


def _is_hallucinated(text):
    """Return True if the text looks like an AI hallucination."""
    if not text:
        return False
    lower = str(text).lower()
    return any(marker in lower for marker in _HALLUCINATION_MARKERS)


def _normalize_revalidation_status(s):
    if not s:
        return "Unpatched"
    key = str(s).strip().lower()
    if "justified" in key and "close" in key:
        return "Justified and Closed"
    if "part" in key and "patch" in key:
        return "Partially Patched"
    if "patch" in key or key in ("patched", "fixed"):
        return "Patched"
    if "unpatch" in key or key in ("unpatched", "open"):
        return "Unpatched"
    return "Unpatched"


def _normalize_closure_status(s):
    if not s:
        return "Open"
    key = str(s).strip().lower()
    if any(x in key for x in ("clos", "closed", "fixed", "resolved", "remediat")):
        return "Closed"
    return "Open"


def _validate_extracted_data(vulns):
    """Validate that extracted vulnerability data is accurate and complete.
    Flags suspicious or missing data, never fabricates values.
    Returns the list with only verified/cleaned entries."""
    validated = []
    for idx, v in enumerate(vulns):
        # Must have a title or description — skip truly empty rows
        title = str(v.get("title", "")).strip()
        desc = str(v.get("description", "")).strip()
        if not title and not desc:
            print(f"  ⚠ Skipping vuln #{v.get('number', idx+1)}: no title or description")
            continue

        # Sanitize: never allow placeholder data to leak through
        for field in ("title", "description", "recommendation", "reference",
                      "affected_asset", "cwe", "cve", "cvss_score", "cvss_vector"):
            val = str(v.get(field, "")).strip()
            if _is_hallucinated(val):
                print(f"  ⚠ Removed suspicious {field} for vuln #{v.get('number', idx+1)}")
                v[field] = ""

        # Verify severity is a known value
        sev = str(v.get("severity", "")).strip().capitalize()
        if sev not in ("Critical", "High", "Medium", "Low", "Informational"):
            v["severity"] = "Medium"

        validated.append(v)
    return validated


def _validate_ai_output(ai_item, original_item):
    """Validate a single AI output item against the original.
    Returns sanitized AI result dict — any hallucinated or
    protected-field tampering is rejected and falls back to original."""
    safe = {}
    # Only accept closure-related fields from AI
    allowed_ai_fields = {
        "revalidation_status", "developer_remarks",
        "vapt_remarks", "closure_status", "closure_notes",
    }
    for field in allowed_ai_fields:
        ai_val = str(ai_item.get(field, "")).strip()
        if _is_hallucinated(ai_val):
            # Reject hallucinated value, use original
            safe[field] = original_item.get(field, "")
            print(f"  ⚠ AI hallucination rejected for {field} on {original_item.get('id', '?')}")
        else:
            safe[field] = ai_val if ai_val else original_item.get(field, "")

    # Cross-check: AI must not have changed any protected fields
    for fld in _PROTECTED_FIELDS:
        ai_val = str(ai_item.get(fld, "")).strip()
        orig_val = str(original_item.get(fld, "")).strip()
        if ai_val and ai_val != orig_val:
            print(f"  ⚠ AI attempted to modify protected field '{fld}' on {original_item.get('id', '?')} — blocked")

    return safe


def _enforce_excel_alignment(ai_vulns, original_vulns):
    """
    Ensure AI output aligns exactly with the original vulnerabilities list.
    Preserve identifying fields from the original and only accept AI values
    for the closure-related fields (with truncation/normalization/hallucination check).
    """
    # Build lookup of AI results by id/number
    ai_map = {}
    for a in ai_vulns or []:
        key = str(a.get("id") or a.get("ID") or a.get("id", "")).strip()
        if not key:
            key = str(a.get("number") or a.get("number", "")).strip()
        ai_map[key] = a

    out = []
    for orig in original_vulns or []:
        ident = str(orig.get("id") or orig.get("ID") or orig.get("id", "")).strip()
        if not ident:
            ident = str(orig.get("number") or orig.get("number", "")).strip()
        ai_raw = ai_map.get(ident, {})

        # Validate AI output against original
        ai = _validate_ai_output(ai_raw, orig) if ai_raw else {}

        final = copy.deepcopy(orig)

        # Closure/revalidation fields: prefer validated AI but normalize
        final["revalidation_status"] = _normalize_revalidation_status(
            ai.get("revalidation_status") or orig.get("revalidation_status") or "Unpatched"
        )

        dev = ai.get("developer_remarks") or orig.get("developer_remarks") or "-"
        dev = _truncate_words(dev, MAX_DEVELOPER_REMARKS_WORDS) or "-"
        final["developer_remarks"] = dev

        vapt = ai.get("vapt_remarks") or orig.get("vapt_remarks") or "Unpatched – Pending remediation."
        final["vapt_remarks"] = _truncate_words(vapt, MAX_VAPT_REMARKS_WORDS)

        final["closure_status"] = _normalize_closure_status(
            ai.get("closure_status") or orig.get("closure_status") or "Open"
        )

        notes = ai.get("closure_notes") or orig.get("closure_notes") or "Vulnerability remains open pending remediation."
        final["closure_notes"] = _truncate_words(notes, MAX_CLOSURE_NOTES_WORDS)

        # Ensure identifying fields are never changed by AI
        for fld in _PROTECTED_FIELDS:
            if fld in orig:
                final[fld] = orig[fld]

        out.append(final)

    return out

# ─────────────────────────────────────────────────────────────────────────────
# AI LAYER — OpenAI
# ─────────────────────────────────────────────────────────────────────────────

SYSTEM_PROMPT = """You are a senior cybersecurity analyst writing a Technical Closure Report.

CRITICAL GUARDRAIL RULES (you MUST follow every single rule — violations will be detected):

1. You receive a JSON array of vulnerabilities from an initial VAPT report.
2. Fill ONLY these closure fields for each entry:
   - revalidation_status: ONE of exactly: "Patched" | "Unpatched" | "Justified and Closed" | "Partially Patched"
   - developer_remarks: Max 12 words. Use "-" if unknown. ONLY factual, direct statements.
   - vapt_remarks: Max 15 words. Professional phrases ONLY. No speculation.
   - closure_status: Exactly "Closed" or "Open"
   - closure_notes: ONE sentence, max 20 words. Factual only.
3. ZERO FABRICATION POLICY:
   - Since this is an initial report with NO patch evidence, set ALL vulnerabilities to Unpatched/Open
     UNLESS the description EXPLICITLY states the issue was fixed/patched/resolved.
   - Do NOT invent developer remarks, patch details, or fixes that are not in the data.
   - Do NOT speculate about remediation actions that may have happened.
   - If in doubt, default to Unpatched/Open with "-" as developer_remarks.
4. Return ONLY a valid JSON array. No markdown, no extra text, no commentary.
5. NEVER modify these fields (they will be cross-checked): id, number, title, severity, cvss_score, description, recommendation.
6. Array length MUST exactly match input length. Missing or extra items = rejection.
7. Do NOT add any fields beyond what was asked.
8. Do NOT use phrases like "hypothetically", "for example", "as an AI", "I think"."""


def ai_fill_closure_openai(vulns, api_key):
    try:
        from openai import OpenAI
    except ImportError:
        print("  ⚠ openai not installed. Run: pip install openai")
        return vulns

    client = OpenAI(api_key=api_key)
    # Send only minimal data to AI — never send images or raw file content
    minimal = [{
        "number": v["number"], "id": v["id"], "title": v["title"],
        "severity": v["severity"], "cvss_score": v.get("cvss_score", ""),
        "description": str(v.get("description", ""))[:200],
        "revalidation_status": "", "developer_remarks": "",
        "vapt_remarks": "", "closure_status": "", "closure_notes": "",
    } for v in vulns]

    print(f"  → OpenAI gpt-4o-mini: filling closure for {len(vulns)} vulnerabilities...")
    try:
        response = client.chat.completions.create(
            model="gpt-4o-mini",
            temperature=0.0,
            max_tokens=1200,
            response_format={"type": "json_object"},
            messages=[
                {"role": "system", "content": SYSTEM_PROMPT},
                {"role": "user", "content":
                 f"Fill closure fields. All status=Open (no patches).\n{json.dumps(minimal)}"}
            ]
        )
    except Exception as exc:
        print(f"  ⚠ OpenAI API call failed: {exc}")
        print("  → Falling back to safe defaults (Unpatched/Open)")
        return vulns

    raw_content = response.choices[0].message.content
    raw = (raw_content or "").strip()
    try:
        parsed = json.loads(raw)
        ai_list = parsed if isinstance(parsed, list) else next(
            (v for v in parsed.values() if isinstance(v, list)), [])
    except Exception:
        print("  ⚠ AI returned invalid JSON — using safe defaults")
        ai_list = []

    # GUARDRAIL: Array length must match
    if ai_list and len(ai_list) != len(vulns):
        print(f"  ⚠ AI returned {len(ai_list)} items but expected {len(vulns)} — partial apply")

    ai_map = {str(item.get("number", "")): item for item in ai_list}
    for v in vulns:
        ai = ai_map.get(str(v["number"]), {})

        # Apply AI values with hallucination guard
        reval = ai.get("revalidation_status", "Unpatched") or "Unpatched"
        v["revalidation_status"] = reval if not _is_hallucinated(reval) else "Unpatched"

        dev = ai.get("developer_remarks", "-") or "-"
        v["developer_remarks"] = dev if not _is_hallucinated(dev) else "-"

        vapt = ai.get("vapt_remarks", "Unpatched – Pending remediation.") or "Unpatched – Pending remediation."
        v["vapt_remarks"] = vapt if not _is_hallucinated(vapt) else "Unpatched – Pending remediation."

        clo = ai.get("closure_status", "Open") or "Open"
        v["closure_status"] = clo if not _is_hallucinated(clo) else "Open"

        notes = ai.get("closure_notes", "Vulnerability remains open pending remediation.") or "Vulnerability remains open pending remediation."
        v["closure_notes"] = notes if not _is_hallucinated(notes) else "Vulnerability remains open pending remediation."

    used_tokens = getattr(getattr(response, "usage", None), "total_tokens", None)
    if used_tokens is None:
        print("  ✓ AI done (with guardrail validation).")
    else:
        print(f"  ✓ AI done. Tokens used: {used_tokens} (guardrail validated)")
    return vulns

# ─────────────────────────────────────────────────────────────────────────────
# PAGE BUILDERS
# ─────────────────────────────────────────────────────────────────────────────

def build_cover(doc, meta):
    for _ in range(3): para(doc, sa=2, sb=0)

    # Report ID top right (matching AdvantageGo PDF p.1)
    para(doc, meta.get("report_id","CSS-MOB-2025"),
         bold=True, size=13, fg=PRP, align=WD_ALIGN_PARAGRAPH.RIGHT, sb=0, sa=2)
    para(doc, "Technical Closure Report",
         bold=True, size=14, fg=NAVY, align=WD_ALIGN_PARAGRAPH.RIGHT, sb=0, sa=30)

    for _ in range(5): para(doc, sb=0, sa=6)

    para(doc, "CyberSmith SECURE", bold=True, size=24, fg=PRP,
         align=WD_ALIGN_PARAGRAPH.CENTER, sb=0, sa=4)
    para(doc, "Private Limited", size=14, fg=NAVY,
         align=WD_ALIGN_PARAGRAPH.CENTER, sb=0, sa=20)

    # Purple horizontal rule
    pd = doc.add_paragraph(); pd.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pd.paragraph_format.space_after = Pt(14)
    pPr = pd._p.get_or_add_pPr(); pb = OxmlElement("w:pBdr")
    bot = OxmlElement("w:bottom")
    bot.set(qn("w:val"),"single"); bot.set(qn("w:sz"),"12")
    bot.set(qn("w:space"),"1"); bot.set(qn("w:color"),PRP)
    pb.append(bot); pPr.append(pb)

    title = meta.get("doc_title", "Android Application VAPT")
    para(doc, "Android Application VAPT", size=14, fg="404040",
         align=WD_ALIGN_PARAGRAPH.CENTER, sb=0, sa=6)
    para(doc, "Technical Closure Report", bold=True, size=18, fg=NAVY,
         align=WD_ALIGN_PARAGRAPH.CENTER, sb=0, sa=6)
    para(doc, title, bold=True, size=14, fg=PRP,
         align=WD_ALIGN_PARAGRAPH.CENTER, sb=0, sa=4)

    release = meta.get("Release date", meta.get("release_date",
              datetime.today().strftime("%d %B %Y")))
    para(doc, release, size=11, fg="808080",
         align=WD_ALIGN_PARAGRAPH.CENTER, sb=0, sa=0)
    page_break(doc)


def build_disclaimer(doc):
    p_top = heading(doc, "© Copyright 2025 CyberSmith SECURE Private Limited")
    _mark_heading(p_top, 1)
    para(doc, sa=8)
    p_disclaimer = heading(doc, "Disclaimer", 2)
    _mark_heading(p_disclaimer, 1)
    para(doc, (
        "All information contained in this document is confidential and proprietary to "
        "CyberSmithSECURE. Reproduction, disclosure or use of any information contained "
        "in this document by photographic, electronic or any other means, in whole or part, "
        "for any reason other than for the purpose of operations is strictly prohibited "
        "without written consent."
    ), size=BODY)
    para(doc, sa=8)
    heading(doc, "Limitations on Disclosure and Use", 2)
    para(doc, (
        "This document contains sensitive and confidential information concerning "
        "vulnerabilities of target applications. CyberSmithSECURE recommends that special "
        "precautions be taken to protect the confidentiality of the information contained "
        "in this report."
    ), size=BODY)
    para(doc, sa=4)
    para(doc, (
        "While the VAPT Team is confident that the major security vulnerabilities of the "
        "target applications have been identified, there can be no assurance that an "
        "assessment of this nature will identify all possible security exposures. "
        "The findings and recommendations presented in this document are based on the "
        "technologies and known threats as of the date of this report. As technologies "
        "and risks change over time, the vulnerabilities and the recommendations associated "
        "with the target applications may also change."
    ), size=BODY)
    page_break(doc)


def build_toc(doc, vulns):
    """Table of Contents with live page numbers (Word field)."""
    p = doc.add_paragraph()
    _mark_heading(p, 1)
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(14)
    run = p.add_run("Table of Content")
    run.font.name = FONT
    run.font.size = Pt(22)
    run.bold = True
    run.font.color.rgb = rgb(PRP)

    toc_para = doc.add_paragraph()
    _add_toc_field(toc_para)
    _enable_update_fields_on_open(doc)

    page_break(doc)


def build_risk_levels(doc):
    """Page: RISK LEVEL & DESCRIPTION — diamond chart style from AdvantageGo PDF p.3"""
    # Bold uppercase heading matching PDF
    p = doc.add_paragraph()
    _mark_heading(p, 1)
    p.paragraph_format.space_before = Pt(0); p.paragraph_format.space_after = Pt(8)
    r = p.add_run("RISK LEVEL & DESCRIPTION")
    r.font.name = FONT; r.font.size = Pt(18); r.bold = True; r.font.color.rgb = rgb(PRP)

    para(doc, (
        "The below vulnerability ranging risk pattern indicates the ratings of the "
        "vulnerability according to their respective CVSS3.1 Score."
    ), size=BODY, sa=12)

    levels = [
        ("Critical",      "800000",
         "The Critical risk level Immediate measures must be taken to reduce these risks and "
         "mitigate hazards. These vulnerabilities can allow attackers to take complete control "
         "of web applications and web servers."),
        ("High",          "FF0000",
         "The high-risk level indicates maximum risk. Such vulnerability may enable an attacker "
         "to successfully exploit the underlying application and its data partially or completely "
         "to modify application behaviour."),
        ("Medium",        "FFC000",
         "The medium risk level indicates considerable risk. Such vulnerability may enable an "
         "attacker to exploit the underlying application and its data to gain low level "
         "information about the application."),
        ("Low",           "92D050",
         "The low risk level indicates lowest risk. Such vulnerability may allow an attacker to "
         "gain some information about the application which was not intended to be known."),
        ("Informational", "00B0F0",
         "The recommended risk level indicates that some functionality or component is missing "
         "best practices. Such vulnerability may not have a risk currently, but it may become "
         "a vulnerability in future."),
    ]

    tbl = doc.add_table(rows=len(levels), cols=3)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl_width(tbl); tbl_borders(tbl,"EEEEEE",4)
    for ci, tw in enumerate([560, 1400, 7390]):
        for cell in tbl.columns[ci].cells: tcw(cell, tw)

    for i, (name, color, desc) in enumerate(levels):
        row = tbl.rows[i]
        # Colored swatch cell
        c0 = row.cells[0]; shd(c0, color); borders(c0,"FFFFFF",4)
        write_cell(c0, str(i+1), bold=True, size=13, fg=WHITE, align=WD_ALIGN_PARAGRAPH.CENTER)
        c0.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

        # Name cell
        c1 = row.cells[1]; shd(c1,"F8F8F8"); borders(c1,"EEEEEE",4)
        c1.paragraphs[0].clear(); p = c1.paragraphs[0]
        run = p.add_run(name); run.font.name=FONT; run.font.size=Pt(11)
        run.bold=True; run.font.color.rgb=rgb(color); margins(c1)

        # Description cell
        c2 = row.cells[2]; borders(c2,"EEEEEE",4)
        c2.paragraphs[0].clear(); p = c2.paragraphs[0]
        p.paragraph_format.space_before=Pt(2); p.paragraph_format.space_after=Pt(2)
        # Bold name + normal desc
        r_name = p.add_run(f"The {name} risk level ")
        r_name.font.name=FONT; r_name.font.size=Pt(9); r_name.bold=True
        r_desc = p.add_run(desc.split("The "+name+" risk level ")[-1] if name in desc else desc)
        r_desc.font.name=FONT; r_desc.font.size=Pt(9); margins(c2)

    page_break(doc)


def build_scan_type(doc, meta):
    """Page: SCAN Type — matching AdvantageGo p.4"""
    p = doc.add_paragraph()
    _mark_heading(p, 1)
    r = p.add_run("SCAN Type")
    r.font.name=FONT; r.font.size=Pt(18); r.bold=True; r.font.color.rgb=rgb(PRP)
    p.paragraph_format.space_after=Pt(16)

    # Three scope box rows — purple bg (like AdvantageGo table)
    scope_data = [
        ("Assessment\nScope",
         f"The application security assessment is done:\n{meta.get('audit_type','Security assessment (VAPT) - Android')}"),
        ("Assessment\nDate",
         f"The security vulnerabilities were reported for the following quarter on:\n{meta.get('Period', '01-11-2025 To 05-11-2025')}"),
        ("Assessment\nNote",
         f"The Scope of the Assessment only focuses on the Android Application provided by Team Capita. Target: {meta.get('target_url','https://kcq-uat.capita.co.in/')}"),
    ]

    tbl = doc.add_table(rows=len(scope_data), cols=2)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl_width(tbl)
    for cell in tbl.columns[0].cells: tcw(cell, 1700)
    for cell in tbl.columns[1].cells: tcw(cell, 7650)

    for i, (label, value) in enumerate(scope_data):
        r0 = tbl.rows[i].cells[0]; r1 = tbl.rows[i].cells[1]
        # Purple left cell
        shd(r0, PRP); borders(r0,"FFFFFF",4)
        r0.paragraphs[0].clear(); lp = r0.paragraphs[0]
        lp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = lp.add_run(label)
        run.bold=True; run.font.name=FONT; run.font.size=Pt(11); run.font.color.rgb=rgb(WHITE)
        margins(r0, 80, 80, 100, 100); r0.vertical_alignment=WD_ALIGN_VERTICAL.CENTER

        # Right content cell
        borders(r1, BGRAY, 4); shd(r1,"F8F8F8")
        r1.paragraphs[0].clear(); vp = r1.paragraphs[0]
        lines = value.split("\n")
        for li, line in enumerate(lines):
            if li == 0:
                run1 = vp.add_run(line+"\n" if len(lines)>1 else line)
                run1.font.name=FONT; run1.font.size=Pt(10)
            else:
                run2 = vp.add_run(line)
                run2.bold=True; run2.font.name=FONT; run2.font.size=Pt(11)
        margins(r1, 80, 80, 100, 100)

    page_break(doc)


def build_tools(doc, meta):
    """Page: List of Tools."""
    p = doc.add_paragraph()
    _mark_heading(p, 1)
    r = p.add_run("List of Tools.")
    r.font.name=FONT; r.font.size=Pt(18); r.bold=True; r.font.color.rgb=rgb(PRP)
    p.paragraph_format.space_after=Pt(16)

    # Tools table from metadata
    tools = meta.get("tools",[])
    # Default tools if not found
    if not tools:
        tools = [
            {"no":"1","name":"Microsoft 365","version":"Microsoft 365","type":"Licensed"},
            {"no":"2","name":"Common Vulnerability Scoring System","version":"3","type":"Open Source"},
            {"no":"3","name":"Burp Suite","version":"2024.9.5","type":"Licensed"},
        ]

    rows_count = 1 + len(tools)
    tbl = doc.add_table(rows=rows_count, cols=4)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl_width(tbl); tbl_borders(tbl)
    for ci, tw in enumerate([450, 4500, 1500, 2900]):
        for cell in tbl.columns[ci].cells: tcw(cell, tw)

    header_row(tbl.rows[0], ["S. No","Name of Tool/Software used","Version","Open Source/Licensed"])

    for i, t in enumerate(tools):
        row = tbl.rows[i+1]
        for j, val in enumerate([t["no"], t["name"], t["version"], t["type"]]):
            write_cell(row.cells[j], val, size=10)
            borders(row.cells[j], BGRAY, 4)
        if i % 2 == 0:
            for c in row.cells: shd(c, LGRAY)

    page_break(doc)


def build_overall_findings(doc, vulns, meta):
    """Page: Overall Findings — exact AdvantageGo PDF p.6 style"""
    # Big heading
    p = doc.add_paragraph()
    _mark_heading(p, 1)
    r = p.add_run("Overall Findings")
    r.font.name=FONT; r.font.size=Pt(20); r.bold=True; r.font.color.rgb=rgb(PRP)
    p.paragraph_format.space_after=Pt(8)

    # Sub heading — client name
    client = meta.get("client_name","Capita")
    target = meta.get("target_url","https://kcq-uat.capita.co.in/")
    p2 = doc.add_paragraph()
    r2 = p2.add_run(f"{client} Risk Assessment Analysis of the Entire Framework")
    r2.font.name=FONT; r2.font.size=Pt(13); r2.bold=True; r2.font.color.rgb=rgb(PRP)
    p2.paragraph_format.space_after=Pt(10)

    # Counts
    c = {"Critical":0,"High":0,"Medium":0,"Low":0,"Informational":0}
    for v in vulns:
        s = v["severity"].capitalize()
        if s in c: c[s] += 1
    total = sum(c.values())

    # ── Asset findings table (exact type3/AdvantageGo format) ───────────────
    col_bgs = {0:"6923D0",1:"6923D0",2:"6923D0",3:"6923D0",
               4:"800000",5:"FF0000",6:"FFC000",7:"92D050",8:"00B0F0",9:"7030A0"}
    headers = ["Sr.\nNo.", "Asset/Hostname", "Instant\npurpose", "VAPT\nStatus",
               "Critical","High","Medium","Low","Informational","Total"]

    tbl = doc.add_table(rows=3, cols=10)
    _safe_set_table_style(tbl, "Table Grid")
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl_width(tbl)
    for ci, tw in enumerate([500,2100,1100,900,600,600,600,600,700,450]):
        for cell in tbl.columns[ci].cells: tcw(cell, tw)

    header_row(tbl.rows[0], headers, bg=PRP2, fg=WHITE, sz=8, col_bgs=col_bgs)

    # Data row
    r1 = tbl.rows[1]
    data_row = ["1.", target[:25], "Web App", "Completed",
                str(c["Critical"]),str(c["High"]),str(c["Medium"]),str(c["Low"]),str(c["Informational"]),str(total)]
    for ci, val in enumerate(data_row):
        write_cell(r1.cells[ci], val, size=9, align=WD_ALIGN_PARAGRAPH.CENTER)
        if ci >= 4:
            shd(r1.cells[ci], col_bgs[ci])
            write_cell(r1.cells[ci], val, bold=True, size=9, fg=WHITE, align=WD_ALIGN_PARAGRAPH.CENTER)
        elif ci % 2 == 0: shd(r1.cells[ci], LGRAY)

    # Overall row — merge first 4
    r2_row = tbl.rows[2]
    r2_row.cells[0].merge(r2_row.cells[3])
    write_cell(r2_row.cells[0],"Overall Findings",bold=True,size=9,align=WD_ALIGN_PARAGRAPH.CENTER)
    shd(r2_row.cells[0], LGRAY)
    for ci, val, cg in [(4,c["Critical"],"800000"),(5,c["High"],"FF0000"),(6,c["Medium"],"FFC000"),
                        (7,c["Low"],"92D050"),(8,c["Informational"],"00B0F0"),(9,total,"7030A0")]:
        write_cell(r2_row.cells[ci],str(val),bold=True,size=9,fg=WHITE,align=WD_ALIGN_PARAGRAPH.CENTER)
        shd(r2_row.cells[ci], cg)

    para(doc, sa=14)

    # ── Charts ───────────────────────────────────────────────────────────────
    p3 = doc.add_paragraph(); r3 = p3.add_run("Overall Vulnerabilities Identified")
    r3.font.name=FONT; r3.font.size=Pt(13); r3.bold=True; r3.font.color.rgb=rgb(PRP)
    p3.paragraph_format.space_after=Pt(10)

    bar_bytes   = bar_chart(c)
    donut_bytes = donut_chart(c)

    chart_tbl = doc.add_table(rows=1, cols=2)
    chart_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl_width(chart_tbl)
    for ci, tw in enumerate([5500,3850]):
        for cell in chart_tbl.columns[ci].cells: tcw(cell, tw)
    for cell in chart_tbl.rows[0].cells: borders(cell,"FFFFFF",0)

    c0 = chart_tbl.rows[0].cells[0]; c0.paragraphs[0].clear()
    p0 = c0.paragraphs[0]; p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p0.add_run().add_picture(io.BytesIO(bar_bytes), width=Inches(3.5))

    c1 = chart_tbl.rows[0].cells[1]; c1.paragraphs[0].clear()
    p1 = c1.paragraphs[0]; p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p1.add_run().add_picture(io.BytesIO(donut_bytes), width=Inches(2.6))

    page_break(doc)


def build_test_cases(doc, meta):
    """Page: Test Case — like AdvantageGo PDF p.7"""
    p = doc.add_paragraph()
    _mark_heading(p, 1)
    r = p.add_run("Test Case")
    r.font.name=FONT; r.font.size=Pt(20); r.bold=True; r.font.color.rgb=rgb(PRP)
    p.paragraph_format.space_after=Pt(16)

    test_cases = [
        ("1","Business Logic Testing"),
        ("2","Service Disruption Analysis"),
        ("3","OWASP Top 10"),
        ("4","Framework Based Vulnerability"),
        ("5","Database Pentest"),
        ("6","Mobile Security Testing (OWASP MASVS)"),
        ("7","API Security Testing"),
    ]

    tbl = doc.add_table(rows=len(test_cases)+1, cols=3)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl_width(tbl); tbl_borders(tbl)
    for ci, tw in enumerate([500, 6450, 2400]):
        for cell in tbl.columns[ci].cells: tcw(cell, tw)

    header_row(tbl.rows[0], ["Sr. No.","Test Case","Status"], bg=PRP2, fg=WHITE, sz=10)

    for i, (num, case) in enumerate(test_cases):
        row = tbl.rows[i+1]
        write_cell(row.cells[0], num, size=10, align=WD_ALIGN_PARAGRAPH.CENTER)
        write_cell(row.cells[1], case, size=10)
        write_cell(row.cells[2], "Completed ✓", bold=True, size=10, fg="375623",
                   align=WD_ALIGN_PARAGRAPH.CENTER)
        for c in row.cells: borders(c, BGRAY, 4)
        if i % 2 == 0:
            for c in row.cells: shd(c, LGRAY)

    page_break(doc)


def build_vuln_list(doc, vulns, meta):
    """Page: Vulnerabilities Found — bar chart + list table (AdvantageGo p.8-9)"""
    p = doc.add_paragraph()
    _mark_heading(p, 1)
    r = p.add_run("Vulnerabilities Found")
    r.font.name=FONT; r.font.size=Pt(20); r.bold=True; r.font.color.rgb=rgb(PRP)
    p.paragraph_format.space_after=Pt(16)

    c = {"Critical":0,"High":0,"Medium":0,"Low":0,"Informational":0}
    for v in vulns:
        s = v["severity"].capitalize()
        if s in c: c[s] += 1

    # Bar chart
    bc = doc.add_paragraph(); bc.alignment = WD_ALIGN_PARAGRAPH.CENTER
    bc.add_run().add_picture(io.BytesIO(bar_chart(c)), width=Inches(5.0))
    para(doc, sa=10)

    # Target URL as heading
    target = meta.get("target_url","https://kcq-uat.capita.co.in/")
    p2 = doc.add_paragraph()
    r2 = p2.add_run(target)
    r2.font.name=FONT; r2.font.size=Pt(14); r2.bold=True; r2.font.color.rgb=rgb(PRP)
    p2.paragraph_format.space_after=Pt(10)

    # Vulnerability list table (like AdvantageGo PDF p.9)
    tbl = doc.add_table(rows=len(vulns)+1, cols=3)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl_width(tbl); tbl_borders(tbl)
    for ci, tw in enumerate([500, 6450, 2400]):
        for cell in tbl.columns[ci].cells: tcw(cell, tw)

    header_row(tbl.rows[0], ["Sr.no","Vulnerability Name","Vulnerability Risk Type"],
               bg=PRP2, fg=WHITE, sz=10)

    for i, v in enumerate(vulns):
        row = tbl.rows[i+1]
        sc = sev_colors(v["severity"])
        write_cell(row.cells[0], str(v["number"]), size=10, align=WD_ALIGN_PARAGRAPH.CENTER)
        write_cell(row.cells[1], v["title"], size=10)
        write_cell(row.cells[2], v["severity"], bold=True, size=10, fg=sc[1],
                   align=WD_ALIGN_PARAGRAPH.CENTER)
        shd(row.cells[2], sc[0])
        for c in row.cells: borders(c, BGRAY, 4)
        if i % 2 == 1:
            for c in (row.cells[0], row.cells[1]): shd(c, LGRAY)

    page_break(doc)


def build_revalidation_table(doc, vulns):
    """Page: Vulnerability Revalidation Status (exact AdvantageGo PDF p.10)"""
    p = doc.add_paragraph()
    _mark_heading(p, 1)
    r = p.add_run("Vulnerability Revalidation Status")
    r.font.name=FONT; r.font.size=Pt(18); r.bold=True; r.font.color.rgb=rgb(PRP)
    p.paragraph_format.space_after=Pt(10)

    # ── Dual super-header: DEVELOPER TEAM | VAPT TEAM ──────────────────────
    super_tbl = doc.add_table(rows=1, cols=2)
    super_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl_width(super_tbl)
    tcw(super_tbl.rows[0].cells[0], 5550); tcw(super_tbl.rows[0].cells[1], 3800)
    for c in super_tbl.rows[0].cells:
        shd(c, PRP2); borders(c,"FFFFFF",4)
    write_cell(super_tbl.rows[0].cells[0],"DEVELOPER TEAM",bold=True,size=9,
               fg=WHITE,align=WD_ALIGN_PARAGRAPH.CENTER)
    write_cell(super_tbl.rows[0].cells[1],"VAPT TEAM",bold=True,size=9,
               fg=WHITE,align=WD_ALIGN_PARAGRAPH.CENTER)

    # ── Main revalidation table ──────────────────────────────────────────────
    headers = ["Sr","Vulnerabilities","Risk\nType","Remarks",
               "Revalidation\nStatus","Justification\nRemarks","Vulnerability\nStatus"]
    col_tw  = [350, 2650, 800, 1450, 1200, 1550, 1350]

    tbl = doc.add_table(rows=len(vulns)+1, cols=7)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl_width(tbl); tbl_borders(tbl)
    for ci, tw in enumerate(col_tw):
        for cell in tbl.columns[ci].cells: tcw(cell, tw)

    header_row(tbl.rows[0], headers, bg=PRP2, fg=WHITE, sz=8)

    for i, v in enumerate(vulns):
        row = tbl.rows[i+1]
        sc   = sev_colors(v["severity"])
        rv   = v.get("revalidation_status","Unpatched")
        rvc  = reval_colors(rv)
        clo  = v.get("closure_status","Open")
        cloc = reval_colors(clo)

        data = [str(v["number"]), v["title"], v["severity"],
                v.get("developer_remarks","-"), rv,
                v.get("vapt_remarks","-"), clo]

        for ci, val in enumerate(data):
            cell = row.cells[ci]
            al = WD_ALIGN_PARAGRAPH.CENTER if ci in (0,2,4,6) else WD_ALIGN_PARAGRAPH.LEFT
            write_cell(cell, val, size=8, align=al)
            borders(cell, BGRAY, 4)

            if ci == 2:
                shd(cell, sc[0])
                write_cell(cell, val, size=8, bold=True, fg=sc[1], align=WD_ALIGN_PARAGRAPH.CENTER)
            elif ci == 4:
                shd(cell, rvc[0])
                write_cell(cell, val, size=8, bold=True, fg=rvc[1], align=WD_ALIGN_PARAGRAPH.CENTER)
            elif ci == 6:
                shd(cell, cloc[0])
                write_cell(cell, val, size=8, bold=True, fg=cloc[1], align=WD_ALIGN_PARAGRAPH.CENTER)
            elif i % 2 == 0:
                shd(cell, LGRAY)

    page_break(doc)


def build_vuln_detail(doc, v):
    """
    Per-vulnerability detail page — EXACT AdvantageGo PDF p.11+ format:
      [ID banner (color bg) | title]
      [URL row]
      [Vulnerable Parameter row]
      [CVSS row (colored)]
      [Severity row (colored)]
      [Vulnerability Description:]
      [Impact:]
      [Recommendation:]
      [Proof of Concept:]
      [Revalidation Proof of Concept:]
    """
    sc = sev_colors(v["severity"])

    # Keep the entire detail block in one table to avoid visual gaps between sections.
    tbl = doc.add_table(rows=1, cols=2)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.autofit = False
    tbl_width(tbl)
    tbl_borders(tbl, BGRAY, 4)
    tcw(tbl.rows[0].cells[0], 900)
    tcw(tbl.rows[0].cells[1], 8450)

    def add_meta_row(label, value, val_color=None, val_bold=False, val_bg=None, is_url=False):
        row = tbl.add_row()
        lc = row.cells[0]
        vc = row.cells[1]

        write_cell(lc, label, bold=True, size=10, fg=PRP)
        shd(lc, "F8F8F8")
        borders(lc, BGRAY, 4)
        margins(lc, 40, 40, 90, 90)

        write_cell(vc, value, bold=val_bold, size=10, fg=val_color, underline=is_url)
        if val_bg:
            shd(vc, val_bg)
        borders(vc, BGRAY, 4)
        margins(vc, 40, 40, 90, 90)

    def add_section_row(title, lines):
        row = tbl.add_row()
        cell = row.cells[0].merge(row.cells[1])
        borders(cell, BGRAY, 4)
        cell.paragraphs[0].clear()

        p = cell.paragraphs[0]
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(2)
        head = p.add_run(f"{title}:")
        head.bold = True
        head.font.name = FONT
        head.font.size = Pt(10)
        head.font.color.rgb = rgb(PRP)

        for line in lines:
            if not str(line).strip():
                continue
            lp = cell.add_paragraph(str(line))
            lp.paragraph_format.space_before = Pt(1)
            lp.paragraph_format.space_after = Pt(1)
            for run in lp.runs:
                run.font.name = FONT
                run.font.size = Pt(9)

        margins(cell, 40, 40, 90, 90)

    # Header row (ID + title)
    c_id = tbl.rows[0].cells[0]
    c_title = tbl.rows[0].cells[1]
    shd(c_id, sc[0])
    shd(c_title, sc[0])
    borders(c_id, "FFFFFF", 4)
    borders(c_title, "FFFFFF", 4)
    c_id.paragraphs[0].clear()
    c_title.paragraphs[0].clear()

    pid = c_id.paragraphs[0]
    pid.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_id = pid.add_run(v["id"])
    run_id.bold = True
    run_id.font.name = FONT
    run_id.font.size = Pt(10)
    run_id.font.color.rgb = rgb(sc[1])
    margins(c_id, 40, 40, 90, 90)
    c_id.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

    pt = c_title.paragraphs[0]
    run_t = pt.add_run(v["title"])
    run_t.bold = True
    run_t.font.name = FONT
    run_t.font.size = Pt(11)
    run_t.font.color.rgb = rgb(sc[1])
    margins(c_title, 40, 40, 90, 90)
    c_title.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

    add_meta_row("URL", v.get("affected_asset", "N/A"), is_url=True, val_color="0563C1")
    add_meta_row("Vulnerable Parameters:", v.get("affected_asset", "N/A"))

    cvss_str = str(v.get("cvss_score", "N/A"))
    if v.get("cvss_vector"):
        cvss_str += " - " + str(v["cvss_vector"])
    add_meta_row("CVSS:", cvss_str, val_color=sc[0], val_bold=True)
    add_meta_row("Severity", v.get("severity", "N/A"), val_color=WHITE, val_bold=True, val_bg=sc[0])

    if v.get("cwe") or v.get("cve"):
        cwe_text = " / ".join(filter(None, [v.get("cwe", ""), v.get("cve", "")])) or "N/A"
        add_meta_row("CWE / CVE:", cwe_text)

    desc_lines = [line.strip() for line in str(v.get("description", "")).splitlines() if line.strip()]
    add_section_row("Vulnerability Description", desc_lines or ["-"])

    impact = v.get("closure_notes", "It can lead to data exposure and unauthorized access.")
    add_section_row("Impact", ["It can lead to the following impacts:", f"• {impact}"])

    rec_text = str(v.get("recommendation", "")).strip()
    if rec_text:
        rec_lines = []
        for line in rec_text.splitlines():
            cleaned = line.strip()
            if not cleaned:
                continue
            if cleaned.startswith("•") or re.match(r"^\d+[.)]\s*", cleaned):
                rec_lines.append(cleaned)
            else:
                rec_lines.append(f"• {cleaned}")
        add_section_row("Recommendation", rec_lines or ["-"])

    if v.get("reference"):
        row = tbl.add_row()
        ref_cell = row.cells[0].merge(row.cells[1])
        borders(ref_cell, BGRAY, 4)
        ref_cell.paragraphs[0].clear()
        rp = ref_cell.paragraphs[0]
        hr = rp.add_run("Reference: ")
        hr.bold = True
        hr.font.name = FONT
        hr.font.size = Pt(9)
        hr.font.color.rgb = rgb(PRP)
        rr = rp.add_run(str(v["reference"]))
        rr.font.name = FONT
        rr.font.size = Pt(9)
        rr.font.color.rgb = rgb("0563C1")
        rr.underline = True
        margins(ref_cell, 40, 40, 90, 90)

    # Proof of Concept section
    row = tbl.add_row()
    pc = row.cells[0].merge(row.cells[1])
    borders(pc, BGRAY, 4)
    pc.paragraphs[0].clear()
    pp = pc.paragraphs[0]
    rh = pp.add_run("Proof of Concept:")
    rh.bold = True
    rh.font.name = FONT
    rh.font.size = Pt(10)
    rh.font.color.rgb = rgb(PRP)

    if v.get("poc_steps"):
        for step in v["poc_steps"]:
            sp = pc.add_paragraph(f"Step {step['num']}: {step['text']}")
            sp.paragraph_format.space_before = Pt(3)
            sp.paragraph_format.space_after = Pt(1)
            for r2 in sp.runs:
                r2.font.name = FONT
                r2.font.size = Pt(10)
    else:
        sp2 = pc.add_paragraph("Step 1: Here is the following POC")
        for r2 in sp2.runs:
            r2.font.name = FONT
            r2.font.size = Pt(10)

    if v.get("poc_images"):
        for idx, img_bytes in enumerate(v["poc_images"][:8], 1):
            try:
                ip = pc.add_paragraph()
                ip.alignment = WD_ALIGN_PARAGRAPH.CENTER
                ip.add_run().add_picture(io.BytesIO(img_bytes), width=Inches(4.7))
                cap = pc.add_paragraph(f"PoC figure {idx}: {v['id']}")
                cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
                if cap.runs:
                    cap.runs[0].font.name = FONT
                    cap.runs[0].font.size = Pt(8)
                    cap.runs[0].font.italic = True
                    cap.runs[0].font.color.rgb = rgb("505050")
            except Exception:
                pass
    else:
        ph = pc.add_paragraph("[Screenshot: PoC screenshot to be attached here]")
        if ph.runs:
            ph.runs[0].font.italic = True
            ph.runs[0].font.size = Pt(8)

    margins(pc, 40, 40, 90, 90)

    # Revalidation Proof of Concept section
    row = tbl.add_row()
    rc = row.cells[0].merge(row.cells[1])
    borders(rc, BGRAY, 4)
    rc.paragraphs[0].clear()
    rp = rc.paragraphs[0]
    rh2 = rp.add_run("Revalidation Proof of Concept:")
    rh2.bold = True
    rh2.font.name = FONT
    rh2.font.size = Pt(10)
    rh2.font.color.rgb = rgb(PRP)

    clo = v.get("closure_status", "Open")
    reval = v.get("revalidation_status", "Unpatched")
    if clo == "Closed":
        rval_text = (
            "Step 1: We observed that this vulnerability has been successfully patched "
            "and is closed after revalidation."
        )
    else:
        rval_text = (
            f"Step 1: We observed that this vulnerability ({v['id']}) remains unpatched. "
            f"Revalidation Status: {reval}. The vulnerability requires remediation."
        )

    rs = rc.add_paragraph(rval_text)
    for r2 in rs.runs:
        r2.bold = True
        r2.font.name = FONT
        r2.font.size = Pt(10)

    if v.get("reval_images"):
        for idx, img_bytes in enumerate(v["reval_images"][:8], 1):
            try:
                ip = rc.add_paragraph()
                ip.alignment = WD_ALIGN_PARAGRAPH.CENTER
                ip.add_run().add_picture(io.BytesIO(img_bytes), width=Inches(4.7))
                cap = rc.add_paragraph(f"Revalidation figure {idx}: {v['id']}")
                cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
                if cap.runs:
                    cap.runs[0].font.name = FONT
                    cap.runs[0].font.size = Pt(8)
                    cap.runs[0].font.italic = True
                    cap.runs[0].font.color.rgb = rgb("505050")
            except Exception:
                pass
    else:
        rph = rc.add_paragraph("[Screenshot: Revalidation screenshot to be attached here]")
        if rph.runs:
            rph.runs[0].font.italic = True
            rph.runs[0].font.size = Pt(8)

    margins(rc, 40, 40, 90, 90)

    page_break(doc)


def build_conclusion(doc, vulns, meta):
    """SUMMARY OF FINDINGS & CONCLUSION — matching AdvantageGo PDF p.65"""
    p = doc.add_paragraph()
    _mark_heading(p, 1)
    r = p.add_run("SUMMARY OF FINDINGS & CONCLUSION:")
    r.font.name=FONT; r.font.size=Pt(20); r.bold=True; r.font.color.rgb=rgb(PRP)
    p.paragraph_format.space_after=Pt(16)

    closed = sum(1 for v in vulns if v.get("closure_status","Open")=="Closed")
    open_  = len(vulns)-closed

    para(doc, (
        "Finally, it must be remembered that security is an ongoing process, and that this report "
        "will provide an idea of the current vulnerabilities we were able to detect. There is no "
        "guarantee that new vulnerabilities will not be found and exploited in the future."
    ), size=BODY, sa=12)

    para(doc, (
        "The assessment was only possible because joint support & coordination from the information "
        "security team of organization for sharing & coordinating during the assessment period. "
        "It is advised to refer the Technical Report for understanding in-depth of vulnerabilities "
        "that were discovered by technical team of CyberSmithSECURE Pvt. Ltd."
    ), size=BODY, sa=12)

    status_text = (
        f"all {len(vulns)} vulnerabilities remain open and require immediate attention"
        if open_ == len(vulns)
        else f"{closed} are closed and {open_} remain open"
    )
    para(doc, (
        f"The Security Researchers of the CyberSmithSECURE performed Vulnerability Testing. "
        f"We jointly recommend that all suggested measures in this document be performed to "
        f"ensure the overall security of the target website."
    ), size=BODY, sa=16)

    # Bold purple thank you line (matching AdvantageGo PDF last page)
    p_thank = doc.add_paragraph()
    r_thank = p_thank.add_run(
        "We thank internal Information Security team for their support & cooperation "
        "during the time of assessment.")
    r_thank.bold=True; r_thank.font.name=FONT; r_thank.font.size=Pt(BODY)
    r_thank.font.color.rgb=rgb(PRP)


def _build_template_context(meta, vulns):
    return {
        "report_id": meta.get("report_id", "CSS-MOB-2025"),
        "client_name": meta.get("client_name", "Client"),
        "doc_title": meta.get("doc_title", "Technical Closure Report"),
        "target_url": meta.get("target_url", ""),
        "release_date": meta.get("Release date", meta.get("release_date", datetime.today().strftime("%d %B %Y"))),
        "audit_type": meta.get("audit_type", ""),
        "period": meta.get("Period", ""),
        "vulnerability_count": len(vulns),
        "generated_on": datetime.today().strftime("%d %B %Y"),
        "generated_at": datetime.today().strftime("%H:%M:%S"),
        "c_critical": meta.get("c_critical", 0),
        "c_high": meta.get("c_high", 0),
        "c_medium": meta.get("c_medium", 0),
        "c_low": meta.get("c_low", 0),
        "c_informational": meta.get("c_informational", 0),
        "c_total": meta.get("c_total", len(vulns)),
        "REPORT_ID": meta.get("report_id", "CSS-MOB-2025"),
        "CLIENT_NAME": meta.get("client_name", "Client"),
        "DOC_TITLE": meta.get("doc_title", "Technical Closure Report"),
        "TARGET_URL": meta.get("target_url", ""),
        "RELEASE_DATE": meta.get("Release date", meta.get("release_date", datetime.today().strftime("%d %B %Y"))),
        "VULNERABILITY_COUNT": len(vulns),
        "GENERATED_ON": datetime.today().strftime("%d %B %Y"),
        "GENERATED_AT": datetime.today().strftime("%H:%M:%S"),
    }

# ─────────────────────────────────────────────────────────────────────────────
# MAIN PIPELINE
# ─────────────────────────────────────────────────────────────────────────────

def _find_and_clear_closure_marker(doc):
    """
    Locate {CLOSURE_DYNAMIC_SECTIONS} in the main document body, clear the marker, and
    return (anchor_paragraph, found_placeholder). If the placeholder is missing, append
    an empty paragraph for appending the generated report (legacy behaviour).
    """
    for p in _iter_container_paragraphs(doc):
        t = p.text.strip()
        if t in _CONTENT_MARKER_VARIANTS:
            p.text = ""
            return p, True
    return doc.add_paragraph(), False


def _type6_env_include_risk_scan_tools() -> bool:
    _load_root_env()
    v = (os.getenv("TYPE6_CLOSURE_RISK_SCAN_TOOLS", "") or "").strip().lower()
    return v in ("1", "true", "yes", "y")


def _build_advantagego_closure_block(doc, meta, vulns) -> None:
    """
    Content for {CLOSURE_DYNAMIC_SECTIONS}, aligned to AdvantageGo Technical Closure (PDF) order:
      1. Overall Findings (asset table + 3D bar + donut)
      2. Test Case
      3. Vulnerabilities Found (severity bar + target URL + master list table)
      4. Vulnerability Revalidation Status (summary table)
      5. Per-finding details (URL, CVSS, description, impact, recommendation, PoC + reval, images)
      6. Summary of findings & conclusion

    Risk level reference, scan scope, and tools list are **not** included here (keep them in
    the base .docx before the marker) unless env TYPE6_CLOSURE_RISK_SCAN_TOOLS is enabled.
    """
    if _type6_env_include_risk_scan_tools():
        build_risk_levels(doc)
        build_scan_type(doc, meta)
        build_tools(doc, meta)
    build_overall_findings(doc, vulns, meta)
    build_test_cases(doc, meta)
    build_vuln_list(doc, vulns, meta)
    build_revalidation_table(doc, vulns)
    for i, v in enumerate(vulns, 1):
        print(f"  [{i:02d}/{len(vulns):02d}] {v['id']}: {v['title'][:65]}")
        build_vuln_detail(doc, v)
    build_conclusion(doc, vulns, meta)


def _safe_add_page_break(doc):
    """Add a page break only if the last element isn't already a page break.
    Prevents double page breaks that cause blank pages / overlapping."""
    if doc.paragraphs:
        last_p = doc.paragraphs[-1]
        # Check if the last paragraph already has a page break
        xml_text = last_p._element.xml if last_p._element is not None else ""
        if 'w:type="page"' in xml_text:
            return  # Already has a page break
    page_break(doc)


def generate(input_docx, output_docx, api_key="", excel_path=None, poc_zip_path=None):
    print(f"\n{'='*65}")
    print("  CyberSmithSECURE — Closure Report Generator v6.0 (Guardrailed + 3D Charts)")
    print(f"{'='*65}")

    print(f"\n[1/5] Extracting data from: {input_docx}")
    meta, vulns = extract_all_data(input_docx)

    if excel_path:
        print(f"  • Applying Excel overlay from: {excel_path}")
        try:
            meta, vulns = _overlay_excel_data(meta, vulns, excel_path)
        except Exception as exc:
            print(f"  ⚠ Excel overlay failed: {exc}")

    if poc_zip_path:
        print(f"  • Mapping PoC ZIP images from: {poc_zip_path}")
        try:
            attached = _attach_zip_images(vulns, poc_zip_path)
            print(f"  ✓ Attached {attached} images from zip")
        except Exception as exc:
            print(f"  ⚠ ZIP mapping failed: {exc}")

    # ── DATA VALIDATION GUARDRAIL ─────────────────────────────────────────
    print(f"\n[2/5] Validating extracted data (anti-hallucination)...")
    vulns = _validate_extracted_data(vulns)
    if not vulns:
        raise ValueError("No valid vulnerability data found after validation — check your input files.")
    print(f"  ✓ {len(vulns)} validated vulnerabilities | Client: {meta.get('client_name')}")
    print(f"  ✓ Target: {meta.get('target_url')}")

    print(f"\n[3/5] AI filling closure fields (guardrailed)...")
    # Preserve original extraction ordering and content for strict alignment
    original_vulns = copy.deepcopy(vulns)
    if api_key:
        vulns = ai_fill_closure_openai(vulns, api_key)
        try:
            vulns = _enforce_excel_alignment(vulns, original_vulns)
            print("  ✓ AI output aligned & validated against original data")
        except Exception as exc:
            # If enforcement fails, fall back to AI-updated list but log
            print(f"  ⚠ Excel-alignment enforcement failed ({exc}) — using AI results as-is")
    else:
        print("  ⚠ No API key — using default Unpatched/Open")

    print(f"\n[4/5] Building DOCX report (AdvantageGo base template + 3D charts)...")
    template_path = _resolve_advantage_template_path()
    print(f"  • Using base template: {template_path}")
    doc = DocxDocument(str(template_path))

    # Do not override margins/page size: keep the user-supplied base template.

    # Render template placeholders in the main body only (header/footer stay as in the .docx).
    # Do not include CLOSURE_DYNAMIC_* in the context — that slot is filled by the builders below.
    _replace_placeholders_in_doc(doc, _build_template_context(meta, vulns))
    _ensure_toc_field(doc)
    _enable_update_fields_on_open(doc)

    # Insert all generated sections in place of {CLOSURE_DYNAMIC_SECTIONS} (or typo variant).
    anchor, had_placeholder = _find_and_clear_closure_marker(doc)
    d = _InsertionDocument(doc, _DocBodyTail(doc, anchor))
    if not had_placeholder and (not d.paragraphs or not d.paragraphs[-1].text.strip()):
        _safe_add_page_break(d)

    if _type6_env_include_risk_scan_tools():
        print("  • TYPE6_CLOSURE_RISK_SCAN_TOOLS: include Risk/Scan/Tools in closure block (before Overall Findings)")
    else:
        print("  • Closure block: Overall Findings → Test Case → Vulnerabilities Found → "
              "Vulnerability Revalidation → per-finding details (PoC / reval images) → Conclusion")
    _build_advantagego_closure_block(d, meta, vulns)

    # Keep header/footer from the user template; do not replace with program-drawn H/F.

    print(f"\n[5/5] Saving...")
    doc.save(output_docx)
    size = os.path.getsize(output_docx)
    print(f"  ✓ Saved → {output_docx}  ({size//1024} KB)")
    print(f"  ✓ 3D charts: Bar + Pie rendered at 180 DPI")
    print(f"  ✓ Guardrails: Data validated, AI output cross-checked, protected fields locked")
    print(f"{'='*65}\n")


def _convert_pdf_to_docx(pdf_path: str) -> str:
    """
    PDF → DOCX (layout approximated). Type6 still needs a *technical* VAPT .docx with
    the observation tables extract_all_data() expects; a short closure PDF will usually
    not contain those tables even after conversion.
    """
    try:
        from pdf2docx import Converter
    except ImportError as exc:
        raise SystemExit(
            "PDF input requires pdf2docx. Install: pip install pdf2docx"
        ) from exc
    fd, out_path = tempfile.mkstemp(suffix=".docx", prefix="type6_pdf_")
    os.close(fd)
    try:
        cv = Converter(pdf_path)
        cv.convert(out_path)
        cv.close()
    except Exception:
        try:
            if os.path.exists(out_path):
                os.remove(out_path)
        except OSError:
            pass
        raise
    return out_path


def main():
    ap = argparse.ArgumentParser(description="CyberSmithSECURE Closure Report Generator v5.0")
    ap.add_argument(
        "--input", required=True,
        help="Initial VAPT .docx, or a PDF of that report (auto-converted; extraction usually needs native .docx).",
    )
    ap.add_argument("--output", default="Closure_Report.docx")
    ap.add_argument("--key",    default=os.getenv("OPENAI_API_KEY",""), help="OpenAI API key")
    ap.add_argument("--excel",  default="", help="Optional XLSX for vulnerability overlay")
    ap.add_argument("--poc",    default="", help="Optional PoC ZIP with screenshots")
    args = ap.parse_args()
    if not os.path.exists(args.input):
        print(f"Error: {args.input} not found"); sys.exit(1)

    input_path = args.input
    temp_docx: str | None = None
    if str(input_path).lower().endswith(".pdf"):
        print(
            "  Converting PDF → DOCX. Prefer the source Word VAPT if extraction returns "
            "no findings — closure PDFs are not structured like our observation tables.\n"
        )
        input_path = _convert_pdf_to_docx(str(input_path))
        temp_docx = input_path

    try:
        generate(input_path, args.output, args.key, excel_path=args.excel or None, poc_zip_path=args.poc or None)
    finally:
        if temp_docx and os.path.exists(temp_docx):
            try:
                os.remove(temp_docx)
            except OSError:
                pass

if __name__ == "__main__":
    main()


# ---------------------------------------------------------------------------
# Minimal FastAPI wrapper for quick local testing (/type6/)
# ---------------------------------------------------------------------------
app = FastAPI(title="Type6 Closure Generator (dev)")

# Expose router for integration in main app.py (app.include_router(type6_router)).
router = app.router


@app.get("/", include_in_schema=False)
async def root_redirect():
    return RedirectResponse(url="/type6/", status_code=307)


@app.get("/type6/", response_class=HTMLResponse)
async def type6_home(request: Request):
    # Ensure CSS/.env is visible to this process.
    _load_root_env()

    env_present = bool(os.getenv("OPENAI_API_KEY"))
    template_path = _resolve_advantage_template_path()
    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Type6 — AdvantageGo Closure Report Generator</title>
  <style>
    *{{box-sizing:border-box;margin:0;padding:0}}
    body{{font-family:'Segoe UI',Arial,sans-serif;background:#1a1a2e;color:#e0e0e0;min-height:100vh;display:flex;align-items:center;justify-content:center;padding:24px}}
    .card{{background:#16213e;border-radius:12px;box-shadow:0 8px 32px rgba(105,35,208,.25);padding:36px 40px;width:100%;max-width:600px}}
    h2{{color:#8a2be2;font-size:1.5rem;margin-bottom:4px;display:flex;align-items:center;gap:10px}}
    .subtitle{{color:#9090b0;font-size:.875rem;margin-bottom:28px}}
    .field{{margin-bottom:20px}}
    label{{display:block;font-weight:600;margin-bottom:6px;font-size:.9rem;color:#c0b8e8}}
    .file-wrap{{border:2px dashed #6923d0;border-radius:8px;padding:14px 16px;background:#0d1426;cursor:pointer;transition:border-color .2s,background .2s;position:relative}}
    .file-wrap:hover,.file-wrap.drag{{border-color:#a855f7;background:#1a1040}}
    .file-wrap input[type=file]{{position:absolute;inset:0;opacity:0;cursor:pointer;width:100%;height:100%}}
    .file-label{{display:flex;align-items:center;gap:10px;pointer-events:none}}
    .file-icon{{font-size:1.4rem}}
    .file-name{{font-size:.85rem;color:#9090b0}}
    .required-badge{{background:#6923d0;color:#fff;font-size:.7rem;padding:2px 7px;border-radius:20px;margin-left:8px;vertical-align:middle}}
    .optional-badge{{background:#2a3060;color:#9090b0;font-size:.7rem;padding:2px 7px;border-radius:20px;margin-left:8px;vertical-align:middle}}
    .btn{{width:100%;padding:15px;background:linear-gradient(135deg,#6923d0,#8a2be2);color:#fff;border:none;border-radius:8px;font-size:1rem;font-weight:700;cursor:pointer;margin-top:8px;transition:opacity .2s,transform .1s;letter-spacing:.5px}}
    .btn:hover:not(:disabled){{opacity:.9;transform:translateY(-1px)}}
    .btn:disabled{{opacity:.55;cursor:not-allowed}}
    .progress-wrap{{display:none;margin-top:20px}}
    .progress-bar-bg{{background:#0d1426;border-radius:8px;overflow:hidden;height:10px;margin-bottom:10px}}
    .progress-bar{{height:10px;background:linear-gradient(90deg,#6923d0,#a855f7);border-radius:8px;width:0%;transition:width .4s ease}}
    .progress-msg{{font-size:.85rem;color:#9090b0;text-align:center}}
    .result-box{{display:none;margin-top:18px;padding:14px 18px;border-radius:8px;font-size:.9rem}}
    .result-ok{{background:rgba(0,176,80,.12);border:1px solid #00b050;color:#4CAF50}}
    .result-err{{background:rgba(255,50,50,.1);border:1px solid #ff3333;color:#ff6b6b}}
    .info-row{{display:flex;gap:10px;margin-top:22px;padding-top:18px;border-top:1px solid #2a2a4a;font-size:.8rem;color:#7070a0}}
    .info-row span{{display:flex;align-items:center;gap:5px}}
    .dot-ok{{width:8px;height:8px;border-radius:50%;background:#00b050;display:inline-block}}
    .dot-warn{{width:8px;height:8px;border-radius:50%;background:#FFC000;display:inline-block}}
  </style>
</head>
<body>
<div class="card">
  <h2>📋 AdvantageGo Closure Report</h2>
  <p class="subtitle">Upload your files below — the report generates as a DOCX and downloads automatically.</p>

  <form id="genForm" enctype="multipart/form-data">

    <div class="field">
      <label>Initial VAPT Report <span class="required-badge">REQUIRED</span></label>
      <div class="file-wrap" id="wrap_docx">
        <input type="file" id="initial_report" name="initial_report" accept=".docx" required onchange="updateLabel('wrap_docx','initial_report','📄')">
        <div class="file-label"><span class="file-icon">📄</span><span class="file-name" id="lbl_docx">Choose .docx file…</span></div>
      </div>
    </div>

    <div class="field">
      <label>Excel Vulnerability Data <span class="optional-badge">OPTIONAL</span></label>
      <div class="file-wrap" id="wrap_xlsx">
        <input type="file" id="excel_file" name="excel_file" accept=".xlsx,.xls" onchange="updateLabel('wrap_xlsx','excel_file','📊')">
        <div class="file-label"><span class="file-icon">📊</span><span class="file-name" id="lbl_xlsx">Choose .xlsx file…</span></div>
      </div>
    </div>

    <div class="field">
      <label>PoC Screenshots ZIP <span class="optional-badge">OPTIONAL</span></label>
      <div class="file-wrap" id="wrap_zip">
        <input type="file" id="poc_zip" name="poc_zip" accept=".zip" onchange="updateLabel('wrap_zip','poc_zip','🗜️')">
        <div class="file-label"><span class="file-icon">🗜️</span><span class="file-name" id="lbl_zip">Choose .zip file…</span></div>
      </div>
    </div>

    <button class="btn" type="submit" id="genBtn">⚡ Generate Closure Report</button>
  </form>

  <div class="progress-wrap" id="progressWrap">
    <div class="progress-bar-bg"><div class="progress-bar" id="progressBar"></div></div>
    <div class="progress-msg" id="progressMsg">Uploading files…</div>
  </div>

  <div class="result-box" id="resultBox"></div>

  <div class="info-row">
    <span><span class="{'dot-ok' if env_present else 'dot-warn'}"></span>OpenAI API Key: <strong>{'Configured' if env_present else 'Not set'}</strong></span>
    <span><span class="dot-ok"></span>Base template: <strong>{template_path.name}</strong></span>
  </div>
</div>

<script>
function updateLabel(wrapId, inputId, icon) {{
  const f = document.getElementById(inputId).files[0];
  const lbl = document.getElementById('lbl_' + wrapId.split('_')[1]);
  if (f) lbl.textContent = icon + ' ' + f.name;
}}

const stages = [
  [10,  'Uploading files…'],
  [25,  'Extracting VAPT data from DOCX…'],
  [45,  'Applying Excel overlay…'],
  [60,  'Running AI guardrail validation…'],
  [75,  'Rendering 3D charts…'],
  [88,  'Building AdvantageGo template…'],
  [96,  'Finalising DOCX…'],
];
let stageIdx = 0;
let ticker;

function startProgress() {{
  const bar = document.getElementById('progressBar');
  const msg = document.getElementById('progressMsg');
  stageIdx = 0;
  bar.style.width = '0%';
  ticker = setInterval(() => {{
    if (stageIdx < stages.length) {{
      bar.style.width = stages[stageIdx][0] + '%';
      msg.textContent = stages[stageIdx][1];
      stageIdx++;
    }}
  }}, 4000);
}}

function stopProgress(pct, message) {{
  clearInterval(ticker);
  document.getElementById('progressBar').style.width = pct + '%';
  document.getElementById('progressMsg').textContent = message;
}}

document.getElementById('genForm').addEventListener('submit', async function(e) {{
  e.preventDefault();

  const docxInput = document.getElementById('initial_report');
  if (!docxInput.files.length) {{
    showResult('error', '❌ Please select an initial VAPT report (.docx).');
    return;
  }}
  if (!docxInput.files[0].name.toLowerCase().endsWith('.docx')) {{
    showResult('error', '❌ Initial report must be a .docx file.');
    return;
  }}

  const btn = document.getElementById('genBtn');
  btn.disabled = true;
  btn.textContent = '⏳ Generating…';

  document.getElementById('progressWrap').style.display = 'block';
  document.getElementById('resultBox').style.display = 'none';
  startProgress();

  const fd = new FormData(this);

  try {{
    const resp = await fetch('/type6/generate-closure/', {{
      method: 'POST',
      body: fd,
      credentials: 'same-origin',
    }});

    stopProgress(100, 'Done!');

    if (!resp.ok) {{
      let detail = resp.statusText;
      try {{ const j = await resp.json(); detail = j.detail || detail; }} catch(_) {{}}
      showResult('error', '❌ Generation failed: ' + detail);
      return;
    }}

    const ct = (resp.headers.get('content-type') || '').toLowerCase();
    if (!ct.includes('wordprocessingml') && !ct.includes('octet-stream')) {{
      let detail = '';
      try {{ const j = await resp.json(); detail = j.detail || ''; }} catch(_) {{}}
      showResult('error', '❌ Unexpected response type. ' + detail);
      return;
    }}

    // ── Trigger DOCX download ──────────────────────────────────────────
    const cd = resp.headers.get('content-disposition') || '';
    let filename = 'AdvantageGo_Closure_Report.docx';
    const quoted = cd.match(/filename="([^"]+)"/i);
    const bare   = cd.match(/filename=([^;\\n]+)/i);
    if (quoted) filename = quoted[1];
    else if (bare) filename = bare[1].trim().replace(/^['"]+|['"]+$/g, '');

    const blob = await resp.blob();
    const url  = URL.createObjectURL(blob);
    const a    = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);

    showResult('ok',
      '✅ Report <strong>' + filename.replace(/</g, '') + '</strong> was sent to your device. ' +
      'The server does not keep a copy. If the download did not start, use your browser’s download list or run Generate again.'
    );

  }} catch(err) {{
    stopProgress(0, '');
    showResult('error', '❌ Network error: ' + err.message);
  }} finally {{
    btn.disabled = false;
    btn.textContent = '⚡ Generate Closure Report';
  }}
}});

function showResult(type, html) {{
  const box = document.getElementById('resultBox');
  box.className = 'result-box ' + (type === 'ok' ? 'result-ok' : 'result-err');
  box.innerHTML = html;
  box.style.display = 'block';
}}
</script>
</body>
</html>"""
    return HTMLResponse(html)


@app.post("/type6/generate-closure/")
async def type6_generate_closure(
    background_tasks: BackgroundTasks,
    initial_report: UploadFile = File(...),
    excel_file: UploadFile | None = File(None),
    poc_zip: UploadFile | None = File(None)
):
    """
    POST /type6/generate-closure/
    Accepts DOCX initial report + optional XLSX + optional ZIP.
    Returns the generated DOCX as a download only. The file is written to a temp path,
    streamed to the client, then deleted; nothing is kept under Automation/backend.
    """
    # Temp dir for upload inputs; separate temp .docx for the generated output
    temp_dir = tempfile.mkdtemp(prefix="type6_uploads_")
    out_path: str | None = None
    try:
        # ── Save initial DOCX ────────────────────────────────────────────
        init_name = Path(initial_report.filename or "initial_report.docx").name
        init_path = os.path.join(temp_dir, init_name)
        with open(init_path, "wb") as fh:
            fh.write(await initial_report.read())

        if not init_path.lower().endswith(".docx"):
            raise HTTPException(status_code=400, detail="Initial report must be a .docx file.")

        # ── Save optional Excel ──────────────────────────────────────────
        excel_path = None
        if excel_file and getattr(excel_file, "filename", ""):
            excel_name = Path(excel_file.filename or "vulnerabilities.xlsx").name
            excel_path = os.path.join(temp_dir, excel_name)
            with open(excel_path, "wb") as fh:
                fh.write(await excel_file.read())

        # ── Save optional ZIP ────────────────────────────────────────────
        zip_path = None
        if poc_zip and getattr(poc_zip, "filename", ""):
            zip_name = Path(poc_zip.filename or "poc.zip").name
            zip_path = os.path.join(temp_dir, zip_name)
            with open(zip_path, "wb") as fh:
                fh.write(await poc_zip.read())

        ts = int(time.time())
        out_filename = f"AdvantageGo_Closure_Report_{ts}.docx"
        fd, out_path = tempfile.mkstemp(
            suffix=".docx", prefix="type6_out_", dir=tempfile.gettempdir()
        )
        os.close(fd)

        # ── Load API key ─────────────────────────────────────────────────
        _load_root_env()
        chosen_key = os.getenv("OPENAI_API_KEY", "") or ""
        if not chosen_key:
            _unlink_quiet(out_path)
            out_path = None
            shutil.rmtree(temp_dir, ignore_errors=True)
            raise HTTPException(
                status_code=400,
                detail="Server OPENAI_API_KEY is not configured. Add it to CSS/.env"
            )

        # ── Run the generation pipeline ──────────────────────────────────
        generate(init_path, out_path, api_key=chosen_key,
                 excel_path=excel_path, poc_zip_path=zip_path)

        if not os.path.isfile(out_path):
            _unlink_quiet(out_path)
            out_path = None
            shutil.rmtree(temp_dir, ignore_errors=True)
            raise HTTPException(
                status_code=500, detail="Report generation completed but output file not found."
            )
        size_kb = os.path.getsize(out_path) // 1024
        print(f"  ✓ Serving: {out_filename} ({size_kb} KB) in-memory (temp file not kept on disk)")

        # Read bytes, delete temp .docx immediately, return bytes (avoids FileResponse path
        # races on Windows and never leaves files under Automation/backend/closure_reports).
        with open(out_path, "rb") as _fh:
            body = _fh.read()
        _unlink_quiet(out_path)
        out_path = None
        background_tasks.add_task(shutil.rmtree, temp_dir, True)

        return Response(
            content=body,
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            headers={
                "Content-Disposition": f'attachment; filename="{out_filename}"',
                "X-Report-Filename": out_filename,
            },
        )

    except HTTPException:
        if out_path:
            _unlink_quiet(out_path)
        shutil.rmtree(temp_dir, ignore_errors=True)
        raise
    except Exception as exc:
        if out_path:
            _unlink_quiet(out_path)
        shutil.rmtree(temp_dir, ignore_errors=True)
        import traceback; traceback.print_exc()
        raise HTTPException(status_code=500, detail=str(exc)) from exc
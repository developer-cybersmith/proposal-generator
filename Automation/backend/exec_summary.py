"""
Executive Summary report generator — FastAPI sub-app.

Uploads MitKat-style Excel + form metadata; builds canonical JSON; calls OpenAI once;
assembles a Word document for C-suite readers.
"""

from __future__ import annotations

import logging
import io
import itertools
import json
import os
import re
import shutil
import tempfile
import time
import traceback
import zipfile
from pathlib import Path
from typing import Any

from dotenv import load_dotenv

# Load repo root .env first, then Automation/backend/.env (override) so keys work either place.
_EXEC_SUMMARY_DIR = Path(__file__).resolve().parent
_REPO_ROOT = _EXEC_SUMMARY_DIR.parent.parent
load_dotenv(_REPO_ROOT / ".env", override=False)
load_dotenv(_EXEC_SUMMARY_DIR / ".env", override=True)

from fastapi import FastAPI, File, Form, HTTPException, Request, UploadFile
from fastapi.responses import FileResponse, HTMLResponse
from openai import OpenAI
from openpyxl import load_workbook
from sqlalchemy.orm import Session
from starlette.background import BackgroundTask

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT, WD_TAB_LEADER
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor

from db import AuditLog, SessionLocal

from .exec_summary_prompts import (
    SYSTEM_PROMPT,
    build_canonical_user_payload,
    build_user_message,
    parse_exec_summary_json,
    validate_no_foreign_identifiers,
)

logger = logging.getLogger(__name__)

app = FastAPI()

VALID_REPORT_TYPES = frozenset({"ip", "web", "certin", "mitkat"})
RATE_LIMIT_WINDOW_SECONDS = 600
RATE_LIMIT_MAX_REQUESTS = 5
_REQUEST_TIMESTAMPS_BY_USER: dict[str, list[float]] = {}

REQUIRED_HEADERS = [
    "Asset/Hostname",
    "Instant purpose",
    "VAPT Status",
    "Critical",
    "High",
    "Medium",
    "Low",
    "Informational",
    "Total",
    "Tester_Name",
    "Project",
    "Sr no.",
    "Observation",
    "Severity",
    "Status",
    "New or Repeat Observation",
    "CVE/CWE",
    "CVSS",
    "CVSS Version Ref",
    "Affected Asset i.e. IP/URL/Application etc.",
    "Observation/Vulnerability Title",
    "Detailed Observation/ Vulnerable Point",
    "Recommendation",
    "Reference",
    "Evidence / Proof of Concept",
]

SEVERITY_ORDER = {
    "critical": 0,
    "high": 1,
    "medium": 2,
    "low": 3,
    "informational": 4,
    "info": 4,
    "unknown": 5,
}

DISCLAIMER_PARAGRAPH_1 = (
    "All information contained in this document is confidential and proprietary to "
    "CyberSmithSECURE and Reproduction, disclosure or use of any information contained in this "
    "document by photographic, electronic or any other means, in whole or part, for any reason "
    "other than for the purpose of operations is strictly prohibited without written consent."
)
DISCLAIMER_PARAGRAPH_2 = (
    "This report is intended to provide a high level overview of the current ongoing VAPT Assessment; "
    "this report describes the latest findings identified by the team CyberSmithSECURE. This report is "
    "an executive basis report; the indicators mentioned in this report will be mentioned separately "
    "in the later stage technical report as well."
)

# Accent colors for Heading 2 subsections (run-level; paragraph style stays Heading 2)
_CSS_GREEN = RGBColor(0x00, 0x83, 0x3B)
_CSS_RED = RGBColor(0x60, 0x00, 0x00)
_CSS_BLUE = RGBColor(0x00, 0x37, 0x60)
_CSS_PURPLE = RGBColor(0x70, 0x30, 0xA0)

THANK_YOU_LINE = (
    "We thank the internal Information Security team for their support and cooperation during the time of assessment."
)

# Branded DOCX template (built by build_exec_template.py)
EXEC_SUMMARY_TEMPLATE = _EXEC_SUMMARY_DIR / "exec_summary_template.docx"

# Severity column header colors (match CSS Executive Summary template)
_SEV_COL_FILLS = ["8EA9DB", "8EA9DB", "8EA9DB", "8EA9DB", "C00000", "FF0000", "FFC000", "00B050", "0070C0", "7030A0"]
def _safe_style(doc: Document, name: str):
    try:
        return doc.styles[name]
    except KeyError:
        return doc.styles["Normal"]


def _iter_paragraphs_everywhere(doc: Document):
    """Body, tables, headers, footers — for placeholder replacement."""
    for p in doc.paragraphs:
        yield p
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    yield p
    for section in doc.sections:
        for p in section.header.paragraphs:
            yield p
        for p in section.footer.paragraphs:
            yield p


def _replace_placeholders(doc: Document, mapping: dict[str, str]) -> None:
    for p in _iter_paragraphs_everywhere(doc):
        t = p.text
        if not any(k in t for k in mapping):
            continue
        new_t = t
        for k, v in mapping.items():
            new_t = new_t.replace(k, v)
        for r in p.runs:
            r.text = ""
        if p.runs:
            p.runs[0].text = new_t
        else:
            p.add_run(new_t)


def _remove_content_start_marker(doc: Document) -> None:
    for p in list(doc.paragraphs):
        if "{{CONTENT_START}}" in p.text:
            p._element.getparent().remove(p._element)
            return


def _set_cell_shading(cell, fill_hex: str) -> None:
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    for el in list(tcPr):
        if el.tag == qn("w:shd"):
            tcPr.remove(el)
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), fill_hex.replace("#", ""))
    shd.set(qn("w:val"), "clear")
    tcPr.append(shd)


def _set_cell_text_color(cell, rgb: RGBColor) -> None:
    for p in cell.paragraphs:
        for r in p.runs:
            r.font.color.rgb = rgb


def _style_objectives_table(tbl) -> None:
    """Objectives / Status table: gray header, alternating rows."""
    for ci in range(2):
        _set_cell_shading(tbl.rows[0].cells[ci], "D9D9D9")
        for p in tbl.rows[0].cells[ci].paragraphs:
            for r in p.runs:
                r.bold = True
    for r in range(1, len(tbl.rows)):
        fill = "FFFFFF" if r % 2 == 1 else "F2F2F2"
        for ci in range(2):
            _set_cell_shading(tbl.rows[r].cells[ci], fill)


def _style_risk_table_header_and_data(tbl) -> None:
    """Risk table: blue label columns, severity-colored columns, light blue data."""
    nrows = len(tbl.rows)
    for ci, fill in enumerate(_SEV_COL_FILLS):
        cell = tbl.rows[0].cells[ci]
        _set_cell_shading(cell, fill)
        if ci >= 4:
            _set_cell_text_color(cell, RGBColor(0xFF, 0xFF, 0xFF))
    for ri in range(1, nrows):
        is_reval = ri == nrows - 1
        is_overall = ri == nrows - 2
        for ci in range(10):
            if is_reval:
                _set_cell_shading(tbl.rows[ri].cells[ci], "BDD7EE")
                continue
            if ci < 4:
                _set_cell_shading(tbl.rows[ri].cells[ci], "BDD7EE" if not is_overall else "D9E1F2")
            else:
                _set_cell_shading(tbl.rows[ri].cells[ci], _SEV_COL_FILLS[ci])
                _set_cell_text_color(tbl.rows[ri].cells[ci], RGBColor(0xFF, 0xFF, 0xFF))


def _apply_risk_table_bold(tbl) -> None:
    """Bold header row and numeric columns (4–9) for all data rows."""
    for cell in tbl.rows[0].cells:
        for p in cell.paragraphs:
            for r in p.runs:
                r.bold = True
    for ri in range(1, len(tbl.rows)):
        for ci in range(4, 10):
            for p in tbl.rows[ri].cells[ci].paragraphs:
                for r in p.runs:
                    r.bold = True
    for ci in range(4):
        for p in tbl.rows[-2].cells[ci].paragraphs:
            for r in p.runs:
                r.bold = True
    for ci in range(4):
        for p in tbl.rows[-1].cells[ci].paragraphs:
            for r in p.runs:
                r.bold = True


def _apply_branded_heading_styles(doc: Document) -> None:
    """Visually match CSS template on built-in Heading 1 / Heading 2 (TOC uses these styles)."""
    try:
        h1 = doc.styles["Heading 1"]
        h1.font.name = "Calibri"
        h1.font.size = Pt(24)
        h1.font.bold = True
        h1.font.color.rgb = _CSS_PURPLE
    except KeyError:
        pass
    try:
        h2 = doc.styles["Heading 2"]
        h2.font.name = "Calibri"
        h2.font.size = Pt(18)
        h2.font.bold = True
        h2.font.color.rgb = _CSS_PURPLE
    except KeyError:
        pass


# Unique bookmark ids for PAGEREF / TOC (avoid clashing with template bookmarks)
_BOOKMARK_ID_SEQ = itertools.count(50_000)


def _next_bookmark_id() -> int:
    return next(_BOOKMARK_ID_SEQ)


def _add_bookmark_to_paragraph(paragraph, name: str, bid: int) -> None:
    """Wrap paragraph with w:bookmarkStart/End so PAGEREF can resolve page numbers."""
    p = paragraph._element
    p_pr = p.find(qn("w:pPr"))
    start = OxmlElement("w:bookmarkStart")
    start.set(qn("w:id"), str(bid))
    start.set(qn("w:name"), name)
    end = OxmlElement("w:bookmarkEnd")
    end.set(qn("w:id"), str(bid))
    if p_pr is not None:
        p_pr.addnext(start)
    else:
        p.insert(0, start)
    p.append(end)


def _append_pageref_field(paragraph, bookmark_name: str) -> None:
    """PAGE ref for a bookmark; separate runs (same rule as TOC/PAGE fields)."""
    r1 = paragraph.add_run()
    fld_begin = OxmlElement("w:fldChar")
    fld_begin.set(qn("w:fldCharType"), "begin")
    r1._r.append(fld_begin)

    r2 = paragraph.add_run()
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = f" PAGEREF {bookmark_name} \\h "
    r2._r.append(instr)

    r3 = paragraph.add_run()
    fld_sep = OxmlElement("w:fldChar")
    fld_sep.set(qn("w:fldCharType"), "separate")
    r3._r.append(fld_sep)

    # Empty field result: a literal "1" here makes every PAGEREF show page 1 until Word recalculates.
    r4 = paragraph.add_run()
    t = OxmlElement("w:t")
    t.set(qn("xml:space"), "preserve")
    t.text = ""
    r4._r.append(t)

    r5 = paragraph.add_run()
    fld_end = OxmlElement("w:fldChar")
    fld_end.set(qn("w:fldCharType"), "end")
    r5._r.append(fld_end)


def _add_manual_toc(doc: Document, client: str) -> None:
    """Always-visible TOC lines (titles are plain text). Page numbers use PAGEREF + bookmarks on headings.

    The old single ``TOC \\\\o`` field often stays empty in Protected View or until Word updates fields.
    """
    entries: list[tuple[str, bool, str]] = [
        ("Executive Summary of The Assessment", False, "bm_es"),
        ("Summary", False, "bm_sum"),
        (f"{client} Risk Assessment Analysis of The Entire Framework", False, "bm_risk"),
        ("Overall Observation", False, "bm_obs"),
        ("Phase 1", True, "bm_p1"),
        ("Phase 2", True, "bm_p2"),
        ("Overall Recommendation", False, "bm_rec"),
        ("Best Practises", True, "bm_best"),
        ("Must Have", True, "bm_must"),
        ("Should Have", True, "bm_should"),
        ("SUMMARY OF FINDINGS & CONCLUSION", False, "bm_concl"),
    ]
    for label, is_sub, bm in entries:
        p = doc.add_paragraph()
        p.style = _safe_style(doc, "CSSBody")
        if is_sub:
            p.paragraph_format.left_indent = Inches(0.35)
        p.paragraph_format.tab_stops.add_tab_stop(
            Inches(6.2), WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS
        )
        p.add_run(label)
        p.add_run("\t")
        _append_pageref_field(p, bm)


def _set_paragraph_outline_level(paragraph, level: int | None) -> None:
    """Map headings into Word's TOC (TOC \\o \"1-3\"). OOXML uses 0-based outlineLvl (0=level1, 1=level2)."""
    pPr = paragraph._element.get_or_add_pPr()
    for el in list(pPr):
        if el.tag == qn("w:outlineLvl"):
            pPr.remove(el)
    if level is None:
        return
    outline = OxmlElement("w:outlineLvl")
    outline.set(qn("w:val"), str(int(level)))
    pPr.append(outline)


def _heading(
    doc: Document,
    text: str,
    level: int = 1,
    *,
    run_color: RGBColor | None = None,
    include_in_toc: bool = True,
    alignment: WD_ALIGN_PARAGRAPH | None = None,
    bookmark: str | None = None,
) -> None:
    """Section titles use Word built-in Heading 1 / Heading 2; optional bookmark for manual TOC PAGEREF."""
    style_key = "Heading 1" if int(level) == 1 else "Heading 2"
    p = doc.add_paragraph()
    p.style = _safe_style(doc, style_key)
    run = p.add_run(text)
    if run_color is not None:
        run.font.color.rgb = run_color
    if alignment is not None:
        p.alignment = alignment
    if include_in_toc:
        _set_paragraph_outline_level(p, 0 if int(level) == 1 else 1)
    else:
        _set_paragraph_outline_level(p, None)
    if bookmark is not None:
        _add_bookmark_to_paragraph(p, bookmark, _next_bookmark_id())


def _get_checkmark_png() -> bytes:
    """Small green checkmark PNG for objectives table (generated; no asset file required)."""
    import matplotlib

    matplotlib.use("Agg")
    import matplotlib.pyplot as plt

    fig, ax = plt.subplots(figsize=(0.16, 0.16), dpi=100)
    ax.plot([0.05, 0.42, 0.92], [0.38, 0.08, 0.88], color="#00B050", linewidth=2.2, solid_capstyle="round")
    ax.set_xlim(0, 1)
    ax.set_ylim(0, 1)
    ax.axis("off")
    fig.subplots_adjust(0, 0, 1, 1)
    buf = io.BytesIO()
    fig.savefig(buf, format="png", bbox_inches="tight", pad_inches=0, transparent=True)
    plt.close(fig)
    buf.seek(0)
    return buf.getvalue()


def _fill_status_cell_with_checkmark(doc: Document, cell, status_text: str) -> None:
    """Inline checkmark image + status text (objectives table)."""
    cell.text = ""
    p = cell.paragraphs[0]
    p.style = _safe_style(doc, "CSSBody")
    pic = p.add_run()
    pic.add_picture(io.BytesIO(_get_checkmark_png()), width=Pt(10.5))
    p.add_run(f" {status_text.strip()}")


def _body_paragraph(doc: Document, text: str, bold: bool = False) -> None:
    name = "CSSBodyBold" if bold else "CSSBody"
    p = doc.add_paragraph()
    p.style = _safe_style(doc, name)
    p.add_run(text)


def _normalize_header(value: Any) -> str:
    return str(value or "").strip()


def _get_openai_client() -> OpenAI:
    """Prefer EXEC_SUMMARY_OPENAI_* keys; fall back to OPENAI_API_KEY if unset (same file as AI Layer)."""
    api_key = (
        os.getenv("EXEC_SUMMARY_OPENAI_API_KEY")
        or os.getenv("OPENAI_EXEC_SUMMARY_API_KEY")
        or os.getenv("OPENAI_API_KEY")
        or ""
    ).strip()
    if not api_key:
        raise HTTPException(
            status_code=500,
            detail=(
                "No OpenAI API key found. Set EXEC_SUMMARY_OPENAI_API_KEY in Automation/backend/.env "
                "(dedicated key), or ensure OPENAI_API_KEY is set there for shared use."
            ),
        )
    kwargs: dict[str, str] = {"api_key": api_key}
    org = (
        os.getenv("EXEC_SUMMARY_OPENAI_ORGANIZATION")
        or os.getenv("EXEC_SUMMARY_OPENAI_ORG_ID")
        or os.getenv("OPENAI_ORGANIZATION")
        or os.getenv("OPENAI_ORG_ID")
        or ""
    ).strip()
    if org:
        kwargs["organization"] = org
    project = (
        os.getenv("EXEC_SUMMARY_OPENAI_PROJECT")
        or os.getenv("EXEC_SUMMARY_OPENAI_PROJECT_ID")
        or os.getenv("OPENAI_PROJECT")
        or os.getenv("OPENAI_PROJECT_ID")
        or ""
    ).strip()
    if project:
        kwargs["project"] = project
    return OpenAI(**kwargs)


def _is_insufficient_quota(exc: BaseException) -> bool:
    text = str(exc).lower()
    if "insufficient_quota" in text:
        return True
    body = getattr(exc, "body", None)
    if isinstance(body, dict):
        err = body.get("error") or {}
        if err.get("code") == "insufficient_quota":
            return True
    return False


def _read_html_template() -> str:
    html_path = Path(__file__).resolve().parent.parent / "exec_summary.html"
    with open(html_path, encoding="utf-8") as handle:
        return handle.read()


def _enforce_rate_limit(request: Request) -> None:
    user = request.session.get("user") or {}
    user_key = (user.get("email") or (request.client.host if request.client else None) or "anonymous")
    now = time.time()
    window_start = now - RATE_LIMIT_WINDOW_SECONDS
    history = _REQUEST_TIMESTAMPS_BY_USER.setdefault(user_key, [])
    history[:] = [ts for ts in history if ts >= window_start]
    if len(history) >= RATE_LIMIT_MAX_REQUESTS:
        raise HTTPException(
            status_code=429,
            detail=(
                f"Rate limit exceeded. Max {RATE_LIMIT_MAX_REQUESTS} executive summary generations per "
                f"{RATE_LIMIT_WINDOW_SECONDS // 60} minutes."
            ),
        )
    history.append(now)


def _log_audit(request: Request, metadata: dict[str, Any]) -> None:
    db: Session = SessionLocal()
    try:
        user = request.session.get("user") or {}
        db.add(
            AuditLog(
                user_email=user.get("email"),
                user_name=user.get("name"),
                action="exec-summary-generate",
                metadata_json=json.dumps(metadata),
                ip_address=request.client.host if request.client else None,
                user_agent=request.headers.get("user-agent"),
            )
        )
        db.commit()
    except Exception:
        db.rollback()
    finally:
        db.close()


def _norm_header_key(value: Any) -> str:
    """Normalise a column title for matching (case, spacing)."""
    s = _normalize_header(value)
    return " ".join(s.lower().replace("_", " ").replace("/", " ").split())


def _build_header_map(worksheet) -> dict[str, int]:
    """Map canonical REQUIRED_HEADERS names to 1-based column indexes.

    Row 1 may use slightly different spacing or casing than REQUIRED_HEADERS; we match loosely.
    """
    col_by_norm: dict[str, int] = {}
    for column_index, cell in enumerate(worksheet[1], start=1):
        raw = _normalize_header(cell.value)
        if not raw:
            continue
        col_by_norm[_norm_header_key(raw)] = column_index

    header_map: dict[str, int] = {}
    missing: list[str] = []
    for h in REQUIRED_HEADERS:
        nk = _norm_header_key(h)
        if nk in col_by_norm:
            header_map[h] = col_by_norm[nk]
        else:
            missing.append(h)
    if missing:
        raise HTTPException(
            status_code=400,
            detail=f"Excel is missing required columns: {', '.join(missing)}",
        )
    return header_map


def _cell(worksheet, row: int, header_map: dict[str, int], header: str) -> str:
    col = header_map.get(header)
    if col is None:
        return ""
    return str(worksheet.cell(row=row, column=col).value or "").strip()


def _parse_int_cell(val: Any) -> int:
    if val is None or val == "":
        return 0
    try:
        return int(float(str(val).strip()))
    except (TypeError, ValueError):
        return 0


def _sheet_has_vuln_headers(ws) -> bool:
    """True if row 1 has all REQUIRED_HEADERS (after normalisation)."""
    try:
        _build_header_map(ws)
        return True
    except HTTPException:
        return False


def _pick_worksheet(wb, report_type: str):
    """Prefer named sheets (IP / Web / Cert-IN); otherwise use the first sheet that has all vuln columns."""
    names = wb.sheetnames
    rt = report_type.lower()
    if rt == "ip" and "IP" in names:
        cand = wb["IP"]
        if _sheet_has_vuln_headers(cand):
            return cand
    if rt == "certin" and "Cert-IN" in names:
        cand = wb["Cert-IN"]
        if _sheet_has_vuln_headers(cand):
            return cand
    if rt == "web" and "Web" in names:
        cand = wb["Web"]
        if _sheet_has_vuln_headers(cand):
            return cand
    if rt == "mitkat" and "MitKat" in names:
        cand = wb["MitKat"]
        if _sheet_has_vuln_headers(cand):
            return cand
    # Scan all sheets for the standard 29-column vulnerability layout
    for name in names:
        ws = wb[name]
        if _sheet_has_vuln_headers(ws):
            return ws
    # Last resort: first sheet (may raise missing columns in _build_header_map)
    return wb[names[0]]


def _norm_severity(raw: str) -> str:
    s = (raw or "").strip().lower()
    if s in ("info", "informational"):
        return "informational"
    if s not in SEVERITY_ORDER:
        return "medium"
    return s


def _parse_excel_to_payload(
    excel_path: str,
    report_type: str,
    metadata: dict[str, Any],
) -> tuple[list[dict[str, Any]], dict[str, int], list[dict[str, Any]], str]:
    wb = load_workbook(excel_path, data_only=True)
    try:
        ws = _pick_worksheet(wb, report_type)
        header_map = _build_header_map(ws)
        assets_map: dict[str, dict[str, Any]] = {}
        vulnerabilities: list[dict[str, Any]] = []
        cve_parts: list[str] = []

        for row_idx in range(2, ws.max_row + 1):
            host = _cell(ws, row_idx, header_map, "Asset/Hostname")
            title = _cell(ws, row_idx, header_map, "Observation/Vulnerability Title")
            if title:
                sev = _norm_severity(_cell(ws, row_idx, header_map, "Severity"))
                rec = _cell(ws, row_idx, header_map, "Recommendation")[:200]
                asset_affected = _cell(ws, row_idx, header_map, "Affected Asset i.e. IP/URL/Application etc.")
                cve = _cell(ws, row_idx, header_map, "CVE/CWE")
                if cve:
                    cve_parts.append(cve)
                vulnerabilities.append(
                    {
                        "title": title,
                        "severity": sev,
                        "affected_asset": asset_affected or host,
                        "recommendation_summary": rec,
                    }
                )

            if host:
                key = host.strip().lower()
                row_asset = {
                    "hostname": _cell(ws, row_idx, header_map, "Asset/Hostname"),
                    "purpose": _cell(ws, row_idx, header_map, "Instant purpose"),
                    "vapt_status": _cell(ws, row_idx, header_map, "VAPT Status"),
                    "critical": _parse_int_cell(ws.cell(row=row_idx, column=header_map["Critical"]).value),
                    "high": _parse_int_cell(ws.cell(row=row_idx, column=header_map["High"]).value),
                    "medium": _parse_int_cell(ws.cell(row=row_idx, column=header_map["Medium"]).value),
                    "low": _parse_int_cell(ws.cell(row=row_idx, column=header_map["Low"]).value),
                    "informational": _parse_int_cell(ws.cell(row=row_idx, column=header_map["Informational"]).value),
                    "total": _parse_int_cell(ws.cell(row=row_idx, column=header_map["Total"]).value),
                }
                if key not in assets_map:
                    assets_map[key] = row_asset
                else:
                    prev = assets_map[key]
                    for f in ("critical", "high", "medium", "low", "informational", "total"):
                        prev[f] = max(prev[f], row_asset[f])

        assets = sorted(assets_map.values(), key=lambda x: x["hostname"].lower())
        total_counts = {"critical": 0, "high": 0, "medium": 0, "low": 0, "informational": 0, "total": 0}
        for v in vulnerabilities:
            k = v["severity"]
            if k in total_counts:
                total_counts[k] += 1
        total_counts["total"] = len(vulnerabilities)

        vulns_sorted = sorted(
            vulnerabilities,
            key=lambda x: SEVERITY_ORDER.get(x["severity"], 99),
        )
        top_v = vulns_sorted[:20]

        allowed_cve = " ".join(cve_parts)
        if not vulnerabilities:
            raise HTTPException(
                status_code=400,
                detail="No vulnerability rows found (Observation/Vulnerability Title empty for all rows).",
            )

        if not assets:
            assets = [
                {
                    "hostname": "— (no asset rows in sheet)",
                    "purpose": "",
                    "vapt_status": "Completed",
                    "critical": total_counts["critical"],
                    "high": total_counts["high"],
                    "medium": total_counts["medium"],
                    "low": total_counts["low"],
                    "informational": total_counts["informational"],
                    "total": total_counts["total"],
                }
            ]

        return assets, total_counts, top_v, allowed_cve
    finally:
        wb.close()


def _call_openai_exec_summary(user_payload_str: str) -> dict[str, Any]:
    client = _get_openai_client()
    model = (
        os.getenv("EXEC_SUMMARY_OPENAI_MODEL") or os.getenv("OPENAI_MODEL") or "gpt-4o"
    ).strip() or "gpt-4o"
    response = client.chat.completions.create(
        model=model,
        temperature=0.25,
        max_tokens=4000,
        response_format={"type": "json_object"},
        messages=[
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": build_user_message({}, user_payload_str)},
        ],
    )
    content = response.choices[0].message.content or ""
    parsed = parse_exec_summary_json(content)
    return parsed


def _generate_with_retries(user_payload_str: str, allowed_cve: str, allowed_hosts: set[str], retries: int = 2) -> dict[str, Any]:
    last_err: Exception | None = None
    for attempt in range(retries + 1):
        try:
            parsed = _call_openai_exec_summary(user_payload_str)
            if allowed_cve.strip():
                validate_no_foreign_identifiers(parsed, allowed_cve, allowed_hosts)
            return parsed
        except Exception as exc:
            if _is_insufficient_quota(exc):
                raise HTTPException(
                    status_code=503,
                    detail="OpenAI insufficient_quota. Check billing and OPENAI_API_KEY.",
                ) from exc
            last_err = exc
            if attempt < retries:
                time.sleep(1.0)
    raise HTTPException(status_code=502, detail=f"AI generation failed: {last_err}")


def _severity_chart_png(total_counts: dict[str, int]) -> bytes:
    import matplotlib

    matplotlib.use("Agg")
    import matplotlib.pyplot as plt

    labels = ["Critical", "High", "Medium", "Low", "Info"]
    keys = ["critical", "high", "medium", "low", "informational"]
    vals = [total_counts.get(k, 0) for k in keys]
    colors = ["#C00000", "#FF0000", "#FFC000", "#00B050", "#0070C0"]
    fig, ax = plt.subplots(figsize=(6.2, 3.6))
    ax.bar(labels, vals, color=colors, edgecolor="#333333", linewidth=0.5)
    ax.set_ylabel("Count")
    ax.set_title("Vulnerability count by severity")
    fig.tight_layout()
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=120)
    plt.close(fig)
    return buf.getvalue()


def _severity_pie_chart_png(total_counts: dict[str, int]) -> bytes:
    """Donut chart with CSS severity colors and percentage labels."""
    import matplotlib

    matplotlib.use("Agg")
    import matplotlib.pyplot as plt

    labels = ["Critical", "High", "Medium", "Low", "Info"]
    keys = ["critical", "high", "medium", "low", "informational"]
    vals = [total_counts.get(k, 0) for k in keys]
    colors = ["#C00000", "#FF0000", "#FFC000", "#00B050", "#0070C0"]
    total = sum(vals)
    if total == 0:
        fig, ax = plt.subplots(figsize=(4.2, 4.2))
        ax.text(0.5, 0.5, "No vulnerabilities", ha="center", va="center", fontsize=12)
        ax.axis("off")
        buf = io.BytesIO()
        fig.savefig(buf, format="png", dpi=120)
        plt.close(fig)
        return buf.getvalue()
    fig, ax = plt.subplots(figsize=(4.2, 4.2))
    ax.pie(
        vals,
        labels=labels,
        colors=colors,
        autopct=lambda pct: f"{pct:.1f}%",
        startangle=90,
        wedgeprops=dict(width=0.45, edgecolor="white"),
        textprops={"fontsize": 8},
    )
    ax.set_title("Severity distribution")
    fig.tight_layout()
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=120)
    plt.close(fig)
    return buf.getvalue()


def _build_document(
    metadata: dict[str, Any],
    assets: list[dict[str, Any]],
    total_counts: dict[str, int],
    ai: dict[str, Any],
) -> Document:
    org = metadata.get("org_name") or "CyberSmithSECURE"
    client = metadata.get("client_name") or "Client"
    assessment = metadata.get("assessment_type") or "VAPT Assessment"
    rdate = metadata.get("report_date") or ""
    doc_id = (metadata.get("document_id") or "").strip()

    if EXEC_SUMMARY_TEMPLATE.is_file():
        doc = Document(str(EXEC_SUMMARY_TEMPLATE))
        _replace_placeholders(
            doc,
            {
                "{{ORG_NAME}}": org,
                "{{ASSESSMENT_TYPE}}": assessment,
                "{{CLIENT_NAME}}": client,
                "{{REPORT_DATE}}": rdate,
                "{{DOCUMENT_ID}}": doc_id,
            },
        )
        _remove_content_start_marker(doc)
    else:
        doc = Document()
        sec = doc.sections[0]
        sec.left_margin = Inches(1)
        sec.right_margin = Inches(1)
        p = doc.add_paragraph()
        p.add_run(f"{org} — {assessment} — For {client} — {rdate}")

    _apply_branded_heading_styles(doc)

    # Disclaimer: not a heading; two body paragraphs (no embedded line breaks)
    p_dis = doc.add_paragraph()
    p_dis.style = _safe_style(doc, "CSSBodyBold")
    p_dis.add_run("Disclaimer").bold = True
    p_d1 = doc.add_paragraph()
    p_d1.style = _safe_style(doc, "CSSBody")
    p_d1.paragraph_format.space_after = Pt(10)
    p_d1.add_run(DISCLAIMER_PARAGRAPH_1)
    p_d2 = doc.add_paragraph()
    p_d2.style = _safe_style(doc, "CSSBody")
    p_d2.add_run(DISCLAIMER_PARAGRAPH_2)

    doc.add_page_break()
    p_toc = doc.add_paragraph()
    p_toc.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_toc.style = _safe_style(doc, "CSSTOCTitle")
    p_toc.add_run("Table of Content")
    _add_manual_toc(doc, client)

    doc.add_page_break()

    _heading(doc, "Executive Summary of The Assessment", level=1, bookmark="bm_es")
    for block in ai["executive_narrative"].split("\n\n"):
        block = block.strip()
        if block:
            _body_paragraph(doc, block)

    doc.add_page_break()

    _heading(doc, "Summary", level=1, bookmark="bm_sum")
    tbl = doc.add_table(rows=5, cols=2)
    tbl.style = "Table Grid"
    hdr = ("Objectives", "Status")
    for i, h in enumerate(hdr):
        tbl.rows[0].cells[i].text = h
    objectives = [
        "Identify Critical Vulnerability Inside Web Application",
        "Identify Methods for Bypassing Present Security Controls",
        "Detect Vulnerabilities causing Data-Theft or Intrusion",
        "Reporting Vulnerabilities with Recommendation and Proof of Concept",
    ]
    for r, obj in enumerate(objectives, start=1):
        tbl.rows[r].cells[0].text = obj
        _fill_status_cell_with_checkmark(doc, tbl.rows[r].cells[1], "Completed")
    _style_objectives_table(tbl)

    doc.add_paragraph()
    p_bar = doc.add_paragraph()
    p_bar.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_bar.add_run().add_picture(io.BytesIO(_severity_chart_png(total_counts)), width=Inches(5.8))
    doc.add_paragraph()
    p_pie = doc.add_paragraph()
    p_pie.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_pie.add_run().add_picture(io.BytesIO(_severity_pie_chart_png(total_counts)), width=Inches(4.0))

    doc.add_page_break()

    _heading(
        doc,
        f"{client} Risk Assessment Analysis of The Entire Framework",
        level=1,
        alignment=WD_ALIGN_PARAGRAPH.CENTER,
        bookmark="bm_risk",
    )

    ncols = 10
    n_data = len(assets)
    t2 = doc.add_table(rows=3 + n_data, cols=ncols)
    t2.style = "Table Grid"
    headers = [
        "Sr. no.",
        "Hostname",
        "Instant purpose",
        "VAPT Status",
        "Critical",
        "High",
        "Medium",
        "Low",
        "Informational",
        "Total",
    ]
    for i, h in enumerate(headers):
        t2.rows[0].cells[i].text = h
    for idx, a in enumerate(assets, start=1):
        row = t2.rows[idx]
        vals = [
            str(idx),
            a["hostname"],
            a["purpose"],
            a["vapt_status"],
            str(a["critical"]),
            str(a["high"]),
            str(a["medium"]),
            str(a["low"]),
            str(a["informational"]),
            str(a["total"]),
        ]
        for i, v in enumerate(vals):
            row.cells[i].text = v
    sum_row = t2.rows[1 + n_data]
    sum_row.cells[0].text = "Overall Findings"
    sum_row.cells[1].text = ""
    sum_row.cells[2].text = ""
    sum_row.cells[3].text = ""
    sum_row.cells[4].text = str(sum(a["critical"] for a in assets))
    sum_row.cells[5].text = str(sum(a["high"] for a in assets))
    sum_row.cells[6].text = str(sum(a["medium"] for a in assets))
    sum_row.cells[7].text = str(sum(a["low"] for a in assets))
    sum_row.cells[8].text = str(sum(a["informational"] for a in assets))
    sum_row.cells[9].text = str(sum(a["total"] for a in assets))

    rev_row = t2.rows[2 + n_data]
    rev_row.cells[0].merge(rev_row.cells[3])
    rev_row.cells[4].merge(rev_row.cells[9])
    rev_row.cells[0].text = "Overall Status of Revalidation"
    for p in rev_row.cells[0].paragraphs:
        for r in p.runs:
            r.bold = True
    rev_row.cells[4].text = (
        "Refer to the technical report for per-finding revalidation status. "
        "Executive summary reflects assessment totals at generation time."
    )

    _style_risk_table_header_and_data(t2)
    _apply_risk_table_bold(t2)

    doc.add_page_break()

    _heading(doc, "Overall Observation", level=1, bookmark="bm_obs")
    _heading(doc, "Phase 1", level=2, bookmark="bm_p1")
    _body_paragraph(
        doc,
        "Prior to remediation, the following themes were recorded during the assessment based on the findings provided.",
    )
    for i, obs in enumerate(ai["phase1_observations"], start=1):
        _body_paragraph(doc, f"{i}. {obs['title']}", bold=True)
        _body_paragraph(doc, obs["description"])
        _body_paragraph(doc, f"Threat Perception: {obs['threat_perception']}", bold=True)

    _heading(doc, "Phase 2", level=2, bookmark="bm_p2")
    _body_paragraph(doc, ai["phase2_summary"])

    doc.add_page_break()

    _heading(doc, "Overall Recommendation", level=1, bookmark="bm_rec")
    _body_paragraph(doc, ai["overall_recommendation_intro"])

    def write_reco_section(title: str, accent: RGBColor, items: list[dict[str, str]], bm: str) -> None:
        _heading(doc, title, level=2, run_color=accent, bookmark=bm)
        for idx, it in enumerate(items, start=1):
            p_it = doc.add_paragraph()
            p_it.style = _safe_style(doc, "CSSBody")
            p_it.paragraph_format.left_indent = Inches(0.5)
            p_it.paragraph_format.first_line_indent = Inches(-0.25)
            p_it.add_run(f"{idx}. ").bold = True
            p_it.add_run(it["title"]).bold = True
            _body_paragraph(doc, it["description"])
            p_req = doc.add_paragraph()
            p_req.style = _safe_style(doc, "CSSBody")
            p_req.paragraph_format.left_indent = Inches(1.0)
            p_req.paragraph_format.first_line_indent = Inches(-0.22)
            p_req.add_run("a. ").bold = True
            p_req.add_run(f"Requirement: {it['requirement']}")
            p_bi = doc.add_paragraph()
            p_bi.style = _safe_style(doc, "CSSBody")
            p_bi.paragraph_format.left_indent = Inches(1.0)
            p_bi.paragraph_format.first_line_indent = Inches(-0.22)
            p_bi.add_run("b. ").bold = True
            p_bi.add_run(f"Business Impact: {it['business_impact']}")
            doc.add_paragraph()

    write_reco_section("Best Practises", _CSS_GREEN, ai["best_practices"], "bm_best")
    write_reco_section("Must Have", _CSS_RED, ai["must_have"], "bm_must")
    write_reco_section("Should Have", _CSS_BLUE, ai["should_have"], "bm_should")

    doc.add_page_break()
    _heading(doc, "SUMMARY OF FINDINGS & CONCLUSION", level=1, bookmark="bm_concl")
    concl = (ai.get("conclusion") or "").strip()
    parts = concl.split("\n\n", 1)
    if len(parts) == 2:
        _body_paragraph(doc, parts[0])
        p_lab = doc.add_paragraph()
        p_lab.style = _safe_style(doc, "CSSBodyBold")
        p_lab.add_run("CONCLUSION:").bold = True
        _body_paragraph(doc, parts[1])
    else:
        p_lab = doc.add_paragraph()
        p_lab.style = _safe_style(doc, "CSSBodyBold")
        p_lab.add_run("CONCLUSION:").bold = True
        _body_paragraph(doc, parts[0])
    _body_paragraph(doc, THANK_YOU_LINE, bold=True)

    return doc


def _patch_docx_settings_update_fields(docx_path: str) -> None:
    """Ensure Word refreshes fields (TOC, PAGE) on open. Protected View still blocks until user enables editing."""
    try:
        buf = io.BytesIO()
        with zipfile.ZipFile(docx_path, "r") as zin:
            st = zin.read("word/settings.xml").decode("utf-8")
            if re.search(r"<w:updateFields\b", st):
                st = re.sub(
                    r"<w:updateFields[^>]*/>",
                    '<w:updateFields w:val="true"/>',
                    st,
                    count=1,
                )
            else:
                st = re.sub(
                    r"(<w:settings[^>]*>)",
                    r'\1<w:updateFields w:val="true"/>',
                    st,
                    count=1,
                )
            with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    data = zin.read(item.filename)
                    if item.filename == "word/settings.xml":
                        data = st.encode("utf-8")
                    zout.writestr(item, data)
        with open(docx_path, "wb") as f:
            f.write(buf.getvalue())
    except Exception:
        logger.exception("exec_summary: failed to patch word/settings.xml (updateFields) in %s", docx_path)


def _patch_docx_document_pageref_placeholders(docx_path: str) -> None:
    """Remove cached PAGEREF result digits (e.g. placeholder '1') so Word shows real page numbers after update."""
    try:
        buf = io.BytesIO()
        with zipfile.ZipFile(docx_path, "r") as zin:
            if "word/document.xml" not in zin.namelist():
                return
            doc = zin.read("word/document.xml").decode("utf-8")
            doc = re.sub(
                r'(<w:instrText[^>]*>\s*PAGEREF\s+\w+\s+\\h\s*</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r>)<w:t>\d+</w:t>(</w:r><w:r><w:fldChar w:fldCharType="end"/>)',
                r'\1<w:t xml:space="preserve"></w:t>\2',
                doc,
            )
            with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    data = zin.read(item.filename)
                    if item.filename == "word/document.xml":
                        data = doc.encode("utf-8")
                    zout.writestr(item, data)
        with open(docx_path, "wb") as f:
            f.write(buf.getvalue())
    except Exception:
        logger.exception("exec_summary: failed to patch PAGEREF placeholders in %s", docx_path)


def _patch_docx_after_save(docx_path: str) -> None:
    """Settings (update fields) + strip bogus PAGEREF placeholders so page numbers are dynamic."""
    _patch_docx_settings_update_fields(docx_path)
    _patch_docx_document_pageref_placeholders(docx_path)


@app.get("/", response_class=HTMLResponse)
async def exec_summary_home() -> HTMLResponse:
    return HTMLResponse(_read_html_template())


@app.get("/health")
async def health():
    return {"status": "ok", "service": "exec-summary"}


@app.post("/generate/")
async def generate_exec_summary(
    request: Request,
    excel_file: UploadFile = File(...),
    report_type: str = Form(...),
    client_name: str = Form(...),
    org_name: str = Form("CyberSmithSECURE"),
    report_date: str = Form(...),
    assessment_type: str = Form(...),
    document_id: str = Form(""),
):
    report_type = report_type.strip().lower()
    if report_type not in VALID_REPORT_TYPES:
        raise HTTPException(status_code=400, detail="Invalid report type")

    if not excel_file.filename or not excel_file.filename.lower().endswith(".xlsx"):
        raise HTTPException(status_code=400, detail="Please upload a valid .xlsx file")

    _enforce_rate_limit(request)

    metadata = {
        "client_name": client_name.strip(),
        "org_name": org_name.strip() or "CyberSmithSECURE",
        "report_date": report_date.strip(),
        "assessment_type": assessment_type.strip(),
        "document_id": document_id.strip(),
        "report_type_label": report_type,
    }

    temp_dir = tempfile.mkdtemp(prefix="exec_summary_")
    excel_path = os.path.join(temp_dir, excel_file.filename)

    try:
        with open(excel_path, "wb") as f:
            f.write(await excel_file.read())

        assets, total_counts, top_v, allowed_cve = _parse_excel_to_payload(excel_path, report_type, metadata)
        user_json = build_canonical_user_payload(metadata, assets, total_counts, top_v)
        allowed_hosts = {a.get("hostname", "").lower() for a in assets if a.get("hostname")}

        ai_out = _generate_with_retries(user_json, allowed_cve, allowed_hosts)

        doc = _build_document(metadata, assets, total_counts, ai_out)
        out_name = f"{re.sub(r'[^a-zA-Z0-9_-]+', '_', client_name) or 'Client'}_Executive_Summary.docx"
        out_path = os.path.join(temp_dir, out_name)
        doc.save(out_path)
        _patch_docx_after_save(out_path)

        _log_audit(
            request,
            {
                "client_name": metadata["client_name"],
                "report_type": report_type,
                "asset_rows": len(assets),
                "vuln_rows": total_counts.get("total", 0),
                "output": out_name,
            },
        )

        def _cleanup_dir():
            shutil.rmtree(temp_dir, ignore_errors=True)

        return FileResponse(
            path=out_path,
            filename=out_name,
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            background=BackgroundTask(_cleanup_dir),
        )
    except HTTPException:
        __import__("shutil").rmtree(temp_dir, ignore_errors=True)
        raise
    except Exception as exc:
        traceback.print_exc()
        __import__("shutil").rmtree(temp_dir, ignore_errors=True)
        raise HTTPException(status_code=500, detail=f"Executive summary generation failed: {exc}") from exc

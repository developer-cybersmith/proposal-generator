"""
mitkat2.py
==========

Mitkat2 - Single-file VAPT report automation module.

Reads a single multi-sheet Excel workbook (MitKat_Report_Template.xlsx)
plus a POC zip and produces a Silbern-style .docx report. Designed to be
mounted alongside the existing /mitkat router without touching main.py:

    # In main.py (when ready):
    from mitkat2 import mitkat2_router
    app.include_router(mitkat2_router)

The module is intentionally self-contained: no shared state with main.py,
its own router, its own helpers, its own programmatic base template.

Public surface:
    - mitkat2_router          FastAPI APIRouter with /mitkat2/* endpoints
    - generate_report(...)    pure-python entrypoint (used by tests + CLI)
    - load_report_data(...)   Excel parser
    - load_poc_archive(...)   POC zip parser
    - validate_report(...)    structured validator
    - Mitkat2Error            base exception
    - ValidationIssue           structured validation / fidelity warnings

CLI smoke usage:
    python mitkat2.py --excel MitKat_Report_Template.xlsx \\
                      --poc   MitKat_POC_Sample.zip       \\
                      --out   report.docx

Standalone browser tester (no main.py):
    python mitkat2.py --serve
    # Open http://127.0.0.1:8765/mitkat2/ui — Validate + Generate + Health

Notes:
    - python-docx 1.1.x and openpyxl 3.x are required (already in
      requirements.txt). Pillow is optional but strongly recommended for
      large image inputs (resize/compress).

**Hallucination / content fidelity (“guardrails” in product sense)** —
MitKat2 maps workbook cells and POC assets into Word output without inventing
**vulnerability rows**, CVE identifiers, or POC screenshots. Optional
validation warns on draft markers (`[Insert …]`, `[TBD]`, …) left in source
cells so analysts catch mistakes before release.

The Executive Summary **Overview** subsections are assembled from the
``EXECUTIVE_SUMMARY`` worksheet (keys such as ``org_name``, ``overview_text``,
``findings_summary``, …), optionally polished by OpenAI when
``OPENAI_API_KEY`` is configured.

**Infrastructure limits** (upload size, zip-bomb checks) are separate —
they protect the server process; they do **not** cap how many vulnerability
rows or screenshots you may submit for legitimate reports (see env vars).

    - VULNERABILITIES is the master sheet for all vulnerability rows.
      All other sheets contribute to non-vuln sections of the report.
"""

from __future__ import annotations

# ---------------------------------------------------------------------------
# Imports
# ---------------------------------------------------------------------------

import os
import io
import re
import sys
import json
import time
from datetime import datetime
import shutil
import logging
import tempfile
import zipfile
import traceback
from dataclasses import dataclass, field, asdict
from pathlib import Path
from typing import Any, Iterable, Iterator, List, Dict, Optional, Tuple, Set, Sequence, cast

# Word
from docx import Document
from docx.document import Document as _DocxDocument
from docx.shared import Pt, Inches, Cm, Emu, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.enum.section import WD_SECTION
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn, nsdecls

# Excel
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

# Optional Pillow for image processing
try:
    from PIL import Image  # type: ignore
    _PIL_OK = True
    _PIL_RESAMPLE_LANCZOS = getattr(Image, "LANCZOS", 3)
except Exception:  # pragma: no cover
    _PIL_OK = False

# FastAPI is optional at import time so the module can be used as a library
# (e.g. CLI / tests) even where fastapi is unavailable.
try:
    from fastapi import APIRouter, UploadFile, File, HTTPException, BackgroundTasks
    from fastapi.responses import FileResponse, JSONResponse, HTMLResponse
    from starlette.responses import RedirectResponse, Response
    _FASTAPI_OK = True
except Exception:  # pragma: no cover
    _FASTAPI_OK = False
    class _DummyResponse:  # pragma: no cover
        def __init__(self, *args: Any, **kwargs: Any) -> None:
            pass

    HTMLResponse = _DummyResponse  # type: ignore[assignment]
    RedirectResponse = _DummyResponse  # type: ignore[assignment]
    Response = _DummyResponse  # type: ignore[assignment]


__version__ = "0.1.0"
log = logging.getLogger("mitkat2")
if not log.handlers:
    _h = logging.StreamHandler()
    _h.setFormatter(logging.Formatter("[%(asctime)s] %(name)s %(levelname)s | %(message)s"))
    log.addHandler(_h)
log.setLevel(os.environ.get("MITKAT2_LOG_LEVEL", "INFO").upper())


# ---------------------------------------------------------------------------
# Constants, resource limits & content-fidelity policy
# ---------------------------------------------------------------------------

# Sheet names (the workbook is the source of truth)
SHEET_REPORT_META = "REPORT_META"
SHEET_EXEC_SUMMARY = "EXECUTIVE_SUMMARY"
SHEET_TECH_MANPOWER = "TECH_MANPOWER"
SHEET_FRAMEWORKS = "STANDARD_FRAMEWORKS"
SHEET_TOOLS = "TOOLS_SCRIPTS"
SHEET_QUESTIONS = "QUESTIONS"
SHEET_VULNS = "VULNERABILITIES"
SHEET_POC_MAP = "POC_MAP"
SHEET_INSTRUCTIONS = "INSTRUCTIONS"

# Authoritative VULNERABILITIES headers (row 2). Order matters - this defines
# the geometry of the detail table layout used for each vuln.
VULN_COLUMNS: Tuple[str, ...] = (
    "sno",                    # A
    "vuln_name",              # B
    "vuln_location_full",     # C
    "cve_cwe",                # D
    "vuln_location",          # E
    "vuln_path",              # F
    "vuln_parameter",         # G
    "cvss_score",             # H
    "epss_score",             # I
    "severity",               # J
    "description",            # K
    "poc_steps_text",         # L
    "workaround",             # M
    "references",             # N
    "additional_obs",         # O
    "poc_folder",             # P
    "status",                 # Q
    "new_or_repeat",          # R
)

# 10-row detail table per vuln. The "Vulnerable Point / Error Location" row
# has a *nested* 3-row sub-table on the value side (Vulnerable Location /
# Vulnerable Path/Port/URL / Vulnerable Parameter) — see the user's reference
# screenshot for the visual layout.
# Special source-keys handled inline by `build_detail_report`:
#   "_serial_"                  : running 1-based vulnerability number
#   "_cvss_epss_combined_"      : "CVSS: x.y / EPSS: a.b" string
#   "_vulnerable_point_block_"  : nested sub-table with location/path/parameter
#   "_poc_steps_"               : numbered step-by-step PoC + screenshots
DETAIL_ROW_SPEC: Tuple[Tuple[str, str], ...] = (
    ("S. No.",                                          "_serial_"),
    ("Name of Vulnerability",                           "vuln_name"),
    ("Vulnerable Point / Error Location",               "_vulnerable_point_block_"),
    ("CVE/CWE",                                         "cve_cwe"),
    ("CVSS / EPSS Score",                               "_cvss_epss_combined_"),
    ("Description",                                     "description"),
    ("Proof of concept and Steps to Reproduce",         "_poc_steps_"),
    ("Workaround / Solutions / Recommendations",        "workaround"),
    ("References (Optional)",                           "references"),
    ("Additional observations",                         "additional_obs"),
)

ALLOWED_SEVERITIES = {"Critical", "High", "Medium", "Low", "Informational"}
SEVERITY_ORDER = ["Critical", "High", "Medium", "Low", "Informational"]

# Severity → background fill color (Silbern-ish palette)
SEVERITY_COLOR = {
    "Critical":      "C00000",
    "High":          "E97132",
    "Medium":        "FFC000",
    "Low":           "70AD47",
    "Informational": "5B9BD5",
}

ALLOWED_IMAGE_EXT = {".png", ".jpg", ".jpeg"}

# Style-related constants (Silbern-parity, measured from reference docx)
BODY_FONT_NAME = "Calibri"
BODY_FONT_PT = 11
HEADING1_PT = 20
HEADING2_PT = 16
HEADING_COLOR = "0F4761"     # accent1 majorHAnsi shade (Mitkat headings)
# Heading 2 font (per user request — overrides the Aptos Display theme major
# font that OLPST otherwise inherits). If "Play" is not installed on the
# rendering machine, Word will substitute its closest match automatically.
H2_FONT_NAME = "Play"
# Tables — match the visual reference shared by the user (1.png): white body
# cells, **light-grey header / label cells**, black borders, black bold text.
TABLE_LABEL_FILL = "F2F2F2"             # light grey for header / label cells
TABLE_BODY_FILL: Optional[str] = None    # None = no fill (white) for value cells
TABLE_LABEL_TEXT = "000000"              # black bold labels & header text
TABLE_BORDER_COLOR = "000000"            # black borders (per user request)
TABLE_HEADER_FILL = TABLE_LABEL_FILL     # legacy aliases (kept for back-compat)
TABLE_HEADER_TEXT = TABLE_LABEL_TEXT
FOOTER_BORDER_COLOR = "D9D9D9"
FOOTER_TEXT_COLOR = "7F7F7F"

# Page geometry (A4) - twips
A4_WIDTH_TWIPS = 11906
A4_HEIGHT_TWIPS = 16838
PAGE_MARGIN_LR_TWIPS = 1440
PAGE_MARGIN_TOP_TWIPS = 1440
PAGE_MARGIN_BOTTOM_TWIPS = 1134

# --- Resource limits (upload / unzip abuse protection only — not vuln caps) ---

def _env_int(key: str, default: int) -> int:
    try:
        v = os.environ.get(key)
        return int(v) if v else default
    except Exception:
        return default


def _env_bool(key: str, default: bool = False) -> bool:
    v = os.environ.get(key)
    if v is None:
        return default
    return str(v).strip().lower() in {"1", "true", "yes", "on"}


RL_MAX_EXCEL_BYTES = _env_int("MITKAT2_MAX_EXCEL_BYTES", 50 * 1024 * 1024)            # 50 MB
RL_MAX_ZIP_BYTES = _env_int("MITKAT2_MAX_ZIP_BYTES", 200 * 1024 * 1024)               # 200 MB
RL_MAX_UNZIP_BYTES = _env_int("MITKAT2_MAX_UNZIP_BYTES", 1024 * 1024 * 1024)          # 1 GB unpacked
RL_MAX_ZIP_RATIO = _env_int("MITKAT2_MAX_ZIP_RATIO", 200)                             # uncompressed/compressed
# POC images: no artificial cap on legitimate engagement scale — defaults are huge.
# Set MITKAT2_TRUNCATE_POC_IMAGES=1 only if you intentionally cap embed count per folder.
RL_MAX_IMAGES_PER_FOLDER = _env_int("MITKAT2_MAX_IMAGES_PER_FOLDER", 10**9)
RL_TRUNCATE_POC_IMAGES = _env_bool("MITKAT2_TRUNCATE_POC_IMAGES", False)
RL_MAX_TOTAL_IMAGES = _env_int("MITKAT2_MAX_TOTAL_IMAGES", 10**9)
RL_SOFT_WARN_TOTAL_IMAGES = _env_int("MITKAT2_SOFT_WARN_TOTAL_IMAGES", 50000)
RL_MAX_IMAGE_WIDTH_PX = _env_int("MITKAT2_MAX_IMAGE_WIDTH_PX", 1200)
RL_GENERATE_TIMEOUT_S = _env_int("MITKAT2_GENERATE_TIMEOUT_S", 600)                   # 10 min (document only)
RL_INSERT_INSTRUCTIONS_AS_ANNEX = bool(int(os.environ.get("MITKAT2_INCLUDE_INSTRUCTIONS", "0")))

# Back-compat aliases (older naming); prefer RL_* in new code.
GR_MAX_EXCEL_BYTES = RL_MAX_EXCEL_BYTES
GR_MAX_ZIP_BYTES = RL_MAX_ZIP_BYTES
GR_MAX_UNZIP_BYTES = RL_MAX_UNZIP_BYTES
GR_MAX_ZIP_RATIO = RL_MAX_ZIP_RATIO
GR_MAX_IMAGES_PER_FOLDER = RL_MAX_IMAGES_PER_FOLDER
GR_MAX_TOTAL_IMAGES = RL_MAX_TOTAL_IMAGES
GR_MAX_IMAGE_WIDTH_PX = RL_MAX_IMAGE_WIDTH_PX
GR_GENERATE_TIMEOUT_S = RL_GENERATE_TIMEOUT_S
GR_INSERT_INSTRUCTIONS_AS_ANNEX = RL_INSERT_INSTRUCTIONS_AS_ANNEX


# ---------------------------------------------------------------------------
# Exceptions
# ---------------------------------------------------------------------------

class Mitkat2Error(Exception):
    """Base error for Mitkat2."""


class ResourceLimitExceeded(Mitkat2Error):
    """An infrastructure upload/unzip size limit was exceeded."""

    def __init__(self, code: str, message: str, limit: Optional[int] = None,
                 actual: Optional[int] = None):
        super().__init__(message)
        self.code = code
        self.message = message
        self.limit = limit
        self.actual = actual

    def to_dict(self) -> Dict[str, Any]:
        return {"code": self.code, "message": self.message,
                "limit": self.limit, "actual": self.actual}


GuardrailError = ResourceLimitExceeded  # backwards-compat alias


@dataclass
class ValidationIssue:
    code: str
    message: str
    sheet: Optional[str] = None
    row: Optional[int] = None
    column: Optional[str] = None
    severity: str = "error"     # error | warning


# ---------------------------------------------------------------------------
# Schema dataclasses
# ---------------------------------------------------------------------------

@dataclass
class ReportMeta:
    fields: Dict[str, str] = field(default_factory=dict)

    def get(self, key: str, default: str = "") -> str:
        v = self.fields.get(key)
        return str(v) if v not in (None, "") else default

    @property
    def report_title(self) -> str:
        return self.get("report_title", "MitKat VAPT Report")

    @property
    def client_name(self) -> str:
        return self.get("client_name", "Client")

    @property
    def report_date(self) -> str:
        return self.get("report_date", _default_report_date())

    @property
    def assessment_code(self) -> str:
        return self.get("assessment_code", "")


@dataclass
class ExecSummaryBlock:
    """Generic key/value entry from EXECUTIVE_SUMMARY.

    Column A = FIELD_KEY (snake_case identifier), Column B = VALUE.
    The doc_builder groups blocks by key prefix to mirror Silbern layout.
    """
    key: str
    value: str


@dataclass
class ManpowerEntry:
    sno: str = ""
    name: str = ""
    designation: str = ""
    mobile: str = ""


@dataclass
class FrameworkEntry:
    sno: str = ""
    framework: str = ""
    severity_system: str = ""
    references: str = ""


@dataclass
class ToolEntry:
    sno: str = ""
    category: str = ""
    tools: str = ""
    purpose: str = ""


@dataclass
class QuestionEntry:
    qno: str = ""
    question: str = ""
    answer: str = ""


@dataclass
class Vulnerability:
    sno: str = ""
    vuln_name: str = ""
    vuln_location_full: str = ""
    cve_cwe: str = ""
    vuln_location: str = ""
    vuln_path: str = ""
    vuln_parameter: str = ""
    cvss_score: str = ""
    epss_score: str = ""
    severity: str = ""
    description: str = ""
    poc_steps_text: str = ""
    workaround: str = ""
    references: str = ""
    additional_obs: str = ""
    poc_folder: str = ""
    status: str = ""
    new_or_repeat: str = ""
    excel_row: int = 0


@dataclass
class PocMapEntry:
    sno: str = ""
    vuln_name: str = ""
    poc_folder: str = ""
    screenshot_count: int = 0
    expected_files: str = ""
    notes: str = ""


@dataclass
class ReportData:
    meta: ReportMeta = field(default_factory=ReportMeta)
    exec_blocks: List[ExecSummaryBlock] = field(default_factory=list)
    manpower: List[ManpowerEntry] = field(default_factory=list)
    frameworks: List[FrameworkEntry] = field(default_factory=list)
    tools: List[ToolEntry] = field(default_factory=list)
    questions: List[QuestionEntry] = field(default_factory=list)
    vulnerabilities: List[Vulnerability] = field(default_factory=list)
    poc_map: List[PocMapEntry] = field(default_factory=list)
    instructions_text: str = ""

    def severity_counts(self) -> Dict[str, int]:
        out = {sev: 0 for sev in SEVERITY_ORDER}
        for v in self.vulnerabilities:
            sev = (v.severity or "").strip().capitalize()
            if sev in out:
                out[sev] += 1
        return out


@dataclass
class PocArchive:
    root: Path
    folders: Dict[str, List[Path]] = field(default_factory=dict)

    def images_for(self, folder_key: str) -> List[Path]:
        return list(self.folders.get(folder_key, []))


# ---------------------------------------------------------------------------
# Excel loader
# ---------------------------------------------------------------------------

def _cell_str(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def _iter_rows(ws, *, min_row: int = 1, max_row: Optional[int] = None,
               max_col: Optional[int] = None) -> Iterator[Tuple[int, Tuple[Any, ...]]]:
    for idx, row in enumerate(
        ws.iter_rows(min_row=min_row, max_row=max_row, max_col=max_col, values_only=True),
        start=min_row,
    ):
        yield idx, row


def _open_workbook(xlsx_path: Path):
    try:
        return load_workbook(xlsx_path, read_only=True, data_only=True)
    except Exception as ex:
        raise Mitkat2Error(f"Failed to open Excel workbook: {ex}") from ex


def _load_report_meta(wb) -> ReportMeta:
    if SHEET_REPORT_META not in wb.sheetnames:
        return ReportMeta()
    ws = wb[SHEET_REPORT_META]
    out: Dict[str, str] = {}
    for r_idx, row in _iter_rows(ws, min_row=2, max_col=3):
        if not row:
            continue
        key = _cell_str(row[0] if len(row) > 0 else None)
        val = _cell_str(row[1] if len(row) > 1 else None)
        if not key:
            continue
        out[key.lower()] = val
    return ReportMeta(fields=out)


def _load_executive_summary(wb) -> List[ExecSummaryBlock]:
    if SHEET_EXEC_SUMMARY not in wb.sheetnames:
        return []
    ws = wb[SHEET_EXEC_SUMMARY]
    blocks: List[ExecSummaryBlock] = []
    for r_idx, row in _iter_rows(ws, min_row=2, max_col=2):
        if not row:
            continue
        key = _cell_str(row[0] if len(row) > 0 else None)
        val = _cell_str(row[1] if len(row) > 1 else None)
        if not key and not val:
            continue
        blocks.append(ExecSummaryBlock(key=key or "", value=val or ""))
    return blocks


def _load_tech_manpower(wb) -> List[ManpowerEntry]:
    if SHEET_TECH_MANPOWER not in wb.sheetnames:
        return []
    ws = wb[SHEET_TECH_MANPOWER]
    out: List[ManpowerEntry] = []
    for r_idx, row in _iter_rows(ws, min_row=2, max_col=4):
        if not row or all((c is None or str(c).strip() == "") for c in row):
            continue
        out.append(ManpowerEntry(
            sno=_cell_str(row[0] if len(row) > 0 else None),
            name=_cell_str(row[1] if len(row) > 1 else None),
            designation=_cell_str(row[2] if len(row) > 2 else None),
            mobile=_cell_str(row[3] if len(row) > 3 else None),
        ))
    return out


def _load_frameworks(wb) -> List[FrameworkEntry]:
    if SHEET_FRAMEWORKS not in wb.sheetnames:
        return []
    ws = wb[SHEET_FRAMEWORKS]
    out: List[FrameworkEntry] = []
    for r_idx, row in _iter_rows(ws, min_row=2, max_col=4):
        if not row or all((c is None or str(c).strip() == "") for c in row):
            continue
        out.append(FrameworkEntry(
            sno=_cell_str(row[0] if len(row) > 0 else None),
            framework=_cell_str(row[1] if len(row) > 1 else None),
            severity_system=_cell_str(row[2] if len(row) > 2 else None),
            references=_cell_str(row[3] if len(row) > 3 else None),
        ))
    return out


def _load_tools(wb) -> List[ToolEntry]:
    if SHEET_TOOLS not in wb.sheetnames:
        return []
    ws = wb[SHEET_TOOLS]
    out: List[ToolEntry] = []
    for r_idx, row in _iter_rows(ws, min_row=2, max_col=4):
        if not row or all((c is None or str(c).strip() == "") for c in row):
            continue
        out.append(ToolEntry(
            sno=_cell_str(row[0] if len(row) > 0 else None),
            category=_cell_str(row[1] if len(row) > 1 else None),
            tools=_cell_str(row[2] if len(row) > 2 else None),
            purpose=_cell_str(row[3] if len(row) > 3 else None),
        ))
    return out


def _load_questions(wb) -> List[QuestionEntry]:
    if SHEET_QUESTIONS not in wb.sheetnames:
        return []
    ws = wb[SHEET_QUESTIONS]
    out: List[QuestionEntry] = []
    for r_idx, row in _iter_rows(ws, min_row=2, max_col=3):
        if not row or all((c is None or str(c).strip() == "") for c in row):
            continue
        out.append(QuestionEntry(
            qno=_cell_str(row[0] if len(row) > 0 else None),
            question=_cell_str(row[1] if len(row) > 1 else None),
            answer=_cell_str(row[2] if len(row) > 2 else None),
        ))
    return out


def _load_vulnerabilities(wb) -> List[Vulnerability]:
    if SHEET_VULNS not in wb.sheetnames:
        raise Mitkat2Error(f"Required sheet '{SHEET_VULNS}' not found in workbook")
    ws = wb[SHEET_VULNS]
    rows: List[Vulnerability] = []
    # Headers live on row 2 per template; data starts row 3.
    for r_idx, row in _iter_rows(ws, min_row=3, max_col=len(VULN_COLUMNS)):
        if not row or all((c is None or str(c).strip() == "") for c in row):
            continue
        kwargs: Dict[str, Any] = {"excel_row": r_idx}
        for i, key in enumerate(VULN_COLUMNS):
            kwargs[key] = _cell_str(row[i]) if i < len(row) else ""
        rows.append(Vulnerability(**kwargs))
    return rows


def _load_poc_map(wb) -> List[PocMapEntry]:
    if SHEET_POC_MAP not in wb.sheetnames:
        return []
    ws = wb[SHEET_POC_MAP]
    out: List[PocMapEntry] = []
    for r_idx, row in _iter_rows(ws, min_row=2, max_col=6):
        if not row or all((c is None or str(c).strip() == "") for c in row):
            continue
        try:
            sc = int(row[3]) if (len(row) > 3 and row[3] is not None and str(row[3]).strip() != "") else 0
        except Exception:
            sc = 0
        out.append(PocMapEntry(
            sno=_cell_str(row[0] if len(row) > 0 else None),
            vuln_name=_cell_str(row[1] if len(row) > 1 else None),
            poc_folder=_cell_str(row[2] if len(row) > 2 else None),
            screenshot_count=sc,
            expected_files=_cell_str(row[4] if len(row) > 4 else None),
            notes=_cell_str(row[5] if len(row) > 5 else None),
        ))
    return out


def _load_instructions(wb) -> str:
    if SHEET_INSTRUCTIONS not in wb.sheetnames:
        return ""
    ws = wb[SHEET_INSTRUCTIONS]
    lines: List[str] = []
    for r_idx, row in _iter_rows(ws, min_row=1, max_col=1):
        if not row:
            continue
        s = _cell_str(row[0] if len(row) > 0 else None)
        lines.append(s)
    return "\n".join(lines).strip()


def load_report_data(xlsx_path: str | os.PathLike) -> ReportData:
    """Load the entire workbook into a typed `ReportData`.

    Raises `Mitkat2Error` when the workbook cannot be opened or required
    sheets/headers are missing. `ResourceLimitExceeded` when **upload**
    size limits are exceeded (not vulnerability row counts).
    """
    p = Path(xlsx_path)
    if not p.exists():
        raise Mitkat2Error(f"Excel file not found: {p}")
    if p.stat().st_size > RL_MAX_EXCEL_BYTES:
        raise ResourceLimitExceeded(
            code="excel_too_large",
            message=f"Excel file exceeds maximum allowed size ({RL_MAX_EXCEL_BYTES} bytes).",
            limit=RL_MAX_EXCEL_BYTES,
            actual=p.stat().st_size,
        )
    wb = _open_workbook(p)
    try:
        data = ReportData(
            meta=_load_report_meta(wb),
            exec_blocks=_load_executive_summary(wb),
            manpower=_load_tech_manpower(wb),
            frameworks=_load_frameworks(wb),
            tools=_load_tools(wb),
            questions=_load_questions(wb),
            vulnerabilities=_load_vulnerabilities(wb),
            poc_map=_load_poc_map(wb),
            instructions_text=_load_instructions(wb),
        )
        return data
    finally:
        try:
            wb.close()
        except Exception:
            pass


# ---------------------------------------------------------------------------
# POC zip loader
# ---------------------------------------------------------------------------

def _safe_extract_zip(zip_path: Path, dest: Path) -> None:
    """Extract `zip_path` into `dest`, refusing zip-bombs and traversal."""
    if zip_path.stat().st_size > RL_MAX_ZIP_BYTES:
        raise ResourceLimitExceeded(
            code="zip_too_large",
            message=f"POC zip exceeds maximum size ({RL_MAX_ZIP_BYTES} bytes).",
            limit=RL_MAX_ZIP_BYTES,
            actual=zip_path.stat().st_size,
        )
    dest.mkdir(parents=True, exist_ok=True)
    total_uncompressed = 0
    with zipfile.ZipFile(zip_path, "r") as zf:
        for info in zf.infolist():
            # Path traversal guard
            tgt = (dest / info.filename).resolve()
            if not str(tgt).startswith(str(dest.resolve())):
                raise Mitkat2Error(f"Refused unsafe zip entry: {info.filename}")
            if info.is_dir():
                tgt.mkdir(parents=True, exist_ok=True)
                continue
            # Per-entry compression ratio guard
            if info.file_size > 0 and info.compress_size > 0:
                ratio = info.file_size / max(info.compress_size, 1)
                if ratio > RL_MAX_ZIP_RATIO:
                    raise ResourceLimitExceeded(
                        code="zip_bomb_ratio",
                        message=f"Suspicious compression ratio in zip entry {info.filename}.",
                        limit=RL_MAX_ZIP_RATIO,
                        actual=int(ratio),
                    )
            total_uncompressed += info.file_size
            if total_uncompressed > RL_MAX_UNZIP_BYTES:
                raise ResourceLimitExceeded(
                    code="zip_unpacked_too_large",
                    message=("POC zip would unpack beyond maximum allowed total bytes "
                             f"({RL_MAX_UNZIP_BYTES})."),
                    limit=RL_MAX_UNZIP_BYTES,
                    actual=total_uncompressed,
                )
            tgt.parent.mkdir(parents=True, exist_ok=True)
            with zf.open(info, "r") as src, open(tgt, "wb") as dst:
                shutil.copyfileobj(src, dst)


def _find_poc_root(extracted_dir: Path) -> Path:
    """Locate a 'POC' directory inside the extracted tree, if present.

    Returns the directory considered the POC root. Falls back to the extracted
    root if no nested POC folder is found.
    """
    candidates: List[Path] = []
    for child in extracted_dir.iterdir():
        if child.is_dir():
            if child.name.lower() == "poc":
                candidates.append(child)
            else:
                # also accept first-level wrapper containing POC
                inner = child / "POC"
                if inner.exists() and inner.is_dir():
                    candidates.append(inner)
    if candidates:
        return candidates[0]
    return extracted_dir


def _index_poc_folders(poc_root: Path) -> Dict[str, List[Path]]:
    """Return {folder_name: [image_paths sorted numerically]} index.

    Skips README.txt and any non-image files. Skips folder names that
    don't look like POC folders only when they contain no images.
    """
    out: Dict[str, List[Path]] = {}
    if not poc_root.exists():
        return out
    total_images = 0
    for entry in sorted(poc_root.iterdir()):
        if not entry.is_dir():
            continue
        imgs: List[Path] = []
        for f in entry.iterdir():
            if not f.is_file():
                continue
            if f.suffix.lower() not in ALLOWED_IMAGE_EXT:
                continue
            imgs.append(f)
        if not imgs:
            continue
        # Numeric-ish sort: split into (numeric_prefix, lower_name)
        def _key(p: Path):
            m = re.match(r"^(\d+)", p.stem)
            return (int(m.group(1)) if m else 10**9, p.name.lower())
        imgs.sort(key=_key)
        n_img = len(imgs)
        if RL_TRUNCATE_POC_IMAGES and n_img > RL_MAX_IMAGES_PER_FOLDER:
            log.warning(
                "Folder %s has %d images; truncating to %d (MITKAT2_TRUNCATE_POC_IMAGES=1)",
                entry.name, n_img, RL_MAX_IMAGES_PER_FOLDER,
            )
            imgs = imgs[:RL_MAX_IMAGES_PER_FOLDER]
        elif not RL_TRUNCATE_POC_IMAGES and n_img > RL_MAX_IMAGES_PER_FOLDER:
            log.warning(
                "Folder %s has %d images (no truncation — MITKAT2_TRUNCATE_POC_IMAGES=0)",
                entry.name, n_img,
            )
        out[entry.name] = imgs
        prev_total = total_images
        total_images += len(imgs)
        if prev_total < RL_SOFT_WARN_TOTAL_IMAGES <= total_images:
            log.warning(
                "POC archive embed index crossed %d images; expect large docx / CPU time.",
                RL_SOFT_WARN_TOTAL_IMAGES,
            )
        if total_images > RL_MAX_TOTAL_IMAGES:
            raise ResourceLimitExceeded(
                code="too_many_images",
                message=(
                    f"POC archive exceeds MITKAT2_MAX_TOTAL_IMAGES ({RL_MAX_TOTAL_IMAGES}) "
                    "— raise the env value if this is a legitimate engagement."
                ),
                limit=RL_MAX_TOTAL_IMAGES,
                actual=total_images,
            )
    return out


def load_poc_archive(zip_path: str | os.PathLike, work_dir: str | os.PathLike) -> PocArchive:
    """Extract and index `zip_path` into `work_dir`. Returns a `PocArchive`."""
    z = Path(zip_path)
    w = Path(work_dir)
    if not z.exists():
        raise Mitkat2Error(f"POC zip not found: {z}")
    extracted = w / "_poc_extracted"
    if extracted.exists():
        shutil.rmtree(extracted)
    _safe_extract_zip(z, extracted)
    root = _find_poc_root(extracted)
    folders = _index_poc_folders(root)
    return PocArchive(root=root, folders=folders)


# ---------------------------------------------------------------------------
# Validator & content-fidelity (hallucination-safe emission rules)
# ---------------------------------------------------------------------------

# Draft/template markers left by analysts — MitKat2 does not “fill these in”.
_DRAFT_PLACEHOLDER_RE = re.compile(
    r"\[\s*(?:Insert|Enter|TBD|TODO|FIX\s*ME)\b|\(\s*(?:TBD|click here)",
    re.IGNORECASE,
)


def validate_content_fidelity(data: ReportData) -> List[ValidationIssue]:
    """Warn when workbook cells still look like drafts — **not** LLM hallucinations.

    MitKat2 is deterministic: it never invents vulnerability prose. These checks
    flag suspicious source text so analysts fix Excel before external release.
    """
    issues: List[ValidationIssue] = []

    auto_seen = any(
        "[AUTO-CALCULATED" in (b.value or "").upper() or
        "[AUTO" in (b.value or "") and "CALCUL" in (b.value or "").upper()
        for b in data.exec_blocks
    )
    if auto_seen:
        issues.append(ValidationIssue(
            code="fidelity_auto_count_substitution",
            message=(
                "[AUTO-CALCULATED…] markers found — MitKat2 substitutes counts deterministically "
                "from the VULNERABILITIES sheet at render time (no AI synthesis)."
            ),
            sheet=SHEET_EXEC_SUMMARY,
            severity="info",
        ))

    for blk in data.exec_blocks:
        for issue in _scan_text_fidelity(blk.value or "", sheet=SHEET_EXEC_SUMMARY,
                                         column=None, row=None):
            issues.append(issue)

    field_cols = {"description": "K", "poc_steps_text": "L", "workaround": "M",
                    "vuln_name": "B", "references": "N"}
    for v in data.vulnerabilities:
        for fname, col in field_cols.items():
            txt = getattr(v, fname, "") or ""
            for issue in _scan_text_fidelity(txt, sheet=SHEET_VULNS,
                                             row=v.excel_row, column=col):
                issues.append(issue)

    # Dedupe (same finding multiple cells)
    seen: Set[Tuple[str, str, Optional[int], Optional[str]]] = set()
    deduped: List[ValidationIssue] = []
    for iss in issues:
        sig = (iss.code, iss.message, iss.row, iss.column)
        if sig in seen:
            continue
        seen.add(sig)
        deduped.append(iss)
    return deduped


def _scan_text_fidelity(text: str, *, sheet: str,
                        row: Optional[int], column: Optional[str]) -> List[ValidationIssue]:
    issues: List[ValidationIssue] = []
    if not text.strip():
        return issues
    if _DRAFT_PLACEHOLDER_RE.search(text):
        issues.append(ValidationIssue(
            code="fidelity_draft_marker_in_source_cell",
            message=(
                "Draft/template marker detected ([Insert]/[TBD]/…). MitKat2 emits cells verbatim — "
                "review source Excel before distribution (anti–hallucination hygiene)."
            ),
            sheet=sheet,
            row=row,
            column=column,
            severity="warning",
        ))
    return issues


def validate_report(data: ReportData, poc: Optional[PocArchive] = None) -> List[ValidationIssue]:
    """Run cheap pre-generation checks. Returns a list of structured issues.

    Caller decides whether to abort (any item with severity == 'error') or
    proceed with warnings. Includes **content-fidelity** scans (draft markers)
    plus structural validation — does **not** cap vulnerability volume.
    """
    issues: List[ValidationIssue] = []
    issues.extend(validate_content_fidelity(data))

    # Meta
    if not data.meta.client_name:
        issues.append(ValidationIssue(
            code="meta_missing_client_name",
            message="REPORT_META.client_name is empty.",
            sheet=SHEET_REPORT_META,
        ))

    # Vuln-level checks
    seen_sno: Dict[str, int] = {}
    for v in data.vulnerabilities:
        if not v.sno:
            issues.append(ValidationIssue(
                code="vuln_missing_sno",
                message="Vulnerability row missing 'sno'.",
                sheet=SHEET_VULNS, row=v.excel_row, column="A",
            ))
        else:
            if v.sno in seen_sno:
                issues.append(ValidationIssue(
                    code="vuln_duplicate_sno",
                    message=f"Duplicate sno '{v.sno}' (also at row {seen_sno[v.sno]}).",
                    sheet=SHEET_VULNS, row=v.excel_row, column="A",
                ))
            else:
                seen_sno[v.sno] = v.excel_row
        if not v.vuln_name:
            issues.append(ValidationIssue(
                code="vuln_missing_name",
                message="Vulnerability row missing 'vuln_name'.",
                sheet=SHEET_VULNS, row=v.excel_row, column="B",
            ))
        if v.severity and v.severity.capitalize() not in ALLOWED_SEVERITIES:
            issues.append(ValidationIssue(
                code="vuln_invalid_severity",
                message=(f"Invalid severity '{v.severity}'. "
                         f"Allowed: {sorted(ALLOWED_SEVERITIES)}"),
                sheet=SHEET_VULNS, row=v.excel_row, column="J",
            ))
        if not v.poc_folder:
            issues.append(ValidationIssue(
                code="vuln_missing_poc_folder",
                message="Vulnerability row missing 'poc_folder'.",
                sheet=SHEET_VULNS, row=v.excel_row, column="P",
                severity="warning",
            ))

    # POC vs disk
    if poc is not None:
        for v in data.vulnerabilities:
            if v.poc_folder and v.poc_folder not in poc.folders:
                issues.append(ValidationIssue(
                    code="poc_folder_missing_on_disk",
                    message=f"POC folder '{v.poc_folder}' not found in archive.",
                    sheet=SHEET_VULNS, row=v.excel_row, column="P",
                    severity="warning",
                ))
        # POC_MAP cross-check (warnings only)
        for entry in data.poc_map:
            if entry.poc_folder and entry.poc_folder in poc.folders:
                actual = len(poc.folders[entry.poc_folder])
                if entry.screenshot_count and actual != entry.screenshot_count:
                    issues.append(ValidationIssue(
                        code="poc_count_mismatch",
                        message=(f"POC_MAP says folder '{entry.poc_folder}' should have "
                                 f"{entry.screenshot_count} images; archive has {actual}."),
                        sheet=SHEET_POC_MAP,
                        severity="warning",
                    ))
            elif entry.poc_folder:
                issues.append(ValidationIssue(
                    code="poc_map_folder_missing",
                    message=f"POC_MAP references folder '{entry.poc_folder}' not in archive.",
                    sheet=SHEET_POC_MAP,
                    severity="warning",
                ))

    return issues


# ---------------------------------------------------------------------------
# Image utils
# ---------------------------------------------------------------------------

def safe_image_for_docx(src: Path, work_dir: Path,
                        max_width_px: int = RL_MAX_IMAGE_WIDTH_PX) -> Optional[Path]:
    """Return a Path the docx writer can safely embed.

    If Pillow is available, copy/resize the image to a JPEG inside `work_dir`.
    On any failure, returns None (caller substitutes a placeholder paragraph).
    """
    if not src.exists():
        return None
    if not _PIL_OK:
        # No Pillow: pass through, but only if format matches docx's needs.
        return src if src.suffix.lower() in ALLOWED_IMAGE_EXT else None
    try:
        with Image.open(src) as im:
            im.load()
            if im.mode in ("RGBA", "P"):
                im = im.convert("RGB")
            w, h = im.size
            if w > max_width_px:
                new_h = int(h * (max_width_px / float(w)))
                im = im.resize((max_width_px, new_h), _PIL_RESAMPLE_LANCZOS)
            work_dir.mkdir(parents=True, exist_ok=True)
            out = work_dir / (src.stem + "_compressed.jpg")
            im.save(out, format="JPEG", quality=82, optimize=True)
            return out
    except Exception as ex:
        log.warning("Image processing failed for %s: %s", src, ex)
        return None


# ---------------------------------------------------------------------------
# Low-level docx helpers
# ---------------------------------------------------------------------------

def _set_cell_shading(cell, color_hex: Optional[str]) -> None:
    """Apply a fill color to a cell. ``None`` = no shading (transparent / white)."""
    if not color_hex:
        return
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), color_hex)
    tc_pr.append(shd)


def _set_table_layout_fixed(table) -> None:
    """Force fixed table layout so explicit column widths are honoured."""
    tbl_pr = table._tbl.tblPr
    layout = tbl_pr.find(qn("w:tblLayout"))
    if layout is None:
        layout = OxmlElement("w:tblLayout")
        tbl_pr.append(layout)
    layout.set(qn("w:type"), "fixed")


def _set_table_full_width(table, width_in: float = 6.5) -> None:
    """Force the table width in twips so it fills the printable area."""
    tbl_pr = table._tbl.tblPr
    w = tbl_pr.find(qn("w:tblW"))
    if w is None:
        w = OxmlElement("w:tblW")
        tbl_pr.append(w)
    w.set(qn("w:w"), str(int(width_in * 1440)))
    w.set(qn("w:type"), "dxa")


def _set_cell_no_wrap(cell) -> None:
    """Hint Word to prefer keeping cell text on a single line (auto-grow column)."""
    tc_pr = cell._tc.get_or_add_tcPr()
    nw = tc_pr.find(qn("w:noWrap"))
    if nw is None:
        tc_pr.append(OxmlElement("w:noWrap"))


def _set_repeat_header(row) -> None:
    tr_pr = row._tr.get_or_add_trPr()
    tbl_header = OxmlElement("w:tblHeader")
    tbl_header.set(qn("w:val"), "true")
    tr_pr.append(tbl_header)


def _set_cell_borders(cell, color_hex: Optional[str] = None, sz: int = 4) -> None:
    if color_hex is None:
        color_hex = TABLE_BORDER_COLOR
    tc_pr = cell._tc.get_or_add_tcPr()
    tc_borders = OxmlElement("w:tcBorders")
    for edge in ("top", "left", "bottom", "right"):
        b = OxmlElement(f"w:{edge}")
        b.set(qn("w:val"), "single")
        b.set(qn("w:sz"), str(sz))
        b.set(qn("w:color"), color_hex)
        tc_borders.append(b)
    tc_pr.append(tc_borders)


def _add_field(paragraph, instr: str, default_text: str = "") -> None:
    """Append a Word field (e.g., 'PAGE', 'TOC \\o ..') to a paragraph."""
    run = paragraph.add_run()
    fld_begin = OxmlElement("w:fldChar")
    fld_begin.set(qn("w:fldCharType"), "begin")
    run._r.append(fld_begin)

    instr_text = OxmlElement("w:instrText")
    instr_text.set(qn("xml:space"), "preserve")
    instr_text.text = instr
    run._r.append(instr_text)

    fld_sep = OxmlElement("w:fldChar")
    fld_sep.set(qn("w:fldCharType"), "separate")
    run._r.append(fld_sep)

    if default_text:
        t = OxmlElement("w:t")
        t.text = default_text
        run._r.append(t)

    fld_end = OxmlElement("w:fldChar")
    fld_end.set(qn("w:fldCharType"), "end")
    run._r.append(fld_end)


def _add_date_field(paragraph, format_code: str = "DD-MM-YYYY",
                     italic: bool = False, size_pt: int = BODY_FONT_PT) -> None:
    """Append a dynamic DATE field to a paragraph.

    The date field updates automatically when the document is opened or refreshed,
    even in view mode. Format codes: 'D' (default), 'DD-MM-YYYY', 'MMMM D, YYYY', etc.
    """
    run = paragraph.add_run()
    if italic:
        run.italic = True
    if size_pt:
        run.font.size = Pt(size_pt)
    run.font.name = BODY_FONT_NAME

    fld_begin = OxmlElement("w:fldChar")
    fld_begin.set(qn("w:fldCharType"), "begin")
    run._r.append(fld_begin)

    instr_text = OxmlElement("w:instrText")
    instr_text.set(qn("xml:space"), "preserve")
    instr_text.text = f"DATE \\@ \"{format_code}\""
    run._r.append(instr_text)

    fld_sep = OxmlElement("w:fldChar")
    fld_sep.set(qn("w:fldCharType"), "separate")
    run._r.append(fld_sep)

    fld_end = OxmlElement("w:fldChar")
    fld_end.set(qn("w:fldCharType"), "end")
    run._r.append(fld_end)


def _add_bookmark(paragraph, name: str, bookmark_id: int) -> None:
    """Wrap the paragraph's content with a bookmark range.

    The bookmark anchors must sit **after** ``w:pPr`` — placing
    ``bookmarkStart`` before ``w:pPr`` can confuse Word's layout / field
    engine and yields broken ``PAGEREF`` outputs (everything stuck at the
    placeholder page number).
    """
    p = paragraph._p
    start = OxmlElement("w:bookmarkStart")
    start.set(qn("w:id"), str(bookmark_id))
    start.set(qn("w:name"), name)
    end = OxmlElement("w:bookmarkEnd")
    end.set(qn("w:id"), str(bookmark_id))
    insert_at = 0
    for idx, child in enumerate(p):
        if child.tag == qn("w:pPr"):
            insert_at = idx + 1
            break
    p.insert(insert_at, start)
    p.append(end)


def _set_keep_with_next(paragraph) -> None:
    """Force the paragraph to stay on the same page as the next one."""
    p_pr = paragraph._p.get_or_add_pPr()
    if p_pr.find(qn("w:keepNext")) is None:
        p_pr.append(OxmlElement("w:keepNext"))


def _set_keep_lines(paragraph) -> None:
    """Force all lines of the paragraph to stay together on one page."""
    p_pr = paragraph._p.get_or_add_pPr()
    if p_pr.find(qn("w:keepLines")) is None:
        p_pr.append(OxmlElement("w:keepLines"))


def _set_page_break_before(paragraph) -> None:
    """Mark this paragraph to start on a new page (used for new sections)."""
    p_pr = paragraph._p.get_or_add_pPr()
    if p_pr.find(qn("w:pageBreakBefore")) is None:
        p_pr.append(OxmlElement("w:pageBreakBefore"))


def _set_cant_split(row) -> None:
    """Prevent a table row from splitting across pages."""
    tr_pr = row._tr.get_or_add_trPr()
    if tr_pr.find(qn("w:cantSplit")) is None:
        tr_pr.append(OxmlElement("w:cantSplit"))


def _set_paragraph_top_border(paragraph, color_hex: str = "D9D9D9") -> None:
    p_pr = paragraph._p.get_or_add_pPr()
    pbdr = p_pr.find(qn("w:pBdr"))
    if pbdr is None:
        pbdr = OxmlElement("w:pBdr")
        p_pr.append(pbdr)
    top = pbdr.find(qn("w:top"))
    if top is None:
        top = OxmlElement("w:top")
        pbdr.append(top)
    top.set(qn("w:val"), "single")
    top.set(qn("w:sz"), "4")
    top.set(qn("w:space"), "1")
    top.set(qn("w:color"), color_hex)


def _add_text_run(paragraph, text: str, *, bold: bool = False, italic: bool = False,
                  color_hex: Optional[str] = None, size_pt: Optional[int] = None,
                  font_name: Optional[str] = None) -> Any:
    run = paragraph.add_run(text or "")
    if bold:
        run.bold = True
    if italic:
        run.italic = True
    if color_hex:
        run.font.color.rgb = RGBColor.from_string(color_hex)
    if size_pt:
        run.font.size = Pt(size_pt)
    if font_name:
        run.font.name = font_name
    return run


def _set_default_font(doc: _DocxDocument) -> None:
    style = cast(Any, doc.styles["Normal"])
    style.font.name = BODY_FONT_NAME
    style.font.size = Pt(BODY_FONT_PT)
    rpr = style.element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.append(rfonts)
    for attr in ("ascii", "hAnsi", "cs", "eastAsia"):
        rfonts.set(qn(f"w:{attr}"), BODY_FONT_NAME)


def _restart_page_numbering(doc: _DocxDocument) -> None:
    """Force the first section of the document to restart page numbering at 1."""
    section = doc.sections[0]
    sectPr = section._sectPr
    pgNumType = sectPr.find(qn("w:pgNumType"))
    if pgNumType is None:
        pgNumType = OxmlElement("w:pgNumType")
        sectPr.append(pgNumType)
    pgNumType.set(qn("w:start"), "1")


def _configure_heading(doc: _DocxDocument, level: int, size_pt: int, color_hex: str) -> None:
    name = f"Heading {level}"
    if name not in doc.styles:
        return
    style = cast(Any, doc.styles[name])
    style.font.name = BODY_FONT_NAME
    style.font.size = Pt(size_pt)
    style.font.bold = True
    style.font.color.rgb = RGBColor.from_string(color_hex)


def _set_page_geometry(section) -> None:
    section.page_height = Emu(int(A4_HEIGHT_TWIPS * 635))   # 1 twip = 635 EMU
    section.page_width = Emu(int(A4_WIDTH_TWIPS * 635))
    section.left_margin = Emu(int(PAGE_MARGIN_LR_TWIPS * 635))
    section.right_margin = Emu(int(PAGE_MARGIN_LR_TWIPS * 635))
    section.top_margin = Emu(int(PAGE_MARGIN_TOP_TWIPS * 635))
    section.bottom_margin = Emu(int(PAGE_MARGIN_BOTTOM_TWIPS * 635))


def _set_update_fields_on_open(doc: _DocxDocument) -> None:
    settings = doc.settings.element
    if settings.find(qn("w:updateFields")) is None:
        upd = OxmlElement("w:updateFields")
        upd.set(qn("w:val"), "true")
        settings.append(upd)


def _default_report_date() -> str:
    return datetime.now().strftime("%d-%m-%Y")


_HEADING_STYLE_NAMES: Dict[int, str] = {
    1: "Heading 1",
    2: "Heading 2",
    3: "Heading 3",
    4: "Heading 4",
}


def _override_heading_font(doc: _DocxDocument, *, level: int, font_name: str) -> None:
    """Pin a Heading-N style to a literal font (no theme inheritance).

    The OLPST chrome we clone uses ``majorHAnsi`` for Heading 2 — which
    resolves to whatever the theme defines (``Aptos Display`` in OLPST). The
    user wants ``Play`` instead, so we strip the theme references and set
    explicit ``ascii`` / ``hAnsi`` / ``cs`` / ``eastAsia`` font names. Every
    H2 emitted by ``doc.add_heading(..., level=2)`` then picks up Play
    automatically without per-run overrides.
    """
    style_name = _HEADING_STYLE_NAMES.get(level)
    if not style_name:
        return
    try:
        style = doc.styles[style_name]
    except KeyError:
        log.warning("Heading style %s not found — cannot override font", style_name)
        return
    elem = style.element
    rpr = elem.find(qn("w:rPr"))
    if rpr is None:
        rpr = OxmlElement("w:rPr")
        elem.append(rpr)
    existing = rpr.find(qn("w:rFonts"))
    if existing is not None:
        rpr.remove(existing)
    rfonts = OxmlElement("w:rFonts")
    for attr in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        rfonts.set(qn(attr), font_name)
    # Insert at start of rPr so it visually leads, matching Word's convention.
    rpr.insert(0, rfonts)


# ---------------------------------------------------------------------------
# Base template builder
# ---------------------------------------------------------------------------

# Cloned styling chrome from the Silbern reference. Build by running
# `python _build_mitkat2_base_template.py` (commit the resulting docx).
BASE_TEMPLATE_PATH = Path(__file__).resolve().parent / "Mitkat2_Base_Template.docx"

# Static first-page cover template (page1.docx lives beside this module).
# When present, it is prepended as-is to every generated report so all
# reports share a single, designed cover page without any dynamic re-drawing.
COVER_PAGE_1_PATH = Path(__file__).resolve().parent / "page 1.docx"


def _substitute_header_client_name(doc: _DocxDocument, client_name: str) -> None:
    """Replace static reference client placeholders inside cloned headers.

    Walks every header/footer part in the package directly (rather than
    going through ``section.header`` / ``even_page_header`` which would
    cause python-docx to lazily *create* additional even-page header parts
    we don't want).

    Both reference-derived (``Silbern``) and our placeholder text
    (``CLIENT NAME``, ``{client_name}``) are accepted so the same code can
    chew on either source-template variant.
    """
    if not client_name:
        return
    candidates = ("Silbern", "{client_name}", "CLIENT NAME")
    pkg = getattr(doc.part, "package", None)
    if pkg is None:
        return
    for part in pkg.iter_parts():
        partname = str(getattr(part, "partname", ""))
        if not (partname.startswith("/word/header") or partname.startswith("/word/footer")):
            continue
        # python-docx exposes the lxml root via .element on HeaderPart/FooterPart
        el = getattr(part, "element", None)
        if el is None:
            continue
        for t in el.iter(qn("w:t")):
            if not t.text:
                continue
            txt = t.text
            for needle in candidates:
                if needle in txt:
                    txt = txt.replace(needle, client_name)
            if txt != t.text:
                t.text = txt


def build_base_document(meta: ReportMeta) -> _DocxDocument:
    """Open Mitkat2_Base_Template.docx (Silbern chrome) or fall back.

    Preferred: the cloned reference template provides exact styles, theme,
    fonts, header (with logo) and footer (page field, border, gray "Page"),
    matching the Silbern golden master. We only substitute the header
    client-name placeholders at runtime.
    """
    if BASE_TEMPLATE_PATH.exists():
        doc = Document(str(BASE_TEMPLATE_PATH))
        _substitute_header_client_name(doc, meta.client_name or "Client")
        _set_update_fields_on_open(doc)
        _override_heading_font(doc, level=2, font_name=H2_FONT_NAME)
        # Remove the seed empty paragraph so generated content starts at top.
        try:
            body = doc.element.body
            for p in body.iterchildren(qn("w:p")):
                if not p.findall(qn("w:r")) and not p.findall(qn("w:hyperlink")):
                    body.remove(p)
                    break
        except Exception:
            pass
        return doc

    log.warning(
        "Base template missing: %s — falling back to programmatic styling. "
        "Run `python _build_mitkat2_base_template.py` to generate it.",
        BASE_TEMPLATE_PATH,
    )

    doc = Document()
    _set_default_font(doc)
    _configure_heading(doc, 1, HEADING1_PT, HEADING_COLOR)
    _configure_heading(doc, 2, HEADING2_PT, HEADING_COLOR)
    if "Heading 3" in doc.styles:
        _configure_heading(doc, 3, 13, HEADING_COLOR)

    section = doc.sections[0]
    _set_page_geometry(section)
    section.different_first_page_header_footer = True
    _set_update_fields_on_open(doc)

    hdr = section.header
    p = hdr.paragraphs[0] if hdr.paragraphs else hdr.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    client = meta.client_name or "Client"
    _add_text_run(p, f"{client}  |  {client}", bold=True,
                  color_hex=HEADING_COLOR, size_pt=10, font_name=BODY_FONT_NAME)
    first_hdr = section.first_page_header
    if first_hdr.paragraphs:
        first_hdr.paragraphs[0].clear()

    ftr = section.footer
    fp = ftr.paragraphs[0] if ftr.paragraphs else ftr.add_paragraph()
    fp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    _set_paragraph_top_border(fp, FOOTER_BORDER_COLOR)
    _add_field(fp, " PAGE   \\* MERGEFORMAT ", default_text="1")
    _add_text_run(fp, "  |  Page", color_hex=FOOTER_TEXT_COLOR,
                  size_pt=10, font_name=BODY_FONT_NAME)
    first_ftr = section.first_page_footer
    if first_ftr.paragraphs:
        first_ftr.paragraphs[0].clear()
    return doc


# ---------------------------------------------------------------------------
# Static cover-page prepend helper
# ---------------------------------------------------------------------------

def _iter_doc_paragraphs(doc: _DocxDocument) -> Iterator[Any]:
    for p in doc.paragraphs:
        yield p
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    yield p


def _replace_cover_placeholders(doc: _DocxDocument, values: Dict[str, str]) -> None:
    if not values:
        return
    normalized = {str(k).strip().lower(): str(v) for k, v in values.items()}
    placeholder_re = re.compile(r"\{\{\s*([^}]+?)\s*\}\}|\{([^}]+?)\}")
    for p in _iter_doc_paragraphs(doc):
        if not p.text:
            continue
        new_text = p.text
        for match in placeholder_re.finditer(p.text):
            key = (match.group(1) or match.group(2) or "").strip().lower()
            if key and key in normalized:
                new_text = new_text.replace(match.group(0), normalized[key])
        if new_text == p.text:
            continue
        if p.runs:
            p.runs[0].text = new_text
            for r in p.runs[1:]:
                r.text = ""
        else:
            p.add_run(new_text)


def _replace_first_date_token(doc: _DocxDocument, date_text: str) -> None:
    if not date_text:
        return
    date_re = re.compile(r"\b\d{1,2}[-/]\d{1,2}[-/]\d{2,4}\b")
    for p in _iter_doc_paragraphs(doc):
        if not p.text:
            continue
        if date_re.search(p.text):
            new_text = date_re.sub(date_text, p.text, count=1)
            if p.runs:
                p.runs[0].text = new_text
                for r in p.runs[1:]:
                    r.text = ""
            else:
                p.add_run(new_text)
            break


def _prepend_static_cover_page(report_path: Path, meta: ReportMeta) -> None:
    """Prepend ``page 1.docx`` as a standalone cover page before the report.

    Strategy:
    1. Open ``page 1.docx`` as the **master** document.
    2. Add a section-break paragraph (nextPage) at the end of the cover
       document so Word treats the cover and the report as separate sections
       on separate pages.
    3. Set the new section to restart page numbering at 1.
    4. Use docxcompose to append the generated report body after the cover.
    5. Save the merged result back to ``report_path``.

    If ``page 1.docx`` is missing, logs a warning and the report is
    delivered without a cover.
    """
    if not COVER_PAGE_1_PATH.exists():
        log.warning(
            "Static cover page not found: %s — skipping cover prepend. "
            "Place 'page 1.docx' beside mitkat2.py to use a static cover page.",
            COVER_PAGE_1_PATH,
        )
        return
    try:
        import copy
        from docxcompose.composer import Composer  # type: ignore

        # ---- 1. Open the cover doc as master ----
        cover_doc = Document(str(COVER_PAGE_1_PATH))
        cover_values = {
            "report_title": meta.report_title,
            "client_name": meta.client_name,
            "report_date": meta.report_date,
            "assessment_code": meta.assessment_code,
            "date": meta.report_date,
        }
        _replace_cover_placeholders(cover_doc, cover_values)
        _replace_first_date_token(cover_doc, meta.report_date)

        # ---- 2. Clone the cover's own sectPr to create a section-break paragraph ----
        # In OOXML, a section break between two sections is stored as a
        # <w:sectPr> inside the *last paragraph's* <w:pPr> of that section.
        # We clone the cover document's body sectPr (page size, margins, etc.)
        # so the cover section keeps its original layout.
        cover_body = cover_doc.element.body
        body_sectPr = cover_body.find(qn("w:sectPr"))
        if body_sectPr is not None:
            break_sectPr = copy.deepcopy(body_sectPr)
        else:
            break_sectPr = OxmlElement("w:sectPr")

        # Set section type to nextPage — hard page break
        sect_type = break_sectPr.find(qn("w:type"))
        if sect_type is None:
            sect_type = OxmlElement("w:type")
            break_sectPr.append(sect_type)
        sect_type.set(qn("w:val"), "nextPage")

        # Remove any header/footer references from the break sectPr so they
        # don't clash with the report's own header/footer references.
        for tag in ("w:headerReference", "w:footerReference"):
            for ref in break_sectPr.findall(qn(tag)):
                break_sectPr.remove(ref)

        # Build a paragraph that carries this sectPr (the section-break paragraph)
        break_para = OxmlElement("w:p")
        break_pPr = OxmlElement("w:pPr")
        break_pPr.append(break_sectPr)
        break_para.append(break_pPr)

        # Insert this paragraph just before the body's own sectPr (i.e. at the
        # end of the body content, but before the trailing sectPr element).
        if body_sectPr is not None:
            body_sectPr.addprevious(break_para)
        else:
            cover_body.append(break_para)

        # ---- 3. Set page numbering on the report's section to start at 1 ----
        # The body sectPr (now the *report's* inherited section) should restart
        # page numbering.  After docxcompose appends the report, the report's
        # own sectPr becomes authoritative — but we also set pgNumType on the
        # cover-body sectPr (which becomes the boundary sectPr for the appended
        # content) so the restart is explicit.
        if body_sectPr is not None:
            pgNum = body_sectPr.find(qn("w:pgNumType"))
            if pgNum is None:
                pgNum = OxmlElement("w:pgNumType")
                body_sectPr.append(pgNum)
            pgNum.set(qn("w:start"), "1")

        # ---- 4. Append the generated report after the cover ----
        report_doc = Document(str(report_path))
        composer = Composer(cover_doc)
        composer.append(report_doc)

        # ---- 5. Save ----
        composer.save(str(report_path))

        # Ensure the merged document updates fields on open (TOC page numbers).
        merged = Document(str(report_path))
        _set_update_fields_on_open(merged)
        merged.save(str(report_path))
        log.info("Static cover page prepended from %s", COVER_PAGE_1_PATH)

    except Exception as exc:
        log.warning(
            "Failed to prepend static cover page (%s): %s — report saved without cover.",
            COVER_PAGE_1_PATH, exc,
        )



# ---------------------------------------------------------------------------
# Section builders
# ---------------------------------------------------------------------------

class _BookmarkRegistry:
    """Tracks bookmark IDs **and** the running list of headings for the static TOC."""

    def __init__(self) -> None:
        self._next_id = 1
        self.headings: List[Tuple[int, str, str]] = []   # (level, text, bookmark_name)
        self.toc_anchor: Any = None                       # placeholder paragraph (filled by build_toc)

    def next(self) -> int:
        v = self._next_id
        self._next_id += 1
        return v

    def record_heading(self, level: int, text: str, bookmark_name: str) -> None:
        self.headings.append((level, text, bookmark_name))


def _add_heading_with_bookmark(doc: _DocxDocument, text: str, *,
                               level: int, bookmarks: _BookmarkRegistry,
                               bookmark_name: Optional[str] = None,
                               page_break_before: bool = False) -> str:
    para = doc.add_heading(text, level=level)
    para.alignment = WD_ALIGN_PARAGRAPH.LEFT
    _set_keep_with_next(para)
    _set_keep_lines(para)
    if page_break_before:
        _set_page_break_before(para)
    name = bookmark_name or _slugify(text)
    _add_bookmark(para, name, bookmarks.next())
    bookmarks.record_heading(level, text, name)
    return name


def _slugify(text: str) -> str:
    s = re.sub(r"[^A-Za-z0-9]+", "_", (text or "").strip())
    s = re.sub(r"_+", "_", s).strip("_")
    if not s:
        s = f"bm{int(time.time()*1000) % 10**8}"
    return s[:40]


def _add_para(doc, text: str, *, bold: bool = False, italic: bool = False,
              size_pt: Optional[int] = None, color_hex: Optional[str] = None,
              alignment=None) -> Any:
    p = doc.add_paragraph()
    if alignment is not None:
        p.alignment = alignment
    _add_text_run(p, text or "", bold=bold, italic=italic,
                  size_pt=size_pt, color_hex=color_hex,
                  font_name=BODY_FONT_NAME)
    return p


def _build_table(doc: _DocxDocument, headers: List[str], rows: List[List[str]], *,
                 col_widths_in: Optional[List[float]] = None,
                 repeat_header: bool = False,
                 zebra: bool = False,                  # noqa: ARG001 — kept for API stability
                 single_line_headers: bool = True,
                 header_size_pt: int = BODY_FONT_PT) -> Any:
    """Render a table that visually matches the user-provided reference image.

    Style spec (per `1.png`):
      * White (transparent) cell fill - no solid blue / grey backgrounds.
      * Thin navy single-line borders on every cell (``TABLE_BORDER_COLOR``).
      * Header row: black bold labels, left aligned.
      * Body rows: black, left aligned, no zebra striping.

    Layout: fixed table width (``col_widths_in`` honoured), and any column whose
    header text would otherwise wrap is set to ``noWrap`` so the column auto-
    grows to keep header labels on a single line.
    """
    n_cols = len(headers)
    table = doc.add_table(rows=1 + len(rows), cols=n_cols)
    try:
        table.style = "Table Grid"
    except KeyError:
        pass
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    _set_table_layout_fixed(table)
    if col_widths_in:
        _set_table_full_width(table, sum(col_widths_in))

    head_row = table.rows[0]
    for i, h in enumerate(headers):
        cell = head_row.cells[i]
        cell.text = ""
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        _add_text_run(p, h, bold=True, color_hex=TABLE_LABEL_TEXT,
                      size_pt=header_size_pt, font_name=BODY_FONT_NAME)
        _set_cell_shading(cell, TABLE_LABEL_FILL)   # light-grey header fill
        _set_cell_borders(cell)
        if single_line_headers:
            _set_cell_no_wrap(cell)
    if repeat_header:
        _set_repeat_header(head_row)

    for r_idx, row_vals in enumerate(rows):
        tr = table.rows[r_idx + 1]
        for c_idx, val in enumerate(row_vals):
            cell = tr.cells[c_idx]
            cell.text = ""
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            _add_text_run(p, val if val is not None else "",
                          size_pt=BODY_FONT_PT, font_name=BODY_FONT_NAME)
            _set_cell_borders(cell)
            _set_cell_shading(cell, TABLE_BODY_FILL)    # white body cells

    if col_widths_in:
        for i, w in enumerate(col_widths_in):
            for row in table.rows:
                if i < len(row.cells):
                    row.cells[i].width = Inches(w)
    return table


COVER_BANNER_FILL = "02385A"   # Mitkat brand dark-blue (matches Mitkat_Base_template.docx)


def _set_row_height(row, twips: int, exact: bool = True) -> None:
    tr_pr = row._tr.get_or_add_trPr()
    h = OxmlElement("w:trHeight")
    h.set(qn("w:val"), str(twips))
    h.set(qn("w:hRule"), "exact" if exact else "atLeast")
    tr_pr.append(h)


def _set_cell_v_align(cell, val: str = "center") -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    v_align = OxmlElement("w:vAlign")
    v_align.set(qn("w:val"), val)
    tc_pr.append(v_align)


def build_cover(doc: _DocxDocument, meta: ReportMeta,
                bookmarks: _BookmarkRegistry) -> None:
    """Page-1 cover - matches Mitkat_Base_template.docx visual style.

    A full-width dark-blue banner across the top (Mitkat brand color
    ``#02385A``), centered white title text, then a centered title block
    underneath: report title, "for", client name, date, assessment code.
    """
    # ---- Top banner (single-cell table for full-width fill control) ----
    banner = doc.add_table(rows=1, cols=1)
    banner.alignment = WD_TABLE_ALIGNMENT.CENTER
    banner.autofit = False
    bcell = banner.cell(0, 0)
    bcell.width = Inches(6.5)
    _set_cell_shading(bcell, COVER_BANNER_FILL)
    _set_cell_v_align(bcell, "center")
    _set_row_height(banner.rows[0], twips=4320, exact=True)  # ~3 inches
    # Strip default cell borders for clean banner look
    tc_pr = bcell._tc.get_or_add_tcPr()
    tcBorders = OxmlElement("w:tcBorders")
    for edge in ("top", "left", "bottom", "right"):
        b = OxmlElement(f"w:{edge}")
        b.set(qn("w:val"), "nil")
        tcBorders.append(b)
    tc_pr.append(tcBorders)

    bp = bcell.paragraphs[0]
    bp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _add_text_run(bp, "MitKat Advisory Services Pvt. Ltd.",
                  bold=True, color_hex="FFFFFF", size_pt=20,
                  font_name=BODY_FONT_NAME)
    sub = bcell.add_paragraph()
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _add_text_run(sub, "Confidential - For Authorised Use Only",
                  italic=True, color_hex="C7D2DC", size_pt=11,
                  font_name=BODY_FONT_NAME)

    # ---- Spacer ----
    for _ in range(3):
        doc.add_paragraph()

    # ---- Report title block ----
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _add_text_run(p, "REPORT", bold=True, size_pt=14,
                  color_hex=HEADING_COLOR, font_name=BODY_FONT_NAME)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _add_text_run(p, meta.report_title or "VAPT Report",
                  bold=True, size_pt=30, color_hex=HEADING_COLOR,
                  font_name=BODY_FONT_NAME)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _add_text_run(p, "Prepared for", italic=True, size_pt=14,
                  color_hex="595959", font_name=BODY_FONT_NAME)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _add_text_run(p, meta.client_name or "Client", bold=True,
                  size_pt=26, color_hex=HEADING_COLOR,
                  font_name=BODY_FONT_NAME)

    if meta.report_date:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        _add_text_run(p, meta.report_date, italic=True, size_pt=14,
                      color_hex="595959", font_name=BODY_FONT_NAME)
        # Inject a DATE field so the date updates dynamically even in view mode
        _add_date_field(p, format_code="DD-MM-YYYY", italic=True, size_pt=14)
    if meta.assessment_code:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        _add_text_run(p, f"Assessment Code: {meta.assessment_code}",
                      size_pt=12, color_hex="595959",
                      font_name=BODY_FONT_NAME)

    doc.add_page_break()


def build_toc(doc: _DocxDocument, bookmarks: _BookmarkRegistry) -> None:
    """Emit the ``Contents`` heading and reserve a sentinel paragraph.

    The actual TOC entries are injected by :func:`_finalize_toc` *after* every
    section is built, so we know all H1/H2 headings up-front. The TOC is
    rendered as **static hyperlinked paragraphs** with a tab + dot leader and
    a ``PAGEREF`` field for each page number. This makes the topics + sub-
    topics visible immediately when the document is opened (even in Protected
    View, before the user clicks "Enable Editing"); the page numbers update
    silently the first time field calculation runs (driven by
    ``<w:updateFields w:val="true"/>`` in settings.xml).
    """
    h = doc.add_heading("Contents", level=1)
    h.alignment = WD_ALIGN_PARAGRAPH.LEFT
    _add_bookmark(h, "bm_toc", bookmarks.next())
    sentinel = doc.add_paragraph()
    bookmarks.toc_anchor = sentinel


# Tab stop position for the right-aligned dot-leader page number (in twips).
# 6.3 inches = 9072 twips, fits inside the 6.5 in. printable area.
_TOC_TAB_POS_TWIPS = 9072


def _build_toc_entry_xml(level: int, text: str, bookmark_name: str) -> Any:
    """Return a single ``<w:p>`` element for one TOC entry.

    Layout: ``<heading text>\t<page>`` where the heading is a hyperlink to the
    target bookmark and the page number is a ``PAGEREF`` field realised as a
    ``w:fldSimple`` (better field-update behaviour across Word versions than
    the multi-``fldChar`` form).

    Exact page figures are computed by Word's layout engine once fields are
    updated — see :func:`_try_refresh_word_fields` for an optional COM pass
    that forces an update immediately after saving.
    """
    p = OxmlElement("w:p")

    p_pr = OxmlElement("w:pPr")
    p.append(p_pr)
    pStyle = OxmlElement("w:pStyle")
    pStyle.set(qn("w:val"), f"TOC{level}")
    p_pr.append(pStyle)
    tabs = OxmlElement("w:tabs")
    tab = OxmlElement("w:tab")
    tab.set(qn("w:val"), "right")
    tab.set(qn("w:leader"), "dot")
    tab.set(qn("w:pos"), str(_TOC_TAB_POS_TWIPS))
    tabs.append(tab)
    p_pr.append(tabs)
    # Indent level-2 entries slightly for visual hierarchy
    if level >= 2:
        ind = OxmlElement("w:ind")
        ind.set(qn("w:left"), "360")  # 0.25"
        p_pr.append(ind)

    # Hyperlink (clickable to the heading bookmark)
    hyper = OxmlElement("w:hyperlink")
    hyper.set(qn("w:anchor"), bookmark_name)
    hyper.set(qn("w:history"), "1")
    p.append(hyper)

    r1 = OxmlElement("w:r")
    rpr = OxmlElement("w:rPr")
    rstyle = OxmlElement("w:rStyle")
    rstyle.set(qn("w:val"), "Hyperlink")
    rpr.append(rstyle)
    r1.append(rpr)
    t1 = OxmlElement("w:t")
    t1.text = text
    t1.set(qn("xml:space"), "preserve")
    r1.append(t1)
    hyper.append(r1)

    # Tab (jumps to right-aligned dot-leader stop)
    r2 = OxmlElement("w:r")
    tab_elem = OxmlElement("w:tab")
    r2.append(tab_elem)
    p.append(r2)

    # PAGEREF via fldSimple — Word refreshes this reliably on open / F9.
    fs = OxmlElement("w:fldSimple")
    fs.set(qn("w:instr"), f" PAGEREF {bookmark_name} \\h ")
    pr = OxmlElement("w:r")
    pt = OxmlElement("w:t")
    pt.text = "\u00a0"  # nbsp — visible gap until Word replaces with page no.
    pr.append(pt)
    fs.append(pr)
    p.append(fs)

    return p


def _finalize_toc(doc: _DocxDocument, bookmarks: _BookmarkRegistry) -> None:
    """Replace the TOC sentinel paragraph with real static TOC entries.

    Runs *after* every section is built so we know exactly which H1 / H2
    headings exist (including the dynamic per-question and per-vulnerability
    headings).
    """
    sentinel = bookmarks.toc_anchor
    if sentinel is None:
        return
    sentinel_p = sentinel._p
    parent = sentinel_p.getparent()
    if parent is None:
        return
    insert_idx = list(parent).index(sentinel_p)

    new_paras: List[Any] = []
    for level, text, bm_name in bookmarks.headings:
        if bm_name == "bm_toc":
            continue          # Don't TOC the Contents heading itself
        if level not in (1, 2):
            continue          # Match \o "1-2" scope from the user-provided new.png
        new_paras.append(_build_toc_entry_xml(level, text, bm_name))

    for offset, np in enumerate(new_paras):
        parent.insert(insert_idx + offset, np)
    parent.remove(sentinel_p)


def build_title_page(doc: _DocxDocument, meta: ReportMeta,
                     bookmarks: _BookmarkRegistry) -> None:
    _add_heading_with_bookmark(doc, "Report Information", level=1,
                               bookmarks=bookmarks, bookmark_name="bm_title")
    rows = []
    for key, label in (
        ("submitted_by", "Submitted By"),
        ("client_name", "Client"),
        ("assessment_code", "Assessment Code"),
        ("report_date", "Report Date"),
        ("test_duration", "Test Duration"),
        ("test_bed", "Test Bed"),
        ("contact_person", "Contact Person"),
        ("report_title", "Title"),
    ):
        val = meta.get(key)
        if val:
            rows.append([label, val])
    if rows:
        _build_table(doc, ["Field", "Value"], rows,
                     col_widths_in=[3.1, 3.15], zebra=True)
    doc.add_page_break()


def _add_bullet_para(doc_or_cell, text: str, *, bold_prefix: str = "") -> None:
    """Add a bulleted paragraph using whichever bullet style the template provides.

    Accepts either a ``Document`` or any container (cell / section) that exposes
    ``add_paragraph``. ``bold_prefix`` is rendered bold *before* the regular text.
    """
    add_para = doc_or_cell.add_paragraph
    styles = getattr(doc_or_cell, "styles", None)
    style_name: Optional[str] = None
    if styles is not None:
        for candidate in ("List Bullet", "List Paragraph"):
            if candidate in styles:
                style_name = candidate
                break
    if style_name == "List Bullet":
        p = add_para(style="List Bullet")
    elif style_name == "List Paragraph":
        p = add_para(style="List Paragraph")
    else:
        p = add_para()
    if not (style_name == "List Bullet"):
        _add_text_run(p, "•  ", size_pt=BODY_FONT_PT, font_name=BODY_FONT_NAME)
    if bold_prefix:
        _add_text_run(p, bold_prefix, bold=True,
                      size_pt=BODY_FONT_PT, font_name=BODY_FONT_NAME)
    _add_text_run(p, text, size_pt=BODY_FONT_PT, font_name=BODY_FONT_NAME)


# ---------------------------------------------------------------------------
# Overview generator (OpenAI assist + deterministic fallback)
# ---------------------------------------------------------------------------

OVERVIEW_SUBSECTIONS: Tuple[str, ...] = (
    "Name of the organization",
    "Overview of Test Setup and Assessment Approach",
    "Summary of Findings",
    "Key Observations & Critical Risk Areas",
    "Testing Methodology",
)

# Words / phrases the AI is told to emphasize (rendered bold). Conservative,
# domain-driven list - matches what is typically highlighted in OLPST/Silbern.
OVERVIEW_HIGHLIGHT_TERMS: Tuple[str, ...] = (
    "Critical", "High", "Medium", "Low", "Informational",
    "OWASP", "OWASP Top 10", "OWASP ASVS", "NIST", "CVSS", "CWE", "CVE",
    "Authentication Bypass", "SQL Injection", "Cross-Site Scripting", "XSS",
    "IDOR", "Sensitive Information Disclosure", "Privilege Escalation",
    "Remote Code Execution", "RCE", "Misconfiguration",
    "Confidentiality", "Integrity", "Availability",
    "Vulnerability Assessment", "Penetration Testing", "OLPST",
)


def _natural_sort_keys(keys: Sequence[str]) -> List[str]:
    """Sorted ``methodology_phase_2`` after ``methodology_phase_10`` correctly."""

    def key_tuple(k: str) -> Tuple[Any, ...]:
        parts = re.split(r"(\d+)", k)
        out: List[Any] = []
        for p in parts:
            if p.isdigit():
                out.append(int(p))
            elif p:
                out.append(p.lower())
        return tuple(out)

    return sorted(keys, key=key_tuple)


def _substitute_exec_placeholders(text: str, data: ReportData) -> str:
    """Expand MitKat template placeholders inside EXECUTIVE_SUMMARY cells."""
    if not text:
        return ""
    n = len(data.vulnerabilities)
    counts = data.severity_counts()
    sev_line = ", ".join(f"{k}: {counts[k]}" for k in SEVERITY_ORDER if counts[k])

    t = text
    t = t.replace("[AUTO-CALCULATED FROM VULNERABILITIES SHEET]", str(n))
    t = t.replace("[AUTO-CALCULATED]", str(n))
    # Common shorthand seen in templates
    if "[TOTAL_VULNS]" in t:
        t = t.replace("[TOTAL_VULNS]", str(n))
    if "[SEVERITY_BREAKDOWN]" in t and sev_line:
        t = t.replace("[SEVERITY_BREAKDOWN]", sev_line)
    return t


def _exec_blocks_blob(data: ReportData) -> Dict[str, str]:
    return {
        b.key.strip().lower(): (b.value or "")
        for b in data.exec_blocks
        if (b.key or "").strip()
    }


def _overview_sections_from_exec_sheet(data: ReportData) -> Dict[str, str]:
    """Map ``EXECUTIVE_SUMMARY`` rows → the five Overview subsection bodies."""
    blob = _exec_blocks_blob(data)
    sub: Dict[str, str] = {k: "" for k in OVERVIEW_SUBSECTIONS}

    org = (blob.get("org_name") or "").strip()
    if org:
        sub["Name of the organization"] = _substitute_exec_placeholders(org, data)

    setup_bits: List[str] = []
    for key in ("overview_text", "testing_approach", "environment_text"):
        v = (blob.get(key) or "").strip()
        if v:
            setup_bits.append(_substitute_exec_placeholders(v, data))
    if setup_bits:
        sub["Overview of Test Setup and Assessment Approach"] = "\n\n".join(
            setup_bits)

    find_parts: List[str] = []
    fs = (blob.get("findings_summary") or "").strip()
    if fs:
        find_parts.append(_substitute_exec_placeholders(fs, data))
    fb_keys = _natural_sort_keys([k for k in blob if k.startswith("findings_bullet")])
    for bk in fb_keys:
        v = (blob.get(bk) or "").strip()
        if v:
            find_parts.append(f"• {_substitute_exec_placeholders(v, data)}")
    if find_parts:
        sub["Summary of Findings"] = "\n".join(find_parts)

    ko_parts: List[str] = []
    intro = (blob.get("key_obs_intro") or "").strip()
    if intro:
        ko_parts.append(_substitute_exec_placeholders(intro, data))
    ko_keys = _natural_sort_keys([
        k for k in blob if re.match(r"^key_obs_\d+$", k)
    ])
    for kk in ko_keys:
        v = (blob.get(kk) or "").strip()
        if v:
            ko_parts.append(f"• {_substitute_exec_placeholders(v, data)}")
    if ko_parts:
        sub["Key Observations & Critical Risk Areas"] = "\n".join(ko_parts)

    meth_parts: List[str] = []
    mi = (blob.get("methodology_intro") or "").strip()
    if mi:
        meth_parts.append(_substitute_exec_placeholders(mi, data))
    meth_keys = _natural_sort_keys([
        k for k in blob
        if k.startswith("methodology_phase") or k.startswith("methodology_step")
    ])
    for mk in meth_keys:
        v = (blob.get(mk) or "").strip()
        if v:
            meth_parts.append(_substitute_exec_placeholders(v, data))
    if meth_parts:
        sub["Testing Methodology"] = "\n\n".join(meth_parts)

    return sub


def _overview_base_from_workbook(data: ReportData) -> Dict[str, str]:
    """Excel EXECUTIVE_SUMMARY first, deterministic overview fills gaps."""
    from_sheet = _overview_sections_from_exec_sheet(data)
    det = _deterministic_overview(data)
    return {
        k: ((from_sheet.get(k) or "").strip() or det[k])
        for k in OVERVIEW_SUBSECTIONS
    }


def _try_call_openai_for_overview(data: ReportData,
                                  workbook_draft: Dict[str, str],
                                  api_key: str,
                                  model: str) -> Optional[Dict[str, str]]:
    """Ask OpenAI for a structured Overview; return ``{subsection: text}`` or ``None``.

    Best-effort: any failure (no `openai` package, network error, malformed
    response, JSON parse failure) returns ``None`` so the deterministic
    fallback can run. **Never raises** to the caller.
    """
    try:
        from openai import OpenAI  # type: ignore
    except Exception:
        log.info("openai package not installed - skipping AI overview "
                 "(pip install openai to enable).")
        return None

    counts = data.severity_counts()
    vuln_titles = [v.vuln_name for v in data.vulnerabilities[:25] if v.vuln_name]
    framework_names = [f.framework for f in data.frameworks if f.framework][:10]
    tools_used = [t.tools for t in data.tools if t.tools][:10]

    user_prompt = json.dumps({
        "workbook_executive_summary": {
            "name_of_the_organization": workbook_draft.get(
                "Name of the organization", ""),
            "overview_of_test_setup_and_assessment_approach": workbook_draft.get(
                "Overview of Test Setup and Assessment Approach", ""),
            "summary_of_findings": workbook_draft.get(
                "Summary of Findings", ""),
            "key_observations_and_critical_risk_areas": workbook_draft.get(
                "Key Observations & Critical Risk Areas", ""),
            "testing_methodology": workbook_draft.get(
                "Testing Methodology", ""),
        },
        "engagement": {
            "client_name": data.meta.client_name,
            "report_title": data.meta.report_title,
            "assessment_code": data.meta.assessment_code,
            "report_date": data.meta.report_date,
            "test_bed_details": data.meta.get("test_bed_details"),
            "report_type": data.meta.get("report_type", "Online PST (OLPST)"),
        },
        "severity_counts": counts,
        "total_findings": sum(counts.values()),
        "frameworks_used": framework_names,
        "tools_used": tools_used,
        "top_vulnerability_titles": vuln_titles,
        "highlight_terms": list(OVERVIEW_HIGHLIGHT_TERMS),
    }, ensure_ascii=False)

    system_prompt = (
        "You are a senior cybersecurity consultant polishing the Executive "
        "Summary Overview for a MitKat Advisory VAPT (Online PST) report.\n"
        "The JSON field ``workbook_executive_summary`` carries paragraphs "
        "typed directly into the client's Excel ``EXECUTIVE_SUMMARY`` sheet — "
        "that text is the factual SOURCE OF TRUTH. You MUST preserve every "
        "substantive claim it contains (counts, dates, scope targets, bullet "
        "themes). You MAY tighten wording for clarity and consistent tone.\n"
        "You may cross-check severity_counts / total_findings from the payload "
        "— never invent CVE IDs, hosts, or counts not supported by the Excel "
        "draft or those structured facts.\n"
        "Return STRICT JSON with exactly these keys: "
        "'name_of_the_organization', "
        "'overview_of_test_setup_and_assessment_approach', "
        "'summary_of_findings', "
        "'key_observations_and_critical_risk_areas', "
        "'testing_methodology'. "
        "Each value should mirror the corresponding workbook paragraph(s); "
        "allow multi-sentence answers where the workbook content requires it. "
        "Wrap highlight_terms that appear with **double asterisks**. "
        "Do NOT output markdown headings or JSON outside the object."
    )

    try:
        client = OpenAI(api_key=api_key)
        response = client.chat.completions.create(
            model=model,
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user",   "content": user_prompt},
            ],
            response_format={"type": "json_object"},
            temperature=0.4,
            timeout=45,
        )
        raw = response.choices[0].message.content or ""
    except Exception as ex:
        log.warning("OpenAI Overview call failed (%s) - using deterministic fallback.", ex)
        return None

    try:
        parsed = json.loads(raw)
    except Exception as ex:
        log.warning("OpenAI Overview returned non-JSON (%s) - falling back.", ex)
        return None

    mapping = {
        "Name of the organization":                       parsed.get("name_of_the_organization", ""),
        "Overview of Test Setup and Assessment Approach": parsed.get("overview_of_test_setup_and_assessment_approach", ""),
        "Summary of Findings":                            parsed.get("summary_of_findings", ""),
        "Key Observations & Critical Risk Areas":         parsed.get("key_observations_and_critical_risk_areas", ""),
        "Testing Methodology":                            parsed.get("testing_methodology", ""),
    }
    if not any(mapping.values()):
        return None
    return {k: str(v) for k, v in mapping.items()}


def _deterministic_overview(data: ReportData) -> Dict[str, str]:
    """Fallback Overview text built strictly from workbook cells (no AI synthesis)."""
    meta = data.meta
    counts = data.severity_counts()
    total = sum(counts.values())
    sev_phrase = ", ".join(
        f"**{k}**: {counts[k]}" for k in SEVERITY_ORDER if counts.get(k)
    ) or "no findings of any severity"
    fw = ", ".join(f.framework for f in data.frameworks if f.framework) or "industry-standard frameworks"
    tools_used = ", ".join(t.tools for t in data.tools if t.tools) or "industry-standard tools"

    return {
        "Name of the organization":
            f"This report has been prepared for **{meta.client_name or 'the Client'}**"
            + (f" as part of the engagement *{meta.report_title}*." if meta.report_title else "."),
        "Overview of Test Setup and Assessment Approach":
            (f"MitKat Advisory Services Pvt. Ltd. conducted an "
             f"**{meta.get('report_type', 'Online Penetration Security Testing (OLPST)')}** "
             f"on the assets identified in the test bed "
             f"({meta.get('test_bed_details', '—')}). "
             f"The engagement was executed between "
             f"{meta.get('test_start_date', '—')} and {meta.get('test_end_date', '—')}, "
             f"using a black-box / grey-box methodology aligned with **{fw}**."),
        "Summary of Findings":
            (f"A total of **{total}** vulnerabilities were identified across the "
             f"in-scope assets - {sev_phrase}. Each finding is documented in "
             f"**Part 2: Detailed Report** with reproduction steps, supporting "
             f"evidence, and remediation guidance."),
        "Key Observations & Critical Risk Areas":
            (f"Critical and High severity issues require **immediate** attention "
             f"as they pose direct exposure of **Confidentiality**, **Integrity** "
             f"or **Availability**. Medium / Low findings, while less urgent, "
             f"contribute to overall attack surface and should be tracked to "
             f"closure as part of the routine security backlog."),
        "Testing Methodology":
            (f"The assessment followed MitKat's standard VAPT methodology - "
             f"reconnaissance, vulnerability identification, exploitation "
             f"validation, and reporting - executed using **{tools_used}** "
             f"and benchmarked against **{fw}**."),
    }


def generate_overview(data: ReportData) -> Dict[str, str]:
    """Return overview content for each fixed sub-heading.

    1. Loads narrative text from ``EXECUTIVE_SUMMARY`` (when populated),
       falling back to ``_deterministic_overview`` per subsection.
    2. When ``OPENAI_API_KEY`` is set and the API call succeeds, OpenAI
       polishes that workbook-derived draft (never replacing it with empty text).

    Env: ``OPENAI_API_KEY``, ``MITKAT2_OPENAI_MODEL`` (default ``gpt-4o-mini``).
    """
    api_key = os.environ.get("OPENAI_API_KEY", "").strip()
    model = os.environ.get("MITKAT2_OPENAI_MODEL", "gpt-4o-mini").strip()
    base = _overview_base_from_workbook(data)

    if not api_key:
        log.info("OPENAI_API_KEY not set — Overview uses EXECUTIVE_SUMMARY "
                 "(Excel) with deterministic gap-fill only.")
        return base

    ai = _try_call_openai_for_overview(
        data, workbook_draft=base, api_key=api_key, model=model,
    )
    if not ai:
        return base

    merged: Dict[str, str] = {}
    for key in OVERVIEW_SUBSECTIONS:
        v = (ai.get(key) or "").strip()
        merged[key] = v or base[key]
    return merged


def _add_rich_paragraph(container, text: str, *,
                        size_pt: int = BODY_FONT_PT) -> Any:
    """Add a paragraph that supports inline ``**bold**`` markers from the AI."""
    p = container.add_paragraph()
    parts = re.split(r"(\*\*[^*]+\*\*)", text or "")
    for part in parts:
        if not part:
            continue
        if part.startswith("**") and part.endswith("**"):
            _add_text_run(p, part[2:-2], bold=True,
                          size_pt=size_pt, font_name=BODY_FONT_NAME)
        else:
            _add_text_run(p, part, size_pt=size_pt, font_name=BODY_FONT_NAME)
    return p


def build_overview_section(doc: _DocxDocument, data: ReportData,
                           bookmarks: _BookmarkRegistry) -> None:
    """The single H2 ``Overview`` block — five fixed sub-points (bold labels).

    Sub-point labels are rendered as bold paragraph leaders (NOT Heading 3),
    so they don't appear in the Table of Contents (which is constrained to
    ``\\o "1-2"``). This matches the user-provided ``new.png`` exactly.
    """
    _add_heading_with_bookmark(doc, "Overview", level=2,
                               bookmarks=bookmarks, bookmark_name="bm_overview")
    overview = generate_overview(data)
    for label in OVERVIEW_SUBSECTIONS:
        body = (overview.get(label) or "").strip()
        if not body:
            continue
        # Bold sub-label as the first paragraph for visual separation.
        lp = doc.add_paragraph()
        _set_keep_with_next(lp)
        _add_text_run(lp, label, bold=True,
                      color_hex=HEADING_COLOR,
                      size_pt=BODY_FONT_PT, font_name=BODY_FONT_NAME)
        # Body paragraph(s)
        for line in body.split("\n"):
            line = line.strip()
            if not line:
                continue
            _add_rich_paragraph(doc, line, size_pt=BODY_FONT_PT)


def build_executive_summary(doc: _DocxDocument, data: ReportData,
                            bookmarks: _BookmarkRegistry) -> None:
    """Part 1 wrapper: H1 + Overview built from ``EXECUTIVE_SUMMARY`` + AI polish."""
    _add_heading_with_bookmark(doc, "Part 1: Executive Summary", level=1,
                               bookmarks=bookmarks, bookmark_name="bm_part1",
                               page_break_before=True)
    build_overview_section(doc, data, bookmarks)


def build_manpower(doc, data: ReportData, bookmarks: _BookmarkRegistry) -> None:
    """Heading + table per OLPST: 'List of technical manpower participated in the OLPST:'."""
    if not data.manpower:
        return
    _add_heading_with_bookmark(
        doc, "List of technical manpower participated in the OLPST:",
        level=2, bookmarks=bookmarks, bookmark_name="bm_manpower",
        page_break_before=True,
    )
    rows = [[m.sno, m.name, m.designation, m.mobile] for m in data.manpower]
    _build_table(doc, ["S. No.", "Name", "Designation", "Mobile No."], rows,
                 col_widths_in=[0.6, 2.05, 1.95, 1.4])


def build_frameworks(doc, data: ReportData, bookmarks: _BookmarkRegistry) -> None:
    """List of Standards, Frameworks, Severity Scoring Systems, and Other References:.

    Columns (per user spec): S.no | Standard / Framework used |
    Severity Score System Used | Other References.
    """
    if not data.frameworks:
        return
    _add_heading_with_bookmark(
        doc,
        "List of Standards, Frameworks, Severity Scoring Systems, and Other References:",
        level=2, bookmarks=bookmarks, bookmark_name="bm_frameworks",
        page_break_before=True,
    )
    rows = [[f.sno, f.framework, f.severity_system, f.references]
            for f in data.frameworks]
    _build_table(doc,
                 ["S. No.",
                  "Standard / Framework used",
                  "Severity Score System Used",
                  "Other References"],
                 rows, col_widths_in=[0.5, 2.0, 1.95, 1.55],
                 repeat_header=True)


def build_tools(doc, data: ReportData, bookmarks: _BookmarkRegistry) -> None:
    """List of tools, script, frameworks, etc. used for carrying out the tests."""
    if not data.tools:
        return
    _add_heading_with_bookmark(
        doc,
        "List of tools, script, frameworks, etc. used for carrying out the tests",
        level=2, bookmarks=bookmarks, bookmark_name="bm_tools",
        page_break_before=True,
    )
    rows = [[t.sno, t.category, t.tools, t.purpose] for t in data.tools]
    _build_table(doc, ["S. No.", "Category", "Tools", "Purpose"], rows,
                 col_widths_in=[0.5, 1.55, 1.95, 2.0], repeat_header=True)


def _vuln_location_combined(v: Vulnerability) -> str:
    """Pack location/parameter/path/link into one cell, suppressing empties."""
    bits: List[str] = []
    if v.vuln_location:
        bits.append(f"Location: {v.vuln_location}")
    if v.vuln_path:
        bits.append(f"Path: {v.vuln_path}")
    if v.vuln_parameter:
        bits.append(f"Parameter: {v.vuln_parameter}")
    if v.vuln_location_full and v.vuln_location_full not in v.vuln_location:
        # Already partially expressed via location; only append if it adds info.
        if not bits:
            bits.append(v.vuln_location_full)
    if not bits and v.vuln_location_full:
        bits.append(v.vuln_location_full)
    return "\n".join(bits)


def build_summary_table(doc, data: ReportData,
                        bookmarks: _BookmarkRegistry) -> None:
    """List of Vulnerable Parameter, Location discovered.

    Columns (per user spec): S.no | Vulnerable location, parameter, path, link |
    Name of the Vulnerability | References (CVE-ID, CWE-ID etc.).
    """
    _add_heading_with_bookmark(
        doc, "List of Vulnerable Parameter, Location discovered",
        level=2, bookmarks=bookmarks, bookmark_name="bm_part1_summary_table",
        page_break_before=True,
    )
    if not data.vulnerabilities:
        _add_para(doc, "No vulnerabilities were reported during this assessment.",
                  italic=True, color_hex="595959")
        return
    rows = [
        [v.sno,
         _vuln_location_combined(v),
         v.vuln_name,
         v.cve_cwe]
        for v in data.vulnerabilities
    ]
    _build_table(doc,
                 ["S. No.",
                  "Vulnerable location, parameter, path, link",
                  "Name of the Vulnerability",
                  "References (CVE-ID, CWE-ID etc.)"],
                 rows, col_widths_in=[0.45, 2.55, 1.85, 1.65],
                 repeat_header=True, header_size_pt=10)


def build_questions(doc, data: ReportData, bookmarks: _BookmarkRegistry) -> None:
    """Questions section.

    Layout rules (per requirements):
      * Whole section starts on a NEW page.
      * Each question (heading + answer) starts on a NEW page.
      * Question heading + answer paragraphs are kept on the **same page**
        (uses ``keepNext`` / ``keepLines`` plus a 1-row keep-together table
        so Word never page-breaks between Q and A).
    """
    if not data.questions:
        return

    section_h = doc.add_heading("Questions", level=1)
    section_h.alignment = WD_ALIGN_PARAGRAPH.LEFT
    _add_bookmark(section_h, "bm_questions", bookmarks.next())
    bookmarks.record_heading(1, "Questions", "bm_questions")
    _set_page_break_before(section_h)

    for idx, q in enumerate(data.questions):
        title = q.question or "(untitled question)"
        heading_text = f"Q{q.qno or (idx + 1)}. {title}".strip()
        bm_name = _slugify(f"q{q.qno or idx + 1}")
        h = doc.add_heading(heading_text, level=2)
        h.alignment = WD_ALIGN_PARAGRAPH.LEFT
        _add_bookmark(h, bm_name, bookmarks.next())
        bookmarks.record_heading(2, heading_text, bm_name)
        if idx > 0:
            _set_page_break_before(h)
        _set_keep_with_next(h)
        _set_keep_lines(h)

        # Wrap the answer in a single-cell, no-border, keep-together table so
        # Word never splits the answer block from its question.
        wrap = doc.add_table(rows=1, cols=1)
        wrap.autofit = True
        _set_cant_split(wrap.rows[0])
        cell = wrap.cell(0, 0)
        # Remove the default Table Grid borders for the wrap cell only.
        tc_pr = cell._tc.get_or_add_tcPr()
        bdrs = OxmlElement("w:tcBorders")
        for edge in ("top", "left", "bottom", "right"):
            b = OxmlElement(f"w:{edge}")
            b.set(qn("w:val"), "nil")
            bdrs.append(b)
        tc_pr.append(bdrs)

        first_para = cell.paragraphs[0]
        first_para.text = ""
        ans_lines = (q.answer or "").split("\n") if q.answer else [""]
        for li, line in enumerate(ans_lines):
            target = first_para if li == 0 else cell.add_paragraph()
            _add_text_run(target, line, size_pt=BODY_FONT_PT,
                          font_name=BODY_FONT_NAME)
            if li < len(ans_lines) - 1:
                _set_keep_with_next(target)
            _set_keep_lines(target)


def _detail_value(v: Vulnerability, key: str, *, idx: int = 0) -> str:
    if key == "_serial_":
        return str(idx) if idx else (v.sno or "")
    if key == "_cvss_epss_combined_":
        parts = []
        if v.cvss_score:
            parts.append(f"CVSS: {v.cvss_score}")
        if v.epss_score:
            parts.append(f"EPSS: {v.epss_score}")
        return " / ".join(parts) if parts else ""
    return getattr(v, key, "") or ""


def _render_vulnerable_point_block_in_cell(cell, v: Vulnerability) -> None:
    """Render a 3-row nested sub-table inside the value cell.

    Layout (matches the user's reference screenshot)::

        Vulnerable Location        | <URL/IP>
        Vulnerable Path/ Port/ URL | <path or port>
        Vulnerable Parameter       | <name or NA>
    """
    cell.text = ""
    # Squash the leading empty paragraph the cell auto-creates so the nested
    # table lines up flush with the parent cell's top border.
    lead = cell.paragraphs[0]
    pf = lead.paragraph_format
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.line_spacing = 1.0

    sub_rows = [
        ("Vulnerable Location",        v.vuln_location or "NA"),
        ("Vulnerable Path/ Port/ URL", v.vuln_path or "NA"),
        ("Vulnerable Parameter",       v.vuln_parameter or "NA"),
    ]

    nested = cell.add_table(rows=len(sub_rows), cols=2)
    try:
        nested.style = "Table Grid"
    except Exception:
        pass
    _set_table_layout_fixed(nested)
    _set_table_full_width(nested)

    for i, (sub_label, sub_value) in enumerate(sub_rows):
        lc = nested.cell(i, 0)
        vc = nested.cell(i, 1)
        lc.text = ""
        vc.text = ""

        lp = lc.paragraphs[0]
        _add_text_run(lp, sub_label, bold=True, color_hex=TABLE_LABEL_TEXT,
                      size_pt=BODY_FONT_PT, font_name=BODY_FONT_NAME)
        _set_cell_shading(lc, TABLE_LABEL_FILL)
        _set_cell_borders(lc)

        vp = vc.paragraphs[0]
        _add_text_run(vp, sub_value or "NA",
                      size_pt=BODY_FONT_PT, font_name=BODY_FONT_NAME)
        _set_cell_borders(vc)
        _set_cell_shading(vc, TABLE_BODY_FILL)

        # Sub-table column widths: ~40 / 60 split inside the parent value cell.
        lc.width = Inches(1.5)
        vc.width = Inches(2.3)


def _render_poc_steps_in_cell(cell, poc_text: str, images: List[Path],
                              work_dir: Path,
                              *, image_width_in: float = 3.6) -> None:
    """Render the PoC value cell as numbered ``Step N: <text> <image>`` blocks.

    * Step text comes from the workbook's ``poc_steps_text`` column, split
      into one step per non-empty line.
    * Screenshots come from the corresponding POC folder, already sorted
      numerically by filename (``1.png``, ``2.png``, …).
    * If the number of text lines and images differs, the longer of the two
      drives the step count and the shorter one is padded transparently
      (text-only or image-only steps).
    """
    text_lines = [ln.strip() for ln in (poc_text or "").split("\n") if ln.strip()]
    n_steps = max(len(text_lines), len(images))

    first_para = cell.paragraphs[0]
    first_para.text = ""

    if n_steps == 0:
        _add_text_run(first_para, "(No proof-of-concept supplied)",
                      italic=True, color_hex="7F7F7F",
                      size_pt=10, font_name=BODY_FONT_NAME)
        return

    use_first = True
    for i in range(n_steps):
        step_text = text_lines[i] if i < len(text_lines) else ""
        target = first_para if use_first else cell.add_paragraph()
        use_first = False
        _add_text_run(target, f"Step {i + 1}: ", bold=True,
                      size_pt=BODY_FONT_PT, font_name=BODY_FONT_NAME)
        _add_text_run(target, step_text or "(see screenshot below)",
                      size_pt=BODY_FONT_PT, font_name=BODY_FONT_NAME)

        if i < len(images):
            img = images[i]
            prepared = safe_image_for_docx(img, work_dir)
            ip = cell.add_paragraph()
            ip.alignment = WD_ALIGN_PARAGRAPH.CENTER
            if prepared is None:
                _add_text_run(ip, f"[Screenshot not embedded: {img.name}]",
                              italic=True, color_hex="C00000", size_pt=10,
                              font_name=BODY_FONT_NAME)
            else:
                try:
                    run = ip.add_run()
                    run.add_picture(str(prepared), width=Inches(image_width_in))
                except Exception as ex:
                    log.warning("add_picture failed for %s: %s", img.name, ex)
                    _add_text_run(ip, f"[Failed to embed: {img.name}]",
                                  italic=True, color_hex="C00000", size_pt=10,
                                  font_name=BODY_FONT_NAME)


def build_detail_report(doc, data: ReportData, poc: PocArchive,
                        work_dir: Path,
                        bookmarks: _BookmarkRegistry) -> None:
    """Part 2: Detailed Report.

    The whole section starts on a new page, and every individual vulnerability
    detail block also starts on its own new page.
    """
    section_h = doc.add_heading("Part 2: Detailed Report", level=1)
    section_h.alignment = WD_ALIGN_PARAGRAPH.LEFT
    _add_bookmark(section_h, "bm_part2_detail", bookmarks.next())
    bookmarks.record_heading(1, "Part 2: Detailed Report", "bm_part2_detail")
    _set_page_break_before(section_h)

    if not data.vulnerabilities:
        _add_para(doc, "No detailed vulnerability findings.",
                  italic=True, color_hex="595959")
        return

    for idx, v in enumerate(data.vulnerabilities, start=1):
        # Per user request: no Heading-2 per vulnerability (S. No. in table is enough).
        # Still force each vulnerability to start on its own new page.
        if idx > 1:
            pb = doc.add_paragraph()
            pb.add_run().add_break(WD_BREAK.PAGE)

        # Detail table layout to match reference screenshot:
        # - 3 columns overall: [main label] [sub-label] [value]
        # - For normal rows, columns 2+3 are merged (so it looks like 2 columns).
        # - For the Vulnerable-Point block, the main label spans 3 rows, and the
        #   sub-label + value stay split (to show inner grid lines).
        #
        # Row plan (10 logical rows):
        #   1  S. No.
        #   2  Name of Vulnerability
        #   3-5 Vulnerable Point / Error Location (3 sub-rows)
        #   6  CVE/CWE
        #   7  CVSS / EPSS Score
        #   8  Description
        #   9  Proof of concept and Steps to Reproduce
        #   10 Workaround / Solutions / Recommendations
        #   11 References (Optional)
        #   12 Additional observations
        #
        # (We implement as 12 rows with the vulnerable-point block consuming 3.)
        table = doc.add_table(rows=12, cols=3)
        try:
            table.style = "Table Grid"
        except KeyError:
            pass
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.autofit = False
        _set_table_layout_fixed(table)
        _set_table_full_width(table, 6.5)

        # Fixed column widths: label / sublabel / value
        for row in table.rows:
            row.cells[0].width = Inches(2.2)
            row.cells[1].width = Inches(1.6)
            row.cells[2].width = Inches(2.7)

        def _set_label_cell(cell, text: str) -> None:
            cell.text = ""
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            _add_text_run(p, text, bold=True, color_hex=TABLE_LABEL_TEXT,
                          size_pt=BODY_FONT_PT, font_name=BODY_FONT_NAME)
            _set_cell_shading(cell, TABLE_LABEL_FILL)
            _set_cell_borders(cell)

        def _set_sub_label_cell(cell, text: str) -> None:
            cell.text = ""
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            _add_text_run(p, text, bold=True, color_hex=TABLE_LABEL_TEXT,
                          size_pt=BODY_FONT_PT, font_name=BODY_FONT_NAME)
            _set_cell_shading(cell, TABLE_LABEL_FILL)
            _set_cell_borders(cell)

        def _set_value_cell(cell, text: str) -> None:
            cell.text = ""
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            first = True
            for line in (text or "").split("\n"):
                if first:
                    _add_text_run(p, line, size_pt=BODY_FONT_PT,
                                  font_name=BODY_FONT_NAME)
                    first = False
                else:
                    pp = cell.add_paragraph()
                    _add_text_run(pp, line, size_pt=BODY_FONT_PT,
                                  font_name=BODY_FONT_NAME)
            _set_cell_borders(cell)
            _set_cell_shading(cell, TABLE_BODY_FILL)

        def _merge_value_cols(r: int) -> Any:
            # Merge col1+col2 (0-based: 1 and 2) into one wide value cell.
            c = table.cell(r, 1).merge(table.cell(r, 2))
            _set_cell_borders(c)
            _set_cell_shading(c, TABLE_BODY_FILL)
            return c

        # Row 0: S. No.
        _set_label_cell(table.cell(0, 0), "S. No.")
        vcell = _merge_value_cols(0)
        _set_value_cell(vcell, str(idx))

        # Row 1: Name of Vulnerability
        _set_label_cell(table.cell(1, 0), "Name of Vulnerability:")
        vcell = _merge_value_cols(1)
        _set_value_cell(vcell, v.vuln_name or "")

        # Rows 2-4: Vulnerable Point / Error Location (merge main label vertically)
        main = table.cell(2, 0)
        main2 = table.cell(3, 0)
        main3 = table.cell(4, 0)
        main.merge(main2).merge(main3)
        _set_label_cell(main, "Vulnerable Point /\nError Location")
        _set_cell_v_align(main, "center")

        sub_rows = [
            ("Vulnerable Location",        v.vuln_location or "NA"),
            ("Vulnerable Path/ Port/\nURL", v.vuln_path or "NA"),
            ("Vulnerable Parameter",       v.vuln_parameter or "NA"),
        ]
        for off, (sl, sv) in enumerate(sub_rows):
            r = 2 + off
            _set_sub_label_cell(table.cell(r, 1), sl)
            _set_value_cell(table.cell(r, 2), sv)

        # Row 5: CVE/CWE
        _set_label_cell(table.cell(5, 0), "CVE/CWE")
        vcell = _merge_value_cols(5)
        _set_value_cell(vcell, v.cve_cwe or "")

        # Row 6: CVSS / EPSS Score
        _set_label_cell(table.cell(6, 0), "CVSS / EPSS\nScore")
        vcell = _merge_value_cols(6)
        _set_value_cell(vcell, _detail_value(v, "_cvss_epss_combined_", idx=idx))

        # Row 7: Description (can split)
        _set_label_cell(table.cell(7, 0), "Description")
        vcell = _merge_value_cols(7)
        _set_value_cell(vcell, v.description or "")

        # Row 8: Proof of concept and Steps to Reproduce (can split)
        _set_label_cell(table.cell(8, 0), "Proof of concept and\nSteps to Reproduce")
        vcell = _merge_value_cols(8)
        images = poc.images_for(v.poc_folder) if v.poc_folder else []
        _render_poc_steps_in_cell(vcell, v.poc_steps_text, images, work_dir / "_img")
        _set_cell_borders(vcell)
        _set_cell_shading(vcell, TABLE_BODY_FILL)

        # Row 9: Workaround / Solutions / Recommendations
        _set_label_cell(table.cell(9, 0), "Workaround /\nSolutions /\nRecommendations")
        vcell = _merge_value_cols(9)
        _set_value_cell(vcell, v.workaround or "")

        # Row 10: References (Optional)
        _set_label_cell(table.cell(10, 0), "References\n(Optional)")
        vcell = _merge_value_cols(10)
        _set_value_cell(vcell, v.references or "")

        # Row 11: Additional observations
        _set_label_cell(table.cell(11, 0), "Additional\nobservations")
        vcell = _merge_value_cols(11)
        _set_value_cell(vcell, v.additional_obs or "")

        # Row splitting rules: keep short rows together; allow Description/PoC to split.
        SPLIT_ALLOWED_ROWS = {7, 8}
        for ri, row in enumerate(table.rows):
            if ri not in SPLIT_ALLOWED_ROWS:
                _set_cant_split(row)

        # extra metadata footer line
        meta_bits = []
        if v.status:
            meta_bits.append(f"Status: {v.status}")
        if v.new_or_repeat:
            meta_bits.append(f"Type: {v.new_or_repeat}")
        if meta_bits:
            mp = doc.add_paragraph()
            _add_text_run(mp, "  •  ".join(meta_bits), italic=True,
                          color_hex="595959", size_pt=10,
                          font_name=BODY_FONT_NAME)


def build_poc_map_section(doc, data: ReportData,
                          bookmarks: _BookmarkRegistry) -> None:
    if not data.poc_map:
        return
    h = doc.add_heading("Annex A: POC Manifest", level=1)
    _add_bookmark(h, "bm_annex_pocmap", bookmarks.next())
    _set_page_break_before(h)
    rows = [
        [m.sno, m.vuln_name, m.poc_folder, str(m.screenshot_count),
         m.expected_files, m.notes]
        for m in data.poc_map
    ]
    _build_table(doc,
                 ["S. No.", "Vulnerability", "Folder",
                  "Screenshots", "Expected Files", "Notes"],
                 rows, col_widths_in=[0.4, 2.0, 0.7, 0.65, 1.15, 1.35],
                 zebra=True, repeat_header=True)


def build_instructions_annex(doc, data: ReportData,
                             bookmarks: _BookmarkRegistry) -> None:
    if not data.instructions_text:
        return
    h = doc.add_heading("Annex B: Author Instructions", level=1)
    _add_bookmark(h, "bm_annex_instructions", bookmarks.next())
    _set_page_break_before(h)
    for line in data.instructions_text.split("\n"):
        _add_para(doc, line, size_pt=10)


def _try_refresh_word_fields(doc_path: Path) -> bool:
    """Paginate + refresh Word fields so ``PAGEREF`` resolves to real pages.

    Pagination is impossible inside python-docx alone. After saving, we optionally
    invoke Microsoft Word via COM (Windows + ``pywin32`` + locally installed
    Word) so TOC page numbers match the printed layout **before** delivery.

    Disable with ``MITKAT2_SKIP_WORD_COM=1`` if automation is undesirable.
    """
    if os.environ.get("MITKAT2_SKIP_WORD_COM", "").strip().lower() in {"1", "true", "yes"}:
        log.info("Word COM refresh disabled via MITKAT2_SKIP_WORD_COM=1")
        return False
    if sys.platform != "win32":
        log.info("Word COM field refresh not available on non-Windows platform (%s)", sys.platform)
        return False
    try:
        import win32com.client  # type: ignore
    except ImportError:
        log.info(
            "Optional: ``pip install pywin32`` so TOC PAGEREF fields refresh "
            "automatically after generation (requires Microsoft Word on Windows). "
            "TOC will still be visible when opened in Word, but page numbers may show as 1 "
            "until user enables editing or opens in Print Preview."
        )
        return False

    word = None
    wd_doc = None
    abs_path = str(doc_path.resolve())
    try:
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        wd_doc = word.Documents.Open(
            abs_path, ConfirmConversions=False, ReadOnly=False, AddToRecentFiles=False)
        try:
            # Force print layout view for pagination
            wd_doc.ActiveWindow.View.Type = 3  # wdPrintView
        except Exception:
            pass
        # Repaginate and update ALL fields (including PAGEREF)
        wd_doc.Repaginate()
        wd_doc.Fields.Update()
        wd_doc.Save()
        log.info("✓ Word refreshed TOC page numbers and fields: %s", doc_path.name)
        return True
    except Exception as ex:
        log.warning("⚠ Word COM field refresh failed (TOC page numbers may show 1 until user opens in edit/print preview): %s", ex)
        return False
    finally:
        try:
            if wd_doc is not None:
                wd_doc.Close(SaveChanges=False)
        except Exception:
            pass
        try:
            if word is not None:
                word.Quit(SaveChanges=False)
        except Exception:
            pass


# ---------------------------------------------------------------------------
# Orchestrator
# ---------------------------------------------------------------------------

@dataclass
class GenerateOptions:
    include_instructions: bool = RL_INSERT_INSTRUCTIONS_AS_ANNEX
    include_poc_map: bool = False         # Annex A removed per user request
    fail_on_warnings: bool = False


def generate_report(xlsx_path: str | os.PathLike,
                    poc_zip_path: Optional[str | os.PathLike],
                    output_path: str | os.PathLike,
                    *,
                    options: Optional[GenerateOptions] = None,
                    work_dir: Optional[str | os.PathLike] = None) -> Dict[str, Any]:
    """End-to-end generation. Returns a dict report (for tests/CLI)."""
    options = options or GenerateOptions()
    started = time.time()
    out = Path(output_path)
    out.parent.mkdir(parents=True, exist_ok=True)
    cleanup_dir: Optional[Path] = None
    if work_dir is None:
        cleanup_dir = Path(tempfile.mkdtemp(prefix="mitkat2_"))
        wd = cleanup_dir
    else:
        wd = Path(work_dir)
        wd.mkdir(parents=True, exist_ok=True)

    log.info("Mitkat2 generation start: xlsx=%s zip=%s out=%s",
             xlsx_path, poc_zip_path, output_path)
    try:
        data = load_report_data(xlsx_path)
        poc: PocArchive
        if poc_zip_path:
            poc = load_poc_archive(poc_zip_path, wd)
        else:
            poc = PocArchive(root=wd, folders={})

        issues = validate_report(data, poc)
        errors = [i for i in issues if i.severity == "error"]
        warnings = [i for i in issues if i.severity != "error"]
        if errors:
            raise Mitkat2Error("Validation failed: " + "; ".join(
                f"{i.code}@row{i.row}" for i in errors[:10]))

        doc = build_base_document(data.meta)
        bookmarks = _BookmarkRegistry()

        # Page 1 is now the static 1.docx cover — build_cover() is NOT called here.
        # The cover is prepended after doc.save() via _prepend_static_cover_page().

        # Page 2 (after static cover): Contents (functional TOC field)
        build_toc(doc, bookmarks)

        # Part 1 — section order mirrors the user-provided new.png TOC:
        #   Executive Summary  (page break)
        #     Overview                                    (no page break - same page as Part 1)
        #     List of technical manpower ...              (page break)
        #     List of Standards, Frameworks ...           (page break)
        #     List of Vulnerable Parameter, Location ...  (page break)
        #     List of tools, script, frameworks ...       (page break)
        build_executive_summary(doc, data, bookmarks)
        build_manpower(doc, data, bookmarks)
        build_frameworks(doc, data, bookmarks)
        build_summary_table(doc, data, bookmarks)
        build_tools(doc, data, bookmarks)

        # Questions (whole section new page; each Q after Q1 new page; Q+A together).
        build_questions(doc, data, bookmarks)

        # Part 2: Detailed Report (whole section new page, each vuln new page).
        build_detail_report(doc, data, poc, wd, bookmarks)

        # Annex A (POC manifest) removed - user does not want it in the report.
        if options.include_instructions:
            build_instructions_annex(doc, data, bookmarks)

        # All headings have been emitted - inject the static, hyperlinked TOC.
        _finalize_toc(doc, bookmarks)

        doc.save(str(out))

        # Prepend the static 1.docx cover page to the front of the saved report.
        _prepend_static_cover_page(out, data.meta)

        # Force Word field updates TWICE: once for initial doc, once after cover prepend
        # (page numbers change after cover is prepended).
        refreshed1 = _try_refresh_word_fields(out)
        refreshed = refreshed1

        elapsed = time.time() - started

        report = {
            "ok": True,
            "output": str(out),
            "elapsed_seconds": round(elapsed, 3),
            "vuln_count": len(data.vulnerabilities),
            "manpower_count": len(data.manpower),
            "frameworks_count": len(data.frameworks),
            "tools_count": len(data.tools),
            "questions_count": len(data.questions),
            "exec_blocks": len(data.exec_blocks),
            "poc_folders": len(poc.folders),
            "warnings": [asdict(w) for w in warnings],
            "severity_counts": data.severity_counts(),
            "word_fields_refreshed": refreshed,
        }
        log.info("Mitkat2 generation complete in %.2fs (vulns=%d)",
                 elapsed, report["vuln_count"])
        return report
    finally:
        if cleanup_dir is not None:
            shutil.rmtree(cleanup_dir, ignore_errors=True)


def _mitkat2_web_page_html() -> str:
    """Single-page tester UI (embedded in this module — no external static files)."""
    return """<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8"/>
<meta name="viewport" content="width=device-width, initial-scale=1"/>
<title>MitKat2 — Test Console</title>
<style>
  :root { --bg:#0f1419; --panel:#1a2332; --text:#e7ecf3; --muted:#8b9cae;
         --accent:#0f4761; --ok:#2e7d32; --warn:#ed6c02; --err:#c62828; --border:#2d3d52; }
  * { box-sizing: border-box; }
  body { font-family: Calibri, "Segoe UI", sans-serif; margin:0; background:var(--bg); color:var(--text);
         line-height:1.45; }
  header { background:var(--panel); border-bottom:1px solid var(--border); padding:1rem 1.25rem; }
  header h1 { margin:0; font-size:1.35rem; color:#fff; }
  header p { margin:0.35rem 0 0; font-size:0.9rem; color:var(--muted); }
  main { max-width:920px; margin:0 auto; padding:1.25rem; }
  section { background:var(--panel); border:1px solid var(--border); border-radius:8px; padding:1rem 1.25rem; margin-bottom:1rem; }
  section h2 { margin:0 0 0.75rem; font-size:1.05rem; color:#fff; }
  label { display:block; margin:0.5rem 0 0.25rem; font-size:0.85rem; color:var(--muted); }
  input[type=file] { width:100%; padding:0.4rem; background:#0d1218; border:1px solid var(--border); color:var(--text); border-radius:4px; }
  .row { display:flex; flex-wrap:wrap; gap:0.6rem; align-items:center; margin-top:0.75rem; }
  button { cursor:pointer; padding:0.55rem 1rem; border-radius:6px; border:none; font-weight:600; font-size:0.9rem; }
  button.primary { background:var(--accent); color:#fff; }
  button.secondary { background:#243044; color:#fff; border:1px solid var(--border); }
  button:disabled { opacity:0.45; cursor:not-allowed; }
  pre { background:#0d1218; border:1px solid var(--border); border-radius:6px; padding:0.85rem; overflow:auto; max-height:420px;
        font-size:0.78rem; white-space:pre-wrap; word-break:break-word; }
  .pill { display:inline-block; padding:0.15rem 0.5rem; border-radius:4px; font-size:0.75rem; margin-right:0.35rem; }
  .pill.ok { background:#1b3320; color:#81c784; }
  .pill.bad { background:#3d1818; color:#ef9a9a; }
  .hint { font-size:0.8rem; color:var(--muted); margin-top:0.35rem; }
  a { color:#64b5f6; }
</style>
</head>
<body>
<header>
  <h1>MitKat2 — Test console</h1>
  <p>Validate Excel + optional POC zip, fetch health JSON, or generate a .docx (same logic as production API).</p>
</header>
<main>
  <section>
    <h2>Files</h2>
    <label for="excel">MitKat workbook (.xlsx) — required</label>
    <input id="excel" type="file" accept=".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"/>
    <label for="poc">POC archive (.zip) — optional</label>
    <input id="poc" type="file" accept=".zip,application/zip"/>
    <label style="display:flex;align-items:center;gap:0.5rem;margin-top:0.75rem;">
      <input id="instr" type="checkbox"/> Include INSTRUCTIONS sheet as Annex B in Word output
    </label>
    <p class="hint">Uses endpoints under <code id="epPrefix"></code>. Open DevTools (F12) if downloads fail.</p>
  </section>
  <section>
    <h2>Actions</h2>
    <div class="row">
      <button type="button" class="secondary" id="btnHealth">GET /health</button>
      <button type="button" class="secondary" id="btnVal">POST /validate</button>
      <button type="button" class="primary" id="btnGen">POST /generate (.docx)</button>
    </div>
    <p class="hint" id="status"></p>
  </section>
  <section>
    <h2>Response</h2>
    <div id="meta"></div>
    <pre id="out">(no response yet)</pre>
  </section>
</main>
<script>
(function(){
  var PREFIX = "/mitkat2";
  document.getElementById("epPrefix").textContent = PREFIX + "/*";
  var statusEl = document.getElementById("status");
  var outEl = document.getElementById("out");
  var metaEl = document.getElementById("meta");
    var CSRF_TOKEN = null;

  function setStatus(msg, ok) {
    statusEl.textContent = msg || "";
    statusEl.style.color = ok === true ? "#81c784" : (ok === false ? "#ef9a9a" : "var(--muted)");
  }
  function showMeta(ok, extra) {
    metaEl.innerHTML = "";
    var span = document.createElement("span");
    span.className = "pill " + (ok ? "ok" : "bad");
    span.textContent = ok === undefined ? "…" : (ok ? "OK" : "Error");
    metaEl.appendChild(span);
    if (extra) metaEl.appendChild(document.createTextNode(" " + extra));
  }

    async function ensureCsrfToken() {
        if (CSRF_TOKEN) return CSRF_TOKEN;
        try {
            var res = await fetch("/csrf-token", { credentials: "same-origin" });
            if (res.ok) {
                var data = await res.json();
                CSRF_TOKEN = data.csrf_token || data.token || data.csrf || null;
            }
        } catch (e) {
            // No-op: CSRF token may be unavailable in some test contexts.
        }
        return CSRF_TOKEN;
    }

  document.getElementById("btnHealth").onclick = async function() {
    setStatus("Loading…", undefined); showMeta(undefined, "");
    try {
      var r = await fetch(PREFIX + "/health");
      var j = await r.json();
      outEl.textContent = JSON.stringify(j, null, 2);
      setStatus("HTTP " + r.status, r.ok);
      showMeta(r.ok, "version " + (j.version || "?"));
    } catch (e) {
      outEl.textContent = String(e);
      setStatus("Failed", false); showMeta(false, "");
    }
  };

  document.getElementById("btnVal").onclick = async function() {
    var f = document.getElementById("excel").files[0];
    if (!f) { alert("Choose an Excel file."); return; }
        await ensureCsrfToken();
    var fd = new FormData();
    fd.append("excel", f);
    var p = document.getElementById("poc").files[0];
    if (p) fd.append("poc", p);
        if (CSRF_TOKEN) fd.append("csrf_token", CSRF_TOKEN);
    setStatus("Validating…", undefined); showMeta(undefined, "");
    try {
            var headers = {};
            if (CSRF_TOKEN) headers["X-CSRF-Token"] = CSRF_TOKEN;
            var r = await fetch(PREFIX + "/validate", {
                method: "POST",
                body: fd,
                headers: headers,
                credentials: "same-origin"
            });
      var txt = await r.text();
      try { outEl.textContent = JSON.stringify(JSON.parse(txt), null, 2); }
      catch (_) { outEl.textContent = txt; }
      setStatus("HTTP " + r.status, r.ok);
      var okj = false;
      try { okj = JSON.parse(txt).ok; } catch(_){}
      showMeta(okj, "issues — see JSON");
    } catch (e) {
      outEl.textContent = String(e);
      setStatus("Failed", false); showMeta(false, "");
    }
  };

  document.getElementById("btnGen").onclick = async function() {
    var f = document.getElementById("excel").files[0];
    if (!f) { alert("Choose an Excel file."); return; }
        await ensureCsrfToken();
    var fd = new FormData();
    fd.append("excel", f);
    var p = document.getElementById("poc").files[0];
    if (p) fd.append("poc", p);
    fd.append("include_instructions", document.getElementById("instr").checked ? "1" : "0");
        if (CSRF_TOKEN) fd.append("csrf_token", CSRF_TOKEN);
    setStatus("Generating Word…", undefined); showMeta(undefined, "");
    try {
            var headers = {};
            if (CSRF_TOKEN) headers["X-CSRF-Token"] = CSRF_TOKEN;
            var r = await fetch(PREFIX + "/generate", {
                method: "POST",
                body: fd,
                headers: headers,
                credentials: "same-origin"
            });
      if (!r.ok) {
        var errTxt = await r.text();
        try { outEl.textContent = JSON.stringify(JSON.parse(errTxt), null, 2); }
        catch (_) { outEl.textContent = errTxt; }
        setStatus("HTTP " + r.status, false); showMeta(false, "");
        return;
      }
      var blob = await r.blob();
      outEl.textContent = "[Binary .docx received — " + blob.size + " bytes — downloading…]\\n\\n" +
        "Headers: vulns=" + (r.headers.get("x-mitkat2-vulns")||"?") +
        ", elapsed_s=" + (r.headers.get("x-mitkat2-elapsed")||"?") +
        ", warnings=" + (r.headers.get("x-mitkat2-warnings")||"?");
      var url = URL.createObjectURL(blob);
      var a = document.createElement("a");
      a.href = url;
      a.download = "Mitkat2_Report.docx";
      document.body.appendChild(a);
      a.click();
      a.remove();
      URL.revokeObjectURL(url);
      setStatus("Download started.", true); showMeta(true, blob.size + " bytes");
    } catch (e) {
      outEl.textContent = String(e);
      setStatus("Failed", false); showMeta(false, "");
    }
  };
})();
</script>
</body>
</html>
"""


# ---------------------------------------------------------------------------
# FastAPI router (only constructed if FastAPI is available)
# ---------------------------------------------------------------------------

if _FASTAPI_OK:
    mitkat2_router = APIRouter(prefix="/mitkat2", tags=["mitkat2"])

    @mitkat2_router.get("/ui", response_class=HTMLResponse)  # type: ignore[arg-type]
    async def mitkat2_test_ui() -> Any:
        """Embedded browser tester — open `/mitkat2/ui` when using `--serve`."""
        return HTMLResponse(content=_mitkat2_web_page_html())

    @mitkat2_router.get("/", response_class=RedirectResponse)  # type: ignore[arg-type]
    async def mitkat2_router_root_redirect():
        return RedirectResponse(url="/mitkat2/ui")

    @mitkat2_router.get("/health")
    async def mitkat2_health() -> Dict[str, Any]:
        return {
            "ok": True,
            "version": __version__,
            "content_fidelity": {
                "policy": "deterministic_source_only",
                "sources_of_truth": [
                    "excel_cell_values",
                    "poc_zip_images_and_filenames",
                ],
                "never_synthesize": [
                    "vulnerability_findings",
                    "cve_ids_not_present_in_workbook",
                    "inferred_exploit_steps",
                ],
                "checks": [
                    "draft_placeholder_scan ([Insert]/[TBD]/…)",
                    "auto_calculated_count_substitution_docstring",
                ],
            },
            "resource_limits": {
                "max_excel_bytes": RL_MAX_EXCEL_BYTES,
                "max_zip_bytes": RL_MAX_ZIP_BYTES,
                "max_unzip_bytes": RL_MAX_UNZIP_BYTES,
                "max_zip_compression_ratio": RL_MAX_ZIP_RATIO,
                "max_total_images_indexed": RL_MAX_TOTAL_IMAGES,
                "soft_warn_total_images": RL_SOFT_WARN_TOTAL_IMAGES,
                "truncate_poc_images": RL_TRUNCATE_POC_IMAGES,
                "max_image_width_px_resize": RL_MAX_IMAGE_WIDTH_PX,
                "generate_timeout_s_documentation_only": RL_GENERATE_TIMEOUT_S,
            },
            "pillow": _PIL_OK,
        }

    async def _read_upload_with_limit(upload: UploadFile, limit: int,
                                      label: str) -> bytes:
        chunk_size = 1024 * 1024
        total = 0
        buf = bytearray()
        while True:
            chunk = await upload.read(chunk_size)
            if not chunk:
                break
            total += len(chunk)
            if total > limit:
                raise HTTPException(
                    status_code=413,
                    detail={"code": f"{label}_too_large",
                            "message": f"{label} exceeds maximum allowed size",
                            "limit": limit, "actual": total},
                )
            buf.extend(chunk)
        return bytes(buf)

    @mitkat2_router.post("/validate")
    async def mitkat2_validate(
        excel: UploadFile = File(...),
        poc: Optional[UploadFile] = File(None),
    ) -> JSONResponse:
        tmp = Path(tempfile.mkdtemp(prefix="mitkat2_validate_"))
        try:
            xls_path = tmp / (excel.filename or "report.xlsx")
            xls_bytes = await _read_upload_with_limit(excel, RL_MAX_EXCEL_BYTES, "excel")
            xls_path.write_bytes(xls_bytes)

            poc_archive: Optional[PocArchive] = None
            if poc is not None:
                zip_path = tmp / (poc.filename or "poc.zip")
                zip_bytes = await _read_upload_with_limit(poc, RL_MAX_ZIP_BYTES, "zip")
                zip_path.write_bytes(zip_bytes)
                try:
                    poc_archive = load_poc_archive(zip_path, tmp)
                except GuardrailError as ge:
                    return JSONResponse(status_code=413, content={"ok": False,
                                                                  "error": ge.to_dict()})
            try:
                data = load_report_data(xls_path)
            except GuardrailError as ge:
                return JSONResponse(status_code=413, content={"ok": False,
                                                              "error": ge.to_dict()})
            except Mitkat2Error as me:
                return JSONResponse(status_code=422, content={"ok": False,
                                                              "error": str(me)})

            issues = validate_report(data, poc_archive)
            return JSONResponse(content={
                "ok": not any(i.severity == "error" for i in issues),
                "vuln_count": len(data.vulnerabilities),
                "severity_counts": data.severity_counts(),
                "poc_folders": list((poc_archive.folders.keys()
                                     if poc_archive else [])),
                "issues": [asdict(i) for i in issues],
            })
        finally:
            shutil.rmtree(tmp, ignore_errors=True)

    @mitkat2_router.post("/generate")
    async def mitkat2_generate(
        background_tasks: BackgroundTasks,
        excel: UploadFile = File(...),
        poc: Optional[UploadFile] = File(None),
        include_instructions: int = 0,
    ) -> Any:
        tmp = Path(tempfile.mkdtemp(prefix="mitkat2_run_"))
        try:
            xls_path = tmp / (excel.filename or "report.xlsx")
            xls_bytes = await _read_upload_with_limit(excel, RL_MAX_EXCEL_BYTES, "excel")
            xls_path.write_bytes(xls_bytes)

            zip_path: Optional[Path] = None
            if poc is not None:
                zip_path = tmp / (poc.filename or "poc.zip")
                zip_bytes = await _read_upload_with_limit(poc, RL_MAX_ZIP_BYTES, "zip")
                zip_path.write_bytes(zip_bytes)

            out_path = tmp / "Mitkat2_Report.docx"
            options = GenerateOptions(
                include_instructions=bool(include_instructions),
                include_poc_map=True,
            )
            try:
                report = generate_report(xls_path, zip_path, out_path,
                                         options=options, work_dir=tmp)
            except GuardrailError as ge:
                return JSONResponse(status_code=413,
                                    content={"ok": False, "error": ge.to_dict()})
            except Mitkat2Error as me:
                return JSONResponse(status_code=422,
                                    content={"ok": False, "error": str(me)})

            # Schedule cleanup AFTER FileResponse streams the body.
            background_tasks.add_task(shutil.rmtree, str(tmp), True)
            filename = (
                f"{(report.get('output') or 'Mitkat2_Report')}").rsplit("\\", 1)[-1].rsplit("/", 1)[-1]
            return FileResponse(
                path=out_path,
                filename=filename,
                media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                headers={
                    "X-Mitkat2-Vulns": str(report["vuln_count"]),
                    "X-Mitkat2-Elapsed": str(report["elapsed_seconds"]),
                    "X-Mitkat2-Warnings": str(len(report["warnings"])),
                },
            )
        except HTTPException:
            shutil.rmtree(tmp, ignore_errors=True)
            raise
        except Exception as ex:
            shutil.rmtree(tmp, ignore_errors=True)
            log.error("Mitkat2 /generate failure: %s\n%s", ex, traceback.format_exc())
            raise HTTPException(status_code=500, detail=f"Mitkat2 internal error: {ex}")

else:  # pragma: no cover
    mitkat2_router = None  # type: ignore[assignment]


def create_standalone_app():
    """Minimal FastAPI app for local testing: ``python mitkat2.py --serve``.

    Mounts only ``mitkat2_router`` and redirects ``GET /`` → ``/mitkat2/ui``.
    """
    if not _FASTAPI_OK:
        raise RuntimeError("FastAPI is required — pip install fastapi uvicorn")
    from fastapi import FastAPI
    from fastapi.middleware.cors import CORSMiddleware

    app = FastAPI(title="MitKat2 Test Server", version=__version__)
    app.add_middleware(
        CORSMiddleware,
        allow_origins=["*"],
        allow_credentials=True,
        allow_methods=["*"],
        allow_headers=["*"],
    )
    app.include_router(mitkat2_router)  # type: ignore[arg-type]

    @app.get("/")
    async def _browser_root():
        return RedirectResponse(url="/mitkat2/ui")

    return app


def _serve_web(host: str, port: int) -> None:
    import uvicorn

    app = create_standalone_app()
    log.info("MitKat2 test server — open http://%s:%s/mitkat2/ui in your browser", host, port)
    uvicorn.run(app, host=host, port=port, log_level="info")

def _cli() -> int:
    import argparse

    ap = argparse.ArgumentParser(prog="mitkat2", description="MitKat2 report generator")
    ap.add_argument(
        "--serve", action="store_true",
        help="Run standalone web server + embedded browser UI (Validate / Generate / Health)",
    )
    ap.add_argument("--host", default=os.environ.get("MITKAT2_HOST", "127.0.0.1"))
    ap.add_argument("--port", type=int, default=int(os.environ.get("MITKAT2_PORT", "8765")))
    ap.add_argument("--excel", help="Path to MitKat workbook (.xlsx)")
    ap.add_argument("--poc", default=None, help="Path to POC.zip (optional)")
    ap.add_argument("--out", help="Output .docx path")
    ap.add_argument("--include-instructions", action="store_true")
    ap.add_argument("--validate-only", action="store_true")
    args = ap.parse_args()

    if args.serve:
        if not _FASTAPI_OK:
            print("Error: FastAPI not available — pip install fastapi uvicorn", file=sys.stderr)
            return 1
        try:
            _serve_web(args.host, args.port)
        except KeyboardInterrupt:
            print("", file=sys.stderr)
            return 0
        return 0

    if args.validate_only:
        if not args.excel:
            ap.error("--excel is required with --validate-only")
        try:
            data = load_report_data(args.excel)
            poc_arch = None
            if args.poc:
                tmp = Path(tempfile.mkdtemp(prefix="mitkat2_cli_"))
                try:
                    poc_arch = load_poc_archive(args.poc, tmp)
                    issues = validate_report(data, poc_arch)
                finally:
                    shutil.rmtree(tmp, ignore_errors=True)
            else:
                issues = validate_report(data, None)
            print(json.dumps({
                "vuln_count": len(data.vulnerabilities),
                "severity_counts": data.severity_counts(),
                "issues": [asdict(i) for i in issues],
            }, indent=2))
            return 0 if not any(i.severity == "error" for i in issues) else 2

        except GuardrailError as ge:
            print(json.dumps({"ok": False, "error": ge.to_dict()}, indent=2),
                  file=sys.stderr)
            return 3
        except Mitkat2Error as me:
            print(json.dumps({"ok": False, "error": str(me)}, indent=2),
                  file=sys.stderr)
            return 2

    if not args.excel or not args.out:
        ap.error("--excel and --out are required unless you use --serve")

    try:
        report = generate_report(
            args.excel,
            args.poc,
            args.out,
            options=GenerateOptions(include_instructions=bool(args.include_instructions)),
        )
        print(json.dumps(report, indent=2))
        return 0
    except GuardrailError as ge:
        print(json.dumps({"ok": False, "error": ge.to_dict()}, indent=2),
              file=sys.stderr)
        return 3
    except Mitkat2Error as me:
        print(json.dumps({"ok": False, "error": str(me)}, indent=2),
              file=sys.stderr)
        return 2
    except Exception as ex:
        traceback.print_exc()
        print(json.dumps({"ok": False, "error": f"unhandled: {ex}"}, indent=2),
              file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(_cli())

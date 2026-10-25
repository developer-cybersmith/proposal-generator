#!/usr/bin/env python3
"""
CSS Proposal Generator — standalone single-file module.

Does not modify app.py or any other project files. Run directly:

    python css_proposal_generator.py              # web UI on http://0.0.0.0:8010/
    python css_proposal_generator.py --build-template   # rebuild base .docx from source

Assets expected in the same folder as this script:
    CSS_Proposal_Source.docx       (the organisation's master + 2 extra annexures; produced by
                                    build_source_template.py, then run --build-template)
    CSS_Proposal_Base_Template.docx (prepared docxtpl template; auto-built if missing)
    default_logo.png               (no longer placed on the cover; kept for reference)
"""
from __future__ import annotations

import argparse
import asyncio
import copy
import json
import os
import re
import shutil
import sys
import tempfile
import time
import traceback
import zipfile
from dataclasses import dataclass, field
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor
from docxtpl import DocxTemplate, InlineImage
from fastapi import Depends, FastAPI, File, Form, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, HTMLResponse

try:
    from PIL import Image

    PIL_AVAILABLE = True
except ImportError:
    PIL_AVAILABLE = False

try:
    from pydantic import BaseModel
    PYDANTIC_AVAILABLE = True
except ImportError:
    PYDANTIC_AVAILABLE = False
    BaseModel = object  # type: ignore[assignment,misc]

# ---------------------------------------------------------------------------
# Paths (all relative to this file — no imports from other project modules)
# ---------------------------------------------------------------------------
HERE = Path(__file__).resolve().parent
# The sibling proposal_* modules are imported by plain name so this file works both standalone
# and when mounted as Automation.backend.css_proposal_generator by the root app.
if str(HERE) not in sys.path:
    sys.path.insert(0, str(HERE))

import proposal_fields  # noqa: E402
import proposal_lint  # noqa: E402
import proposal_pdf  # noqa: E402
import proposal_contents  # noqa: E402
import proposal_pricing  # noqa: E402
import proposal_style  # noqa: E402
import proposal_tokens  # noqa: E402
from proposal_ratelimit import rate_limit  # noqa: E402

SOURCE_TEMPLATE = HERE / "CSS_Proposal_Source.docx"
BASE_TEMPLATE = HERE / "CSS_Proposal_Base_Template.docx"
DEFAULT_LOGO = HERE / "default_logo.png"
UPLOAD_DIR = HERE / "uploads" / "css_proposal"
MASTER_BD_PROMPT_FILE = HERE / "master_bd_prompt.txt"
VERSION = "1.5.0"
DEFAULT_PORT = 8010

# ---------------------------------------------------------------------------
# CyberSmithSECURE brand formatting constants — LOCKED, DO NOT CHANGE
# Verified against CyberSmithSECURE Annexures 2.docx
# ---------------------------------------------------------------------------
# Values are defined once, in proposal_tokens.py (Formatting Specification); these names remain
# for the existing code and tests.
CSS_FONT_NAME = proposal_tokens.FONT
CSS_ANNEXURE_HEADING_SIZE = Pt(proposal_tokens.SIZE_ANNEXURE_TITLE)   # Bold + Italic
CSS_SECTION_LABEL_SIZE = Pt(proposal_tokens.SIZE_HEADING_3)           # Bold
CSS_BODY_SIZE = Pt(proposal_tokens.SIZE_BODY)                         # Regular
CSS_HEADING_COLOR = RGBColor(*proposal_tokens.rgb(proposal_tokens.BRAND_PURPLE))  # #7030A0
CSS_BODY_COLOR = RGBColor(*proposal_tokens.rgb(proposal_tokens.BODY_BLACK))       # body text and bullets
CSS_MAIN_BULLET = proposal_tokens.LIST_GLYPHS[0]                      # U+2022
CSS_SUB_BULLET = proposal_tokens.LIST_GLYPHS[1]                       # lowercase o (per Master BD Prompt)
CSS_MAIN_BULLET_INDENT = str(int(proposal_tokens.LIST_LEVEL_INDENT_IN[0] * 1440))  # twips
CSS_SUB_BULLET_INDENT = str(int(proposal_tokens.LIST_LEVEL_INDENT_IN[1] * 1440))   # twips
CSS_BULLET_HANGING = str(int(proposal_tokens.LIST_HANGING_IN * 1440))              # twips

# ---------------------------------------------------------------------------
# AI annexure key pattern and required output sections (from Master BD Prompt)
# ---------------------------------------------------------------------------
OTHER_KEY_RE = re.compile(r"^__other__\d+_[a-z0-9]+$")

AI_REQUIRED_SECTIONS = [
    "Sub-Activity:",
    "Methodologies:",
    "Deliverables:",
    "Expectation:",
    "Objectives:",
    "Out of Scope",
    "Requirements for Starting",
    "Comments:",
]

VALID_PAYMENT_MODELS = {"milestone_v1", "hybrid_v2", "retainer_v3", "tier_v4"}
PAYMENT_MODEL_DESCRIPTIONS = {
    "milestone_v1": (
        "20% Advance Payment on PO release | "
        "30% on submission of initial/draft technical report | "
        "30% on initiation of remediation & verification (first revalidation window) | "
        "20% on submission of final closure report"
    ),
    "hybrid_v2": (
        "40% advance | "
        "20% at end of week 2 of testing | "
        "20% on detailed technical report submission | "
        "20% at day 30 from project start or after first revalidation (whichever is earlier)"
    ),
    "retainer_v3": (
        "30% advance | "
        "Remaining 70% split into 3 equal monthly instalments, due irrespective of client delays"
    ),
    "tier_v4": "Per-tier payment breakup (Advanced vs Offensive etc.), derived logically",
}

# Layout-based page estimation (lines per printed page; tuned for A4 proposal layout).
LINES_PER_PAGE = 38.0
CHARS_PER_LINE = 82
PARAS_PER_PAGE = LINES_PER_PAGE  # backwards compatibility for callers
# The Contents itself is written by proposal_contents.py (a real TOC field built from the headings).
TERMS_MARKER = "Terms & Conditions"
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
_DEFAULT_TEXT_WIDTH_TWIPS = 10914  # A4 minus the Base Template's 1.0 / 0.75 cm side margins

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
PIC_TAG = f"{{{W_NS}}}drawing"

TEXT_REPLACEMENTS: tuple[tuple[str, str], ...] = (
    ("XYZ LLC", "{{ client_name }}"),
    ("Company name", "{{ client_name }}"),
    ("EN-01-12-25", "{{ proposal_ref }}"),
    ("15-09-2025", "{{ proposal_date }}"),          # the master's cover date
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

# ``annexure_number`` is the annexure's position in the Source template (the organisation's master
# document has Annexures 1-11; Post Breach and Active Directory Assessment follow as 12 and 13).
ASSESSMENT_CATALOG: tuple[dict[str, Any], ...] = (
    {"key": "network_pt", "title": "Network Penetration Testing Assessment", "toc_title": "Network Penetration Testing Assessment", "annexure_number": 1},
    {"key": "web_app_pt", "title": "Web Application Penetration Testing Assessment", "toc_title": "Web Application Penetration Testing Assessment", "annexure_number": 2},
    {"key": "network_config_review", "title": "Network Configuration Review Assessment", "toc_title": "Network Configuration Review Assessment", "annexure_number": 3},
    {"key": "insider_threat", "title": "Insider Threat Assessment for Data Infiltration & Exfiltration", "toc_title": "Insider Threat Assessment for Data Infiltration & Exfiltration", "annexure_number": 4},
    {"key": "mobile_app_pt", "title": "Mobile Application Penetration Testing", "toc_title": "Mobile Application Penetration Testing", "annexure_number": 5},
    {"key": "ad_security", "title": "Active Directory Security Assessment", "toc_title": "Active Directory Security Assessment", "annexure_number": 6},
    {"key": "soc_simulation", "title": "SOC Simulation Exercise", "toc_title": "SOC Simulation Exercise", "annexure_number": 7},
    {"key": "cloud_pt", "title": "Cloud Penetration Testing", "toc_title": "Cloud Penetration Testing", "annexure_number": 8},
    {"key": "api_pt", "title": "API Penetration Testing Assessment", "toc_title": "API Penetration Testing Assessment", "annexure_number": 9},
    {"key": "offensive_defensive_training", "title": "Offensive & Defensive Security Training", "toc_title": "Offensive & Defensive Security Training", "annexure_number": 10},
    {"key": "source_code_review", "title": "Source Code Review", "toc_title": "Source Code Review", "annexure_number": 11},
    {"key": "post_breach", "title": "Post Breach Scenario Assessment", "toc_title": "Post Breach Scenario Assessment", "annexure_number": 12},
    {"key": "active_directory", "title": "Active Directory Assessment", "toc_title": "Active Directory Assessment", "annexure_number": 13},
)

CATALOG_BY_KEY: dict[str, dict[str, Any]] = {item["key"]: dict(item) for item in ASSESSMENT_CATALOG}
CATALOG_BY_NUMBER: dict[int, dict[str, Any]] = {
    int(item["annexure_number"]): dict(item) for item in ASSESSMENT_CATALOG
}

# Per-request mutable copies (keeps thread safety by creating a fresh dict per call)
_BASE_CATALOG_BY_KEY: dict[str, dict[str, Any]] = CATALOG_BY_KEY


def _make_request_catalog() -> dict[str, dict[str, Any]]:
    """Return a fresh copy of the catalog safe for per-request mutation."""
    return {k: dict(v) for k, v in _BASE_CATALOG_BY_KEY.items()}


def _register_custom_annexure(
    catalog: dict[str, dict[str, Any]],
    key: str,
    title: str,
    toc_title: str,
    slot: int,
    ai_elements: list | None = None,
) -> None:
    """Register a runtime AI-generated annexure into a per-request catalog."""
    catalog[key] = {
        "key": key,
        "title": title,
        "toc_title": toc_title[:60],
        "annexure_number": 99 + slot,   # avoids collisions with static 1–13
        "is_custom": True,
        "custom_slot": slot,
        "ai_elements": ai_elements or [],
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


# ---------------------------------------------------------------------------
# Pydantic schemas for AI context input
# ---------------------------------------------------------------------------

if PYDANTIC_AVAILABLE:
    class CountsSchema(BaseModel):  # type: ignore[misc]
        ad_identity: str = ""
        cloud: str = ""
        apps_apis: str = ""
        infra: str = ""
        users: str = ""
        duration: str = ""

    class CommercialsSchema(BaseModel):  # type: ignore[misc]
        total_cost: str = ""
        taxes: str = "18% GST"

    class CustomAnnexureInput(BaseModel):  # type: ignore[misc]
        slot_index: int
        key: str
        title: str
        scope: str
        # Excel flow only (all optional, so the manual form is unchanged)
        count_text: str = ""
        quantity: str = ""
        excel_context: str = ""

    class AIContextSchema(BaseModel):  # type: ignore[misc]
        sector: str = ""
        engagement_summary: str = ""
        primary_tier: str = "Tier 3 \u2013 Advanced"
        tier_notes: str = ""
        counts: CountsSchema = CountsSchema()
        commercials: CommercialsSchema = CommercialsSchema()
        payment_model: str = "hybrid_v2"
        special_constraints: str = ""
        custom_annexures: list[CustomAnnexureInput] = []

else:
    # Fallback dataclass equivalents when pydantic is not installed
    @dataclass
    class CountsSchema:  # type: ignore[no-redef]
        ad_identity: str = ""
        cloud: str = ""
        apps_apis: str = ""
        infra: str = ""
        users: str = ""
        duration: str = ""

    @dataclass
    class CommercialsSchema:  # type: ignore[no-redef]
        total_cost: str = ""
        taxes: str = "18% GST"

    @dataclass
    class CustomAnnexureInput:  # type: ignore[no-redef]
        slot_index: int = 0
        key: str = ""
        title: str = ""
        scope: str = ""
        count_text: str = ""
        quantity: str = ""
        excel_context: str = ""

    @dataclass
    class AIContextSchema:  # type: ignore[no-redef]
        sector: str = ""
        engagement_summary: str = ""
        primary_tier: str = "Tier 3 \u2013 Advanced"
        tier_notes: str = ""
        counts: Any = field(default_factory=CountsSchema)  # type: ignore[misc]
        commercials: Any = field(default_factory=CommercialsSchema)  # type: ignore[misc]
        payment_model: str = "hybrid_v2"
        special_constraints: str = ""
        custom_annexures: list = field(default_factory=list)


def _parse_ai_context(raw: str) -> "AIContextSchema":
    """Parse and validate the ai_context_json form field."""
    if not raw or not raw.strip():
        if PYDANTIC_AVAILABLE:
            return AIContextSchema()
        return AIContextSchema()  # type: ignore[call-arg]
    data = json.loads(raw)
    if PYDANTIC_AVAILABLE:
        return AIContextSchema.model_validate(data)

    # Manual fallback
    ctx = AIContextSchema()
    ctx.sector = str(data.get("sector", ""))
    ctx.engagement_summary = str(data.get("engagement_summary", ""))
    ctx.primary_tier = str(data.get("primary_tier", ctx.primary_tier))
    ctx.tier_notes = str(data.get("tier_notes", ""))
    ctx.payment_model = str(data.get("payment_model", ctx.payment_model))
    ctx.special_constraints = str(data.get("special_constraints", ""))
    counts_raw = data.get("counts", {})
    ctx.counts = CountsSchema(  # type: ignore[call-arg]
        ad_identity=str(counts_raw.get("ad_identity", "")),
        cloud=str(counts_raw.get("cloud", "")),
        apps_apis=str(counts_raw.get("apps_apis", "")),
        infra=str(counts_raw.get("infra", "")),
        users=str(counts_raw.get("users", "")),
        duration=str(counts_raw.get("duration", "")),
    )
    comm_raw = data.get("commercials", {})
    ctx.commercials = CommercialsSchema(  # type: ignore[call-arg]
        total_cost=str(comm_raw.get("total_cost", "")),
        taxes=str(comm_raw.get("taxes", "18% GST")),
    )
    ctx.custom_annexures = [
        CustomAnnexureInput(  # type: ignore[call-arg]
            slot_index=int(ca.get("slot_index", 0)),
            key=str(ca.get("key", "")),
            title=str(ca.get("title", "")),
            scope=str(ca.get("scope", "")),
            count_text=str(ca.get("count_text", "")),
            quantity=str(ca.get("quantity", "")),
            excel_context=str(ca.get("excel_context", "")),
        )
        for ca in data.get("custom_annexures", [])
    ]
    return ctx


def _validate_ai_context(ctx: "AIContextSchema", *, require_total_cost: bool = True) -> None:
    """Raise ValueError for missing required AI context fields.

    *require_total_cost* is True for the manual form (unchanged). The Excel flow passes False
    (plan D18): prices are entered later in the pricing preview, so a total may not exist yet.
    """
    if not ctx.sector.strip():
        raise ValueError("AI context: 'sector' is required when generating a custom annexure.")
    if len(ctx.engagement_summary.strip()) < 10:
        raise ValueError("AI context: 'engagement_summary' must be at least 10 characters.")
    if not ctx.counts.duration.strip():
        raise ValueError("AI context: 'counts.duration' is required.")
    if require_total_cost and not ctx.commercials.total_cost.strip():
        raise ValueError("AI context: 'commercials.total_cost' is required.")
    if ctx.payment_model not in VALID_PAYMENT_MODELS:
        raise ValueError(
            f"AI context: 'payment_model' must be one of {sorted(VALID_PAYMENT_MODELS)}."
        )
    for ca in ctx.custom_annexures:
        if not ca.title.strip():
            raise ValueError(f"Custom annexure key={ca.key}: 'title' is required.")
        if len(ca.scope.strip()) < 20:
            raise ValueError(
                f"Custom annexure '{ca.title}': 'scope' must be at least 20 characters."
            )


# ---------------------------------------------------------------------------
# Annexure order / assessments parsing
# ---------------------------------------------------------------------------

def parse_annexure_order_json(
    raw: str | None,
    *,
    allow_custom: bool = True,
) -> list[str] | None:
    """Parse optional annexure order from the form.

    Accepts either a JSON array of catalog keys or
    ``{"enabled": true, "order": ["post_breach", "web_app_pt", ...]}``.
    Custom AI keys matching ``OTHER_KEY_RE`` are allowed when *allow_custom* is True.
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
            if allow_custom and OTHER_KEY_RE.match(k):
                pass  # Custom AI annexure — validated later
            else:
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


# ---------------------------------------------------------------------------
# AI Annexure Engine
# ---------------------------------------------------------------------------

def _load_master_bd_prompt() -> str:
    """Load the locked Master BD Prompt from disk.  Never modify this file."""
    if not MASTER_BD_PROMPT_FILE.is_file():
        raise FileNotFoundError(
            f"Master BD Prompt not found at {MASTER_BD_PROMPT_FILE}. "
            "Ensure master_bd_prompt.txt exists next to css_proposal_generator.py."
        )
    return MASTER_BD_PROMPT_FILE.read_text(encoding="utf-8")


def _build_ai_user_message(
    company_name: str,
    context: "AIContextSchema",
    custom_annexure: "CustomAnnexureInput",
    slot_number: int,
) -> str:
    """Fill in the Master BD Prompt Section A with actual values.

    Only this annexure is requested — T&C and Payment Schedule are
    already present in the static template and must NOT be regenerated.
    """
    c = context.counts
    counts_lines: list[str] = []
    if c.ad_identity:
        counts_lines.append(f"AD / Identity:\n {c.ad_identity}")
    if c.cloud:
        counts_lines.append(f"Cloud:\n {c.cloud}")
    if c.apps_apis:
        counts_lines.append(f"Apps / APIs:\n {c.apps_apis}")
    if c.infra:
        counts_lines.append(f"Infra:\n {c.infra}")
    if c.users:
        counts_lines.append(f"Users / Employees (if Ghost Employee / Insider):\n {c.users}")
    if c.duration:
        counts_lines.append(f"Duration:\n {c.duration}")
    counts_block = "\n".join(counts_lines) if counts_lines else "(not specified)"

    payment_desc = PAYMENT_MODEL_DESCRIPTIONS.get(context.payment_model, context.payment_model)

    excel_block = ""
    count_text = getattr(custom_annexure, "count_text", "") or ""
    if count_text or getattr(custom_annexure, "excel_context", ""):
        excel_block = "\nFrom the client's pre-engagement scoping sheet (state only what is given here; never invent figures):\n"
        if count_text:
            excel_block += (f"Count for this annexure: {count_text}\n"
                            f"Total quantity: {getattr(custom_annexure, 'quantity', '') or 'not specified'}\n"
                            f"The annexure's Count line must read exactly: Count: {count_text}\n")
        if custom_annexure.excel_context:
            excel_block += custom_annexure.excel_context + "\n"

    return f"""A. INPUTS (I WILL FILL THESE BEFORE YOU GENERATE)

Client Name: {company_name}
Client Entity Name to be used in document: {company_name}
Sector / Industry: {context.sector}

Overall Engagement Summary (1\u20132 lines):
{context.engagement_summary}

Annexures to Generate (complete list):
Annexure {slot_number} \u2013 {custom_annexure.title}
High-level scope:
{custom_annexure.scope}
{excel_block}
Tiering Model:
We use these 4 standard tiers (for internal clarity; only mention tier where necessary):
Tier 1 \u2013 Essential
Tier 2 \u2013 Enhanced
Tier 3 \u2013 Advanced
Tier 4 \u2013 Offensive

Primary Tier(s) in Scope for this Proposal:
{context.primary_tier}

Any Tier-specific notes:
{context.tier_notes or "None"}

Counts / Volumes:
{counts_block}

Total Commercials:
TOTAL COST (before tax): {context.commercials.total_cost or "Not yet specified"}
Taxes: {context.commercials.taxes}

Preferred Payment Model (choose ONE):
Selected Payment Model: {context.payment_model}
({payment_desc})

Special Constraints / Client Notes (if any):
{context.special_constraints or "None"}

====================================================================
IMPORTANT OUTPUT INSTRUCTION:
Generate ONLY the annexure content block for Annexure {slot_number}: {custom_annexure.title}.
Do NOT generate the Terms & Conditions section.
Do NOT generate the Payment Schedule section.
Do NOT generate the Conclusion & Pricing section.
Output ONLY the annexure block starting from:
"Annexure {slot_number}: {custom_annexure.title}"
and ending after the "Comments:" section.
Use exactly "\\u2022 " (bullet character + space) for main bullets and "o " for sub-bullets.
Do NOT use markdown headers (no ## or **). Use plain text headings ending with colon.
"""


def _build_ai_retry_message(problem: str) -> str:
    """Corrective follow-up sent once when the first answer fails validation."""
    return (
        "Your previous answer was rejected by the structure check: "
        f"{problem}\n"
        "Regenerate the COMPLETE annexure again. It must include every one of these section "
        f"labels, each on its own line: {', '.join(AI_REQUIRED_SECTIONS)}, plus 'Frequency:' and "
        "'Count:'. Keep the same formatting rules (plain text, bullets as the bullet character, "
        "sub-bullets as 'o', no markdown). Output only the annexure block."
    )


def validate_ai_annexure_output(text: str, title: str) -> None:
    """Lightweight structural check on AI-generated annexure text.

    Checks that all required sections from the Master BD Prompt exist.
    Does NOT alter the text — content quality is enforced by the guardrail prompt.
    """
    missing = [s for s in AI_REQUIRED_SECTIONS if s.lower() not in text.lower()]
    if missing:
        raise ValueError(
            f"AI output for '{title}' is missing required sections: {missing}. "
            "Please try generating again."
        )
    if len(text.strip()) < 500:
        raise ValueError(
            f"AI output for '{title}' is too short ({len(text)} chars). "
            "The model may not have generated complete content. Please try again."
        )


async def generate_custom_annexure_text(
    company_name: str,
    context: "AIContextSchema",
    custom_annexure: "CustomAnnexureInput",
    slot_number: int,
) -> str:
    """Call OpenAI with the locked Master BD Prompt + filled Section A.

    Returns the raw annexure text block (plain text with • / o bullets).
    Raises HTTPException on missing key, quota errors, or timeout.
    """
    from dotenv import load_dotenv
    load_dotenv(HERE / ".env", override=True)

    # A key meant only for the proposal's custom annexures wins; otherwise the shared OPENAI_API_KEY.
    api_key = (os.getenv("PROPOSAL_OPENAI_API_KEY") or os.getenv("OPENAI_API_KEY") or "").strip()
    if not api_key:
        raise HTTPException(
            status_code=500,
            detail="AI generation is not configured. Set PROPOSAL_OPENAI_API_KEY (or OPENAI_API_KEY) in .env.",
        )

    system_prompt = _load_master_bd_prompt()
    user_message = _build_ai_user_message(company_name, context, custom_annexure, slot_number)

    try:
        from openai import AsyncOpenAI, APITimeoutError, RateLimitError
        client = AsyncOpenAI(api_key=api_key)
        model = (os.getenv("PROPOSAL_OPENAI_MODEL") or os.getenv("OPENAI_MODEL") or "gpt-4o").strip()

        messages = [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_message},
        ]
        last_error: ValueError | None = None
        for attempt in (1, 2):                      # one automatic retry on a structural failure
            response = await client.chat.completions.create(
                model=model,
                temperature=0.2,
                max_tokens=3500,
                messages=messages,
            )
            raw_text = (response.choices[0].message.content or "").strip()
            try:
                validate_ai_annexure_output(raw_text, custom_annexure.title)
                return raw_text
            except ValueError as exc:
                last_error = exc
                if attempt == 1:
                    messages = messages + [
                        {"role": "assistant", "content": raw_text},
                        {"role": "user", "content": _build_ai_retry_message(str(exc))},
                    ]
        raise HTTPException(
            status_code=502,
            detail=("The AI output did not meet the required structure, even after one automatic "
                    f"retry. {last_error} Please try again."),
        )

    except ImportError:
        raise HTTPException(
            status_code=500,
            detail="openai package not installed. Run: pip install openai",
        )
    except HTTPException:
        raise
    except Exception as exc:
        err_str = str(exc).lower()
        if "insufficient_quota" in err_str:
            raise HTTPException(
                status_code=429,
                detail=(
                    "OpenAI quota exhausted. Add prepaid credits and update OPENAI_API_KEY in .env. "
                    "See docs/OPENAI_AI_LAYER_PRODUCTION.md"
                ),
            )
        if "timeout" in err_str or "timed out" in err_str:
            raise HTTPException(
                status_code=504,
                detail="AI generation timed out. Please try again in a moment.",
            )
        if "rate_limit" in err_str or "rate limit" in err_str:
            raise HTTPException(
                status_code=429,
                detail="AI service is busy. Please wait 30 seconds and try again.",
            )
        traceback.print_exc()
        raise HTTPException(
            status_code=500,
            detail=f"AI generation failed: {exc}",
        )


# ---------------------------------------------------------------------------
# DOCX element builder for AI annexures
# (produces elements with exact CyberSmithSECURE brand formatting)
# ---------------------------------------------------------------------------

def _css_section_label_re() -> re.Pattern[str]:
    """Regex matching section label lines (Title Case word(s) ending with colon)."""
    return re.compile(
        r"^(?:[A-Z][A-Za-z0-9 &/\-\u2013\u2014]+:|\bFrequency:\b|\bCount:\b)$"
    )


_SECTION_LABEL_RE = _css_section_label_re()


def _parse_annexure_line(text: str) -> dict[str, Any]:
    """Classify a single text line into a formatting style dict."""
    stripped = text.strip()
    if not stripped:
        return {"style": "blank", "text": ""}
    if re.match(r"^Annexure\s+\d+\s*:", stripped, re.IGNORECASE):
        return {"style": "annexure_heading", "text": stripped}
    if _SECTION_LABEL_RE.match(stripped):
        return {"style": "section_label", "text": stripped}
    if stripped.startswith(CSS_MAIN_BULLET + " "):
        return {"style": "main_bullet", "text": stripped[2:].strip()}
    if stripped.startswith("o ") and not stripped.startswith("of ") and not stripped.startswith("on "):
        return {"style": "sub_bullet", "text": stripped[2:].strip()}
    return {"style": "body", "text": stripped}


def _add_formatted_paragraph(doc: Document, parsed: dict[str, Any]) -> None:
    """Add a single paragraph to *doc* with exact CyberSmithSECURE brand formatting."""
    style_name = parsed["style"]
    text = parsed["text"]

    if style_name == "blank":
        return  # spacing comes from paragraph spacing below, never from empty paragraphs

    p = doc.add_paragraph()

    # Bullets are real Word list items (numbering definition owned by proposal_style), not typed glyphs.
    run_text = text

    run = p.add_run(run_text)

    # Apply Corbel to all; purple headings match static template annexures (#7030A0).
    run.font.name = CSS_FONT_NAME
    if style_name in ("annexure_heading", "section_label"):
        run.font.color.rgb = CSS_HEADING_COLOR
    else:
        run.font.color.rgb = CSS_BODY_COLOR

    if style_name == "annexure_heading":
        run.font.size = CSS_ANNEXURE_HEADING_SIZE
        run.font.bold = True
        run.font.italic = True
    elif style_name == "section_label":
        run.font.size = CSS_SECTION_LABEL_SIZE
        run.font.bold = True
    else:
        run.font.size = CSS_BODY_SIZE

    # Bullet list membership (indent and glyph come from the numbering level)
    if style_name in ("main_bullet", "sub_bullet"):
        pPr = p._p.get_or_add_pPr()
        pPr.append(proposal_style.bullet_numpr(0 if style_name == "main_bullet" else 1))

    # Paragraph spacing (Formatting Specification): labels open a new block, lists are tight,
    # body text gets the standard gap.
    before_pt, after_pt = {
        "annexure_heading": (0, 6),
        "section_label": (12, 4),
        "main_bullet": (0, proposal_tokens.LIST_SPACE_AFTER_PT),
        "sub_bullet": (0, proposal_tokens.LIST_SPACE_AFTER_PT),
    }.get(style_name, (proposal_tokens.BODY_SPACE_BEFORE_PT, proposal_tokens.BODY_SPACE_AFTER_PT))
    pPr = p._p.get_or_add_pPr()
    sp = OxmlElement("w:spacing")
    sp.set(qn("w:before"), str(int(before_pt * 20)))
    sp.set(qn("w:after"), str(int(after_pt * 20)))
    pPr.append(sp)
    if style_name in ("annexure_heading", "section_label"):
        pPr.append(OxmlElement("w:keepNext"))        # a heading never ends a page on its own


def build_custom_annexure_block(
    raw_ai_text: str,
    slot_number: int,
    title: str,
    client_name: str,
) -> list:
    """Convert raw AI plain-text annexure into a list of OOXML elements.

    The elements have the exact CyberSmithSECURE brand formatting
    (Corbel font, correct sizes/styles, bullet indents) and can be
    inserted directly into the proposal by ``apply_annexure_order``.
    """
    # Replace any {{Client Name}} tokens the AI may have emitted
    for token in ("{{Client Name}}", "{{client_name}}", "{client_name}", "{{Company Name}}"):
        raw_ai_text = raw_ai_text.replace(token, client_name)

    # Defensive: strip markdown artifacts the model occasionally produces
    text = re.sub(r"(?m)^#{1,3}\s*", "", raw_ai_text)
    text = re.sub(r"\*{1,2}([^*]+)\*{1,2}", r"\1", text)

    # The heading always carries this annexure's real position and title (numbering rule),
    # whatever number the model wrote.
    lines = text.split("\n")
    heading = f"Annexure {slot_number}: {title}"
    first = next((i for i, line in enumerate(lines) if line.strip()), None)
    if first is None:
        lines = [heading]
    elif ANNEXURE_HEADING_RE.match(lines[first].strip()):
        lines[first] = heading
    else:
        lines.insert(first, heading)

    scratch = Document()
    # Remove the default empty paragraph python-docx always adds
    for p in list(scratch.paragraphs):
        p._element.getparent().remove(p._element)

    for line in lines:
        parsed = _parse_annexure_line(line)
        _add_formatted_paragraph(scratch, parsed)

    body = scratch.element.body
    elements: list = []
    for child in body:
        if child.tag == qn("w:sectPr"):
            continue
        elements.append(copy.deepcopy(child))

    return elements


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


def _default_pricing_rows(
    ordered_keys: list[str],
    catalog: dict[str, dict[str, Any]],
) -> list[dict[str, str]]:
    """Manual flow: one row per annexure with the title and default frequency.

    Quantity and cost stay empty (to be typed in Word). The template's old dummy quantities
    (e.g. "1 Web Application", "Three Time") were placeholders, not data, and are no longer used.
    """
    rows: list[dict[str, str]] = []
    for key in ordered_keys:
        item = catalog.get(key) or CATALOG_BY_KEY.get(key, {})
        rows.append({"title": str(item.get("title", key)), "quantity": "", "frequency": "One Time", "cost": ""})
    return rows


def _update_conclusion_pricing_table(
    elements,
    ordered_keys: list[str],
    catalog: dict[str, dict[str, Any]] | None = None,
    pricing_rows: list[dict[str, str]] | None = None,
    text_width_twips: int | None = None,
) -> None:
    """Replace the Conclusion & Pricing table with one built from the formatting tokens.

    *pricing_rows* come from the reviewed Excel draft (Preview 2); without them (manual flow)
    one row per selected annexure is produced. *catalog* may include AI-generated custom
    entries; falls back to CATALOG_BY_KEY.
    """
    effective_catalog = catalog if catalog is not None else CATALOG_BY_KEY

    tbl = None
    for el in elements:
        tbl = _find_assessment_pricing_table(el)
        if tbl is not None:
            break
    if tbl is None:
        return

    rows = pricing_rows if pricing_rows is not None else _default_pricing_rows(ordered_keys, effective_catalog)
    width = text_width_twips or _DEFAULT_TEXT_WIDTH_TWIPS
    if tbl.getparent() is None:
        # The master has the table directly in the body (no content control around it): the cloned
        # block is a plain list, so the table is replaced in that list.
        elements[elements.index(tbl)] = proposal_pricing.build_pricing_table(rows, width)
        return
    proposal_pricing.replace_pricing_table(tbl, rows, width)


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


def _has_page_break_before(ppr) -> bool:
    """True when a paragraph really starts a new page.

    ``<w:pageBreakBefore w:val="0"/>`` (very common in the template) means *off*; only a
    missing val or an explicit true/1/on value turns the break on.
    """
    if ppr is None:
        return False
    node = ppr.find(qn("w:pageBreakBefore"))
    if node is None:
        return False
    return (node.get(qn("w:val")) or "true").strip().lower() not in ("0", "false", "off")


def _ensure_final_section(doc: Document) -> bool:
    """Make sure the body ends with a ``<w:sectPr>`` (the document's last section).

    Assembly removes the template's closing section properties together with the old body
    content. Without them Word falls back to application defaults for the annexures, terms and
    pricing pages (which can be Letter instead of A4). The properties of the last section break
    are copied, as a *continuous* section so no extra page is created. Returns True if added.
    """
    body = doc.element.body
    if body.find(qn("w:sectPr")) is not None:
        return False
    in_paragraph = [s for s in body.iter(qn("w:sectPr")) if s.getparent().tag == qn("w:pPr")]
    if not in_paragraph:
        return False
    final = copy.deepcopy(in_paragraph[-1])
    for tag in ("w:type", "w:titlePg"):
        for node in final.findall(qn(tag)):
            final.remove(node)
    kind = OxmlElement("w:type")
    kind.set(qn("w:val"), "continuous")
    final.insert(len(final.findall(qn("w:headerReference"))) + len(final.findall(qn("w:footerReference"))), kind)
    body.append(final)
    return True


def _is_blank_paragraph(el) -> bool:
    """Empty paragraph: no text, no break, no drawing/picture, no section break."""
    return (
        el.tag == qn("w:p")
        and not _paragraph_text(el).strip()
        and not _has_page_break(el)
        and el.find(f".//{PIC_TAG}") is None
        and el.find(f".//{{{W_NS}}}pict") is None
        and el.find(f".//{{{W_NS}}}sectPr") is None
    )


def _set_page_break_before(p) -> None:
    """Give paragraph *p* ``<w:pageBreakBefore/>`` in schema order (after pStyle/keepNext/keepLines)."""
    ppr = p.find(qn("w:pPr"))
    if ppr is None:
        ppr = OxmlElement("w:pPr")
        p.insert(0, ppr)
    node = ppr.find(qn("w:pageBreakBefore"))
    if node is None:
        node = OxmlElement("w:pageBreakBefore")
        anchor = None
        for tag in ("w:pStyle", "w:keepNext", "w:keepLines"):
            found = ppr.find(qn(tag))
            if found is not None:
                anchor = found
        if anchor is not None:
            anchor.addnext(node)
        else:
            ppr.insert(0, node)
    node.attrib.pop(qn("w:val"), None)  # no val = on


def _normalize_page_breaks(doc: Document) -> dict[str, int]:
    """Make empty pages structurally impossible (hard acceptance rule).

    1. a page-break-only paragraph directly after another break is dropped,
    2. every remaining break paragraph (and the blank paragraphs after it) is replaced by
       ``pageBreakBefore`` on the next real paragraph: a free-standing break paragraph can itself
       spill onto a new page when the previous page is exactly full, which is how blank pages
       appeared between sections,
    3. blank paragraphs right before a page-start paragraph, and at the end of the document, go:
       they have no visible purpose and are the other cause of spill-over pages.
    """
    stats = {"stacked": _remove_stacked_page_breaks(doc), "converted": 0, "blank_removed": 0}
    body = doc.element.body
    children = list(body)
    removed: set[int] = set()
    for i, el in enumerate(children):
        if id(el) in removed or el.tag != qn("w:p") or not _is_page_break_only_paragraph(el):
            continue
        j, blanks = i + 1, []
        while j < len(children) and _is_blank_paragraph(children[j]):
            blanks.append(children[j])
            j += 1
        target = children[j] if j < len(children) else None
        if target is not None and target.tag != qn("w:sectPr"):
            first_p = target if target.tag == qn("w:p") else target.find(f".//{qn('w:p')}")
            if first_p is not None:
                _set_page_break_before(first_p)
        for dead in [el, *blanks]:
            body.remove(dead)
            removed.add(id(dead))
        stats["converted"] += 1

    def preceding_blank_run(start_el):
        run, prev = [], start_el.getprevious()
        while prev is not None and _is_blank_paragraph(prev):
            run.append(prev)
            prev = prev.getprevious()
        return run

    for p in list(body.iter(qn("w:p"))):
        if not _has_page_break_before(p.find(qn("w:pPr"))):
            continue
        top = p                                  # the body-level element that starts the new page
        while top.getparent() is not body and top.getparent() is not None:
            top = top.getparent()
        first_p = top if top.tag == qn("w:p") else top.find(f".//{qn('w:p')}")
        if first_p is not p:
            continue                             # only the first paragraph of a block starts the page
        for blank in preceding_blank_run(top):
            body.remove(blank)
            stats["blank_removed"] += 1
    last = body.find(qn("w:sectPr"))
    tail = preceding_blank_run(last) if last is not None else []
    if tail:
        before_run = tail[-1].getprevious()              # the element the blank run follows
        if before_run is not None and before_run.tag == qn("w:tbl"):
            proposal_style.make_tiny_paragraph(tail.pop())   # Word needs one paragraph after a final table
    for blank in tail:
        body.remove(blank)
        stats["blank_removed"] += 1
    return stats


def _remove_stacked_page_breaks(doc: Document) -> int:
    """Delete page-break-only paragraphs that directly follow another page break.

    Two breaks with no content between them produce an empty page (hard acceptance rule:
    no empty pages). Blank paragraphs between the breaks do not count as content.
    Returns the number of paragraphs removed.
    """
    body = doc.element.body
    pending = False  # a page break has just happened and nothing visible followed it
    removed = 0
    for el in list(body):
        if el.tag != qn("w:p"):
            pending = False  # tables / content controls are content
            continue
        if _is_page_break_only_paragraph(el):
            if pending:
                body.remove(el)
                removed += 1
            else:
                pending = True
            continue
        has_content = (
            bool(_paragraph_text(el))
            or el.find(f".//{PIC_TAG}") is not None
            or el.find(f".//{{{W_NS}}}pict") is not None
            or el.find(f".//{{{W_NS}}}sectPr") is not None
        )
        if has_content:
            pending = _has_page_break(el)
    return removed


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


def _catalog_item_for_key(
    key: str,
    catalog: dict[str, dict[str, Any]] | None = None,
) -> dict[str, Any]:
    """Resolve a catalog entry from the per-request catalog or the static base catalog."""
    effective = catalog if catalog is not None else CATALOG_BY_KEY
    return effective.get(key) or CATALOG_BY_KEY.get(key) or {}


def _toc_title_for_key(
    key: str,
    catalog: dict[str, dict[str, Any]] | None = None,
) -> str:
    """Human-readable Contents line label for an annexure key."""
    item = _catalog_item_for_key(key, catalog)
    return str(item.get("toc_title") or item.get("title") or key)


def _finalize_document(
    docx_path: Path,
    ordered_keys: list[str] | None = None,
    *,
    request_catalog: dict[str, dict[str, Any]] | None = None,
) -> None:
    """Apply final layout fixes and refresh all Contents page numbers."""
    if not ordered_keys:
        ordered_keys = [item["key"] for item in ASSESSMENT_CATALOG]

    doc = Document(str(docx_path))
    _ensure_final_section(doc)
    proposal_style.apply_specification(doc)          # Track B: styles, body text, tables, signature, lists
    _normalize_page_breaks(doc)

    # Contents: a real TOC field written from the headings that are in the body. The page numbers
    # are an estimate here; finalize_with_render replaces them with the measured ones.
    proposal_contents.build(doc)
    pages = _compute_layout_pages(doc)
    estimate = _estimate_heading_pages(doc, pages)
    proposal_contents.build(doc, estimate)

    _strip_all_yellow_highlights(doc)
    doc.save(str(docx_path))
    _set_update_fields_on_open(docx_path)


def _apply_measured_toc_pages(docx_path: Path, measured: dict[str, int]) -> int:
    """Write page numbers measured from a real render into the Contents lines.

    *measured* comes from proposal_pdf.measure_heading_pages. The Contents lines are plain text,
    so this is the only way they can be exact; the layout estimator is only a fallback.
    Returns how many lines changed.
    """
    doc = Document(str(docx_path))
    changed = proposal_contents.refresh_pages(doc, measured)
    if changed:
        doc.save(str(docx_path))
    return changed


def _estimate_heading_pages(doc: Document, pages: dict[int, int]) -> dict[str, int]:
    """Estimated page of every Contents heading (same keys as proposal_pdf.measure_heading_pages)."""
    estimate: dict[str, int] = {}
    paragraphs = doc.paragraphs                       # kept alive so the lxml proxies (and their ids) stay valid
    index_of = {id(p._element): i for i, p in enumerate(paragraphs)}
    for h in proposal_contents.find_headings(doc.element.body):
        i = index_of.get(id(h.paragraph))
        if i is not None:
            estimate[h.key] = int(pages.get(i, 1))
    return estimate


def finalize_with_render(docx_path: Path, ordered_keys: list[str], company_name: str = "") -> "proposal_pdf.PdfCheck":
    """Render the finished proposal, correct Contents from the real pagination, verify.

    Steps: render -> measure where each heading really starts -> rewrite the Contents numbers ->
    render again only if something changed -> report blank pages. Never raises: with no PDF engine
    (or a render failure) the proposal is returned unverified and the Contents keep the estimate.
    The PDF stays next to the .docx as <name>.pdf.
    """
    import proposal_pdf

    pdf_path = docx_path.with_suffix(".pdf")
    check = proposal_pdf.check_document(docx_path, len(ordered_keys), pdf_path=pdf_path)
    if not check.verified:
        return check
    try:
        if _apply_measured_toc_pages(docx_path, check.headings):
            _repair_core_properties(docx_path, title=f"Statement of Work & Proposal - {company_name}")
            check = proposal_pdf.check_document(docx_path, len(ordered_keys), pdf_path=pdf_path)
    except Exception as exc:  # the unverified document is still delivered
        traceback.print_exc()
        check.note = f"contents correction failed: {exc}"[:200]
    return check


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

        if _has_page_break_before(el.find(qn("w:pPr"))):
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


_CORE_NS = {
    "cp": "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
    "dc": "http://purl.org/dc/elements/1.1/",
    "dcterms": "http://purl.org/dc/terms/",
    "xsi": "http://www.w3.org/2001/XMLSchema-instance",
}


def _fit_cover_reference(doc: Document) -> int:
    """Shrink the cover reference so a long one (e.g. CSS-ENG-2026-943) stays on the page.

    The cover's Title paragraph starts 4.6 in from the left at 36 pt (designed for short refs like
    EN-01-12-25). A longer reference overflowed and lost its first characters. The largest size
    that fits the remaining width is chosen; short references keep the template's 36 pt.
    """
    for p in doc.paragraphs:
        text = (p.text or "").strip()
        if not text:
            continue
        ppr = p._p.find(qn("w:pPr"))
        style = ppr.find(qn("w:pStyle")) if ppr is not None else None
        if style is None or style.get(qn("w:val")) != "Title":
            return 0
        sect = doc.sections[0]
        ind = ppr.find(qn("w:ind"))
        first_line = int(ind.get(qn("w:firstLine")) or 0) if ind is not None else 0
        if ind is not None and first_line:
            # The Title style already indents the paragraph (left 6686); the direct first-line
            # indent doubled it and pushed the first character off the page (the reference lost
            # its first letter, e.g. "N-01-12-25"). Zero it: same position, whole text visible.
            ind.set(qn("w:firstLine"), "0")
        usable_in = (sect.page_width - sect.left_margin - sect.right_margin) / 914400 - first_line / 1440 - 0.1
        size = 36
        while size > 16 and len(text) * 0.57 * size / 72 > usable_in:
            size -= 2
        if size == 36:
            return 36
        for r in p._p.iter(qn("w:r")):
            rpr = r.find(qn("w:rPr"))
            if rpr is None:
                rpr = OxmlElement("w:rPr")
                r.insert(0, rpr)
            for tag in ("w:sz", "w:szCs"):
                for node in rpr.findall(qn(tag)):
                    rpr.remove(node)
                el = OxmlElement(tag)
                el.set(qn("w:val"), str(size * 2))
                rpr.append(el)
        return size
    return 0


def _repair_core_properties(docx_path: Path, title: str = "") -> None:
    """Rewrite docProps/core.xml in valid form.

    docxtpl/python-docx create this part when the template has none, and (because the template
    carries docProps/custom.xml) bind ``cp:lastModifiedBy`` and ``cp:revision`` to the
    *custom-properties* namespace. Microsoft Word treats that as a corrupt file and refuses to
    open it. This runs last, after every python-docx save, so the final file is always valid.
    """
    from xml.sax.saxutils import escape

    now = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    core = (
        "<?xml version='1.0' encoding='UTF-8' standalone='yes'?>\n"
        + "<cp:coreProperties " + " ".join(f'xmlns:{k}="{v}"' for k, v in _CORE_NS.items()) + ">"
        + f"<dc:title>{escape(title or 'Statement of Work & Proposal')}</dc:title>"
        + "<dc:creator>CyberSmithSECURE</dc:creator>"
        + "<cp:lastModifiedBy>CyberSmithSECURE</cp:lastModifiedBy>"
        + "<cp:revision>1</cp:revision>"
        + f'<dcterms:created xsi:type="dcterms:W3CDTF">{now}</dcterms:created>'
        + f'<dcterms:modified xsi:type="dcterms:W3CDTF">{now}</dcterms:modified>'
        + "</cp:coreProperties>"
    )
    try:
        with zipfile.ZipFile(docx_path, "r") as zin:
            entries = [(info, zin.read(info.filename)) for info in zin.infolist()]
    except (zipfile.BadZipFile, OSError):
        return
    names = {info.filename for info, _ in entries}
    tmp = docx_path.with_suffix(".core.tmp")
    with zipfile.ZipFile(tmp, "w", compression=zipfile.ZIP_DEFLATED) as zout:
        for info, data in entries:
            if info.filename == "docProps/core.xml":
                data = core.encode("utf-8")
            zout.writestr(info.filename, data)
        if "docProps/core.xml" not in names:       # nothing to repair (and nothing to add)
            pass
    shutil.move(str(tmp), str(docx_path))


def apply_annexure_order(
    docx_path: Path,
    ordered_keys: list[str] | None,
    *,
    company_name: str = "Client",
    request_catalog: dict[str, dict[str, Any]] | None = None,
    annexure_details: dict[str, dict[str, Any]] | None = None,
    pricing_rows: list[dict[str, str]] | None = None,
) -> None:
    """Reorder / subset annexure sections, append Terms & Conclusion, rebuild Contents.

    *request_catalog* extends the base CATALOG_BY_KEY with per-request AI-generated
    entries (keys matching ``OTHER_KEY_RE``).  When None the base catalog is used.

    Excel flow only: *annexure_details* maps an annexure key to
    ``{"count_text": "6 Web Application categories", "frequency": "One Time"}`` and rewrites
    that annexure's Count / Frequency lines; *pricing_rows* feeds the Conclusion & Pricing
    table. Both None leaves the manual flow's annexure text untouched.
    """
    if not ordered_keys:
        return

    catalog = request_catalog if request_catalog is not None else CATALOG_BY_KEY

    doc = Document(str(docx_path))
    last_section = doc.sections[-1]
    text_width_twips = int((last_section.page_width - last_section.left_margin
                            - last_section.right_margin) / 635)
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
            if body_children[i].tag == qn("w:sectPr"):
                continue                      # the document's own (header/footer/page size) section stays
            body.remove(body_children[i])
    else:
        _remove_paragraphs_from(doc, first_idx)
    body = doc.element.body

    # Template annexure number -> position in this proposal, for cross-reference renumbering.
    number_map: dict[int, int] = {}
    for pos, k in enumerate(ordered_keys, start=1):
        it = catalog.get(k) or CATALOG_BY_KEY.get(k) or {}
        if not it.get("is_custom") and it.get("annexure_number") is not None:
            number_map[int(it["annexure_number"])] = pos

    for slot, key in enumerate(ordered_keys, start=1):
        item = catalog.get(key) or CATALOG_BY_KEY.get(key)
        if item is None:
            raise ValueError(f"Unknown annexure key in order: {key!r}")

        # ---- Custom / AI-generated annexure --------------------------------
        if item.get("is_custom"):
            ai_elements = item.get("ai_elements") or []
            if not ai_elements:
                raise ValueError(
                    f"No AI-generated content found for custom annexure '{item.get('title', key)}'. "
                    "This is an internal error — the content should have been generated before this point."
                )
            cloned: list = [_make_page_break_paragraph()]
            for el in ai_elements:
                cloned.append(copy.deepcopy(el))
            _replace_client_literals_in_elements(cloned, company_name)
            if annexure_details and key in annexure_details:
                d = annexure_details[key]
                proposal_fields.apply_count_and_frequency(
                    cloned, count_text=d.get("count_text"), frequency=d.get("frequency"))
            _insert_elements_before_sectpr(body, cloned)
            continue

        # ---- Static template annexure --------------------------------------
        default_num = int(item["annexure_number"])
        block = blocks.get(default_num)
        if not block:
            raise ValueError(f"Missing template section for annexure {default_num} ({key}).")

        cloned = []
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

        if annexure_details and key in annexure_details:
            d = annexure_details[key]
            proposal_fields.apply_count_and_frequency(
                cloned, count_text=d.get("count_text"), frequency=d.get("frequency"))
        proposal_fields.renumber_cross_references(cloned, number_map)
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
    _update_conclusion_pricing_table(
        trailer_clone, ordered_keys, catalog=catalog,
        pricing_rows=pricing_rows, text_width_twips=text_width_twips)
    _replace_client_literals_in_elements(trailer_clone, company_name)
    for el in trailer_clone:
        _strip_yellow_from_element(el)
    _insert_elements_before_sectpr(body, trailer_clone)

    doc.save(str(docx_path))          # the Contents is written from the finished body in _finalize_document


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
    """Return the client-logo slot: the paragraph between "For" and the client name on page 2."""
    after_for = False
    for para in doc.paragraphs:
        txt = (para.text or "").strip()
        if txt == "For":
            after_for = True
            continue
        if not after_for:
            continue
        if "{{ client_logo }}" in (para.text or ""):
            return para
        if para._element.findall(f".//{PIC_TAG}") and not txt:
            return para
        if txt:
            return None             # reached the client name without finding the slot
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
    logo_para.alignment = WD_ALIGN_PARAGRAPH.LEFT
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
    logo_para.alignment = WD_ALIGN_PARAGRAPH.LEFT
    logo_para.add_run("{{ client_logo }}")
    doc.save(str(docx_path))


def _ensure_client_logo_left_alignment(docx_path: Path) -> None:
    """Ensure the client logo slot under 'For' is left-aligned (brand layout)."""
    try:
        doc = Document(str(docx_path))
        logo_para = _find_cover_logo_paragraph(doc)
        if logo_para is None:
            return
        logo_para.alignment = WD_ALIGN_PARAGRAPH.LEFT
        doc.save(str(docx_path))
    except Exception:
        traceback.print_exc()


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
                paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT

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
        _ensure_client_logo_left_alignment(BASE_TEMPLATE)
    return BASE_TEMPLATE


# ---------------------------------------------------------------------------
# Proposal generation
# ---------------------------------------------------------------------------
def default_proposal_ref(dt: datetime | None = None) -> str:
    dt = dt or datetime.now()
    return dt.strftime("%d-%m-%Y")


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
    except Exception:
        traceback.print_exc()
    # No client logo: the master leaves the slot between "For" and the client name empty (a
    # CyberSmithSECURE logo there would read as the client's). An empty string renders as nothing.
    return ""


def generate_proposal(
    *,
    company_name: str,
    proposal_ref: str,
    logo_path: str | Path | None,
    output_path: str | Path,
    annexure_order: list[str] | None = None,
    assessments: list[AssessmentItem] | None = None,
    template_path: str | Path | None = None,
    request_catalog: dict[str, dict[str, Any]] | None = None,
    annexure_details: dict[str, dict[str, Any]] | None = None,
    pricing_rows: list[dict[str, str]] | None = None,
    proposal_date: str | None = None,
) -> Path:
    company_name = (company_name or "").strip() or "Client"
    proposal_ref = (proposal_ref or "").strip() or default_proposal_ref()
    proposal_date = (proposal_date or "").strip() or datetime.now().strftime("%d-%m-%Y")
    template = Path(template_path) if template_path else ensure_base_template()
    output = Path(output_path)

    tpl = DocxTemplate(str(template))
    context: dict[str, Any] = {
        "client_name": company_name,
        "proposal_ref": proposal_ref,
        "proposal_date": proposal_date,
        "client_logo": _resolve_logo_inline(tpl, logo_path),
        "report_date": datetime.now().strftime("%d %B %Y"),
    }
    try:
        # autoescape: a client name such as "L&T" or "Smith & Sons" must not break the Word XML
        tpl.render(context, autoescape=True)
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
        _fit_cover_reference(doc)

        if logo_file:
            if not _insert_logo_on_cover(doc, logo_file):
                _replace_highlighted_placeholder_image(doc, logo_file)

        doc.save(str(output))
    except Exception:
        traceback.print_exc()
    
    try:
        ordered_keys = annexure_order or [item["key"] for item in ASSESSMENT_CATALOG]
        apply_annexure_order(
            output,
            ordered_keys,
            company_name=company_name,
            request_catalog=request_catalog,
            annexure_details=annexure_details,
            pricing_rows=pricing_rows,
        )
    except Exception as exc:
        raise RuntimeError(f"Annexure layout failed: {exc}") from exc

    _finalize_document(output, ordered_keys, request_catalog=request_catalog)
    _repair_core_properties(output, title=f"Statement of Work & Proposal - {company_name}")
    return output


def _safe_filename(name: str, ext: str = ".docx") -> str:
    stem = re.sub(r"[^\w\-]+", "_", name.strip())[:80].strip("_") or "Proposal"     # "Pvt. Ltd." must not leave a trailing "_"
    return f"{stem}{ext}"


# ---------------------------------------------------------------------------
# FastAPI app (self-hosted — not mounted in app.py)
# ---------------------------------------------------------------------------
app = FastAPI(title="CSS Proposal Generator", version=VERSION)
# Credentials are not allowed cross-origin: the page is served from the same origin, and the
# login cookie must never be usable by another site (a wildcard origin with credentials would
# echo any origin). API-key clients are unaffected.
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=False,
    allow_methods=["*"],
    allow_headers=["*"],
    expose_headers=["X-Proposal-Lint", "X-Proposal-Lint-Errors", "X-Proposal-Pdf"],
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
    .container {{ max-width: 860px; margin: 0 auto; background: #16213e; border-radius: 8px; padding: 24px; }}
    h1 {{ color: #8a2be2; text-align: center; margin-top: 0; }}
    h2 {{ color: #b388ff; font-size: 1.05rem; margin: 24px 0 12px; }}
    h3 {{ color: #b388ff; font-size: 0.95rem; margin: 16px 0 8px; }}
    .form-group {{ margin-bottom: 16px; }}
    label {{ display: block; margin-bottom: 6px; font-weight: bold; }}
    input[type="text"], input[type="file"], select, textarea {{ width: 100%; box-sizing: border-box; }}
    input[type="text"], select, textarea {{
      padding: 10px; border: 1px solid #444; border-radius: 4px; background: #0f3460; color: #fff;
    }}
    textarea {{ resize: vertical; min-height: 70px; }}
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
    #status {{ margin-top: 16px; min-height: 1.2em; font-weight: bold; }}
    .error {{ color: #ff6b6b; }}
    .success {{ color: #69db7c; }}
    .info {{ color: #74b9ff; }}
    .annexure-panel {{
      border: 1px solid #394867; border-radius: 6px; padding: 16px; background: #0f2744;
    }}
    .annexure-row {{
      display: grid; grid-template-columns: 120px 1fr auto; gap: 10px; align-items: start;
      margin-bottom: 14px;
    }}
    .annexure-row > label {{ margin: 0; font-weight: 600; padding-top: 10px; }}
    .annexure-slot-body {{ display: flex; flex-direction: column; gap: 6px; }}
    .annexure-actions {{ display: flex; gap: 8px; flex-wrap: wrap; margin-top: 10px; }}
    .checkbox-row {{ display: flex; align-items: center; gap: 10px; }}
    .checkbox-row input {{ width: auto; }}
    .checkbox-row label {{ margin: 0; font-weight: normal; }}
    .hidden {{ display: none; }}
    .ai-panel {{
      border: 1px solid #5e35b1; border-radius: 6px; padding: 18px; background: #130f30;
      margin-top: 20px;
    }}
    .ai-panel h2 {{ color: #ce93d8; margin-top: 0; }}
    .counts-grid {{ display: grid; grid-template-columns: 1fr 1fr; gap: 10px; }}
    .radio-group {{ display: flex; flex-wrap: wrap; gap: 14px; margin-top: 4px; }}
    .radio-group label {{ font-weight: normal; display: flex; align-items: center; gap: 6px; }}
    .radio-group input {{ width: auto; }}
    .custom-slot-fields {{
      background: #1a1060; border: 1px solid #7b1fa2; border-radius: 5px;
      padding: 12px; margin-top: 8px;
    }}
    .custom-slot-fields label {{ color: #ce93d8; font-size: 0.92em; }}
    .step-progress {{ color: #74b9ff; font-style: italic; }}
    .char-count {{ color: #888; font-size: 0.8em; text-align: right; }}
  </style>
</head>
<body>
  <div class="container">
    <h1>CSS Proposal Generator</h1>
    <p class="hint">CyberSmithSECURE &mdash; generates polished client proposals with annexure ordering and AI custom services.</p>
    <p class="hint"><a href="excel/" style="color:#b388ff">Have the filled scoping Excel? Upload it instead &rarr;</a></p>
    <form id="proposalForm">
      <!-- Basic fields -->
      <div class="form-group">
        <label for="company_name">Company name *</label>
        <input type="text" id="company_name" name="company_name" required placeholder="Acme Corp" />
      </div>
      <div class="form-group">
        <label for="proposal_ref">Proposal date (printed on the cover, DD-MM-YYYY)</label>
        <input type="text" id="proposal_ref" name="proposal_ref" value="{today_ref}" placeholder="23-05-2026" />
        <p class="hint">This date will replace any placeholders in the template.</p>
      </div>
      <div class="form-group">
        <label for="client_logo">Client logo (cover page)</label>
        <input type="file" id="client_logo" name="client_logo" accept="image/png,image/jpeg,image/jpg,image/webp" />
        <p class="hint">Upload your company logo. It will appear on the cover page.</p>
      </div>

      <!-- Annexure builder -->
      <h2>Annexures (optional)</h2>
      <div class="annexure-panel">
        <div class="checkbox-row form-group">
          <input type="checkbox" id="customize_annexures" />
          <label for="customize_annexures">Customize which annexures to include and their order</label>
        </div>
        <p class="hint">Leave unchecked to keep all 13 annexures in the template default order.</p>
        <div id="annexureBuilder" class="hidden">
          <p class="hint">Pick Annexure 1 first, then Annexure 2, and so on.
             Select <strong>&#10022; Other / Custom Service </strong> to generate a new annexure using AI.</p>
          <div id="annexureSlots"></div>
          <div class="annexure-actions">
            <button type="button" class="secondary" id="addAnnexureBtn">+ Add annexure slot</button>
            <button type="button" class="secondary" id="clearAnnexuresBtn">Clear all</button>
          </div>
        </div>
      </div>

      <!-- AI Context Panel (shown when any slot is __other__) -->
      <div id="aiContextPanel" class="ai-panel hidden">
        <h2>&#128161; AI Context &mdash; Custom Service Generation</h2>
        <p class="hint">This information feeds the Master BD Prompt to generate the custom annexure.</p>

        <div class="form-group">
          <label for="ai_sector">Sector / Industry *</label>
          <input type="text" id="ai_sector" placeholder="e.g. Banking, Insurance, Healthcare, Fintech" />
        </div>

        <div class="form-group">
          <label for="ai_engagement_summary">Overall Engagement Summary * <span class="hint">(1&ndash;2 lines)</span></label>
          <textarea id="ai_engagement_summary" rows="3"
            placeholder="e.g. Comprehensive cloud security and identity assessment for Acme Corp&rsquo;s hybrid infrastructure."></textarea>
        </div>

        <div class="form-group">
          <label>Primary Tier *</label>
          <div class="radio-group" id="ai_tier_group">
            <label><input type="radio" name="ai_primary_tier" value="Tier 1 \u2013 Essential" /> Tier 1 &ndash; Essential</label>
            <label><input type="radio" name="ai_primary_tier" value="Tier 2 \u2013 Enhanced" /> Tier 2 &ndash; Enhanced</label>
            <label><input type="radio" name="ai_primary_tier" value="Tier 3 \u2013 Advanced" checked /> Tier 3 &ndash; Advanced</label>
            <label><input type="radio" name="ai_primary_tier" value="Tier 4 \u2013 Offensive" /> Tier 4 &ndash; Offensive</label>
          </div>
        </div>

        <div class="form-group">
          <label for="ai_tier_notes">Tier-specific notes <span class="hint">(optional)</span></label>
          <textarea id="ai_tier_notes" rows="2" placeholder="e.g. Client wants Tier 3 for cloud but Tier 2 for endpoint."></textarea>
        </div>

        <h3>Counts / Volumes <span class="hint">(fill what applies)</span></h3>
        <div class="counts-grid">
          <div class="form-group">
            <label for="ai_counts_ad">AD / Identity</label>
            <input type="text" id="ai_counts_ad" placeholder="e.g. 2 AD Domains" />
          </div>
          <div class="form-group">
            <label for="ai_counts_cloud">Cloud</label>
            <input type="text" id="ai_counts_cloud" placeholder="e.g. AWS + Azure" />
          </div>
          <div class="form-group">
            <label for="ai_counts_apps">Apps / APIs</label>
            <input type="text" id="ai_counts_apps" placeholder="e.g. 5 web apps" />
          </div>
          <div class="form-group">
            <label for="ai_counts_infra">Infra</label>
            <input type="text" id="ai_counts_infra" placeholder="e.g. 50 IPs" />
          </div>
          <div class="form-group">
            <label for="ai_counts_users">Users / Employees</label>
            <input type="text" id="ai_counts_users" placeholder="e.g. 200 employees" />
          </div>
          <div class="form-group">
            <label for="ai_counts_duration">Duration *</label>
            <input type="text" id="ai_counts_duration" placeholder="e.g. 4 weeks" />
          </div>
        </div>

        <h3>Commercials</h3>
        <div class="counts-grid">
          <div class="form-group">
            <label for="ai_total_cost">Total Cost (before tax) *</label>
            <input type="text" id="ai_total_cost" placeholder="e.g. 4,33,000 INR" />
          </div>
          <div class="form-group">
            <label for="ai_taxes">Taxes</label>
            <input type="text" id="ai_taxes" value="18% GST" />
          </div>
        </div>

        <div class="form-group">
          <label>Payment Model *</label>
          <div class="radio-group" id="ai_payment_group">
            <label><input type="radio" name="ai_payment_model" value="milestone_v1" checked /> Milestone v1</label>
            <label><input type="radio" name="ai_payment_model" value="hybrid_v2" /> Hybrid v2</label>
            <label><input type="radio" name="ai_payment_model" value="retainer_v3" /> Retainer v3</label>
            <label><input type="radio" name="ai_payment_model" value="tier_v4" /> Tier v4</label>
          </div>
          <p class="hint" id="payment_model_hint">20% advance &rarr; 30% draft report &rarr; 30% revalidation &rarr; 20% closure</p>
        </div>

        <div class="form-group">
          <label for="ai_special_constraints">Special Constraints / Notes <span class="hint">(optional)</span></label>
          <textarea id="ai_special_constraints" rows="2"
            placeholder="e.g. Zero-trust environment, restricted change windows, BCP-sensitive systems."></textarea>
        </div>
      </div>

      <button type="submit" class="submit-btn" id="generateBtn">Generate Proposal</button>
    </form>
    <div id="status"></div>
  </div>

  <script>
    /* ------------------------------------------------------------------ */
    /* Catalog + Other option                                               */
    /* ------------------------------------------------------------------ */
    const BASE_CATALOG = {catalog_json};
    const OTHER_ITEM = {{ key: '__other__', title: '\u2726 Other / Custom Service (AI-Generated)' }};
    const CATALOG = [...BASE_CATALOG, OTHER_ITEM];

    const PAYMENT_HINTS = {{
      milestone_v1: '20% advance \u2192 30% draft report \u2192 30% revalidation \u2192 20% closure',
      hybrid_v2:    '40% advance \u2192 20% week-2 \u2192 20% technical report \u2192 20% at day-30 / revalidation',
      retainer_v3:  '30% advance \u2192 70% in 3 equal monthly instalments (irrespective of client delays)',
      tier_v4:      'Per-tier payment breakup (Advanced vs Offensive etc.)',
    }};

    const slotsEl       = document.getElementById('annexureSlots');
    const customizeEl   = document.getElementById('customize_annexures');
    const builderEl     = document.getElementById('annexureBuilder');
    const aiPanelEl     = document.getElementById('aiContextPanel');
    const paymentHintEl = document.getElementById('payment_model_hint');
    let slotSelections  = [];   // array of strings (catalog key or '' or '__other__<unique>')
    let otherSlotMeta   = {{}};  // key -> {{ title, scope }}

    /* payment model hint */
    document.querySelectorAll('input[name="ai_payment_model"]').forEach(r => {{
      r.addEventListener('change', () => {{ paymentHintEl.textContent = PAYMENT_HINTS[r.value] || ''; }});
    }});

    /* ------------------------------------------------------------------ */
    /* Unique key per "Other" slot                                          */
    /* ------------------------------------------------------------------ */
    function makeOtherKey() {{
      return `__other__${{Date.now()}}_${{Math.random().toString(36).slice(2, 7)}}`;
    }}

    function selectedKeys() {{
      return slotSelections.filter(Boolean);
    }}

    function hasAnyOther() {{
      return slotSelections.some(k => k && k.startsWith('__other__'));
    }}

    /* ------------------------------------------------------------------ */
    /* Available options for a slot (no duplicates; always show Other)     */
    /* ------------------------------------------------------------------ */
    function availableForSlot(slotIndex) {{
      const used = new Set(selectedKeys());
      const current = slotSelections[slotIndex];
      if (current) used.delete(current);
      // Normal catalog entries (no duplicates); Other is always available
      const normal = BASE_CATALOG.filter(item => !used.has(item.key));
      return [...normal, OTHER_ITEM];
    }}

    /* ------------------------------------------------------------------ */
    /* Render all slots                                                     */
    /* ------------------------------------------------------------------ */
    function renderSlots() {{
      slotsEl.innerHTML = '';
      slotSelections.forEach((selected, index) => {{
        const isOther = selected && selected.startsWith('__other__');
        const row = document.createElement('div');
        row.className = 'annexure-row';

        const labelEl = document.createElement('label');
        labelEl.textContent = `Annexure ${{index + 1}}`;

        const body = document.createElement('div');
        body.className = 'annexure-slot-body';

        /* dropdown */
        const select = document.createElement('select');
        select.dataset.index = String(index);
        const emptyOpt = new Option('-- Select assessment --', '');
        select.appendChild(emptyOpt);
        availableForSlot(index).forEach(item => {{
          const isThisOther = item.key === '__other__';
          const displayKey = isThisOther ? (isOther ? selected : '') : item.key;
          const opt = new Option(item.title, isThisOther ? '__other__' : item.key);
          if ((isOther && isThisOther) || selected === item.key) opt.selected = true;
          select.appendChild(opt);
        }});

        select.addEventListener('change', e => {{
          const val = e.target.value;
          if (val === '__other__') {{
            const uniqueKey = makeOtherKey();
            slotSelections[index] = uniqueKey;
            otherSlotMeta[uniqueKey] = {{ title: '', scope: '' }};
          }} else {{
            // If previous was __other__, clean up its meta
            if (slotSelections[index] && slotSelections[index].startsWith('__other__')) {{
              delete otherSlotMeta[slotSelections[index]];
            }}
            slotSelections[index] = val || '';
          }}
          renderSlots();
          aiPanelEl.classList.toggle('hidden', !hasAnyOther());
        }});
        body.appendChild(select);

        /* per-slot custom fields (only when this slot is __other__) */
        if (isOther) {{
          const meta = otherSlotMeta[selected] || {{ title: '', scope: '' }};
          const fields = document.createElement('div');
          fields.className = 'custom-slot-fields';

          const titleLabel = document.createElement('label');
          titleLabel.textContent = 'Service / Annexure Title *';
          const titleInput = document.createElement('input');
          titleInput.type = 'text';
          titleInput.maxLength = 100;
          titleInput.placeholder = 'e.g. Cloud Red Teaming Assessment';
          titleInput.value = meta.title;
          titleInput.addEventListener('input', () => {{
            otherSlotMeta[selected].title = titleInput.value;
          }});

          const scopeLabel = document.createElement('label');
          scopeLabel.textContent = 'High-Level Scope * (20\u2013500 chars)';
          const scopeArea = document.createElement('textarea');
          scopeArea.rows = 3;
          scopeArea.maxLength = 500;
          scopeArea.placeholder =
            'e.g. Simulate adversary attack paths on cloud infrastructure (AWS/GCP), identify privilege escalation paths, lateral movement opportunities, and critical data exposure.';
          scopeArea.value = meta.scope;
          const charCount = document.createElement('div');
          charCount.className = 'char-count';
          charCount.textContent = `${{meta.scope.length}}/500`;
          scopeArea.addEventListener('input', () => {{
            otherSlotMeta[selected].scope = scopeArea.value;
            charCount.textContent = `${{scopeArea.value.length}}/500`;
          }});

          fields.appendChild(titleLabel);
          fields.appendChild(titleInput);
          fields.appendChild(scopeLabel);
          fields.appendChild(scopeArea);
          fields.appendChild(charCount);
          body.appendChild(fields);
        }}

        const removeBtn = document.createElement('button');
        removeBtn.type = 'button';
        removeBtn.className = 'danger';
        removeBtn.textContent = 'Remove';
        removeBtn.style.alignSelf = 'flex-start';
        removeBtn.addEventListener('click', () => {{
          if (slotSelections[index] && slotSelections[index].startsWith('__other__')) {{
            delete otherSlotMeta[slotSelections[index]];
          }}
          slotSelections.splice(index, 1);
          renderSlots();
          aiPanelEl.classList.toggle('hidden', !hasAnyOther());
        }});

        row.appendChild(labelEl);
        row.appendChild(body);
        row.appendChild(removeBtn);
        slotsEl.appendChild(row);
      }});
    }}

    function addSlot() {{
      slotSelections.push('');
      renderSlots();
    }}

    customizeEl.addEventListener('change', () => {{
      builderEl.classList.toggle('hidden', !customizeEl.checked);
      if (customizeEl.checked && slotSelections.length === 0) addSlot();
      if (!customizeEl.checked) aiPanelEl.classList.add('hidden');
    }});

    document.getElementById('addAnnexureBtn').addEventListener('click', addSlot);
    document.getElementById('clearAnnexuresBtn').addEventListener('click', () => {{
      slotSelections = [];
      otherSlotMeta = {{}};
      renderSlots();
      aiPanelEl.classList.add('hidden');
    }});

    /* ------------------------------------------------------------------ */
    /* Build annexure payload (including other_ keys)                      */
    /* ------------------------------------------------------------------ */
    function buildAnnexurePayload() {{
      if (!customizeEl.checked) return null;
      const order = selectedKeys();
      if (!order.length) return null;
      return {{ enabled: true, order }};
    }}

    /* ------------------------------------------------------------------ */
    /* Build AI context JSON from panel                                     */
    /* ------------------------------------------------------------------ */
    function buildAIContext() {{
      if (!hasAnyOther()) return null;
      const tierEl = document.querySelector('input[name="ai_primary_tier"]:checked');
      const paymentEl = document.querySelector('input[name="ai_payment_model"]:checked');
      const customAnnexures = [];
      slotSelections.forEach((key, idx) => {{
        if (key && key.startsWith('__other__')) {{
          const meta = otherSlotMeta[key] || {{}};
          customAnnexures.push({{
            slot_index: idx,
            key: key,
            title: (meta.title || '').trim(),
            scope: (meta.scope || '').trim(),
          }});
        }}
      }});
      return {{
        sector: document.getElementById('ai_sector').value.trim(),
        engagement_summary: document.getElementById('ai_engagement_summary').value.trim(),
        primary_tier: tierEl ? tierEl.value : 'Tier 3 \u2013 Advanced',
        tier_notes: document.getElementById('ai_tier_notes').value.trim(),
        counts: {{
          ad_identity: document.getElementById('ai_counts_ad').value.trim(),
          cloud:       document.getElementById('ai_counts_cloud').value.trim(),
          apps_apis:   document.getElementById('ai_counts_apps').value.trim(),
          infra:       document.getElementById('ai_counts_infra').value.trim(),
          users:       document.getElementById('ai_counts_users').value.trim(),
          duration:    document.getElementById('ai_counts_duration').value.trim(),
        }},
        commercials: {{
          total_cost: document.getElementById('ai_total_cost').value.trim(),
          taxes:      document.getElementById('ai_taxes').value.trim() || '18% GST',
        }},
        payment_model: paymentEl ? paymentEl.value : 'hybrid_v2',
        special_constraints: document.getElementById('ai_special_constraints').value.trim(),
        custom_annexures: customAnnexures,
      }};
    }}

    /* ------------------------------------------------------------------ */
    /* Form validation helpers                                              */
    /* ------------------------------------------------------------------ */
    function validateAIContext(ctx) {{
      if (!ctx) return null;
      if (!ctx.sector) return 'AI Context: Sector / Industry is required.';
      if (ctx.engagement_summary.length < 10) return 'AI Context: Engagement Summary must be at least 10 characters.';
      if (!ctx.counts.duration) return 'AI Context: Duration is required.';
      if (!ctx.commercials.total_cost) return 'AI Context: Total Cost is required.';
      for (const ca of ctx.custom_annexures) {{
        if (!ca.title) return `Annexure slot ${{ca.slot_index + 1}}: Service Title is required.`;
        if (ca.scope.length < 20) return `Annexure "${{ca.title}}": Scope must be at least 20 characters.`;
      }}
      return null;
    }}

    /* ------------------------------------------------------------------ */
    /* Submit                                                               */
    /* ------------------------------------------------------------------ */
    document.getElementById('proposalForm').addEventListener('submit', async (e) => {{
      e.preventDefault();
      const statusEl = document.getElementById('status');
      const btn = document.getElementById('generateBtn');

      const setStatus = (msg, cls='') => {{
        statusEl.textContent = msg;
        statusEl.className = cls;
      }};

      // Validate annexure slots
      if (customizeEl.checked) {{
        const filled = slotSelections.filter(s => s);
        if (filled.length !== slotSelections.length) {{
          setStatus('Please select an assessment for each slot, or remove empty ones.', 'error');
          return;
        }}
        if (!filled.length) {{
          setStatus('Add at least one annexure or uncheck customization to use the default 13.', 'error');
          return;
        }}
      }}

      // Validate AI context if needed
      const aiCtx = buildAIContext();
      if (aiCtx) {{
        const aiErr = validateAIContext(aiCtx);
        if (aiErr) {{ setStatus(aiErr, 'error'); return; }}
      }}

      btn.disabled = true;
      const hasAI = !!aiCtx;
      if (hasAI) {{
        setStatus('Step 1/3: Validating inputs\u2026', 'info step-progress');
      }} else {{
        setStatus('Generating\u2026', '');
      }}

      const formData = new FormData();
      formData.append('company_name', document.getElementById('company_name').value);
      formData.append('proposal_ref', document.getElementById('proposal_ref').value);
      const logo = document.getElementById('client_logo').files[0];
      if (logo) formData.append('client_logo', logo);

      const annexurePayload = buildAnnexurePayload();
      if (annexurePayload) formData.append('annexure_order_json', JSON.stringify(annexurePayload));
      if (aiCtx) formData.append('ai_context_json', JSON.stringify(aiCtx));

      if (hasAI) setStatus('Step 2/3: AI is generating custom annexure\u2026 (this may take 20\u201340 seconds)', 'info step-progress');

      try {{
        const resp = await fetch('generate/', {{ method: 'POST', body: formData }});
        if (!resp.ok) {{
          const err = await resp.json().catch(() => ({{ detail: resp.statusText }}));
          throw new Error(err.detail || 'Generation failed');
        }}
        if (hasAI) setStatus('Step 3/3: Building document\u2026', 'info step-progress');
        const blob = await resp.blob();
        const url  = URL.createObjectURL(blob);
        const a    = document.createElement('a');
        a.href     = url;
        a.download = 'CSS_Proposal.docx';
        document.body.appendChild(a);
        a.click();
        a.remove();
        URL.revokeObjectURL(url);
        setStatus('Proposal downloaded successfully.', 'success');
      }} catch (err) {{
        setStatus(err.message || String(err), 'error');
      }} finally {{
        btn.disabled = false;
      }}
    }});
  </script>
</body>
</html>"""
    return HTMLResponse(content=html)


@app.post("/generate/", dependencies=[Depends(rate_limit("generate"))])
async def generate_proposal_endpoint(
    company_name: str = Form(...),
    proposal_ref: str = Form(""),
    client_logo: UploadFile | None = File(None),
    annexure_order_json: str = Form(""),
    assessments_json: str = Form(""),
    ai_context_json: str = Form(""),
) -> FileResponse:
    tmp_dir: str | None = None
    logo_path: Path | None = None
    try:
        ref = (proposal_ref or "").strip() or default_proposal_ref()
        annexure_order = parse_annexure_order_json(annexure_order_json, allow_custom=True)
        assessments = parse_assessments_json(assessments_json) if not annexure_order else None

        # ---------- AI custom annexure generation ----------------------------
        request_catalog: dict[str, dict[str, Any]] | None = None
        ai_context: "AIContextSchema | None" = None

        if ai_context_json and ai_context_json.strip():
            ai_context = _parse_ai_context(ai_context_json)

        if annexure_order:
            custom_keys = [k for k in annexure_order if OTHER_KEY_RE.match(k)]
            if custom_keys:
                if ai_context is None:
                    raise HTTPException(
                        status_code=400,
                        detail=(
                            "One or more 'Other / Custom Service' slots were selected but "
                            "no AI context was provided. Please fill the AI Context panel."
                        ),
                    )
                _validate_ai_context(ai_context)

                # Index the custom annexures from ai_context by key
                ca_by_key = {ca.key: ca for ca in ai_context.custom_annexures}

                # Validate that every custom_key has a matching CustomAnnexureInput
                for ck in custom_keys:
                    if ck not in ca_by_key:
                        raise HTTPException(
                            status_code=400,
                            detail=f"No title/scope provided for custom annexure key: {ck}",
                        )

                # Build per-request catalog with AI-generated elements
                request_catalog = _make_request_catalog()
                for ck in custom_keys:
                    ca = ca_by_key[ck]
                    slot_number = annexure_order.index(ck) + 1
                    print(
                        f"[AI] Generating annexure {slot_number}: '{ca.title}' for {company_name} …"
                    )
                    raw_text = await generate_custom_annexure_text(
                        company_name=company_name,
                        context=ai_context,
                        custom_annexure=ca,
                        slot_number=slot_number,
                    )
                    elements = build_custom_annexure_block(
                        raw_ai_text=raw_text,
                        slot_number=slot_number,
                        title=ca.title,
                        client_name=company_name,
                    )
                    _register_custom_annexure(
                        catalog=request_catalog,
                        key=ck,
                        title=ca.title,
                        toc_title=ca.title,
                        slot=slot_number,
                        ai_elements=elements,
                    )

        # ---------- Document generation ---------------------------------------
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
            proposal_date=ref,               # the manual form's "reference" field is the cover date
            logo_path=str(logo_path) if logo_path else None,
            output_path=output_path,
            annexure_order=annexure_order,
            assessments=assessments,
            request_catalog=request_catalog,
        )
        keys = annexure_order or [item["key"] for item in ASSESSMENT_CATALOG]
        check = await asyncio.to_thread(finalize_with_render, output_path, keys, company_name)
        final_path = UPLOAD_DIR / out_name
        shutil.copy2(output_path, final_path)
        return FileResponse(
            path=str(final_path),
            filename=out_name,
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            headers={"X-Proposal-Pdf": check.header_value()},
        )
    except HTTPException:
        raise
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
# Excel-driven flow (upload -> two previews -> generate); see proposal_workflow.py
# ---------------------------------------------------------------------------
def _model_dict(obj) -> dict[str, Any]:
    """pydantic model or dataclass -> plain dict."""
    if hasattr(obj, "model_dump"):
        return obj.model_dump()
    import dataclasses
    return dataclasses.asdict(obj)


def _ai_cache_key(company_name: str, context, annexure, slot_number: int) -> str:
    """Identity of one AI annexure request, so a rebuild after a price/frequency edit reuses the text.

    Pricing and payment model are left out on purpose: the annexure prompt forbids pricing content,
    so changing them must not trigger (and pay for) a new generation.
    """
    import hashlib
    ctx = _model_dict(context)
    for volatile in ("custom_annexures", "commercials", "payment_model"):
        ctx.pop(volatile, None)
    ca = _model_dict(annexure)
    ca.pop("key", None)                                   # random per build
    payload = json.dumps({"c": company_name, "s": slot_number, "ctx": ctx, "ca": ca}, sort_keys=True, default=str)
    return hashlib.sha256(payload.encode("utf-8")).hexdigest()


async def build_from_plan(draft, plan, ai_inputs_fn, client_logo, store, draft_id) -> dict[str, Any]:
    """Build the proposal from a reviewed draft (called by the workflow router).

    AI annexures are written first (with Excel context, one automatic retry; texts are cached on the
    draft so a rebuild does not pay again), then the document is assembled through the same
    generate_proposal() the manual flow uses, corrected from a real render and checked.
    Returns {"final_path", "out_name", "check", "report", "headers"} and keeps the .docx (and the PDF
    proof when an engine exists) with the draft for download.
    """
    tmp_dir: str | None = None
    try:
        company_name = plan["company_name"]
        request_catalog: dict[str, dict[str, Any]] | None = None

        if plan["customs"]:
            ai_context = _parse_ai_context(json.dumps(ai_inputs_fn(draft, plan)))
            _validate_ai_context(ai_context, require_total_cost=False)   # D18
            request_catalog = _make_request_catalog()
            by_key = {ca.key: ca for ca in ai_context.custom_annexures}
            cache: dict[str, str] = draft.setdefault("ai_cache", {})
            for c in plan["customs"]:
                ca = by_key[c["key"]]
                slot_number = plan["ordered_keys"].index(c["key"]) + 1
                cache_key = _ai_cache_key(company_name, ai_context, ca, slot_number)
                raw_text = cache.get(cache_key)
                if raw_text is None:
                    raw_text = await generate_custom_annexure_text(
                        company_name=company_name, context=ai_context,
                        custom_annexure=ca, slot_number=slot_number)
                    cache[cache_key] = raw_text
                _register_custom_annexure(
                    catalog=request_catalog, key=c["key"], title=ca.title, toc_title=ca.title,
                    slot=slot_number,
                    ai_elements=build_custom_annexure_block(
                        raw_ai_text=raw_text, slot_number=slot_number, title=ca.title,
                        client_name=company_name))

        tmp_dir = tempfile.mkdtemp(prefix="css_proposal_")
        logo_path: Path | None = None
        if client_logo is not None and client_logo.filename:
            ext = Path(client_logo.filename).suffix.lower()
            if ext not in (".png", ".jpg", ".jpeg"):
                raise HTTPException(status_code=400, detail="The logo must be a PNG or JPG image.")
            data = await client_logo.read()
            if len(data) > 5 * 1024 * 1024:
                raise HTTPException(status_code=413, detail="The logo is larger than 5 MB.")
            logo_path = Path(tmp_dir) / f"client_logo{ext}"
            logo_path.write_bytes(data)

        out_name = _safe_filename(company_name)
        output_path = Path(tmp_dir) / out_name
        generate_proposal(
            company_name=company_name,
            proposal_ref=plan["reference"],
            proposal_date=plan.get("proposal_date"),
            logo_path=str(logo_path) if logo_path else None,
            output_path=output_path,
            annexure_order=plan["ordered_keys"],
            request_catalog=request_catalog,
            annexure_details=plan["annexure_details"],
            pricing_rows=plan["pricing_rows"],
        )
        # Real pagination: correct the Contents numbers and look for blank pages (needs a PDF engine).
        check = await asyncio.to_thread(finalize_with_render, output_path, plan["ordered_keys"], company_name)
        report = proposal_lint.lint_docx(output_path)
        final_path = UPLOAD_DIR / out_name
        shutil.copy2(output_path, final_path)
        headers = {"X-Proposal-Lint": f"{len(report.errors)} error(s), {len(report.warnings)} warning(s)",
                   "X-Proposal-Pdf": check.header_value()}
        problems = [f.message for f in report.errors[:3]]
        if check.verified and check.blank_pages:
            problems.append(f"blank page(s) {check.blank_pages}")
        if problems:
            headers["X-Proposal-Lint-Errors"] = "; ".join(problems).encode("ascii", "replace").decode("ascii")
        if store is not None:
            store.update(draft_id, draft)                        # keeps the AI texts for a cheap rebuild
            store.attach(draft_id, ".docx", final_path)
            if check.verified and check.pdf_path:
                store.attach(draft_id, ".pdf", check.pdf_path)   # the PDF proof / final preview
        return {"final_path": final_path, "out_name": out_name, "check": check, "report": report, "headers": headers}
    except HTTPException:
        raise
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    except Exception as exc:
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=str(exc)) from exc
    finally:
        if tmp_dir and os.path.isdir(tmp_dir):
            shutil.rmtree(tmp_dir, ignore_errors=True)


async def generate_from_plan(draft, plan, ai_inputs_fn, client_logo, store, draft_id) -> FileResponse:
    """Build the proposal and return the .docx directly (API clients; the page uses build_from_plan)."""
    built = await build_from_plan(draft, plan, ai_inputs_fn, client_logo, store, draft_id)
    return FileResponse(
        path=str(built["final_path"]), filename=built["out_name"], headers=built["headers"],
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document")


import proposal_workflow  # noqa: E402  (after app/helpers exist; the router gets this module)

proposal_router = proposal_workflow.create_router(sys.modules[__name__])
app.include_router(proposal_router)


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

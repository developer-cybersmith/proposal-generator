#!/usr/bin/env python3
"""
CSS Proposal Generator — standalone single-file module.

Does not modify app.py or any other project files. Run directly:

    python css_proposal_generator.py              # web UI on http://0.0.0.0:8010/
    python css_proposal_generator.py --build-template   # rebuild base .docx from source

Assets expected in the same folder as this script:
    CSS_Proposal_Source.docx       (original template; replace and run --build-template to update)
    CSS_Proposal_Base_Template.docx (prepared docxtpl template; auto-built if missing)
    default_logo.png               (optional fallback logo)
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
from datetime import datetime
from pathlib import Path
from typing import Any

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor
from docxtpl import DocxTemplate, InlineImage
from fastapi import FastAPI, File, Form, HTTPException, UploadFile
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
import proposal_pricing  # noqa: E402
import proposal_tokens  # noqa: E402

SOURCE_TEMPLATE = HERE / "CSS_Proposal_Source.docx"
BASE_TEMPLATE = HERE / "CSS_Proposal_Base_Template.docx"
DEFAULT_LOGO = HERE / "default_logo.png"
UPLOAD_DIR = HERE / "uploads" / "css_proposal"
MASTER_BD_PROMPT_FILE = HERE / "master_bd_prompt.txt"
VERSION = "1.3.0"
DEFAULT_PORT = 8010

# ---------------------------------------------------------------------------
# CyberSmithSECURE brand formatting constants — LOCKED, DO NOT CHANGE
# Verified against CyberSmithSECURE Annexures 2.docx
# ---------------------------------------------------------------------------
CSS_FONT_NAME = "Corbel"
CSS_ANNEXURE_HEADING_SIZE = Pt(16)   # Bold + Italic
CSS_SECTION_LABEL_SIZE = Pt(14)      # Bold
CSS_BODY_SIZE = Pt(11)               # Regular
CSS_HEADING_COLOR = RGBColor(0x70, 0x30, 0xA0)  # #7030A0 — matches static template annexures
CSS_BODY_COLOR = RGBColor(0, 0, 0)   # Black — body text and bullets only
CSS_MAIN_BULLET = "\u2022"           # U+2022 •
CSS_SUB_BULLET = "o"                 # lowercase o (per Master BD Prompt)
CSS_MAIN_BULLET_INDENT = "360"       # twips (~0.63 cm)
CSS_SUB_BULLET_INDENT = "720"        # twips (~1.27 cm)
CSS_BULLET_HANGING = "240"           # twips

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
CONTENTS_HEADING = "Contents"
CONCLUSION_TOC_PREFIX = "Conclusion & Pricing"

# Default Contents page numbers for the unchanged intro block (from base template).
DEFAULT_STATIC_TOC_PAGES: tuple[int, ...] = (3, 4, 4, 5, 7, 8)
DEFAULT_FIRST_ANNEXURE_PAGE = 13
TERMS_MARKER = "Terms & Conditions"
TERMS_TOC_LABEL = "Terms & Conditions"
# Compact Contents spacing (twips): keeps all TOC lines on one page.
TOC_LINE_BEFORE_TWIPS = "40"       # ~2pt between entries
TOC_FIRST_LINE_BEFORE_TWIPS = "100"  # ~5pt after "Contents" heading
TOC_TAB_POS = "9887"
STATIC_TOC_ANCHORS: tuple[tuple[str, int], ...] = (
    ("INTRODUCTION", 37),
    ("OUR UNDERSTANDING OF YOUR REQUIREMENT", 37),
    ("METHODOLOGY OVERVIEW", 37),
    ("VULNERABILITY ASSESSMENT METHODOLOGY", 37),
    ("PROPOSED TIMELINES & PROJECT MANAGEMENT", 47),
    ("TEAM OVERVIEW", 55),
)

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

ASSESSMENT_CATALOG: tuple[dict[str, Any], ...] = (
    {"key": "post_breach", "title": "Post Breach Scenario Assessment", "toc_title": "Antivirus Bypass Analysis", "annexure_number": 1},
    {"key": "active_directory", "title": "Active Directory Assessment", "toc_title": "Active Directory Post Breach Security Attacks", "annexure_number": 2},
    {"key": "network_pt", "title": "Network Penetration Testing Assessment", "toc_title": "Network Penetration testing", "annexure_number": 3},
    {"key": "web_app_pt", "title": "Web Application Penetration Testing Assessment", "toc_title": "Web Application Penetration testing", "annexure_number": 4},
    {"key": "network_config_review", "title": "Network Configuration Review Assessment", "toc_title": "Network Penetration Review Assessment", "annexure_number": 5},
    {"key": "insider_threat", "title": "Insider Threat Assessment for Data Infiltration & Exfiltration", "toc_title": "Insider Threat Assessment for Data Infiltration and Exfiltration", "annexure_number": 6},
    {"key": "mobile_app_pt", "title": "Mobile Application Penetration Testing", "toc_title": "Mobile Application Penetration testing", "annexure_number": 7},
    {"key": "ad_security", "title": "Active Directory Security Assessment", "toc_title": "Active Directory Security Assessment", "annexure_number": 8},
    {"key": "soc_simulation", "title": "SOC Simulation Exercise", "toc_title": "SOC Simulation Exercise", "annexure_number": 9},
    {"key": "cloud_pt", "title": "Cloud Penetration Testing", "toc_title": "Cloud Penetration testing", "annexure_number": 10},
    {"key": "api_pt", "title": "API Penetration Testing Assessment", "toc_title": "API Penetration testing Assessment", "annexure_number": 11},
    {"key": "offensive_defensive_training", "title": "Offensive & Defensive Security Training", "toc_title": "Offensive & Defensive security Training", "annexure_number": 12},
    {"key": "source_code_review", "title": "Source Code Review", "toc_title": "Source Code Review", "annexure_number": 13},
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
        )
        for ca in data.get("custom_annexures", [])
    ]
    return ctx


def _validate_ai_context(ctx: "AIContextSchema") -> None:
    """Raise ValueError for missing required AI context fields."""
    if not ctx.sector.strip():
        raise ValueError("AI context: 'sector' is required when generating a custom annexure.")
    if len(ctx.engagement_summary.strip()) < 10:
        raise ValueError("AI context: 'engagement_summary' must be at least 10 characters.")
    if not ctx.counts.duration.strip():
        raise ValueError("AI context: 'counts.duration' is required.")
    if not ctx.commercials.total_cost.strip():
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
TOTAL COST (before tax): {context.commercials.total_cost}
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

    api_key = (os.getenv("OPENAI_API_KEY") or "").strip()
    if not api_key:
        raise HTTPException(
            status_code=500,
            detail="AI generation is not configured. Set OPENAI_API_KEY in .env.",
        )

    system_prompt = _load_master_bd_prompt()
    user_message = _build_ai_user_message(company_name, context, custom_annexure, slot_number)

    try:
        from openai import AsyncOpenAI, APITimeoutError, RateLimitError
        client = AsyncOpenAI(api_key=api_key)
        model = (os.getenv("OPENAI_MODEL") or "gpt-4o").strip()

        response = await client.chat.completions.create(
            model=model,
            temperature=0.2,
            max_tokens=3500,
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_message},
            ],
        )
        raw_text = (response.choices[0].message.content or "").strip()
        validate_ai_annexure_output(raw_text, custom_annexure.title)
        return raw_text

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
        p = doc.add_paragraph()
        # Keep a small spacing to match template blank lines
        pPr = p._p.get_or_add_pPr()
        sp = OxmlElement("w:spacing")
        sp.set(qn("w:before"), "0")
        sp.set(qn("w:after"), "0")
        pPr.append(sp)
        return

    p = doc.add_paragraph()

    if style_name in ("main_bullet", "sub_bullet"):
        bullet_char = CSS_MAIN_BULLET if style_name == "main_bullet" else CSS_SUB_BULLET
        run_text = f"{bullet_char} {text}"
    else:
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

    # Bullet indentation
    if style_name in ("main_bullet", "sub_bullet"):
        left = CSS_MAIN_BULLET_INDENT if style_name == "main_bullet" else CSS_SUB_BULLET_INDENT
        pPr = p._p.get_or_add_pPr()
        ind = OxmlElement("w:ind")
        ind.set(qn("w:left"), left)
        ind.set(qn("w:hanging"), CSS_BULLET_HANGING)
        pPr.append(ind)

    # Zero spacing before/after body text (matches template)
    pPr = p._p.get_or_add_pPr()
    sp = OxmlElement("w:spacing")
    sp.set(qn("w:before"), "0")
    sp.set(qn("w:after"), "0")
    pPr.append(sp)


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

    scratch = Document()
    # Remove the default empty paragraph python-docx always adds
    for p in list(scratch.paragraphs):
        p._element.getparent().remove(p._element)

    for line in text.split("\n"):
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
    proposal_pricing.replace_pricing_table(tbl, rows, text_width_twips or _DEFAULT_TEXT_WIDTH_TWIPS)


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


def _find_contents_region(doc: Document) -> tuple[int | None, int | None, int | None, int | None]:
    """Return (contents_heading_idx, first_assessment_toc_idx, terms_toc_idx, conclusion_toc_idx)."""
    boundaries = _find_annexure_boundaries(doc)
    first_annexure_idx = boundaries[0][0] if boundaries else len(doc.paragraphs)
    contents_idx: int | None = None
    first_assessment_idx: int | None = None
    terms_idx: int | None = None
    conclusion_idx: int | None = None

    for i in range(first_annexure_idx):
        text = (doc.paragraphs[i].text or "").strip()
        if text == CONTENTS_HEADING:
            contents_idx = i
        elif text.startswith("TEAM OVERVIEW") and "\t" in text:
            first_assessment_idx = i + 1
        elif text.startswith(TERMS_TOC_LABEL) and "\t" in text:
            terms_idx = i
        elif text.startswith(CONCLUSION_TOC_PREFIX) and "\t" in text:
            conclusion_idx = i

    if first_assessment_idx is None and contents_idx is not None:
        first_assessment_idx = contents_idx + 1 + len(STATIC_TOC_ANCHORS)

    if terms_idx is None and conclusion_idx is not None and conclusion_idx > 0:
        prev = (doc.paragraphs[conclusion_idx - 1].text or "").strip()
        if prev.startswith(TERMS_TOC_LABEL) and "\t" in prev:
            terms_idx = conclusion_idx - 1

    return contents_idx, first_assessment_idx, terms_idx, conclusion_idx


_TOC_PPR_CACHE: Any = None


def _reference_toc_ppr() -> Any:
    """Cached paragraph properties from a template Contents line (incl. dot leaders)."""
    global _TOC_PPR_CACHE
    if _TOC_PPR_CACHE is not None:
        return _TOC_PPR_CACHE
    if not BASE_TEMPLATE.is_file():
        return None
    ref = Document(str(BASE_TEMPLATE))
    _, first_assessment_idx, _, _ = _find_contents_region(ref)
    sample_idx = min(first_assessment_idx or 25, len(ref.paragraphs) - 1)
    ppr = ref.paragraphs[sample_idx]._element.find(qn("w:pPr"))
    _TOC_PPR_CACHE = copy.deepcopy(ppr) if ppr is not None else None
    return _TOC_PPR_CACHE


def _sanitize_toc_ppr(ppr) -> None:
    """Remove layout/highlight artifacts copied from the template TOC lines."""
    for tag in (qn("w:pageBreakBefore"), qn("w:keepNext"), qn("w:pageBreakAfter")):
        node = ppr.find(tag)
        if node is not None:
            ppr.remove(node)
    rpr = ppr.find(qn("w:rPr"))
    if rpr is not None:
        highlight = rpr.find(qn("w:highlight"))
        if highlight is not None:
            rpr.remove(highlight)
        if not len(rpr):
            ppr.remove(rpr)


def _apply_compact_toc_ppr(p_element, *, first_line: bool = False) -> None:
    """Apply compact Contents paragraph formatting with right-aligned dot leader tab."""
    existing = p_element.find(qn("w:pPr"))
    if existing is not None:
        p_element.remove(existing)

    ref = _reference_toc_ppr()
    if ref is not None:
        ppr = copy.deepcopy(ref)
        _sanitize_toc_ppr(ppr)
    else:
        ppr = OxmlElement("w:pPr")
        tabs = OxmlElement("w:tabs")
        tab = OxmlElement("w:tab")
        tab.set(qn("w:val"), "right")
        tab.set(qn("w:leader"), "dot")
        tab.set(qn("w:pos"), TOC_TAB_POS)
        tabs.append(tab)
        ppr.append(tabs)

    spacing = ppr.find(qn("w:spacing"))
    if spacing is None:
        spacing = OxmlElement("w:spacing")
        ppr.append(spacing)
    spacing.set(qn("w:before"), TOC_FIRST_LINE_BEFORE_TWIPS if first_line else TOC_LINE_BEFORE_TWIPS)
    spacing.set(qn("w:after"), "0")
    spacing.set(qn("w:line"), "240")
    spacing.set(qn("w:lineRule"), "auto")

    _sanitize_toc_ppr(ppr)

    # Ensure dot-leader tab stop is present.
    tabs_el = ppr.find(qn("w:tabs"))
    if tabs_el is None:
        tabs_el = OxmlElement("w:tabs")
        ppr.append(tabs_el)
    if not tabs_el.findall(qn("w:tab")):
        tab = OxmlElement("w:tab")
        tab.set(qn("w:val"), "right")
        tab.set(qn("w:leader"), "dot")
        tab.set(qn("w:pos"), TOC_TAB_POS)
        tabs_el.append(tab)

    p_element.insert(0, ppr)


def _set_toc_line(p_element, title: str, page: int, *, first_line: bool = False) -> None:
    """Rebuild a Contents line as ``Title<TAB>Page`` with compact layout."""
    _apply_compact_toc_ppr(p_element, first_line=first_line)
    for child in list(p_element):
        if child.tag != qn("w:pPr"):
            p_element.remove(child)

    r_title = OxmlElement("w:r")
    t_title = OxmlElement("w:t")
    t_title.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    t_title.text = title
    r_title.append(t_title)
    p_element.append(r_title)

    r_tab = OxmlElement("w:r")
    r_tab.append(OxmlElement("w:tab"))
    p_element.append(r_tab)

    r_page = OxmlElement("w:r")
    t_page = OxmlElement("w:t")
    t_page.text = str(page)
    r_page.append(t_page)
    p_element.append(r_page)


def _parse_toc_line(text: str) -> tuple[str, int]:
    parts = (text or "").split("\t")
    title = parts[0].strip() if parts else ""
    page_raw = parts[1].strip() if len(parts) > 1 else "1"
    page_digits = re.sub(r"[^\d]", "", page_raw)
    return title, int(page_digits) if page_digits else 1


def _find_introduction_body_idx(doc: Document) -> int | None:
    """Return paragraph index of the body INTRODUCTION heading (not the Contents line)."""
    _, _, _, conclusion_idx = _find_contents_region(doc)
    if conclusion_idx is None:
        return None
    for i in range(conclusion_idx + 1, len(doc.paragraphs)):
        text = (doc.paragraphs[i].text or "").strip()
        if text == "INTRODUCTION" and "\t" not in text:
            return i
    return None


def _ensure_page_break_before_introduction(doc: Document) -> None:
    """Start the Introduction section on the page after Contents."""
    intro_idx = _find_introduction_body_idx(doc)
    if intro_idx is None:
        return

    if intro_idx > 0:
        prev_el = doc.paragraphs[intro_idx - 1]._element
        if _is_page_break_only_paragraph(prev_el) or _has_page_break(prev_el):
            return

    intro_el = doc.paragraphs[intro_idx]._element
    if _has_page_break_before(intro_el.find(qn("w:pPr"))):
        return

    _insert_paragraph_before(doc, intro_idx, _make_page_break_paragraph())


def _normalize_contents_section(doc: Document) -> None:
    """Compact Contents layout: tight spacing, no page breaks, single-page TOC block."""
    contents_idx, first_assessment_idx, _, conclusion_idx = _find_contents_region(doc)
    if contents_idx is None or conclusion_idx is None:
        return

    # Drop blank lines and page-break-only paragraphs inside the Contents block.
    for i in range(conclusion_idx - 1, contents_idx, -1):
        if i >= len(doc.paragraphs):
            continue
        para = doc.paragraphs[i]
        if _is_page_break_only_paragraph(para._element):
            _remove_paragraph_at(doc, i)
            continue
        txt = (para.text or "").strip()
        if txt:
            continue
        if para._element.findall(f".//{PIC_TAG}"):
            continue
        _remove_paragraph_at(doc, i)

    _, first_assessment_idx, _, conclusion_idx = _find_contents_region(doc)
    if conclusion_idx is None:
        return

    for i in range(contents_idx + 1, conclusion_idx + 1):
        if i >= len(doc.paragraphs):
            break
        p_element = doc.paragraphs[i]._element
        for br in list(p_element.iter(qn("w:br"))):
            if br.get(qn("w:type")) == "page":
                parent = br.getparent()
                if parent is not None:
                    parent.remove(br)
        title, page_num = _parse_toc_line(doc.paragraphs[i].text or "")
        if title:
            _set_toc_line(
                p_element,
                title,
                page_num,
                first_line=(i == contents_idx + 1),
            )


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
    _normalize_contents_section(doc)
    _ensure_page_break_before_introduction(doc)
    _remove_stacked_page_breaks(doc)

    pages = _compute_layout_pages(doc)
    _refresh_all_toc_pages(doc, ordered_keys, pages, catalog=request_catalog)
    _normalize_contents_section(doc)

    _strip_all_yellow_highlights(doc)
    doc.save(str(docx_path))
    _set_update_fields_on_open(docx_path)


def _clone_toc_paragraph(doc: Document) -> OxmlElement:
    """Create a new compact Contents line paragraph."""
    new_p = OxmlElement("w:p")
    _set_toc_line(new_p, "", 1)
    return new_p


def _insert_paragraph_before(doc: Document, index: int, p_element) -> None:
    anchor = doc.paragraphs[index]._element
    anchor.addprevious(p_element)


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


def _compute_paragraph_pages(doc: Document, paras_per_page: float = PARAS_PER_PAGE) -> dict[int, int]:
    """Estimate page number for each paragraph index."""
    del paras_per_page  # layout model supersedes the old paragraph counter
    return _compute_layout_pages(doc)


def _find_body_heading_idx(
    doc: Document,
    label_prefix: str,
    *,
    min_index: int = 0,
    exact: bool = False,
) -> int | None:
    prefix = label_prefix.strip().lower()
    for i, para in enumerate(doc.paragraphs):
        if i < min_index:
            continue
        text = (para.text or "").strip()
        if not text or "\t" in text:
            continue
        lower = text.lower()
        if exact:
            if lower == prefix:
                return i
        elif lower == prefix or lower.startswith(prefix):
            return i
    return None


def _find_body_anchor_page(
    doc: Document,
    label_prefix: str,
    pages: dict[int, int],
    min_index: int = 30,
    *,
    exact: bool = False,
) -> int:
    idx = _find_body_heading_idx(doc, label_prefix, min_index=min_index, exact=exact)
    if idx is not None:
        return pages.get(idx, 1)
    return 1


def _find_terms_body_page(doc: Document, pages: dict[int, int]) -> int:
    boundaries = _find_annexure_boundaries(doc)
    search_from = boundaries[-1][0] if boundaries else 0
    for i, para in enumerate(doc.paragraphs):
        if i < search_from:
            continue
        text = (para.text or "").strip()
        if text == TERMS_MARKER:
            return pages.get(i, 1)
        if text == TRAILER_MARKER:
            break
    return _find_body_anchor_page(doc, TERMS_TOC_LABEL, pages)


def _find_trailer_page(doc: Document, pages: dict[int, int]) -> int:
    for i, para in enumerate(doc.paragraphs):
        if (para.text or "").strip() == TRAILER_MARKER:
            return pages.get(i, 1)
    return _find_body_anchor_page(doc, CONCLUSION_TOC_PREFIX, pages)


def _find_annexure_heading_pages(doc: Document, pages: dict[int, int]) -> dict[int, int]:
    """Map annexure slot number -> page (Annexure 1, Annexure 2, ...)."""
    slot_pages: dict[int, int] = {}
    for i, para in enumerate(doc.paragraphs):
        text = (para.text or "").strip()
        match = re.match(r"^Annexure\s+(\d+)\s*:", text, re.IGNORECASE)
        if match:
            slot_pages[int(match.group(1))] = pages.get(i, 1)
    return slot_pages


def _update_contents_toc_structure(
    doc: Document,
    ordered_keys: list[str],
    catalog: dict[str, dict[str, Any]] | None = None,
) -> None:
    """Rebuild assessment lines in Contents (titles only; pages filled later).

    *catalog* may include AI-generated custom entries; falls back to CATALOG_BY_KEY.
    """
    contents_idx, first_assessment_idx, _, conclusion_idx = _find_contents_region(doc)
    if contents_idx is None or conclusion_idx is None or first_assessment_idx is None:
        return

    for i in range(conclusion_idx - 1, first_assessment_idx - 1, -1):
        _remove_paragraph_at(doc, i)

    _, first_assessment_idx, _, conclusion_idx = _find_contents_region(doc)
    if conclusion_idx is None or first_assessment_idx is None:
        return

    insert_at = first_assessment_idx
    for key in ordered_keys:
        toc_title = _toc_title_for_key(key, catalog)
        toc_line = _clone_toc_paragraph(doc)
        _set_toc_line(toc_line, toc_title, 1)
        _insert_paragraph_before(doc, insert_at, toc_line)
        insert_at += 1

    terms_line = _clone_toc_paragraph(doc)
    _set_toc_line(terms_line, TERMS_TOC_LABEL, 1)
    _insert_paragraph_before(doc, insert_at, terms_line)


def _refresh_all_toc_pages(
    doc: Document,
    ordered_keys: list[str],
    pages: dict[int, int],
    catalog: dict[str, dict[str, Any]] | None = None,
) -> None:
    """Update every Contents line page number from the final document layout.

    *catalog* must include per-request AI custom entries so TOC labels stay
    human-readable instead of falling back to ``__other__...`` internal keys.
    """
    contents_idx, first_assessment_idx, terms_idx, conclusion_idx = _find_contents_region(doc)
    if contents_idx is None:
        return

    body_min = (conclusion_idx if conclusion_idx is not None else contents_idx or 0) + 1
    intro_idx = _find_introduction_body_idx(doc)
    intro_page = int(pages.get(intro_idx, 1)) if intro_idx is not None else 1

    for offset, (label, _anchor) in enumerate(STATIC_TOC_ANCHORS):
        toc_idx = contents_idx + 1 + offset
        if toc_idx >= len(doc.paragraphs):
            break
        page = int(_find_body_anchor_page(doc, label, pages, min_index=body_min, exact=True))
        if page <= 1 and intro_page > 1:
            page = intro_page
        _set_toc_line(
            doc.paragraphs[toc_idx]._element,
            label,
            page,
            first_line=(offset == 0),
        )

    if first_assessment_idx is None or conclusion_idx is None:
        return

    annexure_pages = _find_annexure_heading_pages(doc, pages)
    for slot, key in enumerate(ordered_keys, start=1):
        toc_idx = first_assessment_idx + slot - 1
        if terms_idx is not None and toc_idx >= terms_idx:
            break
        if toc_idx >= conclusion_idx:
            break
        toc_title = _toc_title_for_key(key, catalog)
        page = int(annexure_pages.get(slot, DEFAULT_FIRST_ANNEXURE_PAGE))
        _set_toc_line(doc.paragraphs[toc_idx]._element, toc_title, page)

    _, _, terms_idx, conclusion_idx = _find_contents_region(doc)
    if terms_idx is not None:
        terms_page = int(_find_terms_body_page(doc, pages))
        _set_toc_line(doc.paragraphs[terms_idx]._element, TERMS_TOC_LABEL, terms_page)
    if conclusion_idx is not None:
        conclusion_page = int(_find_trailer_page(doc, pages))
        _set_toc_line(
            doc.paragraphs[conclusion_idx]._element,
            CONCLUSION_TOC_PREFIX,
            conclusion_page,
        )


def _update_contents_toc(
    doc: Document,
    ordered_keys: list[str],
    pages: dict[int, int] | None = None,
    catalog: dict[str, dict[str, Any]] | None = None,
) -> None:
    """Rebuild Contents assessment entries (page numbers refreshed in finalize)."""
    del pages
    _update_contents_toc_structure(doc, ordered_keys, catalog=catalog)


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

    _update_contents_toc(doc, ordered_keys, {}, catalog=catalog)
    doc.save(str(docx_path))


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
    """Return the cover paragraph that holds the client logo slot."""
    seen_heading = False
    for para in doc.paragraphs:
        txt = (para.text or "").strip()
        if "Statement of Work" in txt:
            seen_heading = True
            continue
        if not seen_heading:
            continue
        if "{{ client_logo }}" in (para.text or ""):
            return para
        if para._element.findall(f".//{PIC_TAG}") and not txt:
            return para
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
    return dt.strftime("%d/%m/%Y")


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
        if DEFAULT_LOGO.is_file():
            return InlineImage(tpl, str(DEFAULT_LOGO), width=Inches(calculate_optimal_logo_size(DEFAULT_LOGO)))
    except Exception:
        traceback.print_exc()
    return None


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
) -> Path:
    company_name = (company_name or "").strip() or "Client"
    proposal_ref = (proposal_ref or "").strip() or default_proposal_ref()
    template = Path(template_path) if template_path else ensure_base_template()
    output = Path(output_path)

    tpl = DocxTemplate(str(template))
    context: dict[str, Any] = {
        "client_name": company_name,
        "proposal_ref": proposal_ref,
        "client_logo": _resolve_logo_inline(tpl, logo_path),
        "report_date": datetime.now().strftime("%d %B %Y"),
    }
    try:
        tpl.render(context)
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
    return output


def _safe_filename(name: str, ext: str = ".docx") -> str:
    stem = re.sub(r"[^\w\-]+", "_", name.strip())[:80] or "Proposal"
    return f"{stem}{ext}"


# ---------------------------------------------------------------------------
# FastAPI app (self-hosted — not mounted in app.py)
# ---------------------------------------------------------------------------
app = FastAPI(title="CSS Proposal Generator", version=VERSION)
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
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
    <form id="proposalForm">
      <!-- Basic fields -->
      <div class="form-group">
        <label for="company_name">Company name *</label>
        <input type="text" id="company_name" name="company_name" required placeholder="Acme Corp" />
      </div>
      <div class="form-group">
        <label for="proposal_ref">Proposal reference (date: DD/MM/YYYY)</label>
        <input type="text" id="proposal_ref" name="proposal_ref" value="{today_ref}" placeholder="23/05/2026" />
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


@app.post("/generate/")
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
            logo_path=str(logo_path) if logo_path else None,
            output_path=output_path,
            annexure_order=annexure_order,
            assessments=assessments,
            request_catalog=request_catalog,
        )
        final_path = UPLOAD_DIR / out_name
        shutil.copy2(output_path, final_path)
        return FileResponse(
            path=str(final_path),
            filename=out_name,
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
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

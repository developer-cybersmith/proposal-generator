#!/usr/bin/env python3
"""
generate.py  --  CSS Pre-Engagement Scoping v2 parser test-set generator.

Produces, deterministically (seed 12345):
  samples/            filled dummy .xlsx workbooks (32 named scenarios)
  samples/load/       100 randomised + 5 stress workbooks
  expected/           one <file>.json per workbook (reference-parser output)
  manifest.csv        index: file, scenario, included count, warning count, error
  README.md           is written separately (not by this script)

Rules:
  - Always starts from a FRESH copy of template.xlsx. Never edits sheet names,
    headers, row order, merged ranges, widths or formatting; only writes VALUES
    into input cells, into the TOP-LEFT cell of any merged range.
  - The embedded reference_parser() is the single source of truth for expected
    results: every expected/<file>.json is produced by running it on the saved
    workbook, so the JSONs can never drift from the files.

Re-runnable:  python3 generate.py   (clears and regenerates samples/ + expected/)
Template is resolved next to this script unless --template is given.
"""

import os
import sys

# Freeze string hashing so any incidental set/dict-of-strings iteration order
# is identical across processes -> fully reproducible output. Must happen
# before most imports/work, so we re-exec once with the seed pinned.
if os.environ.get("PYTHONHASHSEED") != "0":
    os.environ["PYTHONHASHSEED"] = "0"
    os.execv(sys.executable, [sys.executable] + sys.argv)

import csv
import json
import shutil
import random
import zipfile
import argparse
import datetime as dt
from copy import copy

import openpyxl
from openpyxl.utils import get_column_letter

SEED = 12345
HERE = os.path.dirname(os.path.abspath(__file__))
MAX_SIZE_BYTES = 5 * 1024 * 1024

# --------------------------------------------------------------------------
# Static description of the template (verified against the blank v2.0 file)
# --------------------------------------------------------------------------
SHEET_COVER = "Cover Page"
SHEET_MATRIX = "Overall Scoping Matrix"
SHEET_GENERAL = "General Scoping Questions"
SHEET_WEB = "Web Application Assessment"
SHEET_MOBILE = "Mobile Application Assessment"
SHEET_CLOUD = "Cloud Security Assessment"
SHEET_AD = "AD Security Assessment"
SHEET_NETSEG = "Network Segmentation Testing"
SHEET_ROE = "Rules of Engagement"

REQUIRED_SHEETS = [SHEET_COVER, SHEET_MATRIX]

# Cover Page client block (label in B/E, response written to the listed cell;
# all response cells are top-left of merged ranges).
COVER_LEFT = {  # column C, rows 16-22
    "client_organization_name": "C16",
    "primary_contact_name": "C17",
    "designation": "C18",
    "email": "C19",
    "phone": "C20",
    "alternate_contact_name": "C21",
    "alternate_contact_email": "C22",
}
COVER_RIGHT = {  # column F, rows 16-22
    "engagement_reference_no": "F16",
    "date_of_submission": "F17",
    "preferred_start_date": "F18",
    "preferred_end_date": "F19",
    "project_classification": "F20",
    "regulatory_driver": "F21",
    "nda_signed": "F22",
}
DATE_FIELDS = {"date_of_submission", "preferred_start_date", "preferred_end_date"}

# Grey placeholder strings baked into the blank template. Any of these present
# in a column-F / known cell counts as EMPTY (not a client value).
PLACEHOLDERS = {
    "CSS-ENG-YYYY-XXX",
    "DD-MMM-YYYY",
    "Confidential / Restricted / Internal",
    "ISO 27001 / DPDPA / SEBI / RBI / CERT-In / Other",
    "Yes / No / In Progress",
    "To be provided on RoE document",
}
REF_PLACEHOLDER = "CSS-ENG-YYYY-XXX"

# In-scope checkbox cells on the Cover Page -> internal type key.
INSCOPE_CELLS = {
    # column C
    "external_network_pt": "C26",
    "internal_network_pt": "C27",
    "cloud_security": "C28",
    "ad_security": "C29",
    "network_segmentation": "C30",
    "saas_security": "C31",
    "config_hardening": "C32",
    # column F
    "web_app_vapt": "F26",
    "mobile_vapt": "F27",
    "source_code_review": "F28",
    "api_security": "F29",
    "social_engineering": "F30",
    "red_team": "F31",
    "compliance_gap": "F32",
}

# Human-readable type labels (for JSON "type").
TYPE_LABELS = {
    "external_network_pt": "External Network Penetration Testing",
    "internal_network_pt": "Internal Network Penetration Testing",
    "cloud_security": "Cloud Security Assessment",
    "ad_security": "Active Directory (AD) Security Assessment",
    "network_segmentation": "Network Segmentation / VLAN Testing",
    "saas_security": "SaaS Security Assessment",
    "config_hardening": "Configuration / Hardening Review",
    "web_app_vapt": "Web Application VAPT",
    "mobile_vapt": "Mobile Application VAPT (Android/iOS)",
    "source_code_review": "Source Code Review / SAST",
    "api_security": "API Security Testing",
    "social_engineering": "Social Engineering / Phishing Simulation",
    "red_team": "Red Team Engagement",
    "compliance_gap": "Compliance Gap Assessment (ISO/DPDPA/SEBI)",
}
ALL_TYPES = list(INSCOPE_CELLS.keys())

# Matrix Quantity column = D. Band-header rows that must never be written.
MATRIX_BAND_ROWS = {3, 16, 23, 34, 42, 48, 54, 61, 66}
MATRIX_DATA_ROWS = (
    list(range(4, 16)) + list(range(17, 23)) + list(range(24, 34)) +
    list(range(35, 42)) + list(range(43, 48)) + list(range(49, 54)) +
    list(range(55, 61)) + list(range(62, 66)) + list(range(67, 71))
)
MATRIX_OPT_COLS = ["E", "F", "G", "H", "I"]  # methodology / service / box / delivery / deadline

# Accepted affirmative tokens (case-insensitive, trimmed).
YES_TOKENS = {"yes", "y", "✓"}


def is_yes(value):
    if value is None:
        return False
    if value is True:
        return False  # boolean TRUE is not an accepted token
    s = str(value).strip().lower()
    return s in YES_TOKENS


# Ordered type rules (D30). Order = order of included_annexures/warnings in the JSON.
# (type key, template annexure or "custom", matrix rows used, Count label)
DEFAULT_FREQUENCY = "One Time"
TYPE_RULES = [
    ("external_network_pt", 1, list(range(17, 23)), "External Infrastructure categories"),
    ("internal_network_pt", 1, list(range(4, 16)), "Internal Infrastructure categories"),
    ("web_app_vapt", 2, list(range(55, 61)), "Web Application categories"),
    ("config_hardening", 3, [4, 5, 6, 9, 10, 11], "Device & Server categories"),
    ("mobile_vapt", 5, list(range(62, 66)), "Mobile Application categories"),
    ("ad_security", 6, list(range(35, 42)), "Active Directory categories"),
    ("cloud_security", 8, list(range(24, 34)), "Cloud Infrastructure categories"),
    ("api_security", 9, [19, 32], "API categories"),
    ("source_code_review", 11, list(range(67, 71)), "Source Code categories"),
    ("network_segmentation", "custom", list(range(43, 48)), "Network Segmentation categories"),
    ("saas_security", "custom", list(range(49, 54)), "SaaS Application categories"),
]
DEFAULT_ONE_TYPES = ("social_engineering", "red_team", "compliance_gap")


# --------------------------------------------------------------------------
# Reference parser  ==  the agreed expected-result rules.
# --------------------------------------------------------------------------
def _num_or_none(value):
    """Return a float/int if the cell holds a real number, else None.

    Real Excel numbers (int/float) count. Booleans and strings do NOT, even
    numeric-looking strings like '10' or '7.0' (spec: text is skipped).
    """
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return value
    return None


def _clean_num(n):
    """Present 7.0 as 7 but keep 2.5 as 2.5."""
    if isinstance(n, float) and n.is_integer():
        return int(n)
    return n


def _block_stats(ws_val, ws_fmla, rows, warnings):
    """Count and sum the Quantity cells (column D) of the given matrix rows (D30-D33).

    Count    = rows holding a number above 0.
    Quantity = sum of those numbers.
    Blank and 0 are skipped silently. Text, negatives and formulas without a cached
    value are skipped with a warning.
    """
    count = 0
    total = 0
    for r in rows:
        coord = f"D{r}"
        v = ws_val[coord].value
        raw = ws_fmla[coord].value
        category = str(ws_val[f"C{r}"].value or "").strip()
        if v is None:
            if isinstance(raw, str) and raw.startswith("="):
                warnings.append(f"formula_no_cached_value: {ws_val.title}!{coord}")
            continue  # blank -> skipped, no warning
        num = _num_or_none(v)
        if num is None:
            warnings.append(
                f"non_numeric_quantity_skipped: {category} ({ws_val.title}!{coord})={v!r}"
            )
            continue
        if num < 0:
            warnings.append(
                f"negative_quantity_skipped: {category} ({ws_val.title}!{coord})={v!r}"
            )
            continue
        if num == 0:
            continue
        count += 1
        total += num
    return count, _clean_num(total)


def _read_client_field(ws, coord, field, warnings):
    v = ws[coord].value
    if v is None:
        return None
    if isinstance(v, dt.datetime):
        return v.date().isoformat()
    if isinstance(v, dt.date):
        return v.isoformat()
    s = str(v).strip()
    if s == "" or s in PLACEHOLDERS:
        return None
    if field in DATE_FIELDS:
        # Real dates are returned above; a surviving string is a text date.
        warnings.append(f"date_not_real_date: {ws.title}!{coord}={s!r}")
    return s


def reference_parser(path):
    """Produce the expected-result dict for one workbook."""
    out = {
        "included_annexures": [],
        "pricing_lines": [],
        "distinct_annexures": [],
        "warnings": [],
        "client": {},
        "reference_no": None,
        "errors": [],
    }
    warnings = out["warnings"]
    errors = out["errors"]

    # --- open / zip-level validation -------------------------------------
    if not os.path.exists(path):
        errors.append("file_not_found")
        return out
    if not zipfile.is_zipfile(path):
        errors.append("invalid_xlsx: not a valid zip/OOXML container")
        return out
    try:
        wb_val = openpyxl.load_workbook(path, data_only=True)
        wb_fmla = openpyxl.load_workbook(path, data_only=False)
    except Exception as exc:  # corrupt / truncated
        errors.append(f"corrupt_file: {type(exc).__name__}")
        return out

    # --- oversize --------------------------------------------------------
    size = os.path.getsize(path)
    if size > MAX_SIZE_BYTES:
        errors.append(f"file_too_large: {size} bytes > {MAX_SIZE_BYTES}")

    # --- required sheets -------------------------------------------------
    names = wb_val.sheetnames
    for req in REQUIRED_SHEETS:
        if req not in names:
            errors.append(f"missing_sheet: {req}")
    if SHEET_COVER not in names or SHEET_MATRIX not in names:
        return out  # cannot proceed without both

    mv = wb_val[SHEET_MATRIX]
    mf = wb_fmla[SHEET_MATRIX]
    cover = wb_val[SHEET_COVER]

    # --- matrix header integrity ----------------------------------------
    if str(mv["D2"].value).strip() != "Quantity":
        errors.append(
            f"quantity_header_changed: expected 'Quantity' at {SHEET_MATRIX}!D2, "
            f"found {mv['D2'].value!r}"
        )
        return out  # column mapping is no longer trustworthy

    # --- client block ----------------------------------------------------
    client = {}
    for field, coord in COVER_LEFT.items():
        client[field] = _read_client_field(cover, coord, field, warnings)
    for field, coord in COVER_RIGHT.items():
        client[field] = _read_client_field(cover, coord, field, warnings)
    out["client"] = client

    # --- reference number ------------------------------------------------
    ref_raw = cover["F16"].value
    ref = None
    if ref_raw is not None:
        s = str(ref_raw).strip()
        if s and s not in PLACEHOLDERS and s != REF_PLACEHOLDER:
            ref = s
    if ref is None:
        warnings.append("reference_required")
    out["reference_no"] = ref

    # --- which types are in scope ---------------------------------------
    in_scope = {t: is_yes(cover[c].value) for t, c in INSCOPE_CELLS.items()}
    if not any(in_scope.values()):
        errors.append("no_assessment_types_in_scope")

    included = []       # per-type annexure objects
    pricing = []        # per billable line
    distinct = []       # unique annexure identifiers (document level)

    def add(type_key, annx, count, quantity, label):
        included.append({
            "type": TYPE_LABELS[type_key],
            "annexure_no_or_custom": annx,
            "count": count,
            "quantity": quantity,
            "count_label": label,
            "frequency": DEFAULT_FREQUENCY,
        })
        # External + Internal share Annexure 3 but keep separate pricing lines.
        pricing.append({
            "label": TYPE_LABELS[type_key],
            "annexure_no_or_custom": annx,
            "quantity": quantity,
            "frequency": DEFAULT_FREQUENCY,
        })
        if annx not in distinct:
            distinct.append(annx)

    for type_key, annx, rows, label in TYPE_RULES:
        if not in_scope[type_key]:
            continue
        count, quantity = _block_stats(mv, mf, rows, warnings)
        add(type_key, annx, count, quantity, label)
        if count == 0:
            warnings.append(f"quantity_zero_confirm: {TYPE_LABELS[type_key]}")
    # Types with no matrix rows: Count 1, Quantity 1 (user edits in the preview).
    for key in DEFAULT_ONE_TYPES:
        if in_scope[key]:
            add(key, "custom", 1, 1, "Engagement categories")

    out["included_annexures"] = included
    out["pricing_lines"] = pricing
    out["distinct_annexures"] = distinct
    return out


# --------------------------------------------------------------------------
# Fictional data pools (all non-real; emails @example.com, phones +91 98…0)
# --------------------------------------------------------------------------
COMPANIES = [
    "Meridian Textiles Pvt. Ltd.", "Blue Harbour Logistics", "Vantage Retail Group",
    "Northwind Pharma Labs", "Crestline Financial Services", "Orbital Tech Solutions",
    "Sunder Agro Industries", "Kaveri Infrastructure Ltd.", "Zephyr Media House",
    "Ironclad Manufacturing Co.", "Pinnacle Health Systems", "Greystone Capital",
    "Alpine Dairy Cooperative", "Lotus Hospitality Group", "Quantum Mining Corp.",
]
FIRST = ["Aarav", "Diya", "Kabir", "Isha", "Rohan", "Meera", "Vikram", "Ananya",
         "Arjun", "Priya", "Sameer", "Neha", "Rahul", "Kavya", "Tarun"]
LAST = ["Nair", "Deshpande", "Kulkarni", "Reddy", "Banerjee", "Chauhan", "Iyer",
        "Malhotra", "Pillai", "Sethi", "Ghosh", "Rao", "Menon", "Bhat", "Joshi"]
DESIGS = ["CISO", "IT Head", "Security Analyst", "VP Engineering", "IT Manager",
          "Head of Infrastructure", "Compliance Officer", "DevOps Lead"]
CLASSIFICATIONS = ["Confidential", "Restricted", "Internal"]
REGULATORS = ["ISO 27001", "DPDPA", "SEBI", "RBI", "CERT-In", "Other"]


def fake_person(rnd):
    return f"{rnd.choice(FIRST)} {rnd.choice(LAST)}"


def fake_email(name, rnd):
    user = name.lower().replace(" ", ".")
    return f"{user}@example.com"


def fake_phone(rnd):
    mid = "".join(rnd.choice("0123456789") for _ in range(7))
    return f"+91 98{mid}0"


def fake_ref(rnd):
    return f"CSS-ENG-2026-{rnd.randint(100, 999)}"


# --------------------------------------------------------------------------
# Workbook writing helpers
# --------------------------------------------------------------------------
def fresh_wb(template_path):
    return openpyxl.load_workbook(template_path, data_only=False)


def set_date(ws, coord, date_obj):
    ws[coord] = date_obj
    ws[coord].number_format = "DD-MMM-YYYY"


def write_client(ws, rnd, company=None, ref=True, dates="real",
                 classification=True, regulator=True, nda=True):
    company = company or rnd.choice(COMPANIES)
    pc = fake_person(rnd)
    ac = fake_person(rnd)
    ws["C16"] = company
    ws["C17"] = pc
    ws["C18"] = rnd.choice(DESIGS)
    ws["C19"] = fake_email(pc, rnd)
    ws["C20"] = fake_phone(rnd)
    ws["C21"] = ac
    ws["C22"] = fake_email(ac, rnd)
    if ref:
        ws["F16"] = fake_ref(rnd)
    if dates == "real":
        start = dt.date(2026, rnd.randint(1, 6), rnd.randint(1, 28))
        set_date(ws, "F17", dt.date(2026, 1, rnd.randint(1, 28)))
        set_date(ws, "F18", start)
        set_date(ws, "F19", start + dt.timedelta(days=rnd.randint(10, 40)))
    elif dates == "text":
        ws["F17"] = "05-Jan-2026"
        ws["F18"] = "2026/01/05"
        ws["F19"] = "Jan 20 2026"
    if classification:
        ws["F20"] = rnd.choice(CLASSIFICATIONS)
    if regulator:
        ws["F21"] = rnd.choice(REGULATORS)
    if nda:
        ws["F22"] = rnd.choice(["Yes", "No", "In Progress"])


def set_scope(ws, yes_keys, token="Yes", no_token="No"):
    """Set In-Scope cells. yes_keys = iterable of type keys to mark affirmative."""
    for key, coord in INSCOPE_CELLS.items():
        ws[coord] = token if key in yes_keys else no_token


def fill_matrix_quantities(ws, values, opt_text=False, rnd=None):
    """values = {row:int-or-value}. Writes to column D. Optionally fills E-I."""
    for row, val in values.items():
        if row in MATRIX_BAND_ROWS:
            raise ValueError(f"refusing to write band-header row {row}")
        ws[f"D{row}"] = val
    if opt_text and rnd is not None:
        methods = ["Black-box", "Grey-box", "White-box"]
        services = ["VAPT", "Config Review", "Threat Emulation", "SAST"]
        boxes = ["Blackbox", "Greybox", "Whitebox"]
        delivery = ["Remote", "On-site", "Hybrid"]
        for row in values:
            if rnd.random() < 0.5:
                ws[f"E{row}"] = rnd.choice(methods)
                ws[f"F{row}"] = rnd.choice(services)
                ws[f"G{row}"] = rnd.choice(boxes)
                ws[f"H{row}"] = rnd.choice(delivery)
                ws[f"I{row}"] = f"2026-{rnd.randint(2,6):02d}-{rnd.randint(1,28):02d}"


# Question-row maps for the AI-context sheets -------------------------------
GENERAL_Q_ROWS = (list(range(4, 12)) + list(range(13, 23)) + list(range(24, 36)) +
                  list(range(37, 47)) + list(range(48, 58)) + list(range(59, 69)) +
                  list(range(70, 78)) + list(range(79, 87)) + list(range(88, 95)) +
                  list(range(96, 104)))
WEB_Q_ROWS = (list(range(4, 9)) + list(range(10, 17)) + list(range(18, 25)) +
              list(range(26, 37)) + list(range(38, 46)) + list(range(47, 57)))
MOBILE_Q_ROWS = (list(range(4, 10)) + list(range(11, 17)) + list(range(18, 27)) +
                 list(range(28, 34)) + list(range(35, 44)) + list(range(45, 54)))
CLOUD_Q_ROWS = (list(range(4, 11)) + list(range(12, 22)) + list(range(23, 29)) +
                list(range(30, 36)) + list(range(37, 43)) + list(range(44, 49)) +
                list(range(50, 56)))
AD_Q_ROWS = (list(range(4, 12)) + list(range(13, 21)) + list(range(22, 32)) +
             list(range(33, 42)) + list(range(43, 47)))
NETSEG_Q_ROWS = (list(range(4, 11)) + list(range(12, 19)) + list(range(20, 27)) +
                 list(range(28, 33)) + list(range(34, 42)))

ANS_SNIPPETS = [
    "Yes", "No", "Not applicable", "To be confirmed", "AWS ap-south-1",
    "Approx. 25", "REST + GraphQL", "Grey-box with standard domain user",
    "Cloudflare WAF in blocking mode", "Production environment only",
    "SSO via Okta (SAML 2.0)", "Available on request", "Internal / VPN only",
]


def fill_context_sheet(ws, q_rows, cols, rnd, density=0.6, text_len=None,
                       all_cols=False):
    """Fill answer cols next to each question row.

    cols: list of answer columns. density: fraction of rows filled.
    text_len: if set, each answer is padded to ~text_len chars.
    all_cols: force every listed column filled (else only first + maybe others).
    """
    for r in q_rows:
        if not all_cols and rnd.random() > density:
            continue
        use_cols = cols if all_cols else ([cols[0]] + [c for c in cols[1:]
                                          if rnd.random() < 0.4])
        for col in use_cols:
            base = rnd.choice(ANS_SNIPPETS)
            if text_len:
                filler = (" Lorem ipsum dolor sit amet consectetur adipiscing "
                          "elit sed do eiusmod tempor.")
                s = base
                while len(s) < text_len:
                    s += filler
                base = s[:text_len]
            ws[f"{col}{r}"] = base


def fill_roe(ws, rnd, text_len=None):
    roles_rows = [5, 6, 7, 8, 9]
    for r in roles_rows:
        nm = fake_person(rnd)
        ws[f"C{r}"] = nm
        ws[f"D{r}"] = fake_phone(rnd)
        ws[f"E{r}"] = fake_phone(rnd)
        ws[f"F{r}"] = fake_email(nm, rnd)
        ws[f"G{r}"] = rnd.choice(["24x7", "Business hours", "On-call"])
    for r in range(13, 26):  # parameters C13:C25 (merged C:G -> write C)
        val = rnd.choice(ANS_SNIPPETS)
        if text_len:
            val = (val + " ") * (text_len // 10)
            val = val[:text_len]
        ws[f"C{r}"] = val
    for r in range(29, 36):  # compliance C29:C35
        ws[f"C{r}"] = rnd.choice(["Yes", "No", "Required", "Not required"])


def fill_all_context(wb, rnd, density=0.6, text_len=None, all_cols=False):
    fill_context_sheet(wb[SHEET_GENERAL], GENERAL_Q_ROWS,
                       ["E", "F", "G", "H", "I"], rnd, density, text_len, all_cols)
    fill_context_sheet(wb[SHEET_WEB], WEB_Q_ROWS, ["D", "E"], rnd, density,
                       text_len, all_cols)
    fill_context_sheet(wb[SHEET_MOBILE], MOBILE_Q_ROWS, ["D", "E"], rnd, density,
                       text_len, all_cols)
    fill_context_sheet(wb[SHEET_CLOUD], CLOUD_Q_ROWS, ["D"], rnd, density,
                       text_len, all_cols)
    fill_context_sheet(wb[SHEET_AD], AD_Q_ROWS, ["D"], rnd, density, text_len,
                       all_cols)
    fill_context_sheet(wb[SHEET_NETSEG], NETSEG_Q_ROWS, ["D"], rnd, density,
                       text_len, all_cols)
    fill_roe(wb[SHEET_ROE], rnd, text_len)


# --------------------------------------------------------------------------
# Convenience: realistic quantity presets per band
# --------------------------------------------------------------------------
def realistic_quantities(rnd, rows=None):
    """Return {row:int} with plausible counts for the requested data rows."""
    rows = rows or MATRIX_DATA_ROWS
    out = {}
    for r in rows:
        out[r] = rnd.randint(1, 40)
    return out


# --------------------------------------------------------------------------
# Scenario builders. Each returns (wb, post) where post is an optional callable
# taking the saved path (used for LibreOffice recalc / corruption / etc.), or
# a special dict for non-workbook outputs.
# --------------------------------------------------------------------------
def build_named_scenarios(template_path, rnd):
    """Yield (filename, scenario_label, builder) tuples."""
    scenarios = []

    def reg(fn):
        scenarios.append(fn)
        return fn

    # ---- 01 all_in_scope_full ----
    @reg
    def s01():
        wb = fresh_wb(template_path)
        cp = wb[SHEET_COVER]
        write_client(cp, rnd, company="Meridian Textiles Pvt. Ltd.")
        set_scope(cp, set(ALL_TYPES))
        q = realistic_quantities(rnd)
        fill_matrix_quantities(wb[SHEET_MATRIX], q, opt_text=True, rnd=rnd)
        fill_all_context(wb, rnd, density=0.9)
        return "01_all_in_scope_full.xlsx", "VALID: all 14 types in scope", wb, None

    # ---- 02 known_only ----
    @reg
    def s02():
        wb = fresh_wb(template_path)
        cp = wb[SHEET_COVER]
        write_client(cp, rnd)
        keys = {"web_app_vapt", "mobile_vapt", "external_network_pt",
                "cloud_security", "source_code_review"}
        set_scope(cp, keys)
        q = realistic_quantities(rnd)
        fill_matrix_quantities(wb[SHEET_MATRIX], q, opt_text=True, rnd=rnd)
        fill_all_context(wb, rnd, density=0.7)
        return "02_known_only.xlsx", "VALID: template-annexure types only", wb, None

    # ---- 03 custom_only ----
    @reg
    def s03():
        wb = fresh_wb(template_path)
        cp = wb[SHEET_COVER]
        write_client(cp, rnd)
        keys = {"network_segmentation", "saas_security", "social_engineering",
                "red_team", "compliance_gap"}
        set_scope(cp, keys)
        fill_matrix_quantities(wb[SHEET_MATRIX], {43: 12, 49: 3})
        fill_all_context(wb, rnd, density=0.5)
        return "03_custom_only.xlsx", "VALID: AI/custom annexure types only", wb, None

    # ---- 04 mixed_known_and_custom ----
    @reg
    def s04():
        wb = fresh_wb(template_path)
        cp = wb[SHEET_COVER]
        write_client(cp, rnd)
        keys = {"web_app_vapt", "ad_security", "api_security",
                "network_segmentation", "red_team"}
        set_scope(cp, keys)
        fill_matrix_quantities(wb[SHEET_MATRIX],
                               {55: 4, 35: 2, 19: 5, 32: 2, 43: 9})
        fill_all_context(wb, rnd, density=0.6)
        return "04_mixed_known_and_custom.xlsx", "VALID: mix of template + custom", wb, None

    # ---- 05 external_and_internal_both ----
    @reg
    def s05():
        wb = fresh_wb(template_path)
        cp = wb[SHEET_COVER]
        write_client(cp, rnd)
        set_scope(cp, {"external_network_pt", "internal_network_pt"})
        # Internal D4:D15, External D17:D22 with distinct quantities.
        q = {4: 2, 5: 6, 6: 3, 7: 4, 8: 1, 9: 10, 10: 8, 11: 5, 12: 2,
             13: 40, 14: 3, 15: 6,          # internal sum = 90
             17: 12, 18: 5, 19: 3, 20: 2, 21: 1, 22: 2}  # external sum = 25
        fill_matrix_quantities(wb[SHEET_MATRIX], q, opt_text=True, rnd=rnd)
        fill_all_context(wb, rnd, density=0.5)
        return ("05_external_and_internal_both.xlsx",
                "VALID: External+Internal -> 1 annexure, 2 pricing lines", wb, None)

    # ---- 06 config_and_api ----
    @reg
    def s06():
        wb = fresh_wb(template_path)
        cp = wb[SHEET_COVER]
        write_client(cp, rnd)
        set_scope(cp, {"config_hardening", "api_security"})
        # Config uses D4,5,6,9,10,11 ; API uses D19 + D32.
        q = {4: 2, 5: 5, 6: 3, 9: 10, 10: 8, 11: 4,   # config sum = 32
             19: 6, 32: 3}                             # api = 9
        fill_matrix_quantities(wb[SHEET_MATRIX], q)
        fill_all_context(wb, rnd, density=0.4)
        return ("06_config_and_api.xlsx",
                "VALID: Config(4-6,9-11)=32, API(19+32)=9", wb, None)

    # ---- 07 mobile_android_only ----
    @reg
    def s07():
        wb = fresh_wb(template_path)
        cp = wb[SHEET_COVER]
        write_client(cp, rnd)
        set_scope(cp, {"mobile_vapt"})
        fill_matrix_quantities(wb[SHEET_MATRIX], {62: 5})  # D63 blank
        fill_all_context(wb, rnd, density=0.4)
        return "07_mobile_android_only.xlsx", "VALID: Android=5, iOS blank", wb, None

    # ---- 08 mobile_both ----
    @reg
    def s08():
        wb = fresh_wb(template_path)
        cp = wb[SHEET_COVER]
        write_client(cp, rnd)
        set_scope(cp, {"mobile_vapt"})
        fill_matrix_quantities(wb[SHEET_MATRIX], {62: 6, 63: 3})
        fill_all_context(wb, rnd, density=0.4)
        return "08_mobile_both.xlsx", "VALID: Android=6, iOS=3", wb, None

    # ---- 09 minimal ----
    @reg
    def s09():
        wb = fresh_wb(template_path)
        cp = wb[SHEET_COVER]
        write_client(cp, rnd)
        set_scope(cp, {"web_app_vapt"})
        fill_matrix_quantities(wb[SHEET_MATRIX], {55: 1})
        # other sheets blank
        return "09_minimal.xlsx", "VALID: one type, one quantity, no context", wb, None

    # ---- 10 zero_quantity ----
    @reg
    def s10():
        wb = fresh_wb(template_path)
        cp = wb[SHEET_COVER]
        write_client(cp, rnd)
        set_scope(cp, {"web_app_vapt", "cloud_security", "ad_security"})
        fill_matrix_quantities(wb[SHEET_MATRIX], {55: 0})  # rest blank
        fill_all_context(wb, rnd, density=0.3)
        return ("10_zero_quantity.xlsx",
                "EDGE: in scope but quantities 0/blank -> zero_confirm", wb, None)

    # ---- 11 placeholders_untouched ----
    @reg
    def s11():
        wb = fresh_wb(template_path)
        cp = wb[SHEET_COVER]
        # client left column filled but column F placeholders untouched
        pc = fake_person(rnd)
        cp["C16"] = "Blue Harbour Logistics"
        cp["C17"] = pc
        cp["C18"] = rnd.choice(DESIGS)
        cp["C19"] = fake_email(pc, rnd)
        cp["C20"] = fake_phone(rnd)
        set_scope(cp, {"web_app_vapt"})
        fill_matrix_quantities(wb[SHEET_MATRIX], {55: 2})
        return ("11_placeholders_untouched.xlsx",
                "EDGE: column F left as grey placeholders -> reference_required", wb, None)

    # ---- 12 yes_no_variants ----
    @reg
    def s12():
        wb = fresh_wb(template_path)
        cp = wb[SHEET_COVER]
        write_client(cp, rnd)
        variants = ["Yes", "YES", "yes", "Y", "y", "✓", " Yes ",
                    "No", "N", "n", None, "maybe", "TRUE", 1]
        # map to the 14 cells in a fixed order
        order = ["external_network_pt", "internal_network_pt", "cloud_security",
                 "ad_security", "network_segmentation", "saas_security",
                 "config_hardening", "web_app_vapt", "mobile_vapt",
                 "source_code_review", "api_security", "social_engineering",
                 "red_team", "compliance_gap"]
        for key, val in zip(order, variants):
            if val is None:
                continue  # leave blank
            cp[INSCOPE_CELLS[key]] = val
        # give every type a quantity so accepted ones resolve
        fill_matrix_quantities(wb[SHEET_MATRIX],
                               {17: 5, 4: 5, 24: 5, 35: 2, 43: 4, 49: 2,
                                55: 3, 62: 2, 63: 1, 67: 1, 19: 2, 32: 1})
        return ("12_yes_no_variants.xlsx",
                "EDGE: accepted {yes,y,checkmark}; TRUE/1/maybe rejected", wb, None)

    # ---- 13 text_in_quantity ----
    @reg
    def s13():
        wb = fresh_wb(template_path)
        cp = wb[SHEET_COVER]
        write_client(cp, rnd)
        set_scope(cp, {"internal_network_pt", "external_network_pt",
                       "cloud_security"})
        mx = wb[SHEET_MATRIX]
        vals = {4: "ten", 5: 6, 6: "N/A", 9: 10, 10: "10.0.0.0/24", 11: 4,
                13: "Y", 14: "-", 15: 8,
                17: 5, 18: "twelve", 19: 3,
                24: 4, 25: "many", 26: 7}
        fill_matrix_quantities(mx, vals)
        return ("13_text_in_quantity.xlsx",
                "EDGE: text mixed with numbers -> skipped + warnings", wb, None)

    # ---- 14 big_numbers ----
    @reg
    def s14():
        wb = fresh_wb(template_path)
        cp = wb[SHEET_COVER]
        write_client(cp, rnd)
        set_scope(cp, {"internal_network_pt", "cloud_security", "web_app_vapt"})
        fill_matrix_quantities(wb[SHEET_MATRIX],
                               {4: 100000, 5: 2350, 6: 999999999,
                                24: 2.5, 55: 5})
        return "14_big_numbers.xlsx", "EDGE: very large ints + a decimal", wb, None

    # ---- 15 negative_and_float ----
    @reg
    def s15():
        wb = fresh_wb(template_path)
        cp = wb[SHEET_COVER]
        write_client(cp, rnd)
        set_scope(cp, {"internal_network_pt", "web_app_vapt", "ad_security"})
        fill_matrix_quantities(wb[SHEET_MATRIX],
                               {4: -3, 5: 0.0, 6: 7.0, 55: 7.0, 35: 7.0})
        return ("15_negative_and_float.xlsx",
                "EDGE: -3, 0.0, 7.0 (float reads as 7)", wb, None)

    # ---- 16 text_dates_vs_real_dates ----
    @reg
    def s16():
        wb = fresh_wb(template_path)
        cp = wb[SHEET_COVER]
        write_client(cp, rnd, dates="real")
        # override with a mix: F17 text, F18 real, F19 other text format
        cp["F17"] = "05-Jan-2026"
        set_date(cp, "F18", dt.date(2026, 1, 5))
        cp["F19"] = "2026/01/05"
        set_scope(cp, {"web_app_vapt"})
        fill_matrix_quantities(wb[SHEET_MATRIX], {55: 2})
        return ("16_text_dates_vs_real_dates.xlsx",
                "EDGE: text dates vs real Excel dates", wb, None)

    # ---- 17 unicode_and_long ----
    @reg
    def s17():
        wb = fresh_wb(template_path)
        cp = wb[SHEET_COVER]
        long_name = ("Société आईए & <Global> \"Holdings\" "
                     "'Ltd' " + "X" * 150)[:160]
        cp["C16"] = long_name
        pc = "Zoë अर्जुन"
        cp["C17"] = pc
        cp["C18"] = "Dir. of <Security> & \"Risk\""
        cp["C19"] = ("verylongemailaddress.department.security.operations."
                     "team." + "x" * 60 + "@example.com")
        cp["C20"] = fake_phone(rnd)
        set_scope(cp, {"web_app_vapt"})
        fill_matrix_quantities(wb[SHEET_MATRIX], {55: 2})
        return ("17_unicode_and_long.xlsx",
                "EDGE: unicode, XML-special chars, 150+ char fields", wb, None)

    # ---- 18 huge_context ----
    @reg
    def s18():
        wb = fresh_wb(template_path)
        cp = wb[SHEET_COVER]
        write_client(cp, rnd)
        set_scope(cp, {"web_app_vapt", "mobile_vapt", "cloud_security"})
        fill_matrix_quantities(wb[SHEET_MATRIX], {55: 3, 62: 2, 63: 2, 24: 5})
        fill_all_context(wb, rnd, density=1.0, text_len=1500, all_cols=True)
        return ("18_huge_context.xlsx",
                "EDGE: ~1500-char answers, all 5 app columns filled", wb, None)

    # ---- 19 formulas_cached ----
    @reg
    def s19():
        wb = fresh_wb(template_path)
        cp = wb[SHEET_COVER]
        write_client(cp, rnd)
        set_scope(cp, {"internal_network_pt", "web_app_vapt"})
        mx = wb[SHEET_MATRIX]
        mx["D4"] = "=5+3"
        mx["D5"] = "=SUM(2,4,6)"
        mx["D55"] = "=10-2"
        return ("19_formulas_cached.xlsx",
                "EDGE: formula quantities, recalced by LibreOffice (cached)", wb,
                "recalc")

    # ---- 20 formulas_uncached ----
    @reg
    def s20():
        wb = fresh_wb(template_path)
        cp = wb[SHEET_COVER]
        write_client(cp, rnd)
        set_scope(cp, {"internal_network_pt", "web_app_vapt"})
        mx = wb[SHEET_MATRIX]
        mx["D4"] = "=5+3"
        mx["D5"] = "=SUM(2,4,6)"
        mx["D55"] = "=10-2"
        return ("20_formulas_uncached.xlsx",
                "EDGE: openpyxl formulas, no cached value -> warnings", wb, None)

    # ---- 21 extra_hidden_sheet ----
    @reg
    def s21():
        wb = fresh_wb(template_path)
        cp = wb[SHEET_COVER]
        write_client(cp, rnd)
        set_scope(cp, {"web_app_vapt", "cloud_security"})
        fill_matrix_quantities(wb[SHEET_MATRIX], {55: 3, 24: 4})
        extra = wb.create_sheet("INTERNAL_NOTES")
        extra["A1"] = "internal pricing notes - ignore"
        extra["A2"] = 42
        extra.sheet_state = "hidden"
        return ("21_extra_hidden_sheet.xlsx",
                "ROBUST: extra hidden sheet must be ignored", wb, None)

    # ---- 22 extra_data_outside_known_cells ----
    @reg
    def s22():
        wb = fresh_wb(template_path)
        cp = wb[SHEET_COVER]
        write_client(cp, rnd)
        set_scope(cp, {"web_app_vapt", "cloud_security"})
        mx = wb[SHEET_MATRIX]
        fill_matrix_quantities(mx, {55: 3, 24: 4})
        cp["K23"] = cp["K23"].value  # leave the stray 's' as-is
        cp["Z99"] = "junk out of range"
        mx["K23"] = "noise"
        mx["Z99"] = 123456
        for r in range(73, 90):
            mx[f"B{r}"] = f"junk row {r}"
            mx[f"D{r}"] = r
        return ("22_extra_data_outside_known_cells.xlsx",
                "ROBUST: junk in K23/Z99/rows>72 must be ignored", wb, None)

    # ---- 23 wrong_sheet_name ----
    @reg
    def s23():
        wb = fresh_wb(template_path)
        cp = wb[SHEET_COVER]
        write_client(cp, rnd)
        set_scope(cp, {"web_app_vapt"})
        fill_matrix_quantities(wb[SHEET_MATRIX], {55: 2})
        wb[SHEET_MATRIX].title = "Scoping Matrix (renamed)"
        return ("23_wrong_sheet_name.xlsx",
                "INVALID: Overall Scoping Matrix renamed -> missing_sheet", wb, None)

    # ---- 24 missing_sheet (delete Cover Page) ----
    @reg
    def s24():
        wb = fresh_wb(template_path)
        fill_matrix_quantities(wb[SHEET_MATRIX], {55: 2})
        del wb[SHEET_COVER]
        return ("24_missing_sheet.xlsx",
                "INVALID: Cover Page deleted -> missing_sheet", wb, None)

    # ---- 25 quantity_header_changed ----
    @reg
    def s25():
        wb = fresh_wb(template_path)
        cp = wb[SHEET_COVER]
        write_client(cp, rnd)
        set_scope(cp, {"web_app_vapt"})
        mx = wb[SHEET_MATRIX]
        fill_matrix_quantities(mx, {55: 2})
        mx["D2"] = "Qty"  # header renamed
        return ("25_quantity_header_changed.xlsx",
                "INVALID: matrix D2 header changed -> quantity_header_changed", wb, None)

    # ---- 26 empty_workbook ----
    @reg
    def s26():
        wb = openpyxl.Workbook()  # brand-new, no template sheets
        wb.active["A1"] = "empty"
        return ("26_empty_workbook.xlsx",
                "INVALID: blank workbook, no scoping sheets", wb, None)

    # ---- 27 csv_renamed_xlsx ----
    @reg
    def s27():
        return ("27_csv_renamed_xlsx.xlsx",
                "INVALID: CSV bytes with .xlsx extension", None, "csv")

    # ---- 28 oversize ----
    @reg
    def s28():
        wb = fresh_wb(template_path)
        cp = wb[SHEET_COVER]
        write_client(cp, rnd)
        set_scope(cp, {"web_app_vapt"})
        fill_matrix_quantities(wb[SHEET_MATRIX], {55: 2})
        pad = wb.create_sheet("PADDING")
        pad.sheet_state = "hidden"
        # Incompressible random content so the zipped file genuinely exceeds 5 MB.
        alphabet = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"
        for r in range(1, 361):
            for c in range(1, 21):
                blob = "".join(rnd.choice(alphabet) for _ in range(1200))
                pad.cell(row=r, column=c, value=blob)
        return ("28_oversize.xlsx",
                "INVALID: > 5 MB via hidden padding sheet", wb, None)

    # ---- 29 corrupt_zip ----
    @reg
    def s29():
        wb = fresh_wb(template_path)
        cp = wb[SHEET_COVER]
        write_client(cp, rnd)
        set_scope(cp, {"web_app_vapt"})
        fill_matrix_quantities(wb[SHEET_MATRIX], {55: 2})
        return ("29_corrupt_zip.xlsx",
                "INVALID: truncated/corrupt OOXML container", wb, "truncate")

    # ---- 30 all_types_No ----
    @reg
    def s30():
        wb = fresh_wb(template_path)
        cp = wb[SHEET_COVER]
        write_client(cp, rnd)
        set_scope(cp, set())  # everything No
        fill_matrix_quantities(wb[SHEET_MATRIX], {55: 2, 24: 3})
        return ("30_all_types_No.xlsx",
                "INVALID: nothing in scope -> no_assessment_types_in_scope", wb, None)

    # ---- 31 image_example_counts (your screenshot: SaaS 3/12, Web 6/23) ----
    @reg
    def s31():
        wb = fresh_wb(template_path)
        cp = wb[SHEET_COVER]
        write_client(cp, rnd)
        set_scope(cp, {"saas_security", "web_app_vapt"})
        fill_matrix_quantities(wb[SHEET_MATRIX],
                               {49: 5, 52: 4, 53: 3,               # SaaS: count 3, qty 12
                                55: 2, 56: 1, 57: 4, 58: 5, 59: 3, 60: 8})  # Web: 6, 23
        return ("31_image_example_counts.xlsx",
                "VALID: screenshot example SaaS count 3/qty 12, Web count 6/qty 23", wb, None)

    # ---- 32 zero_blank_negative_rows ----
    @reg
    def s32():
        wb = fresh_wb(template_path)
        cp = wb[SHEET_COVER]
        write_client(cp, rnd)
        set_scope(cp, {"cloud_security", "mobile_vapt", "source_code_review"})
        fill_matrix_quantities(wb[SHEET_MATRIX],
                               {24: 3, 25: 0, 27: 2, 28: 0, 29: -4,     # Cloud: count 2, qty 5, 1 negative warning
                                62: 4, 63: 0, 64: 10, 65: "Y",           # Mobile: count 2, qty 14, 1 text warning
                                67: 0, 68: 0})                           # Source: count 0 -> zero confirm
        return ("32_zero_blank_negative_rows.xlsx",
                "EDGE: 0/blank rows skipped, negative + text warned, block sum 0 -> zero_confirm", wb, None)

    return scenarios


# --------------------------------------------------------------------------
# Post-processing (recalc / corrupt / csv) and save
# --------------------------------------------------------------------------
def soffice_recalc(path):
    """Open+save via LibreOffice headless so formula cached values exist.

    Converts into a separate temp dir (LibreOffice will not recalc when the
    output path equals the source), then moves the result back over `path`.
    """
    import subprocess
    import tempfile
    env = dict(os.environ)
    tmp = tempfile.mkdtemp(prefix="lo_recalc_")
    profile = os.path.join(tmp, ".lo_profile")
    cmd = [
        "soffice", "--headless", "--calc",
        "--convert-to", "xlsx:Calc MS Excel 2007 XML",
        "--outdir", tmp, path,
        f"-env:UserInstallation=file://{profile}",
    ]
    try:
        subprocess.run(cmd, env=env, check=True,
                       stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
                       timeout=120)
        produced = os.path.join(tmp, os.path.basename(path))
        if os.path.exists(produced):
            shutil.move(produced, path)
        else:
            raise RuntimeError("LibreOffice produced no output file")
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def inject_cached_values(path):
    """Fallback for machines without LibreOffice: write the computed result into
    each formula cell's <v> so the file looks like one saved by Excel.

    Only literal arithmetic and SUM(literals) are evaluated (all this test set
    uses); anything else is left without a cached value.
    """
    import re
    simple_sum = re.compile(r"^SUM\(([\d\s,.]+)\)$")
    arith = re.compile(r"^[\d\s+\-*/().]+$")

    def evaluate(formula):
        m = simple_sum.match(formula)
        if m:
            return sum(float(n) for n in m.group(1).split(",") if n.strip())
        if arith.match(formula):
            return eval(formula, {"__builtins__": {}}, {})  # validated: digits/operators only
        return None

    cell = re.compile(r'(<c r="[A-Z]+\d+"[^>]*>)<f>([^<]*)</f><v></v></c>')

    def fix(match):
        value = evaluate(match.group(2))
        if value is None:
            return match.group(0)
        if float(value).is_integer():
            value = int(value)
        return f"{match.group(1)}<f>{match.group(2)}</f><v>{value}</v></c>"

    tmp = path + ".tmp"
    with zipfile.ZipFile(path) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename.startswith("xl/worksheets/sheet"):
                data = cell.sub(fix, data.decode("utf-8")).encode("utf-8")
            zout.writestr(item, data)
    os.replace(tmp, path)


def save_workbook(wb, path, post):
    if post == "csv":
        with open(path, "w", newline="", encoding="utf-8") as fh:
            w = csv.writer(fh)
            w.writerow(["Client Organization Name", "Acme Example Pvt Ltd"])
            w.writerow(["In Scope", "Web Application VAPT", "Yes"])
            w.writerow(["Quantity", 5])
        return
    wb.save(path)
    if post == "recalc":
        try:
            soffice_recalc(path)
        except Exception as exc:
            print(f"  LibreOffice not usable ({type(exc).__name__}); writing cached values directly")
            inject_cached_values(path)
    elif post == "truncate":
        size = os.path.getsize(path)
        with open(path, "r+b") as fh:
            fh.truncate(size // 2)


# --------------------------------------------------------------------------
# Randomised load set
# --------------------------------------------------------------------------
# Matrix rows that feed each type's quantity (for load-set population).
TYPE_ROWS = {key: rows for key, _a, rows, _l in TYPE_RULES}
TYPE_ROWS.update({key: [] for key in DEFAULT_ONE_TYPES})
TEXT_JUNK = ["ten", "N/A", "10.0.0.0/24", "Y", "-", "many", "TBD"]


def build_load_file(template_path, rnd, idx):
    wb = fresh_wb(template_path)
    cp = wb[SHEET_COVER]
    write_client(cp, rnd)
    n = rnd.randint(1, 14)
    keys = set(rnd.sample(ALL_TYPES, n))
    set_scope(cp, keys)
    mx = wb[SHEET_MATRIX]
    # Populate quantity rows for the chosen types. sorted() makes the order in
    # which rnd is consumed independent of set iteration order (determinism).
    rows = sorted({r for k in keys for r in TYPE_ROWS[k]})
    for r in rows:
        roll = rnd.random()
        if roll < 0.10:
            mx[f"D{r}"] = 0
        elif roll < 0.15:
            mx[f"D{r}"] = rnd.choice(TEXT_JUNK)
        else:
            mx[f"D{r}"] = rnd.randint(0, 500)
    fill_all_context(wb, rnd, density=rnd.choice([0.0, 0.3, 0.6, 0.9]))
    return f"load_{idx:03d}.xlsx", wb


def build_stress_file(template_path, rnd, idx):
    wb = fresh_wb(template_path)
    cp = wb[SHEET_COVER]
    write_client(cp, rnd)
    set_scope(cp, set(ALL_TYPES))
    q = {r: rnd.randint(1, 1000) for r in MATRIX_DATA_ROWS}
    fill_matrix_quantities(wb[SHEET_MATRIX], q, opt_text=True, rnd=rnd)
    fill_all_context(wb, rnd, density=1.0, text_len=1000, all_cols=True)
    return f"stress_{idx:02d}.xlsx", wb


# --------------------------------------------------------------------------
# Main
# --------------------------------------------------------------------------
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--template", default=os.path.join(HERE, "template.xlsx"))
    ap.add_argument("--out", default=HERE)
    args = ap.parse_args()

    template_path = args.template
    if not os.path.exists(template_path):
        sys.exit(f"template not found: {template_path}")

    out = args.out
    samples = os.path.join(out, "samples")
    load_dir = os.path.join(samples, "load")
    expected = os.path.join(out, "expected")
    for d in (samples, expected):
        if os.path.exists(d):
            shutil.rmtree(d)
    os.makedirs(load_dir)
    os.makedirs(expected)

    rnd = random.Random(SEED)
    manifest_rows = []

    def emit(fname, label, path):
        exp = reference_parser(path)
        rel = os.path.relpath(path, samples)
        json_name = rel.replace(os.sep, "__").rsplit(".", 1)[0] + ".json"
        with open(os.path.join(expected, json_name), "w", encoding="utf-8") as fh:
            json.dump(exp, fh, indent=2, ensure_ascii=False)
        primary_err = exp["errors"][0].split(":")[0] if exp["errors"] else ""
        manifest_rows.append({
            "file": rel.replace(os.sep, "/"),
            "scenario": label,
            "expected_included_count": len(exp["included_annexures"]),
            "expected_warnings": len(exp["warnings"]),
            "expected_error": primary_err,
        })

    # ---- named scenarios ----
    print("Generating 32 named scenarios...")
    for builder in build_named_scenarios(template_path, rnd):
        fname, label, wb, post = builder()
        path = os.path.join(samples, fname)
        save_workbook(wb, path, post)
        emit(fname, label, path)
        print(f"  {fname}")

    # ---- load set ----
    print("Generating 100 randomised load files...")
    for i in range(1, 101):
        fname, wb = build_load_file(template_path, rnd, i)
        path = os.path.join(load_dir, fname)
        wb.save(path)
        emit(fname, f"LOAD: randomised subset (seed {SEED})", path)
    print("Generating 5 stress files...")
    for i in range(1, 6):
        fname, wb = build_stress_file(template_path, rnd, i)
        path = os.path.join(load_dir, fname)
        wb.save(path)
        emit(fname, "STRESS: 1000-char cells, max rows, all app columns", path)

    # ---- manifest ----
    manifest_path = os.path.join(out, "manifest.csv")
    with open(manifest_path, "w", newline="", encoding="utf-8") as fh:
        w = csv.DictWriter(fh, fieldnames=["file", "scenario",
                           "expected_included_count", "expected_warnings",
                           "expected_error"])
        w.writeheader()
        w.writerows(manifest_rows)

    # ---- summary table ----
    print("\n==================== SUMMARY ====================")
    print(f"{'FILE':42} {'INCL':>4} {'WARN':>4}  ERROR")
    print("-" * 70)
    for row in manifest_rows:
        if row["file"].startswith("load/"):
            continue
        print(f"{row['file']:42} {row['expected_included_count']:>4} "
              f"{row['expected_warnings']:>4}  {row['expected_error']}")
    nload = sum(1 for r in manifest_rows if r["file"].startswith("load/load_"))
    nstress = sum(1 for r in manifest_rows if r["file"].startswith("load/stress_"))
    print("-" * 70)
    print(f"named scenarios: {len(manifest_rows) - nload - nstress}   load: {nload}   stress: {nstress}   "
          f"total files: {len(manifest_rows)}")
    print(f"manifest: {manifest_path}")


if __name__ == "__main__":
    main()

"""
Unified Excel Parser for Standardised VAPT Report Format  (FIXED - v2)
This module provides a single source of truth for parsing the standardised
29-column Excel format used across all report generators
(main.py, type2.py, type3.py, type4.py).

BUGS FIXED:
  BUG-1  No alias existed for the per-vulnerability "Status" (Open/Closed) column.
         The parser was sharing the vapt_status mapping so every vuln showed
         "Completed".  A dedicated "status" mapping has been added.
  BUG-2  "observation" and "observation_vuln" aliases overlapped with
         "detailed_observation" via partial matching, causing the wrong column
         to be picked.  Alias lists and partial-match logic tightened.
  BUG-3  "Sr no." detection did not distinguish the Cert-IN "#N" style from the
         Web/IP "VUL-NNN" style, so sr_no was sometimes left blank.
         Explicit fallback to the Observation column added.
  BUG-4  Severity counts were computed only over non-blank Observation rows,
         so assets with multiple vulns on blank-asset rows were under-counted.
         Now all rows are counted using forward-filled asset names.
"""

import pandas as pd
import re
from typing import Dict, List, Any, Optional
import os


# ── Column normalisation ───────────────────────────────────────────────────────

def normalize_column_name(col: str) -> str:
    """Lower-case, trim, and replace _ / - with spaces for loose matching."""
    if not isinstance(col, str):
        col = str(col)
    normalized = " ".join(col.lower().strip().split())
    return normalized.replace("_", " ").replace("-", " ").replace("/", " ")


# ── Column mapping dictionary ──────────────────────────────────────────────────
# Lists are checked in order — most specific / common names come first.
# Partial matching is ONLY applied when exact matching fails and only when
# the candidate string is at least 6 characters long to prevent false hits.

COLUMN_MAPPINGS: Dict[str, List[str]] = {
    # ── Asset-level (cols 1-11) ──────────────────────────────────────────────
    "asset": [
        "asset hostname", "asset/hostname",
        "asset", "hostname", "host", "target", "server",
    ],
    "instant_purpose": [
        "instant purpose", "instant_purpose", "purpose",
        "audit purpose", "assessment purpose",
    ],

    # FIX-1: "vapt_status" no longer contains "status" — that is now "status" key
    "vapt_status": [
        "vapt status", "vapt_status", "assessment status", "audit status",
    ],

    "critical_count":       ["critical"],
    "high_count":           ["high"],
    "medium_count":         ["medium"],
    "low_count":            ["low"],
    "informational_count":  ["informational", "info"],
    "total_count":          ["total", "total count", "total vulnerabilities"],
    "tester_name": [
        "tester name", "tester_name", "tester", "reported by",
        "auditor", "assessor", "analyst",
    ],
    "project": ["project", "project name", "engagement"],
    "client":  ["client", "client name", "organization"],

    # ── Vuln-level (cols 12-29) ──────────────────────────────────────────────

    # Sr no. — Web/IP style ("VUL-001")
    "sr_no": [
        "sr no.", "sr no", "sr.no", "sr.no.", "serial no",
        "s.no", "s no", "sno", "sr", "vul id", "vuln id",
    ],

    # Observation — Cert-IN style ("#1", "#2")
    "observation": ["observation"],

    "severity": ["severity", "risk", "risk level", "risk(severity)", "criticality"],

    # FIX-1: dedicated "status" key for per-vulnerability Open/Closed column
    "status": [
        "status", "vuln status", "finding status",
        "vulnerability status", "open closed",
    ],

    "new_or_re": [
        "new or repeat observation", "new or repeat",
        "new/repeat", "observation type", "type of finding",
    ],
    "cve_cwe":  ["cve cwe", "cve/cwe", "cve", "cwe", "cve cwe id"],
    "cvss":     ["cvss", "cvss score", "cvss base score", "score"],
    "cvss_vector": [
        "cvss version ref", "cvss vector", "cvss version",
        "cvss string", "cvss ref", "cvss ver",
    ],
    "affected_asset": [
        "affected asset i.e. ip url application etc.",
        "affected asset", "vulnerable asset",
    ],

    # FIX-2: "observation_vuln" → exact match only for the title column
    "observation_vuln": [
        "observation vulnerability title",
        "observation/vulnerability title",
        "observation/ vulnerability title",
        "vulnerability title",
        "vuln title",
        "finding title",
        "title",
        "finding",
        "vulnerability name",
    ],

    # FIX-2: "detailed_observation" — long specific names first to avoid partial hits
    "detailed_observation": [
        "detailed observation  vulnerable point",
        "detailed observation vulnerable point",
        "detailed observation/vulnerable point",
        "detailed observation/ vulnerable point",
        "detailed observation",
        "vulnerability description",
        "vulnerability details",
        "description",
        "details",
    ],

    "recommendation": [
        "recommendation", "remediation", "fix", "solution",
        "recommended fix", "mitigation",
    ],
    "reference": ["reference", "references", "reference link", "links"],
    "evidence": [
        "evidence   proof of concept",
        "evidence / proof of concept",
        "evidence proof of concept",
        "evidence/proof of concept",
        "proof of concept",
        "evidence",
        "poc",
        "steps to reproduce",
        "steps",
    ],
    "screenshot": ["screenshot", "screen shot", "image", "poc image"],
}


# ── Column finder ──────────────────────────────────────────────────────────────

def find_column(df: pd.DataFrame, standard_key: str) -> Optional[str]:
    """
    Find the actual column name in df that best matches standard_key.
    1. Exact normalised match
    2. Partial match (only for candidates >= 6 chars to prevent false hits)
    Returns the original (un-normalised) column name, or None.
    """
    if standard_key not in COLUMN_MAPPINGS:
        return None

    possible_names = COLUMN_MAPPINGS[standard_key]
    df_cols_norm   = {normalize_column_name(col): col for col in df.columns}

    # Pass 1 — exact
    for name in possible_names:
        norm = normalize_column_name(name)
        if norm in df_cols_norm:
            return df_cols_norm[norm]

    # Pass 2 — partial (longer names only, to reduce false hits)
    for name in possible_names:
        norm = normalize_column_name(name)
        if len(norm) < 6:
            continue
        for df_norm, df_orig in df_cols_norm.items():
            if len(df_norm) < 6:
                continue
            if df_norm.startswith(norm) or norm.startswith(df_norm):
                return df_orig

    return None


def find_screenshot_columns(df: pd.DataFrame) -> List[str]:
    """Return all column names that look like screenshot columns, in order."""
    return [
        col for col in df.columns
        if any(kw in normalize_column_name(col)
               for kw in ("screenshot", "screen shot", "poc image"))
    ]


def find_step_columns(df: pd.DataFrame) -> Dict[int, str]:
    """Return {step_number: column_name} for all Step-N columns."""
    result = {}
    for col in df.columns:
        m = re.search(r"step\s*(\d+)", normalize_column_name(col))
        if m:
            result[int(m.group(1))] = col
    return result


def safe_get(row: pd.Series, col_name: Optional[str], default: Any = "") -> Any:
    """Safely fetch a cell value; return default on None / NaN / missing."""
    if col_name is None or col_name not in row.index:
        return default
    v = row[col_name]
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return default
    if isinstance(v, str) and v.strip() == "":
        return default
    return v


# ── Severity helpers ───────────────────────────────────────────────────────────

def _normalise_severity(raw: Any) -> str:
    s = str(raw).lower()
    if "critical" in s: return "Critical"
    if "high"     in s: return "High"
    if "medium"   in s or "moderate" in s: return "Medium"
    if "low"      in s: return "Low"
    if "info"     in s: return "Informational"
    return str(raw)   # preserve unknown values


def _is_certin_obs(val: str) -> bool:
    return bool(re.fullmatch(r"#\d+", str(val).strip()))


# ── Main parser ────────────────────────────────────────────────────────────────

def parse_excel_data(excel_file_path: str) -> Dict[str, Any]:
    """
    Parse the standardised 29-column VAPT Excel and return structured data.

    Returns
    -------
    {
        'metadata':        { client, project, tester },
        'assets':          [ { asset, purpose, status, critical, high, … } ],
        'vulnerabilities': [ { sr_no, observation, severity, status, … } ],
        'raw_df':          pd.DataFrame
    }
    """
    df = pd.read_excel(excel_file_path, sheet_name=0, dtype=str)
    df.columns = [str(c).strip() for c in df.columns]
    df = df.dropna(how="all").reset_index(drop=True)

    # Avoid emoji in logs: some Windows environments default to cp1252 and crash on Unicode.
    print(f"Loaded: {len(df)} rows x {len(df.columns)} cols - {excel_file_path}")
    print(f"Columns: {list(df.columns)}")

    # ── Resolve all column mappings ───────────────────────────────────────────
    col = {k: find_column(df, k) for k in COLUMN_MAPPINGS}
    screenshot_cols = find_screenshot_columns(df)
    step_cols       = find_step_columns(df)

    print(f"Mapped -> observation:{col['observation']} | severity:{col['severity']} "
          f"| status:{col['status']} | vapt_status:{col['vapt_status']} "
          f"| tester:{col['tester_name']} | project:{col['project']}")
    print(f"   steps:{sorted(step_cols.keys())} | screenshots:{screenshot_cols}")

    # ── Metadata (first non-empty value wins) ─────────────────────────────────
    def _first_val(col_key, placeholder):
        c = col[col_key]
        if not c:
            return placeholder
        for _, row in df.iterrows():
            v = safe_get(row, c)
            if v:
                return v
        return placeholder

    metadata = {
        "client":  _first_val("client",      "Client Name"),
        "project": _first_val("project",      "Project Name"),
        "tester":  _first_val("tester_name",  "Tester Name"),
    }

    # ── BUG-4 FIX: forward-fill asset for counts ──────────────────────────────
    asset_ff = []
    last_asset = ""
    for _, row in df.iterrows():
        raw = safe_get(row, col["asset"])
        if raw:
            last_asset = str(raw).strip()
        asset_ff.append(last_asset)

    # ── Asset summaries ───────────────────────────────────────────────────────
    assets = []
    seen_assets: set = set()

    for idx, row in df.iterrows():
        asset_name = safe_get(row, col["asset"])
        if asset_name and asset_name not in seen_assets:
            seen_assets.add(asset_name)
            assets.append({
                "asset":         asset_name,
                "purpose":       safe_get(row, col["instant_purpose"]),
                "status":        safe_get(row, col["vapt_status"]),
                "critical":      safe_get(row, col["critical_count"],      0),
                "high":          safe_get(row, col["high_count"],           0),
                "medium":        safe_get(row, col["medium_count"],         0),
                "low":           safe_get(row, col["low_count"],            0),
                "informational": safe_get(row, col["informational_count"],  0),
                "total":         safe_get(row, col["total_count"],          0),
            })

    # ── Vulnerability details ─────────────────────────────────────────────────
    vulnerabilities = []

    # Determine identifier column — prefer sr_no for Web/IP, observation for Cert-IN
    has_sr_no  = col["sr_no"] is not None
    has_obs    = col["observation"] is not None

    for df_idx, (_, row) in enumerate(df.iterrows()):
        # ── Find the row identifier ───────────────────────────────────────────
        sr_val  = safe_get(row, col["sr_no"])
        obs_val = safe_get(row, col["observation"])

        # BUG-3 FIX: normalise junk header values to empty
        _junk = {"sr no.", "sr no", "s.no", "serial", "observation", "nan", ""}
        sr_val  = "" if str(sr_val).lower().strip()  in _junk else str(sr_val).strip()
        obs_val = "" if str(obs_val).lower().strip() in _junk else str(obs_val).strip()

        # Use whatever identifier is present
        identifier = sr_val or obs_val
        if not identifier:
            continue  # blank row — skip

        # ── Resolve vuln title ────────────────────────────────────────────────
        obs_title = safe_get(row, col["observation_vuln"])
        if not obs_title:
            obs_title = identifier   # fallback to the ID if title is missing

        # ── Severity & status ─────────────────────────────────────────────────
        severity    = _normalise_severity(safe_get(row, col["severity"], "Medium"))
        # FIX-1: per-vuln status from "Status" col (Open/Closed)
        vuln_status = safe_get(row, col["status"], "Open")
        if not vuln_status:
            vuln_status = "Open"

        # ── Steps ─────────────────────────────────────────────────────────────
        steps: List[Dict[str, Any]] = []
        for step_num in sorted(step_cols.keys()):
            content = safe_get(row, step_cols[step_num])
            if content:
                steps.append({"number": step_num, "content": content})

        # If no dedicated step cols, parse from evidence
        evidence_text = safe_get(row, col["evidence"])
        if not steps and evidence_text:
            for m in re.finditer(
                r"Step\s*(\d+)\s*[:\-\.]?\s*(.*?)(?=Step\s*\d+|$)",
                str(evidence_text),
                re.IGNORECASE | re.DOTALL,
            ):
                content = m.group(2).strip()
                if content:
                    steps.append({"number": int(m.group(1)), "content": content})

        # ── Screenshots ───────────────────────────────────────────────────────
        screenshots = [
            safe_get(row, sc_col)
            for sc_col in screenshot_cols
            if safe_get(row, sc_col)
        ]

        # ── Asset (forward-filled) ────────────────────────────────────────────
        asset = asset_ff[df_idx]

        vuln = {
            "sr_no":               sr_val,
            "observation_number":  obs_val,          # Cert-IN #N
            "observation":         obs_title,         # human-readable title
            "severity":            severity,
            "status":              vuln_status,       # FIX-1
            "new_or_re":           safe_get(row, col["new_or_re"], "New"),
            "cve_cwe":             safe_get(row, col["cve_cwe"]),
            "cvss":                safe_get(row, col["cvss"]),
            "cvss_vector":         safe_get(row, col["cvss_vector"]),
            "affected_asset":      safe_get(row, col["affected_asset"]),
            "observation_summary": safe_get(row, col["observation_vuln"]),
            "detailed_observation":safe_get(row, col["detailed_observation"]),
            "recommendation":      safe_get(row, col["recommendation"]),
            "reference":           safe_get(row, col["reference"]),
            "evidence":            evidence_text,
            "steps":               steps,
            "screenshots":         screenshots,
            "tester":              safe_get(row, col["tester_name"], metadata["tester"]),
            "project":             safe_get(row, col["project"],     metadata["project"]),
            "client":              safe_get(row, col["client"],      metadata["client"]),
            "asset":               asset,
        }

        print(f"   Vuln {df_idx}: {obs_title[:50]!r} | sev={severity} | "
              f"status={vuln_status} | steps={len(steps)}")
        vulnerabilities.append(vuln)

    return {
        "metadata":        metadata,
        "assets":          assets,
        "vulnerabilities": vulnerabilities,
        "raw_df":          df,
    }


# ── Severity count helper ──────────────────────────────────────────────────────

def get_severity_counts(vulnerabilities: List[Dict[str, Any]]) -> Dict[str, int]:
    counts = {"critical": 0, "high": 0, "medium": 0, "low": 0, "informational": 0, "total": 0}
    for v in vulnerabilities:
        sev = str(v.get("severity", "")).lower()
        counts["total"] += 1
        if   "critical" in sev: counts["critical"] += 1
        elif "high"     in sev: counts["high"]     += 1
        elif "medium"   in sev: counts["medium"]   += 1
        elif "low"      in sev: counts["low"]      += 1
        elif "info"     in sev: counts["informational"] += 1
    return counts


# ── Template formatter ─────────────────────────────────────────────────────────

def format_for_template(
    parsed_data: Dict[str, Any],
    template_type: str = "generic",
) -> Dict[str, Any]:
    """Format parsed data for a specific report generator template."""
    meta  = parsed_data["metadata"]
    vulns = parsed_data["vulnerabilities"]
    assets= parsed_data["assets"]
    sev   = get_severity_counts(vulns)

    base = {
        "client_name":   meta["client"],
        "project_name":  meta["project"],
        "tester_name":   meta["tester"],
        "vulnerabilities": vulns,
        "assets":          assets,
        "severity_counts": sev,
    }

    if template_type in ("certin", "type3"):
        return {**base, "CLIENT_NAME": meta["client"], "PROJECT_NAME": meta["project"],
                "TESTER_NAME": meta["tester"]}
    if template_type == "type2":
        return {**base, "severity_summary": sev}
    if template_type in ("type4", "main"):
        return base

    # generic
    return {"metadata": meta, "assets": assets,
            "vulnerabilities": vulns, "severity_counts": sev}
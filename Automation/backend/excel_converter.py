"""
Shared Excel Converter Module
Converts any VAPT tracker Excel (any format) into the standardised 29-column
output format used by IP (type1/type2), Web (type4), and Cert-IN (type3) report generators.

POC ZIP layouts supported:
  "ip"     : POC Screenshots*/<IP>/<Severity>/VUL-NNN/stepN.png
  "web"    : POC Screenshots/<Severity>/VUL-NNN/stepN.png
  "certin" : <root>/#N/stepN.png   (observation-number folders)
"""

import os
import re
import zipfile
import tempfile
import shutil
from datetime import datetime

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# ── Output column schema (29 columns) ─────────────────────────────────────────
OUTPUT_COLUMNS = [
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
    "Screenshot1",
    "Screenshot2",
    "Screenshot3",
    "Screenshot4",
]

SEVERITY_ORDER = ["Critical", "High", "Medium", "Low", "Informational"]

SEVERITY_COLORS = {
    "Critical":      "FF0000",
    "High":          "FF6600",
    "Medium":        "FFC000",
    "Low":           "FFFF00",
    "Informational": "92D050",
}

HEADER_FILL_COLOR = "7030A0"

COLUMN_WIDTHS = {
    "Asset/Hostname": 20,
    "Instant purpose": 18,
    "VAPT Status": 14,
    "Critical": 10,
    "High": 10,
    "Medium": 10,
    "Low": 10,
    "Informational": 14,
    "Total": 10,
    "Tester_Name": 18,
    "Project": 20,
    "Sr no.": 8,
    "Observation": 30,
    "Severity": 14,
    "Status": 14,
    "New or Repeat Observation": 22,
    "CVE/CWE": 18,
    "CVSS": 12,
    "CVSS Version Ref": 16,
    "Affected Asset i.e. IP/URL/Application etc.": 35,
    "Observation/Vulnerability Title": 35,
    "Detailed Observation/ Vulnerable Point": 40,
    "Recommendation": 40,
    "Reference": 30,
    "Evidence / Proof of Concept": 50,
    "Screenshot1": 20,
    "Screenshot2": 20,
    "Screenshot3": 20,
    "Screenshot4": 20,
}

# ── Flexible column name aliases ───────────────────────────────────────────────
_COL_ALIASES = {
    "asset_hostname": [
        "Asset/Hostname", "Asset Hostname", "Hostname", "Asset", "IP Address",
        "IP", "Host", "Target", "Server",
    ],
    "instant_purpose": [
        "Instant purpose", "Purpose", "Audit Purpose", "Assessment Purpose",
        "Type", "Instant Purpose",
    ],
    "vapt_status": [
        "VAPT Status", "Status", "Assessment Status", "Audit Status",
        "Revalidation Status", "Revalidation",
    ],
    "tester_name": [
        "Tester_Name", "Tester Name", "Tester", "Reported By", "Auditor",
        "Assessor", "Analyst",
    ],
    "project": [
        "Project", "Project Name", "Engagement", "Client Project",
        "Assessment Name",
    ],
    "sr_no": [
        "Sr no.", "Sr no", "Sr No", "Sr.No", "Sr. No.", "S.No", "Serial",
        "S. No", "SNo", "Sr", "VUL ID", "Vuln ID", "ID", "S.NO.", "S.No.",
    ],
    "observation": [
        "Observation", "Vulnerability Name", "Name of Vulnerability",
        "Finding", "Title", "Vuln Title", "Vulnerability Title",
        "Observation/Vulnerability Title", "Observation/ Vulnerability Title",
    ],
    "severity": [
        "Severity", "Risk", "Risk Level", "Risk(Severity)", "Criticality",
        "Priority",
    ],
    "new_repeat": [
        "New or Repeat Observation", "New or Repeat", "New/Repeat",
        "Observation Type", "Type of Finding",
    ],
    "cve_cwe": [
        "CVE/CWE", "CVE CWE", "CVE", "CWE", "CVE/CWE ID",
    ],
    "cvss": [
        "CVSS", "CVSS Score", "Score", "CVSS Base Score",
    ],
    "cvss_version": [
        "CVSS Version Ref", "CVSS Version", "CVSS Ref", "CVSS Ver",
    ],
    "affected_asset": [
        "Affected Asset i.e. IP/URL/Application etc.",
        "Affected Asset", "IP/URL/App", "IP URL App", "Target IP",
        "Vulnerable URL", "Vulnerable Parameter", "URL", "Endpoint",
        "Affected URL", "Affected IP",
    ],
    "vuln_title": [
        "Observation/Vulnerability Title", "Observation/ Vulnerability Title",
        "Vulnerability Title", "Vuln Title", "Title",
    ],
    "detailed_observation": [
        "Detailed Observation/ Vulnerable Point",
        "Detailed Observation/Vulnerable Point",
        "Detailed Observation", "Description", "Vulnerability Description",
        "Observation/Vulnerability", "Detailed O", "Detail",
        "Vulnerability Details", "Details",
    ],
    "recommendation": [
        "Recommendation", "Remediation", "Fix", "Solution",
        "Recommended Fix", "Mitigation",
    ],
    "reference": [
        "Reference", "References", "Reference Link", "Links",
    ],
    "evidence": [
        "Evidence / Proof of Concept", "Evidence / Proof Of Concept",
        "Evidence/Proof of Concept", "Evidence", "Proof of Concept",
        "POC", "Steps", "Steps to Reproduce",
    ],
}

# ── Helpers ────────────────────────────────────────────────────────────────────

def _clean(val):
    """Normalise a cell value to a stripped string."""
    if val is None or (isinstance(val, float) and str(val) == "nan"):
        return ""
    return str(val).strip()


def _find_col(df, aliases):
    """Return the first column in df whose name matches any alias (case-insensitive)."""
    lower_cols = {c.lower().strip(): c for c in df.columns}
    for alias in aliases:
        if alias.lower().strip() in lower_cols:
            return lower_cols[alias.lower().strip()]
    return None


def _normalise_severity(raw):
    """Map any severity text to one of the 5 canonical values."""
    s = _clean(raw).lower()
    if "critical" in s:
        return "Critical"
    if "high" in s:
        return "High"
    if "medium" in s or "moderate" in s:
        return "Medium"
    if "low" in s:
        return "Low"
    if "info" in s:
        return "Informational"
    return "Medium"


def _normalise_new_repeat(raw):
    s = _clean(raw).lower()
    if "repeat" in s or "follow" in s or "existing" in s:
        return "Repeat"
    return "New"


# ── Excel reader (flexible, handles row-1 and row-2 headers) ──────────────────

def _read_any_excel(path):
    """
    Read any VAPT Excel file.
    Tries header=0 first; if that yields fewer than 5 recognisable columns,
    tries header=1 (row 2 as header, like MitKat AMCS tracker).
    Returns a pandas DataFrame.
    """
    def _score(df):
        """Count how many of our key aliases match columns in df."""
        key_groups = ["sr_no", "observation", "severity", "affected_asset", "detailed_observation"]
        score = 0
        for g in key_groups:
            if _find_col(df, _COL_ALIASES[g]):
                score += 1
        return score

    try:
        df0 = pd.read_excel(path, header=0, dtype=str)
        df0.columns = [str(c).strip() for c in df0.columns]
        score0 = _score(df0)
    except Exception:
        df0, score0 = None, -1

    try:
        df1 = pd.read_excel(path, header=1, dtype=str)
        df1.columns = [str(c).strip() for c in df1.columns]
        score1 = _score(df1)
    except Exception:
        df1, score1 = None, -1

    if score0 >= score1 and df0 is not None:
        df = df0
    elif df1 is not None:
        df = df1
    else:
        raise ValueError("Could not read Excel file with any header row configuration.")

    # Drop completely empty rows
    df = df.dropna(how="all").reset_index(drop=True)
    return df


# ── Severity counts per asset ──────────────────────────────────────────────────

def _severity_counts(df, asset_col, severity_col):
    """
    Returns a dict: asset -> {Critical:N, High:N, Medium:N, Low:N, Informational:N, Total:N}
    """
    counts = {}
    for _, row in df.iterrows():
        asset = _clean(row.get(asset_col, "")) if asset_col else ""
        sev = _normalise_severity(row.get(severity_col, "")) if severity_col else "Medium"
        if not asset:
            continue
        if asset not in counts:
            counts[asset] = {s: 0 for s in SEVERITY_ORDER}
        counts[asset][sev] = counts[asset].get(sev, 0) + 1
    for asset in counts:
        counts[asset]["Total"] = sum(counts[asset][s] for s in SEVERITY_ORDER)
    return counts


# ── POC ZIP indexers ───────────────────────────────────────────────────────────

def _extract_zip(zip_path):
    """Extract ZIP to a temp dir, return (extract_dir, cleanup_dir)."""
    tmp = tempfile.mkdtemp()
    with zipfile.ZipFile(zip_path, "r") as zf:
        zf.extractall(tmp)
    return tmp, tmp


def _step_num(filename):
    """Extract step number from filename like step1.png, step-2.png, step_3.jpg."""
    m = re.search(r"step[\s_\-]?(\d+)", filename.lower())
    return int(m.group(1)) if m else None


def _index_poc_ip(root):
    """
    IP POC layout: <root>/POC Screenshots*/<IP>/<Severity>/VUL-NNN/stepN.png
    Returns: { (ip, vul_id): {step_num: abs_path} }
    Also returns flat: { sr_no_int: {step_num: abs_path} } keyed by VUL number.
    """
    index = {}  # vul_id_lower -> {step_num: path}
    # Find the POC Screenshots* folder
    poc_root = None
    for entry in os.listdir(root):
        if entry.lower().startswith("poc screenshots") or entry.lower().startswith("poc_screenshots"):
            candidate = os.path.join(root, entry)
            if os.path.isdir(candidate):
                poc_root = candidate
                break
    if not poc_root:
        poc_root = root

    for ip_folder in os.listdir(poc_root):
        ip_path = os.path.join(poc_root, ip_folder)
        if not os.path.isdir(ip_path):
            continue
        for sev_folder in os.listdir(ip_path):
            sev_path = os.path.join(ip_path, sev_folder)
            if not os.path.isdir(sev_path):
                continue
            for vul_folder in os.listdir(sev_path):
                vul_path = os.path.join(sev_path, vul_folder)
                if not os.path.isdir(vul_path):
                    continue
                vul_key = vul_folder.strip().lower()
                if vul_key not in index:
                    index[vul_key] = {}
                for fname in sorted(os.listdir(vul_path)):
                    if fname.lower().endswith((".png", ".jpg", ".jpeg")):
                        sn = _step_num(fname)
                        if sn is not None:
                            index[vul_key][sn] = os.path.join(vul_path, fname)

    # Also build numeric index: VUL-003 -> 3
    numeric_index = {}
    for vul_key, steps in index.items():
        m = re.search(r"(\d+)", vul_key)
        if m:
            numeric_index[int(m.group(1))] = steps
    return numeric_index


def _index_poc_web(root):
    """
    Web POC layout: <root>/POC Screenshots/<Severity>/VUL-NNN/stepN.png
    Returns: { vul_number_int: {step_num: abs_path} }
    """
    index = {}
    poc_root = None
    for entry in os.listdir(root):
        if entry.lower().startswith("poc screenshots") or entry.lower().startswith("poc_screenshots"):
            candidate = os.path.join(root, entry)
            if os.path.isdir(candidate):
                poc_root = candidate
                break
    if not poc_root:
        poc_root = root

    for sev_folder in os.listdir(poc_root):
        sev_path = os.path.join(poc_root, sev_folder)
        if not os.path.isdir(sev_path):
            continue
        for vul_folder in os.listdir(sev_path):
            vul_path = os.path.join(sev_path, vul_folder)
            if not os.path.isdir(vul_path):
                continue
            m = re.search(r"(\d+)", vul_folder)
            if not m:
                continue
            vul_num = int(m.group(1))
            if vul_num not in index:
                index[vul_num] = {}
            for fname in sorted(os.listdir(vul_path)):
                if fname.lower().endswith((".png", ".jpg", ".jpeg")):
                    sn = _step_num(fname)
                    if sn is not None:
                        index[vul_num][sn] = os.path.join(vul_path, fname)
    return index


def _index_poc_certin(root):
    """
    Cert-IN POC layout: <root>/<any_prefix>/#N/stepN.png
    Returns: { obs_number_int: {step_num: abs_path} }
    """
    index = {}

    def _walk(base):
        for entry in os.listdir(base):
            entry_path = os.path.join(base, entry)
            if os.path.isdir(entry_path):
                m = re.fullmatch(r"#(\d+)", entry.strip())
                if m:
                    obs_num = int(m.group(1))
                    index[obs_num] = {}
                    for fname in sorted(os.listdir(entry_path)):
                        if fname.lower().endswith((".png", ".jpg", ".jpeg")):
                            sn = _step_num(fname)
                            if sn is not None:
                                index[obs_num][sn] = os.path.join(entry_path, fname)
                else:
                    _walk(entry_path)

    _walk(root)
    return index


# ── Build POC evidence text and screenshot filenames ──────────────────────────

def _build_poc_info(raw_evidence, vul_number, poc_index):
    """
    Given raw evidence text and a vulnerability number (1-based int),
    return (evidence_text, [ss1, ss2, ss3, ss4]).

    poc_index: { vul_number_int: {step_num_int: abs_path} }
    """
    steps_dict = poc_index.get(vul_number, {}) if poc_index else {}
    sorted_steps = sorted(steps_dict.keys()) if steps_dict else []

    # Build evidence lines
    raw_lines = [l.strip() for l in _clean(raw_evidence).splitlines() if l.strip()]
    evidence_lines = []
    for i, sn in enumerate(sorted_steps, 1):
        # Try to use existing line text if available
        if i <= len(raw_lines):
            line_text = raw_lines[i - 1]
            if not re.match(r"^step\s*\d+", line_text, re.IGNORECASE):
                line_text = f"Step {sn}: {line_text}"
        else:
            line_text = f"Step {sn}: (screenshot)"
        evidence_lines.append(line_text)

    # If no POC images, use raw lines as-is
    if not sorted_steps and raw_lines:
        for i, line in enumerate(raw_lines, 1):
            if not re.match(r"^step\s*\d+", line, re.IGNORECASE):
                evidence_lines.append(f"Step {i}: {line}")
            else:
                evidence_lines.append(line)

    evidence_text = "\n".join(evidence_lines) if evidence_lines else _clean(raw_evidence)

    # Screenshot filenames (just the basename)
    screenshots = []
    for sn in sorted_steps[:4]:
        path = steps_dict[sn]
        screenshots.append(os.path.basename(path))

    while len(screenshots) < 4:
        screenshots.append("")

    return evidence_text, screenshots[:4]


# ── Map any DataFrame to the 29-column output ─────────────────────────────────

def _map_to_29_columns(df, poc_index, poc_layout):
    """
    Map a DataFrame (any column names) to the 29-column output schema.
    poc_index: { vul_number_int: {step_num: abs_path} }
    poc_layout: "ip" | "web" | "certin"
    Returns a list of dicts, one per row.
    """
    # Resolve column names
    c = {k: _find_col(df, v) for k, v in _COL_ALIASES.items()}

    rows = []
    # Track per-asset severity counts
    asset_sev_counts = {}

    # First pass: collect severity counts per asset
    for _, row in df.iterrows():
        asset = _clean(row[c["asset_hostname"]]) if c["asset_hostname"] else ""
        sev = _normalise_severity(row[c["severity"]]) if c["severity"] else "Medium"
        sr = _clean(row[c["sr_no"]]) if c["sr_no"] else ""
        # If sr_no is empty, try observation column (handles #1, #2 style)
        if not sr or sr.lower() in ("sr no.", "sr no", "s.no", "serial", "nan"):
            obs_val = _clean(row[c["observation"]]) if c["observation"] else ""
            if obs_val and obs_val not in ("nan", ""):
                sr = obs_val
            else:
                continue
        if not asset:
            # Try affected_asset as fallback for asset
            asset = _clean(row[c["affected_asset"]]) if c["affected_asset"] else ""
        if not asset:
            continue
        if asset not in asset_sev_counts:
            asset_sev_counts[asset] = {s: 0 for s in SEVERITY_ORDER}
        # Only count real vulnerabilities (not "no vulnerability" rows)
        if not re.search(r"no.{0,10}vuln", sr.lower()):
            asset_sev_counts[asset][sev] = asset_sev_counts[asset].get(sev, 0) + 1

    for asset in asset_sev_counts:
        asset_sev_counts[asset]["Total"] = sum(
            asset_sev_counts[asset][s] for s in SEVERITY_ORDER
        )

    # Second pass: build output rows
    for idx, row in df.iterrows():
        sr = _clean(row[c["sr_no"]]) if c["sr_no"] else ""
        # If sr_no is empty, try observation column (handles #1, #2 style)
        if not sr or sr.lower() in ("sr no.", "sr no", "s.no", "serial", "nan"):
            obs_val = _clean(row[c["observation"]]) if c["observation"] else ""
            if obs_val and obs_val not in ("nan", ""):
                sr = obs_val
            else:
                continue

        asset = _clean(row[c["asset_hostname"]]) if c["asset_hostname"] else ""
        # Fallback: use affected_asset if asset_hostname is empty
        if not asset or asset in ("nan", ""):
            asset = _clean(row[c["affected_asset"]]) if c["affected_asset"] else ""
        sev = _normalise_severity(row[c["severity"]]) if c["severity"] else "Medium"

        # Determine vulnerability number for POC lookup
        vul_num = None
        m = re.search(r"(\d+)", sr)
        if m:
            vul_num = int(m.group(1))

        raw_evidence = _clean(row[c["evidence"]]) if c["evidence"] else ""
        evidence_text, screenshots = _build_poc_info(raw_evidence, vul_num, poc_index)

        # Severity counts for this asset
        sc = asset_sev_counts.get(asset, {s: 0 for s in SEVERITY_ORDER})

        # Build the observation title: prefer vuln_title, fall back to observation
        obs_title = ""
        if c["vuln_title"]:
            obs_title = _clean(row[c["vuln_title"]])
        if not obs_title and c["observation"]:
            obs_title = _clean(row[c["observation"]])

        out = {
            "Asset/Hostname": asset,
            "Instant purpose": _clean(row[c["instant_purpose"]]) if c["instant_purpose"] else "Web Application",
            "VAPT Status": _clean(row[c["vapt_status"]]) if c["vapt_status"] else "Completed",
            "Critical": sc.get("Critical", 0),
            "High": sc.get("High", 0),
            "Medium": sc.get("Medium", 0),
            "Low": sc.get("Low", 0),
            "Informational": sc.get("Informational", 0),
            "Total": sc.get("Total", 0),
            "Tester_Name": _clean(row[c["tester_name"]]) if c["tester_name"] else "",
            "Project": _clean(row[c["project"]]) if c["project"] else "",
            "Sr no.": sr,
            "Observation": _clean(row[c["observation"]]) if c["observation"] else obs_title,
            "Severity": sev,
            "Status": _clean(row[c["vapt_status"]]) if c["vapt_status"] else "Open",
            "New or Repeat Observation": _normalise_new_repeat(
                row[c["new_repeat"]] if c["new_repeat"] else ""
            ),
            "CVE/CWE": _clean(row[c["cve_cwe"]]) if c["cve_cwe"] else "",
            "CVSS": _clean(row[c["cvss"]]) if c["cvss"] else "",
            "CVSS Version Ref": _clean(row[c["cvss_version"]]) if c["cvss_version"] else "",
            "Affected Asset i.e. IP/URL/Application etc.": (
                _clean(row[c["affected_asset"]]) if c["affected_asset"] else asset
            ),
            "Observation/Vulnerability Title": obs_title,
            "Detailed Observation/ Vulnerable Point": (
                _clean(row[c["detailed_observation"]]) if c["detailed_observation"] else ""
            ),
            "Recommendation": _clean(row[c["recommendation"]]) if c["recommendation"] else "",
            "Reference": _clean(row[c["reference"]]) if c["reference"] else "",
            "Evidence / Proof of Concept": evidence_text,
            "Screenshot1": screenshots[0],
            "Screenshot2": screenshots[1],
            "Screenshot3": screenshots[2],
            "Screenshot4": screenshots[3],
        }
        rows.append(out)

    return rows


# ── Write styled Excel ─────────────────────────────────────────────────────────

def _write_styled_excel(rows, output_path):
    """Write the 29-column output as a styled Excel file."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"

    header_font = Font(name="Calibri", bold=True, color="FFFFFF", size=11)
    header_fill = PatternFill("solid", fgColor=HEADER_FILL_COLOR)
    header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    thin = Side(border_style="thin", color="000000")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    wrap_align = Alignment(horizontal="left", vertical="top", wrap_text=True)

    # Write header row
    for col_idx, col_name in enumerate(OUTPUT_COLUMNS, 1):
        cell = ws.cell(row=1, column=col_idx, value=col_name)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_align
        cell.border = border
        ws.column_dimensions[get_column_letter(col_idx)].width = COLUMN_WIDTHS.get(col_name, 18)

    ws.row_dimensions[1].height = 30

    # Write data rows
    count_cols = {"Critical", "High", "Medium", "Low", "Informational", "Total"}
    center_cols = {"Sr no.", "Severity", "Status", "New or Repeat Observation",
                   "CVSS", "CVSS Version Ref", "CVE/CWE"}

    for row_idx, row_data in enumerate(rows, 2):
        sev = row_data.get("Severity", "")
        fill_color = SEVERITY_COLORS.get(sev)
        row_fill = PatternFill("solid", fgColor=fill_color) if fill_color else None

        for col_idx, col_name in enumerate(OUTPUT_COLUMNS, 1):
            value = row_data.get(col_name, "")
            cell = ws.cell(row=row_idx, column=col_idx, value=value)
            cell.border = border

            if col_name in count_cols:
                cell.alignment = center_align
                if row_fill:
                    cell.fill = row_fill
            elif col_name in center_cols:
                cell.alignment = center_align
                if row_fill:
                    cell.fill = row_fill
            else:
                cell.alignment = wrap_align
                if row_fill:
                    cell.fill = row_fill

    # Freeze header row
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(OUTPUT_COLUMNS))}1"

    wb.save(output_path)


# ── Main entry point ───────────────────────────────────────────────────────────

def convert_any_excel_to_format(input_path, poc_zip_path=None, poc_layout="ip"):
    """
    Convert any VAPT Excel to the 29-column standard format.

    Parameters
    ----------
    input_path   : str  - Path to any input VAPT Excel (.xlsx/.xls)
    poc_zip_path : str  - Path to POC ZIP file (optional)
    poc_layout   : str  - One of "ip", "web", "certin"

    Returns
    -------
    (output_path, extract_dir)
      output_path  : str  - Path to the written output Excel
      extract_dir  : str or None - Temp dir to clean up after use
    """
    extract_dir = None
    poc_index = {}

    try:
        df = _read_any_excel(input_path)
    except Exception as e:
        raise ValueError(f"Failed to read input Excel: {e}")

    if poc_zip_path and os.path.isfile(poc_zip_path):
        try:
            extract_dir, _ = _extract_zip(poc_zip_path)
            if poc_layout == "ip":
                poc_index = _index_poc_ip(extract_dir)
            elif poc_layout == "web":
                poc_index = _index_poc_web(extract_dir)
            elif poc_layout == "certin":
                poc_index = _index_poc_certin(extract_dir)
        except Exception:
            poc_index = {}

    rows = _map_to_29_columns(df, poc_index, poc_layout)

    if not rows:
        raise ValueError(
            "No vulnerability rows found in the input Excel. "
            "Please ensure the file has Sr no. and Severity columns."
        )

    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    layout_label = {"ip": "IP", "web": "Web", "certin": "CertIN"}.get(poc_layout, poc_layout.upper())
    output_filename = f"{layout_label}_format_{ts}.xlsx"
    output_path = os.path.join(tempfile.gettempdir(), output_filename)

    _write_styled_excel(rows, output_path)

    return output_path, extract_dir

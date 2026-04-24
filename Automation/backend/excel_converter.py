"""
Shared Excel Converter Module  (FIXED - v3)
Converts any VAPT tracker Excel into the standardised 29-column org format.

KEY CHANGE IN v3
================
Two operating modes are now used:

  MODE A — ORG FORMAT PASSTHROUGH
    Triggered when the input Excel already contains all (or most) of the
    29 standard org-format columns.  In this mode the converter copies
    every cell value EXACTLY as-is (including pre-filled asset metadata,
    severity counts, blank rows, etc.) and only injects screenshot
    filenames from the POC ZIP when one is supplied.
    Recalculating counts or restructuring rows in this mode would corrupt
    correct data that the tester already filled in.

  MODE B — FULL TRANSFORMATION
    Triggered for any external / unknown-format Excel.  Column names are
    matched via aliases, asset metadata is reconstructed, severity counts
    are recalculated, and rows are mapped to the 29-column schema.

IP POC INDEX FIX (all modes)
    The previous indexer used only the VUL folder number as the lookup key,
    causing collisions when multiple IPs each have their own
    VUL-001, VUL-002 … folders (last-write-wins, wrong screenshots).
    The new indexer is keyed by { ip_str: [(vul_id, steps), ...] } and
    screenshot lookup uses asset IP + a per-IP vuln counter, so each row
    always gets exactly the screenshots that belong to it.

POC ZIP layouts supported
    "ip"     : POC Screenshots*/<IP>/<Severity>/<VUL-ID>/stepN.png
    "web"    : POC Screenshots*/<Severity>/<VUL-ID>/stepN.png
    "certin" : <root>/#N/stepN.png
"""

import os
import re
import zipfile
import tempfile
from datetime import datetime
from collections import defaultdict

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# ── Output column schema (29 columns, fixed order) ────────────────────────────
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

# Minimum number of OUTPUT_COLUMNS that must be present to trigger passthrough mode
ORG_FORMAT_MIN_MATCH = 22

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
    "Asset/Hostname": 20, "Instant purpose": 18, "VAPT Status": 14,
    "Critical": 10, "High": 10, "Medium": 10, "Low": 10,
    "Informational": 14, "Total": 10, "Tester_Name": 18, "Project": 20,
    "Sr no.": 10, "Observation": 14, "Severity": 14, "Status": 12,
    "New or Repeat Observation": 22, "CVE/CWE": 18, "CVSS": 12,
    "CVSS Version Ref": 16,
    "Affected Asset i.e. IP/URL/Application etc.": 35,
    "Observation/Vulnerability Title": 35,
    "Detailed Observation/ Vulnerable Point": 40,
    "Recommendation": 40, "Reference": 30,
    "Evidence / Proof of Concept": 50,
    "Screenshot1": 20, "Screenshot2": 20, "Screenshot3": 20, "Screenshot4": 20,
}

# ── Column aliases (Mode B / transformation only) ─────────────────────────────
_COL_ALIASES = {
    # AMCS / government tracker (e.g. "AMCS Route Planner" sheet) uses many non-standard headers; list them explicitly.
    "asset_hostname":       [
        "Asset/Hostname", "Asset Hostname", "Hostname", "Asset", "IP Address", "IP", "Host", "Target", "Server",
        "Audited infrastructure details",
    ],
    "instant_purpose":      ["Instant purpose", "Instant Purpose", "Purpose", "Audit Purpose", "Assessment Purpose", "Type", "Type of Audit"],
    "vapt_status":          ["VAPT Status", "Assessment Status", "Audit Status"],
    "vuln_status":          ["Status", "Vuln Status", "Finding Status", "Vulnerability Status"],
    "tester_name":          ["Tester_Name", "Tester Name", "Tester", "Reported By", "Auditor", "Assessor", "Analyst"],
    "project":              ["Project", "Project Name", "Engagement", "Client Project", "Assessment Name", "Name of Auditee organization"],
    "sr_no":                [
        "Sr no.", "Sr no", "Sr No", "Sr.No", "Sr. No.", "S.No", "S.NO.", "S.No.",
        "Serial", "S. No", "SNo", "Sr", "VUL ID", "Vuln ID", "ID",
    ],
    "observation":          ["Observation", "Name of Vulnerability / Security Issues as per CVE/CWE/Mitre DB Only"],
    "severity":             ["Severity", "Severity of Issue", "Risk", "Risk Level", "Risk(Severity)", "Criticality", "Priority"],
    "new_repeat":           [
        "New or Repeat Observation", "New or Repeat", "New/Repeat", "Observation Type", "Type of Finding",
        "Type of Audit Report(First Audit Report/Follow up Report)", "Type of Audit Report(First Audit Report/follow up Report)",
    ],
    "cve_cwe":              ["CVE/CWE", "CVE CWE", "CVE", "CWE", "CVE/CWE ID", "CVE/CWE Reference Number"],
    "cvss":                 ["CVSS", "CVSS Score", "Score", "CVSS Base Score"],
    "cvss_version":         ["CVSS Version Ref", "CVSS Version", "CVSS Ref", "CVSS Ver", "CVSS Vector", "CVSS String"],
    "affected_asset":       ["Affected Asset i.e. IP/URL/Application etc.", "Affected Asset", "IP/URL/App", "Target IP", "Vulnerable URL", "URL", "Endpoint", "Vulnerable Parameter"],
    "vuln_title":           [
        "Observation/Vulnerability Title", "Observation/ Vulnerability Title", "Vulnerability Title", "Vuln Title", "Finding Title", "Title", "Vulnerability Name", "Finding",
        "Name of Vulnerability / Security Issues as per CVE/CWE/Mitre DB Only",
    ],
    "detailed_observation": ["Detailed Observation/ Vulnerable Point", "Detailed Observation/Vulnerable Point", "Detailed Observation", "Description", "Vulnerability Description", "Vulnerability Details", "Details"],
    "recommendation":       ["Recommendation", "Remediation", "Fix", "Solution", "Recommended Fix", "Mitigation"],
    "reference":            ["Reference", "References", "Reference Link", "Links"],
    "evidence":             ["Evidence / Proof of Concept", "Evidence / Proof Of Concept", "Evidence/Proof of Concept", "Evidence", "Proof of Concept", "POC", "Steps", "Steps to Reproduce"],
}

# ── Utility ────────────────────────────────────────────────────────────────────

def _clean(val):
    if val is None:
        return ""
    s = str(val).strip()
    return "" if s.lower() == "nan" else s


def _find_col(df, aliases):
    lower_cols = {c.lower().strip(): c for c in df.columns}
    for alias in aliases:
        if alias.lower().strip() in lower_cols:
            return lower_cols[alias.lower().strip()]
    return None


def _normalise_severity(raw):
    s = _clean(raw).lower()
    if "critical"          in s: return "Critical"
    if "high"              in s: return "High"
    if "medium" in s or "moderate" in s: return "Medium"
    if "low"               in s: return "Low"
    if "info"              in s: return "Informational"
    return _clean(raw)   # preserve unknown values (e.g. "No vulnerability")


def _normalise_new_repeat(raw):
    s = _clean(raw).lower()
    return "Repeat" if ("repeat" in s or "follow" in s or "existing" in s) else "New"


def _normalise_serial_cell(raw) -> str:
    """Serial numbers from Excel often arrive as 1.0; normalise to '1' for org output."""
    s = _clean(raw)
    if not s:
        return ""
    s2 = s.replace(",", "").strip()
    try:
        f = float(s2)
        if abs(f - int(f)) < 1e-9:
            return str(int(f))
    except (ValueError, TypeError, OverflowError):
        pass
    return s


def _is_certin_obs(sr):
    return bool(re.fullmatch(r"#\d+", str(sr).strip()))


# ── Excel reader ───────────────────────────────────────────────────────────────

def _read_any_excel(path):
    def _score(df):
        key_groups = ["sr_no", "observation", "severity", "vuln_title", "detailed_observation"]
        return sum(1 for g in key_groups if _find_col(df, _COL_ALIASES[g]))

    best_df, best_score = None, -1

    # Try all sheets (some trackers use multiple tabs) and two common header rows.
    sheet_names = [0]
    try:
        xls = pd.ExcelFile(path)
        sheet_names = list(xls.sheet_names) or [0]
    except Exception:
        sheet_names = [0]

    for sheet in sheet_names:
        for hdr in (0, 1, 2):
            try:
                d = pd.read_excel(path, header=hdr, dtype=str, sheet_name=sheet)
                d.columns = [str(c).strip() for c in d.columns]
                score = _score(d)
                if score > best_score:
                    best_df, best_score = d, score
            except Exception:
                pass

    if best_df is None:
        raise ValueError("Could not read Excel file.")
    return best_df.dropna(how="all").reset_index(drop=True)


# ── Org format detection ───────────────────────────────────────────────────────

def _is_org_format(df):
    """
    Return True when enough OUTPUT_COLUMNS are present by exact name,
    meaning the input is already in the org format.
    """
    exact_cols = {str(c).strip() for c in df.columns}
    matched = sum(1 for col in OUTPUT_COLUMNS if col in exact_cols)
    result = matched >= ORG_FORMAT_MIN_MATCH
    print(f"[INFO] Org format check: {matched}/{len(OUTPUT_COLUMNS)} columns matched -> {'PASSTHROUGH' if result else 'TRANSFORM'}")
    return result


# ── POC ZIP helpers ────────────────────────────────────────────────────────────

def _extract_zip(zip_path):
    tmp = tempfile.mkdtemp()
    with zipfile.ZipFile(zip_path, "r") as zf:
        zf.extractall(tmp)
    return tmp


def _step_num(filename):
    m = re.search(r"step[\s_\-]?(\d+)", filename.lower())
    return int(m.group(1)) if m else None


def _find_poc_root(root):
    for entry in os.listdir(root):
        if re.match(r"poc.?screenshots", entry.lower()):
            candidate = os.path.join(root, entry)
            if os.path.isdir(candidate):
                return candidate
    return root


def _screenshots_from_steps(steps_dict):
    """Convert {step_num: abs_path} → list of 4 basenames."""
    ss = [os.path.basename(steps_dict[sn]) for sn in sorted(steps_dict.keys())[:4]]
    while len(ss) < 4:
        ss.append("")
    return ss


# ── IP POC Indexer v3 ─────────────────────────────────────────────────────────
# Keyed by IP so that multiple IPs with same VUL folder name never collide.

def _index_poc_ip(root):
    """
    Returns { ip_str: [(vul_id_str, {step_num: abs_path}), ...] }
    List is sorted ascending by the numeric part of the VUL-ID.

    Lookup is then done by (ip, nth_vuln_for_this_ip) not by global VUL number,
    so the collision where 35.200.155.59/VUL-001 and 172.19.2.21/VUL-001 both
    mapped to key=1 is completely eliminated.
    """
    raw = defaultdict(dict)   # ip -> { vul_id_str -> {step_num: path} }
    poc_root = _find_poc_root(root)

    for ip_folder in os.listdir(poc_root):
        ip_path = os.path.join(poc_root, ip_folder)
        if not os.path.isdir(ip_path):
            continue
        ip_key = ip_folder.strip()

        for sev_folder in os.listdir(ip_path):
            sev_path = os.path.join(ip_path, sev_folder)
            if not os.path.isdir(sev_path):
                continue
            for vul_folder in os.listdir(sev_path):
                vul_path = os.path.join(sev_path, vul_folder)
                if not os.path.isdir(vul_path):
                    continue
                vul_id = vul_folder.strip()
                if vul_id not in raw[ip_key]:
                    raw[ip_key][vul_id] = {}
                for fname in sorted(os.listdir(vul_path)):
                    if fname.lower().endswith((".png", ".jpg", ".jpeg")):
                        sn = _step_num(fname)
                        if sn is not None:
                            raw[ip_key][vul_id][sn] = os.path.join(vul_path, fname)

    def _vul_sort_key(vid):
        m = re.search(r"(\d+)", vid)
        return int(m.group(1)) if m else 0

    return {
        ip: sorted(vul_dict.items(), key=lambda kv: _vul_sort_key(kv[0]))
        for ip, vul_dict in raw.items()
    }


def _index_poc_web(root):
    """Returns { vul_number_int: {step_num: path} }"""
    index = {}
    poc_root = _find_poc_root(root)
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
    """Returns { obs_number_int: {step_num: path} }"""
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


# ── Screenshot lookup ─────────────────────────────────────────────────────────

def _get_screenshots_ip(ip, poc_index, ip_vuln_counter):
    """
    Fetch screenshots for the Nth vuln seen for this IP.
    poc_index : { ip_str: [(vul_id, {step: path}), ...] }
    ip_vuln_counter : { ip_str: int } — caller increments AFTER this call.
    """
    entries = poc_index.get(ip, [])
    n = ip_vuln_counter.get(ip, 0)
    if n < len(entries):
        _, steps_dict = entries[n]
        return _screenshots_from_steps(steps_dict)
    return ["", "", "", ""]


# ══════════════════════════════════════════════════════════════════════════════
# MODE A — ORG FORMAT PASSTHROUGH
# ══════════════════════════════════════════════════════════════════════════════

def _passthrough_org_format(df, poc_index, poc_layout):
    """
    Copy every cell value exactly as-is.
    Only injects screenshot filenames from the POC ZIP if supplied.
    Preserves: asset grouping, blank rows, pre-filled severity counts,
               tester/project fields, evidence text — everything.
    """
    rows_out = []
    ip_vuln_counter = defaultdict(int)
    last_asset = ""

    for _, row in df.iterrows():
        out = {col: _clean(row.get(col, "")) for col in OUTPUT_COLUMNS}

        # Forward-fill asset IP for screenshot lookup (doesn't change cell value)
        if out["Asset/Hostname"]:
            last_asset = out["Asset/Hostname"]
        current_ip = last_asset

        # Skip rows with no identifier at all
        if not out["Sr no."] and not out["Observation"]:
            continue

        # Inject screenshots from ZIP
        if poc_index:
            if poc_layout == "ip" and current_ip:
                screenshots = _get_screenshots_ip(current_ip, poc_index, ip_vuln_counter)
                ip_vuln_counter[current_ip] += 1
            elif poc_layout == "web":
                m = re.search(r"(\d+)", out["Sr no."])
                vul_num = int(m.group(1)) if m else None
                steps = poc_index.get(vul_num, {}) if vul_num else {}
                screenshots = _screenshots_from_steps(steps)
            elif poc_layout == "certin":
                m = re.search(r"(\d+)", out["Observation"])
                obs_num = int(m.group(1)) if m else None
                steps = poc_index.get(obs_num, {}) if obs_num else {}
                screenshots = _screenshots_from_steps(steps)
            else:
                screenshots = ["", "", "", ""]

            # Only overwrite individual screenshot cells that the ZIP provides.
            # Never blank a cell that already has a value in the input Excel.
            for ss_col, ss_val in zip(
                ["Screenshot1", "Screenshot2", "Screenshot3", "Screenshot4"],
                screenshots
            ):
                if ss_val:                           # ZIP gave a filename → use it
                    out[ss_col] = ss_val
                # else: leave the existing input value intact (don't blank it)

        rows_out.append(out)

    return rows_out


# ══════════════════════════════════════════════════════════════════════════════
# MODE B — FULL TRANSFORMATION (external / unknown format)
# ══════════════════════════════════════════════════════════════════════════════

def _build_poc_info_text(raw_evidence, steps_dict):
    sorted_steps = sorted(steps_dict.keys()) if steps_dict else []
    raw_lines = [l.strip() for l in _clean(raw_evidence).splitlines() if l.strip()]
    evidence_lines = []

    for i, sn in enumerate(sorted_steps, 1):
        if i <= len(raw_lines):
            line_text = raw_lines[i - 1]
            if not re.match(r"^step\s*\d+", line_text, re.IGNORECASE):
                line_text = f"Step {sn}: {line_text}"
        else:
            line_text = f"Step {sn}: (screenshot)"
        evidence_lines.append(line_text)

    if not sorted_steps and raw_lines:
        for i, line in enumerate(raw_lines, 1):
            evidence_lines.append(
                line if re.match(r"^step\s*\d+", line, re.IGNORECASE) else f"Step {i}: {line}"
            )

    evidence_text = "\n".join(evidence_lines) if evidence_lines else _clean(raw_evidence)
    screenshots = _screenshots_from_steps({sn: steps_dict[sn] for sn in sorted_steps[:4]})
    return evidence_text, screenshots


def _severity_counts_from_rows(df, asset_col, severity_col):
    counts = {}
    last_asset = ""
    for _, row in df.iterrows():
        raw_asset = _clean(row[asset_col]) if asset_col else ""
        if raw_asset:
            last_asset = raw_asset
        asset = last_asset
        if not asset:
            continue
        sev = _normalise_severity(row[severity_col]) if severity_col else "Medium"
        if asset not in counts:
            counts[asset] = {s: 0 for s in SEVERITY_ORDER}
        if sev in SEVERITY_ORDER:
            counts[asset][sev] += 1
    for a in counts:
        counts[a]["Total"] = sum(counts[a][s] for s in SEVERITY_ORDER)
    return counts


def _build_display_asset_keys(df, c: dict) -> list[str]:
    """
    One stable hostname/IP/org key per data row. Many trackers (e.g. AMCS) leave
    'Audited infrastructure details' blank but fill org name; without this, asset is empty,
    per-asset first-row logic breaks, and the output looks 'missing' in Excel.
    """
    hcol, ocol = c.get("asset_hostname"), c.get("project")
    last_h, last_o = "", ""
    keys: list[str] = []
    for _, row in df.iterrows():
        h = _clean(row[hcol]) if hcol else ""
        o = _clean(row[ocol]) if ocol else ""
        if h:
            last_h = h
        if o:
            last_o = o
        eff_h = h or last_h
        eff_o = o or last_o
        key = eff_h or eff_o or "Unknown"
        keys.append(key)
    return keys


def _severity_counts_by_display_keys(df, c: dict, display_keys: list[str]) -> dict:
    sevcol = c.get("severity")
    counts: dict = {}
    for i, (_, row) in enumerate(df.iterrows()):
        a = display_keys[i]
        sev = _normalise_severity(row[sevcol]) if sevcol else "Medium"
        if a not in counts:
            counts[a] = {s: 0 for s in SEVERITY_ORDER}
        if sev in SEVERITY_ORDER:
            counts[a][sev] += 1
    for a in counts:
        counts[a]["Total"] = sum(counts[a][s] for s in SEVERITY_ORDER)
    return counts


def _map_to_29_columns(df, poc_index, poc_layout):
    """Full transformation for external / unknown-format Excels."""
    c = {k: _find_col(df, v) for k, v in _COL_ALIASES.items()}
    display_keys = _build_display_asset_keys(df, c)
    sev_counts = _severity_counts_by_display_keys(df, c, display_keys)

    rows_out = []
    written_assets = set()
    ip_vuln_counter = defaultdict(int)

    for idx, (_, row) in enumerate(df.iterrows()):
        sr_raw  = _normalise_serial_cell(_clean(row[c["sr_no"]]) if c["sr_no"] else "")
        obs_raw = _clean(row[c["observation"]]) if c["observation"] else ""
        _junk = {"sr no.", "sr no", "s.no", "serial", "observation", "nan", ""}
        sr_l = sr_raw.lower()
        sr_raw = "" if sr_l in _junk or sr_l == "s.no." else sr_raw
        obs_raw = "" if obs_raw.lower() in _junk else obs_raw

        identifier = sr_raw or obs_raw
        if not identifier:
            continue

        if obs_raw and _is_certin_obs(obs_raw) and not sr_raw:
            out_sr_no, out_observation = "", obs_raw
        elif sr_raw and _is_certin_obs(sr_raw):
            out_sr_no, out_observation = "", sr_raw
        elif sr_raw:
            out_sr_no, out_observation = _normalise_serial_cell(sr_raw), ""
        else:
            out_sr_no, out_observation = "", obs_raw

        asset = display_keys[idx]
        sev   = _normalise_severity(row[c["severity"]]) if c["severity"] else "Medium"

        # Screenshot lookup
        if poc_layout == "ip":
            steps_dict = {}
            if asset:
                entries = poc_index.get(asset, [])
                n = ip_vuln_counter[asset]
                if n < len(entries):
                    _, steps_dict = entries[n]
                ip_vuln_counter[asset] += 1
            raw_evidence = _clean(row[c["evidence"]]) if c["evidence"] else ""
            evidence_text, screenshots = _build_poc_info_text(raw_evidence, steps_dict)
        elif poc_layout == "web":
            m = re.search(r"(\d+)", identifier)
            vul_num = int(m.group(1)) if m else None
            steps_dict = poc_index.get(vul_num, {}) if vul_num else {}
            raw_evidence = _clean(row[c["evidence"]]) if c["evidence"] else ""
            evidence_text, screenshots = _build_poc_info_text(raw_evidence, steps_dict)
        else:  # certin
            m = re.search(r"(\d+)", identifier)
            obs_num = int(m.group(1)) if m else None
            steps_dict = poc_index.get(obs_num, {}) if obs_num else {}
            raw_evidence = _clean(row[c["evidence"]]) if c["evidence"] else ""
            evidence_text, screenshots = _build_poc_info_text(raw_evidence, steps_dict)

        sc = sev_counts.get(asset, {s: 0 for s in SEVERITY_ORDER})
        first_for_asset = asset not in written_assets
        if first_for_asset:
            written_assets.add(asset)

        vuln_status = _clean(row[c["vuln_status"]]) if c["vuln_status"] else "Open"
        vapt_status = _clean(row[c["vapt_status"]]) if c["vapt_status"] else "Completed"
        vuln_title  = _clean(row[c["vuln_title"]]) if c["vuln_title"] else ""
        # Many trackers (AMCS) put the finding title in one column; org format also uses "Observation".
        observation_col = out_observation or vuln_title
        proj_val = _clean(row[c["project"]]) if c["project"] else ""
        inst_val = _clean(row[c["instant_purpose"]]) if c["instant_purpose"] else "Web Application"

        out = {
            "Asset/Hostname":   asset,
            "Instant purpose":  inst_val,
            "VAPT Status":      vapt_status,
            "Critical":         sc.get("Critical", 0) if first_for_asset else "",
            "High":             sc.get("High", 0)     if first_for_asset else "",
            "Medium":           sc.get("Medium", 0)   if first_for_asset else "",
            "Low":              sc.get("Low", 0)      if first_for_asset else "",
            "Informational":    sc.get("Informational", 0) if first_for_asset else "",
            "Total":            sc.get("Total", 0)    if first_for_asset else "",
            "Tester_Name":      _clean(row[c["tester_name"]]) if c["tester_name"] else "",
            "Project":          proj_val,
            "Sr no.":           out_sr_no,
            "Observation":      observation_col,
            "Severity":         sev,
            "Status":           vuln_status or "Open",
            "New or Repeat Observation": _normalise_new_repeat(row[c["new_repeat"]] if c["new_repeat"] else ""),
            "CVE/CWE":          _clean(row[c["cve_cwe"]]) if c["cve_cwe"] else "",
            "CVSS":             _clean(row[c["cvss"]]) if c["cvss"] else "",
            "CVSS Version Ref": _clean(row[c["cvss_version"]]) if c["cvss_version"] else "",
            "Affected Asset i.e. IP/URL/Application etc.": _clean(row[c["affected_asset"]]) if c["affected_asset"] else asset,
            "Observation/Vulnerability Title": vuln_title,
            "Detailed Observation/ Vulnerable Point": _clean(row[c["detailed_observation"]]) if c["detailed_observation"] else "",
            "Recommendation":   _clean(row[c["recommendation"]]) if c["recommendation"] else "",
            "Reference":        _clean(row[c["reference"]]) if c["reference"] else "",
            "Evidence / Proof of Concept": evidence_text,
            "Screenshot1": screenshots[0], "Screenshot2": screenshots[1],
            "Screenshot3": screenshots[2], "Screenshot4": screenshots[3],
        }
        rows_out.append(out)

    return rows_out


# ── Write styled Excel ─────────────────────────────────────────────────────────

def _write_styled_excel(rows, output_path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"

    header_font  = Font(name="Calibri", bold=True, color="FFFFFF", size=11)
    header_fill  = PatternFill("solid", fgColor=HEADER_FILL_COLOR)
    header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    thin         = Side(border_style="thin", color="000000")
    border       = Border(left=thin, right=thin, top=thin, bottom=thin)
    center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    wrap_align   = Alignment(horizontal="left", vertical="top", wrap_text=True)

    for col_idx, col_name in enumerate(OUTPUT_COLUMNS, 1):
        cell = ws.cell(row=1, column=col_idx, value=col_name)
        cell.font = header_font; cell.fill = header_fill
        cell.alignment = header_align; cell.border = border
        ws.column_dimensions[get_column_letter(col_idx)].width = COLUMN_WIDTHS.get(col_name, 18)
    ws.row_dimensions[1].height = 30

    count_cols  = {"Critical", "High", "Medium", "Low", "Informational", "Total"}
    center_cols = {"Sr no.", "Observation", "Severity", "Status",
                   "New or Repeat Observation", "CVSS", "CVSS Version Ref", "CVE/CWE"}

    for row_idx, row_data in enumerate(rows, 2):
        sev = row_data.get("Severity", "")
        fill_color = SEVERITY_COLORS.get(sev)
        row_fill = PatternFill("solid", fgColor=fill_color) if fill_color else None

        for col_idx, col_name in enumerate(OUTPUT_COLUMNS, 1):
            value = row_data.get(col_name, "")
            cell  = ws.cell(row=row_idx, column=col_idx, value=value)
            cell.border = border
            cell.alignment = center_align if (col_name in count_cols or col_name in center_cols) else wrap_align
            if row_fill:
                cell.fill = row_fill

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(OUTPUT_COLUMNS))}1"
    wb.save(output_path)


# ── Main entry point ───────────────────────────────────────────────────────────

def convert_any_excel_to_format(input_path, poc_zip_path=None, poc_layout="ip"):
    """
    Convert any VAPT Excel to the 29-column standard org format.

    Parameters
    ----------
    input_path   : str  - Path to input Excel (.xlsx / .xls)
    poc_zip_path : str  - Path to POC ZIP file (optional)
    poc_layout   : str  - "ip" | "web" | "certin"

    Returns
    -------
    (output_path, extract_dir)
    """
    extract_dir = None
    poc_index   = {}

    df = _read_any_excel(input_path)

    if poc_zip_path and os.path.isfile(poc_zip_path):
        try:
            extract_dir = _extract_zip(poc_zip_path)
            if poc_layout == "ip":
                poc_index = _index_poc_ip(extract_dir)
            elif poc_layout == "web":
                poc_index = _index_poc_web(extract_dir)
            elif poc_layout == "certin":
                poc_index = _index_poc_certin(extract_dir)
        except Exception as e:
            print(f"[WARN] POC ZIP indexing failed: {e}")
            poc_index = {}

    if _is_org_format(df):
        print("[INFO] Org format detected - using PASSTHROUGH mode")
        rows = _passthrough_org_format(df, poc_index, poc_layout)
    else:
        print("[INFO] External format detected - using TRANSFORMATION mode")
        rows = _map_to_29_columns(df, poc_index, poc_layout)

    if not rows:
        raise ValueError(
            "No vulnerability rows found. "
            "Ensure the file has Sr no. / Observation and Severity columns."
        )

    ts    = datetime.now().strftime("%Y%m%d_%H%M%S")
    label = {"ip": "IP", "web": "Web", "certin": "CertIN"}.get(poc_layout, poc_layout.upper())
    output_path = os.path.join(tempfile.gettempdir(), f"{label}_format_{ts}.xlsx")
    _write_styled_excel(rows, output_path)
    return output_path, extract_dir
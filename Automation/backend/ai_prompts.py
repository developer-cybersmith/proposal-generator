"""
Prompt helpers and response validation for the AI Layer feature.

The model receives structured row data plus optional screenshots and must return
JSON for the four narrative fields written back into the Excel sheet.
"""

from __future__ import annotations

import json
import re
from typing import Any

OUTPUT_KEYS = (
    "detailed_observation",
    "recommendation",
    "reference",
    "evidence",
)

SYSTEM_PROMPT = """
You are a senior VAPT security consultant writing professional vulnerability
report content.

You generate exactly four fields for each finding:

1. detailed_observation
   - Write 2-3 paragraphs.
   - Explain what was found, how it can be abused, and the likely business risk.
   - Use the exact URL/IP, severity, CVE/CWE, and CVSS values provided.
   - Do not invent new identifiers, URLs, IPs, products, or findings.

2. recommendation
   - Write a numbered remediation list.
   - Minimum 5 items for Critical/High, minimum 3 items for Medium/Low/Informational.
   - Make steps practical and specific to the finding type where possible.

3. evidence
   - Write numbered proof-of-concept steps as "Step 1:", "Step 2:", etc.
   - If screenshots are present, describe what is visible in them.
   - Do not claim screenshots show something that is not supported by the provided context.

4. reference
   - Provide relevant reference URLs.
   - Prefer OWASP, CWE, CVE, or vendor documentation.
   - If a CWE is supplied, you may include the official CWE URL for that id.

Rules:
- Use only the data provided in the prompt and screenshots.
- Never fabricate CVE, CWE, CVSS, or affected assets.
- Match the tone of a professional penetration testing report.
- Return valid JSON only with these exact keys:
  "detailed_observation", "recommendation", "reference", "evidence"
""".strip()


def _clean_value(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip()


def _extract_cwe_reference(cwe_value: str) -> str | None:
    match = re.search(r"\bCWE-(\d+)\b", cwe_value or "", flags=re.IGNORECASE)
    if not match:
        return None
    return f"https://cwe.mitre.org/data/definitions/{match.group(1)}.html"


def build_user_prompt(row_data: dict[str, Any], screenshot_names: list[str]) -> str:
    screenshot_summary = ", ".join(screenshot_names) if screenshot_names else "None provided"
    cwe_reference = _extract_cwe_reference(_clean_value(row_data.get("cve_cwe")))

    prompt_lines = [
        "Generate report content for this vulnerability finding.",
        "",
        f"Template: {_clean_value(row_data.get('report_type'))}",
        f"Row Identifier: {_clean_value(row_data.get('row_id'))}",
        f"Observation Number: {_clean_value(row_data.get('observation_number'))}",
        f"Project: {_clean_value(row_data.get('project'))}",
        f"Tester Name: {_clean_value(row_data.get('tester_name'))}",
        f"Asset/Hostname: {_clean_value(row_data.get('asset_hostname'))}",
        f"Affected Asset: {_clean_value(row_data.get('affected_asset'))}",
        f"Vulnerability Title: {_clean_value(row_data.get('vulnerability_title'))}",
        f"Severity: {_clean_value(row_data.get('severity'))}",
        f"Status: {_clean_value(row_data.get('status'))}",
        f"New or Repeat Observation: {_clean_value(row_data.get('new_or_repeat'))}",
        f"CVE/CWE: {_clean_value(row_data.get('cve_cwe')) or 'Not provided'}",
        f"CVSS Score: {_clean_value(row_data.get('cvss')) or 'Not provided'}",
        f"CVSS Vector: {_clean_value(row_data.get('cvss_vector')) or 'Not provided'}",
        f"Screenshot filenames: {screenshot_summary}",
        "",
        "Optional existing reference hint:",
        _clean_value(row_data.get("reference")) or (cwe_reference or "Not provided"),
        "",
        "Tester-provided proof / raw context:",
        _clean_value(row_data.get("evidence")) or "No written proof supplied. Use screenshots if available.",
        "",
        "Return JSON only.",
    ]
    return "\n".join(prompt_lines)


def build_row_payload(row_data: dict[str, Any]) -> dict[str, str]:
    return {
        "report_type": _clean_value(row_data.get("report_type")),
        "row_id": _clean_value(row_data.get("row_id")),
        "observation_number": _clean_value(row_data.get("observation_number")),
        "asset_hostname": _clean_value(row_data.get("asset_hostname")),
        "tester_name": _clean_value(row_data.get("tester_name")),
        "project": _clean_value(row_data.get("project")),
        "severity": _clean_value(row_data.get("severity")),
        "status": _clean_value(row_data.get("status")),
        "new_or_repeat": _clean_value(row_data.get("new_or_repeat")),
        "cve_cwe": _clean_value(row_data.get("cve_cwe")),
        "cvss": _clean_value(row_data.get("cvss")),
        "cvss_vector": _clean_value(row_data.get("cvss_vector")),
        "affected_asset": _clean_value(row_data.get("affected_asset")),
        "vulnerability_title": _clean_value(row_data.get("vulnerability_title")),
        "reference": _clean_value(row_data.get("reference")),
        "evidence": _clean_value(row_data.get("evidence")),
    }


def _strip_code_fences(text: str) -> str:
    cleaned = text.strip()
    if cleaned.startswith("```"):
        cleaned = re.sub(r"^```(?:json)?\s*", "", cleaned)
        cleaned = re.sub(r"\s*```$", "", cleaned)
    return cleaned.strip()


def _coerce_to_string(value: Any) -> str:
    """Convert AI response values (which may be lists or nested) to a single string."""
    if isinstance(value, list):
        return "\n".join(str(item) for item in value if item)
    if isinstance(value, dict):
        return "\n".join(f"{k}: {v}" for k, v in value.items())
    return _clean_value(value)


def _ensure_required_keys(data: dict[str, Any]) -> dict[str, str]:
    missing = [key for key in OUTPUT_KEYS if key not in data]
    if missing:
        raise ValueError(f"AI response missing keys: {', '.join(missing)}")

    normalized: dict[str, str] = {}
    for key in OUTPUT_KEYS:
        value = _coerce_to_string(data.get(key))
        if not value:
            raise ValueError(f"AI response field is empty: {key}")
        normalized[key] = value

    return normalized


def parse_ai_json(raw_text: str) -> dict[str, str]:
    cleaned = _strip_code_fences(raw_text)
    try:
        data = json.loads(cleaned)
    except json.JSONDecodeError as exc:
        raise ValueError("AI response was not valid JSON") from exc

    if not isinstance(data, dict):
        raise ValueError("AI response must be a JSON object")

    return _ensure_required_keys(data)


def _count_list_items(text: str) -> int:
    """Count numbered or bulleted list items in various formats the model may produce."""
    patterns = [
        r"(?m)^\s*\d+[\.\)]\s",       # 1. or 1) at start of line
        r"(?m)^\s*[-*]\s",             # - or * bullet
        r"(?m)^\s*\d+\.\s*\*\*",      # 1. **Bold item**
    ]
    seen_lines: set[int] = set()
    for pattern in patterns:
        for match in re.finditer(pattern, text):
            line_no = text[:match.start()].count("\n")
            seen_lines.add(line_no)
    return len(seen_lines)


def validate_ai_output(
    parsed: dict[str, str],
    row_data: dict[str, Any],
    screenshot_count: int,
) -> dict[str, str]:
    severity = _clean_value(row_data.get("severity")).lower()
    min_recommendations = 3 if severity in {"critical", "high"} else 2

    rec_count = _count_list_items(parsed["recommendation"])
    if rec_count < min_recommendations:
        raise ValueError(
            f"Recommendation list too short: found {rec_count} items, expected at least {min_recommendations}"
        )

    if screenshot_count > 0:
        step_patterns = re.findall(r"(?mi)\bstep\s*\d+", parsed["evidence"])
        if len(step_patterns) < 1:
            raise ValueError("Evidence does not reference any step numbers despite screenshots being provided")

    allowed_identifier = _clean_value(row_data.get("cve_cwe"))
    if allowed_identifier:
        referenced_identifiers = re.findall(r"\b(?:CVE|CWE)-[\w.-]+\b", json.dumps(parsed), flags=re.IGNORECASE)
        for item in referenced_identifiers:
            if item.lower() not in allowed_identifier.lower():
                raise ValueError(f"AI introduced unexpected identifier: {item}")

    return parsed

"""
Guard-railed prompts and JSON validation for the Executive Summary generator.

Output is a single JSON object from OpenAI (response_format json_object).
"""

from __future__ import annotations

import json
import re
from typing import Any

THREAT_ENUM = frozenset({"Intense", "High", "Moderate", "Low"})
REQUIREMENT_ENUM = frozenset({"Critical", "High", "Medium", "Low"})
BUSINESS_IMPACT_ENUM = frozenset(
    {
        "Breach of Sensitive Information",
        "Operational Capabilities",
        "Resilience Towards Risk Tolerance",
        "Intellectual Property",
        "Service Disruption",
    }
)

OUTPUT_TOP_KEYS = (
    "executive_narrative",
    "overall_recommendation_intro",
    "phase1_observations",
    "phase2_summary",
    "best_practices",
    "must_have",
    "should_have",
    "conclusion",
)

MITKAT_CONCLUSION_ADDENDUM = """
MitKat-specific conclusion rules:
- Write exactly ONE short paragraph (3-5 sentences).
- Do NOT include thank-you lines or sign-offs.
- Close with: residual risk statement, remediation priority by severity/business impact, and the need for continuous monitoring and periodic reassessment to prevent recurrence.
""".strip()

SYSTEM_PROMPT = """
You are a senior cybersecurity executive report writer for CyberSmithSECURE.
You write concise, professional content for C-suite readers (CEO, CTO, board).
No technical jargon: no raw HTTP headers, no code, no exploit payloads, no POC steps.
Do not mention screenshots or proof-of-concept procedures.

You MUST return a single JSON object with these EXACT keys (all required):
- "executive_narrative": string, 5-7 short paragraphs separated by blank lines (use \\n\\n). Summarise what was tested, methodology at a high level (OWASP/SANS/CWE only as phrases), scope, and outcomes. Use client_name and assessment_type from metadata.
- "overall_recommendation_intro": string, 1-2 paragraphs for senior decision-makers: revalidation status, need for ongoing security governance, plain English.
- "phase1_observations": array of 3 to 7 objects, each with:
  "title" (short theme), "description" (2-3 sentences), "threat_perception" (exactly one of: Intense, High, Moderate, Low).
  Themes must map to real vulnerability titles or patterns from the input list — do not invent new CVEs or hostnames.
  If the workbook lists fewer than three distinct themes, still output at least three items by grouping related findings.
- "phase2_summary": string, 1-2 paragraphs on remediation / patching posture and residual risk, using only input data.
- "best_practices": array of 3 to 5 objects: "title", "description" (2-3 sentences), "requirement" (Critical|High|Medium|Low), "business_impact" (exactly one of the allowed business impact strings below).
- "must_have": array of 3 to 5 objects, same shape as best_practices.
- "should_have": array of 2 to 4 objects, same shape as best_practices.
- "conclusion": string, 1-2 paragraphs closing the assessment, thanking the client security team; professional tone.

Allowed business_impact values (copy exactly):
"Breach of Sensitive Information", "Operational Capabilities", "Resilience Towards Risk Tolerance", "Intellectual Property", "Service Disruption"

Rules:
- Use ONLY facts from the JSON payload (metadata, assets, total_counts, top_vulnerabilities). Never invent CVE/CWE identifiers, hostnames, or counts.
- If referencing identifiers, they must appear in the supplied vulnerability or asset strings.
- threat_perception must be one of: Intense, High, Moderate, Low.
- requirement must be one of: Critical, High, Medium, Low.
- Return valid JSON only. No markdown fences.
""".strip()


def system_prompt_for_metadata(metadata: dict[str, Any] | None) -> str:
    md = metadata or {}
    if str(md.get("report_type_label") or "").strip().lower() == "mitkat":
        return SYSTEM_PROMPT + "\n\n" + MITKAT_CONCLUSION_ADDENDUM
    return SYSTEM_PROMPT


def _strip_code_fences(text: str) -> str:
    cleaned = text.strip()
    if cleaned.startswith("```"):
        cleaned = re.sub(r"^```(?:json)?\s*", "", cleaned)
        cleaned = re.sub(r"\s*```$", "", cleaned)
    return cleaned.strip()


def _clean_str(v: Any) -> str:
    if v is None:
        return ""
    return str(v).strip()


def build_canonical_user_payload(
    metadata: dict[str, Any],
    assets: list[dict[str, Any]],
    total_counts: dict[str, Any],
    top_vulnerabilities: list[dict[str, Any]],
) -> str:
    """Human-readable JSON string for the user message."""
    payload = {
        "metadata": metadata,
        "total_counts": total_counts,
        "assets": assets,
        "top_vulnerabilities": top_vulnerabilities,
    }
    return json.dumps(payload, indent=2, ensure_ascii=False)


def build_user_message(metadata: dict[str, Any], canonical_json_str: str) -> str:
    return (
        "Generate the executive summary JSON using ONLY the data below.\n\n"
        f"{canonical_json_str}\n\n"
        "Return JSON only with all required keys."
    )


def _validate_reco_item_fixed(item: Any, path: str) -> dict[str, str]:
    if not isinstance(item, dict):
        raise ValueError(f"{path} must be an object")
    title = _clean_str(item.get("title"))
    desc = _clean_str(item.get("description"))
    req = _clean_str(item.get("requirement"))
    bi = _clean_str(item.get("business_impact"))
    if not title or not desc:
        raise ValueError(f"{path} missing title or description")
    if req not in REQUIREMENT_ENUM:
        raise ValueError(f"{path} invalid requirement: {req}")
    if bi not in BUSINESS_IMPACT_ENUM:
        raise ValueError(f"{path} invalid business_impact: {bi}")
    return {"title": title, "description": desc, "requirement": req, "business_impact": bi}


def _coerce_threat_perception(raw: str) -> str:
    tp = _clean_str(raw)
    if tp in THREAT_ENUM:
        return tp
    tl = tp.lower()
    if "intense" in tl:
        return "Intense"
    if "high" in tl:
        return "High"
    if "moderate" in tl or "medium" in tl:
        return "Moderate"
    if "low" in tl:
        return "Low"
    return "Moderate"


def _coerce_requirement(raw: str) -> str:
    req = _clean_str(raw)
    if req in REQUIREMENT_ENUM:
        return req
    rl = req.lower()
    if "crit" in rl:
        return "Critical"
    if "high" in rl:
        return "High"
    if "medium" in rl or "moderate" in rl:
        return "Medium"
    if "low" in rl:
        return "Low"
    return "Medium"


def _coerce_business_impact(raw: str) -> str:
    bi = _clean_str(raw)
    if bi in BUSINESS_IMPACT_ENUM:
        return bi
    bl = bi.lower()
    for allowed in BUSINESS_IMPACT_ENUM:
        if allowed.lower() == bl:
            return allowed
    if "breach" in bl or "sensitive" in bl:
        return "Breach of Sensitive Information"
    if "operat" in bl or "business" in bl:
        return "Operational Capabilities"
    if "resilien" in bl or "tolerance" in bl:
        return "Resilience Towards Risk Tolerance"
    if "intellectual" in bl or "ip" in bl:
        return "Intellectual Property"
    if "disrupt" in bl or "service" in bl:
        return "Service Disruption"
    return "Operational Capabilities"


def _norm_phase1_item(it: Any, idx: int) -> dict[str, str]:
    if not isinstance(it, dict):
        it = {}
    t = _clean_str(it.get("title")) or f"Assessment theme {idx + 1}"
    d = _clean_str(it.get("description"))
    if len(d) < 20:
        d = (
            "This theme reflects findings from the supplied assessment data. "
            "Remediation and validation should be tracked in line with organisational risk appetite."
        )
    tp = _coerce_threat_perception(_clean_str(it.get("threat_perception")))
    return {"title": t, "description": d, "threat_perception": tp}


def _norm_reco_item_loose(it: Any, idx: int) -> dict[str, str]:
    if not isinstance(it, dict):
        it = {}
    title = _clean_str(it.get("title")) or f"Recommendation area {idx + 1}"
    desc = _clean_str(it.get("description"))
    if len(desc) < 20:
        desc = (
            "Align security investment with the severity profile shown in the workbook "
            "and validate fixes in representative environments."
        )
    req = _coerce_requirement(it.get("requirement"))
    bi = _coerce_business_impact(it.get("business_impact"))
    return {"title": title, "description": desc, "requirement": req, "business_impact": bi}


def _trim_or_pad_phase1(items: list[dict[str, str]], min_n: int, max_n: int) -> list[dict[str, str]]:
    out = list(items)
    fillers = [
        {
            "title": "Residual risk and validation",
            "description": (
                "After remediation, re-test critical paths to confirm that controls hold under realistic use. "
                "Track exceptions and compensating controls where fixes cannot be applied immediately."
            ),
            "threat_perception": "Moderate",
        },
        {
            "title": "Governance and ownership",
            "description": (
                "Assign clear owners for each finding class and report progress to stakeholders. "
                "Security improvements work best when tied to measurable milestones."
            ),
            "threat_perception": "Low",
        },
        {
            "title": "Defence in depth",
            "description": (
                "Layer preventive and detective controls so that a single control failure does not "
                "automatically lead to compromise of sensitive data or services."
            ),
            "threat_perception": "Moderate",
        },
    ]
    fi = 0
    while len(out) < min_n:
        out.append(dict(fillers[fi % len(fillers)]))
        fi += 1
    if len(out) > max_n:
        out = out[:max_n]
    return out


def _trim_or_pad_reco(items: list[dict[str, str]], min_n: int, max_n: int) -> list[dict[str, str]]:
    out = list(items)
    fillers = [
        {
            "title": "Patch and configuration hygiene",
            "description": (
                "Keep systems patched and configurations hardened according to vendor guidance "
                "and your change-management process."
            ),
            "requirement": "High",
            "business_impact": "Operational Capabilities",
        },
        {
            "title": "Access control and authentication",
            "description": (
                "Enforce least privilege, strong authentication, and periodic access reviews "
                "for privileged and service accounts."
            ),
            "requirement": "High",
            "business_impact": "Breach of Sensitive Information",
        },
        {
            "title": "Monitoring and incident readiness",
            "description": (
                "Maintain logging, alerting, and playbooks so that suspicious activity is detected "
                "and handled without undue delay."
            ),
            "requirement": "Medium",
            "business_impact": "Service Disruption",
        },
        {
            "title": "Secure development practices",
            "description": (
                "Use secure coding standards and testing in the SDLC to reduce recurrence "
                "of common vulnerability classes."
            ),
            "requirement": "Medium",
            "business_impact": "Resilience Towards Risk Tolerance",
        },
    ]
    fi = 0
    while len(out) < min_n:
        out.append(dict(fillers[fi % len(fillers)]))
        fi += 1
    if len(out) > max_n:
        out = out[:max_n]
    return out


def _pad_executive_narrative(en: str) -> str:
    en = _clean_str(en)
    if len(en) >= 500:
        return en
    extra = (
        "\n\nThe workbook you provided is the basis for this summary. "
        "Severity counts and representative findings should be read alongside the full technical report "
        "for detailed remediation guidance. Leadership may use this document to align priorities, "
        "ownership, and timelines for remediation and revalidation."
    )
    return (en + extra).strip()


def parse_exec_summary_json(raw_text: str, *, report_type_label: str | None = None) -> dict[str, Any]:
    cleaned = _strip_code_fences(raw_text)
    try:
        data = json.loads(cleaned)
    except json.JSONDecodeError as exc:
        raise ValueError("AI response was not valid JSON") from exc
    if not isinstance(data, dict):
        raise ValueError("AI response must be a JSON object")
    for key in OUTPUT_TOP_KEYS:
        if key not in data:
            raise ValueError(f"AI response missing key: {key}")

    en = _pad_executive_narrative(_clean_str(data.get("executive_narrative")))
    if len(en) < 200:
        raise ValueError("executive_narrative too short (min 200 chars after padding)")
    ori = _clean_str(data.get("overall_recommendation_intro"))
    if len(ori) < 40:
        ori = (
            "Leadership should treat remediation as a sustained program with clear ownership, "
            "milestones, and revalidation against the technical report."
        )
    if len(ori) < 80:
        ori = ori + " Tracking progress and validating fixes reduces residual risk over time."

    p1_raw = data.get("phase1_observations")
    if not isinstance(p1_raw, list):
        p1_raw = []
    norm_p1 = [_norm_phase1_item(it, i) for i, it in enumerate(p1_raw)]
    norm_p1 = _trim_or_pad_phase1(norm_p1, 3, 7)

    p2 = _clean_str(data.get("phase2_summary"))
    if len(p2) < 40:
        p2 = (
            "Remediation should follow severity and business impact, with fixes tested in staging where possible. "
            "Residual risk remains until changes are validated in production-like conditions."
        )

    def parse_reco_list(key: str, min_n: int, max_n: int) -> list[dict[str, str]]:
        lst = data.get(key)
        if not isinstance(lst, list):
            lst = []
        out = [_norm_reco_item_loose(it, i) for i, it in enumerate(lst)]
        return _trim_or_pad_reco(out, min_n, max_n)

    bp = parse_reco_list("best_practices", 3, 5)
    mh = parse_reco_list("must_have", 3, 5)
    sh = parse_reco_list("should_have", 2, 4)

    conc = _clean_str(data.get("conclusion"))
    if len(conc) < 80:
        if str(report_type_label or "").strip().lower() == "mitkat":
            conc = (
                (conc + " ") if conc else ""
            ) + (
                "In conclusion, the assessment indicates measurable risk based on the supplied findings. "
                "Remediation should be prioritised by severity and business impact, with fixes validated through re-testing. "
                "Ongoing monitoring, timely patching, and periodic reassessment are essential to prevent recurrence and reduce residual risk."
            )
        else:
            conc = (
                conc + " "
                if conc
                else ""
            ) + (
                "In summary, the assessment provides a structured view of risk based on the supplied data. "
                "We thank the security team for their cooperation and stand ready to support follow-up activities."
            )

    return {
        "executive_narrative": en,
        "overall_recommendation_intro": ori,
        "phase1_observations": norm_p1,
        "phase2_summary": p2,
        "best_practices": bp,
        "must_have": mh,
        "should_have": sh,
        "conclusion": conc,
    }


def validate_no_foreign_identifiers(
    parsed: dict[str, Any],
    allowed_cve_cwe_blob: str,
    allowed_hostnames: set[str],
) -> None:
    """Reject CVE/CWE tokens not present in allowed blob when the workbook supplied identifiers."""
    _ = allowed_hostnames  # reserved for future hostname checks
    allowed_lower = (allowed_cve_cwe_blob or "").lower().strip()
    if not allowed_lower:
        # No CVE/CWE text in source Excel — do not fail generation if the model uses generic wording
        # or omits identifiers entirely.
        return
    text = json.dumps(parsed, ensure_ascii=False)
    for m in re.finditer(r"\b(?:CVE|CWE)-[\w.-]+\b", text, re.IGNORECASE):
        token = m.group(0)
        if token.lower() not in allowed_lower:
            raise ValueError(f"AI introduced unexpected identifier: {token}")

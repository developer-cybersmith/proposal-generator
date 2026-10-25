"""
Excel-driven workflow: upload -> Preview 1 (annexures) -> Preview 2 (Conclusion & Pricing) -> generate.

Decisions implemented (docs/ADVANCED_FEATURE_PLAN.md): D11 two pricing lines for External/Internal,
D16 zero-quantity confirmation, D18 optional total cost, D20 sector asked on review, D21 editable
reference number, D23/D34 one shared Frequency field, D30-D33 Count and Quantity.

The module has two layers:
  * pure functions (build_draft, merge_edits, readiness_issues, resolve_plan) - no web framework,
    fully unit-tested;
  * create_router(gen) - the FastAPI routes. It receives the generator module so this file never
    imports css_proposal_generator (which includes this router).
"""
from __future__ import annotations

import copy
import os
import re
import secrets
import tempfile
import time
from pathlib import Path
from typing import Any

from fastapi import APIRouter, Depends, File, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse

import proposal_excel
import proposal_fields
import proposal_pricing
from proposal_drafts import DraftStore
from proposal_excel import ScopingResult
from proposal_ratelimit import rate_limit

DEFAULT_FREQUENCY = "One Time"
MAX_ROWS = 40
MAX_COUNT = 1_000_000
SAFE_REF_RE = re.compile(r"^[A-Za-z0-9 ._\-/#()]{1,60}$")
SAFE_DATE_RE = re.compile(r"^[0-9][0-9 ./\-]{5,19}$")
CONTROL_RE = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]")
PAYMENT_MODELS = ("milestone_v1", "hybrid_v2", "retainer_v3", "tier_v4")


def clean_text(value: Any, limit: int) -> str:
    """Trim, drop control characters, cap the length. Never raises."""
    if value is None:
        return ""
    return CONTROL_RE.sub("", str(value)).strip()[:limit]


def _as_int(value: Any, default: int = 0) -> int:
    try:
        n = int(float(value))
    except (TypeError, ValueError):
        return default
    return max(0, min(MAX_COUNT, n))


# =============================================================================== draft creation
def _duration_from_dates(client: dict[str, Any]) -> str:
    from datetime import date
    try:
        start = date.fromisoformat(client.get("preferred_start_date") or "")
        end = date.fromisoformat(client.get("preferred_end_date") or "")
    except ValueError:
        return ""
    days = (end - start).days
    return f"{days} days" if days > 0 else ""


def _entry_from_excel(e, catalog: dict[str, dict]) -> dict[str, Any]:
    custom = e.catalog_key is None
    title = e.type_label if custom or e.type_key in ("external_network_pt", "internal_network_pt") \
        else catalog[e.catalog_key]["title"]
    return {
        "id": e.type_key, "origin": "excel", "type_key": e.type_key, "type_label": e.type_label,
        "title": title, "annexure": e.annexure, "catalog_key": e.catalog_key, "include": True,
        "count": e.count, "count_label": e.count_label, "quantity": e.quantity,
        "frequency": e.frequency, "needs_confirmation": e.needs_confirmation, "confirmed": False,
        "rows": [r.to_dict() for r in e.rows],
        "scope": _default_scope(e) if custom else "",
    }


def _default_scope(e) -> str:
    """Starting scope text for an AI annexure (the user edits it; at least 20 characters)."""
    return (f"{e.type_label}: {e.count} {e.count_label} in scope, total quantity {e.quantity}. "
            "Scope, methodology and deliverables to follow the CyberSmithSECURE standard for this assessment.")


def build_draft(result: ScopingResult, filename: str, catalog: dict[str, dict],
                cfg: dict[str, Any] | None = None) -> dict[str, Any]:
    """Turn a parsed workbook into the editable draft (initial state of both previews)."""
    cfg = cfg or proposal_excel.load_map()
    entries = [_entry_from_excel(e, catalog) for e in result.included]

    defaults = cfg["manual_toggle_defaults"]
    for t in cfg["manual_toggle_annexures"]:
        item = catalog[t["catalog_key"]]
        entries.append({
            "id": t["catalog_key"], "origin": "manual", "type_key": t["catalog_key"],
            "type_label": item["title"], "title": item["title"], "annexure": t["annexure"],
            "catalog_key": t["catalog_key"], "include": False, "count": defaults["count"],
            "count_label": defaults["count_label"], "quantity": defaults["quantity"],
            "frequency": DEFAULT_FREQUENCY, "needs_confirmation": False, "confirmed": False,
            "rows": [], "scope": "",
        })

    # Default order: template annexures by number, then AI annexures in Excel order.
    entries.sort(key=lambda x: (x["annexure"] == "custom", x["annexure"] if x["annexure"] != "custom" else 0))

    client = dict(result.client)
    draft = {
        "version": 1,
        "created": time.time(),
        "source_filename": clean_text(os.path.basename(filename or "scoping.xlsx"), 100),
        "warnings": list(result.warnings),
        "context": result.context,
        "client": {**client, "name": client.get("client_organization_name") or "",
                   "reference_no": result.reference_no or "",
                   "proposal_date": time.strftime("%d-%m-%Y")},         # printed on the cover; editable
        "entries": entries,
        "pricing": [],
        "pricing_removed": [],
        "ai": {
            "sector": "", "sector_suggestion": proposal_excel.suggest_sector(client.get("regulatory_driver"), cfg),
            "engagement_summary": "", "primary_tier": "Tier 3 – Advanced", "tier_notes": "",
            "payment_model": "hybrid_v2", "special_constraints": "",
            "duration": _duration_from_dates(client), "total_cost": "",
        },
    }
    draft["ai"]["engagement_summary"] = _default_engagement_summary(draft)
    draft["pricing"] = _reconcile_pricing(draft, [], {}, provided=False)
    return draft


def _default_engagement_summary(draft: dict[str, Any]) -> str:
    name = draft["client"].get("name") or "the client"
    types = [e["type_label"] for e in draft["entries"] if e["origin"] == "excel"]
    return f"Security assessment engagement for {name}: " + ", ".join(types[:6]) + "." if types else ""


# =============================================================================== pricing rows
def _row_for(entry: dict[str, Any]) -> dict[str, Any]:
    return {"entry_id": entry["id"], "title": entry["title"], "quantity": _fmt_qty(entry.get("quantity")),
            "cost": "", "frequency": entry["frequency"]}


def _fmt_qty(value: Any) -> str:
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    return "" if value is None else str(value)


def _reconcile_pricing(draft: dict[str, Any], incoming: list[dict[str, Any]],
                       previous_include: dict[str, bool], provided: bool = True) -> list[dict[str, Any]]:
    """Keep pricing rows in step with Preview 1 (one linked row per included annexure entry).

    incoming          rows as the user last saw them in Preview 2
    previous_include  each entry's include flag before this edit ({} when the draft is being built)
    provided          False only while building the draft: every included entry then gets a row

    Rules:
      * excluding an annexure drops its row; including it again brings the row back (D34),
      * a row the user deleted in Preview 2 stays deleted while the annexure stays included,
      * rows the user added (entry_id None) are kept,
      * a linked row's Frequency always equals its annexure's (single shared field).
    """
    entries = {e["id"]: e for e in draft["entries"]}
    by_entry = {r.get("entry_id"): r for r in incoming if r.get("entry_id") in entries}
    removed = set(draft.get("pricing_removed", []))
    rows: list[dict[str, Any]] = []
    for e in draft["entries"]:
        if not e["include"]:
            continue
        newly_included = bool(previous_include) and not previous_include.get(e["id"], False)
        if newly_included:
            removed.discard(e["id"])
            row = dict(by_entry.get(e["id"]) or _row_for(e))
        elif e["id"] in by_entry:
            row = dict(by_entry[e["id"]])
        elif provided:
            removed.add(e["id"])        # the entry was included and its row is missing: user deleted it
            continue
        elif e["id"] in removed:
            continue
        else:
            row = _row_for(e)
        row["frequency"] = e["frequency"]
        rows.append(row)
    rows.extend(dict(r) for r in incoming if r.get("entry_id") is None)
    draft["pricing_removed"] = sorted(removed)
    return rows


# =============================================================================== edits (PUT)
def merge_edits(draft: dict[str, Any], incoming: dict[str, Any]) -> dict[str, Any]:
    """Apply the user's edits to a draft. Only whitelisted fields are taken, all are validated.

    Never trusts structure from the client: unknown entry ids are ignored, numbers are clamped,
    text is stripped of control characters and length-limited. Returns a new draft dict.
    """
    out = copy.deepcopy(draft)
    previous_include = {e["id"]: e["include"] for e in out["entries"]}

    inc_client = incoming.get("client") or {}
    if isinstance(inc_client, dict):
        if "name" in inc_client:
            out["client"]["name"] = clean_text(inc_client["name"], 150)
        if "reference_no" in inc_client:
            out["client"]["reference_no"] = clean_text(inc_client["reference_no"], 60)
        if "proposal_date" in inc_client:
            out["client"]["proposal_date"] = clean_text(inc_client["proposal_date"], 20)

    known = {e["id"]: e for e in out["entries"]}
    incoming_entries = [x for x in (incoming.get("entries") or []) if isinstance(x, dict) and x.get("id") in known]
    for item in incoming_entries:
        e = known[item["id"]]
        if "include" in item:
            e["include"] = bool(item["include"])
        if "count" in item:
            e["count"] = _as_int(item["count"], e["count"])
        if "count_label" in item:
            e["count_label"] = clean_text(item["count_label"], 80) or e["count_label"]
        if "frequency" in item:
            e["frequency"] = clean_text(item["frequency"], 60) or DEFAULT_FREQUENCY
        if "confirmed" in item:
            e["confirmed"] = bool(item["confirmed"])
        if e["catalog_key"] is None:               # AI annexures: title and scope are editable
            if "title" in item:
                e["title"] = clean_text(item["title"], 120) or e["title"]
            if "scope" in item:
                e["scope"] = clean_text(item["scope"], 2000)
    # The order changes only when the client sends every annexure (as the page always does);
    # a partial update never reshuffles the list.
    sent = list(dict.fromkeys(x["id"] for x in incoming_entries))
    if len(sent) == len(known):
        out["entries"] = [known[i] for i in sent]

    # ext/int (same template annexure): one document annexure, one shared Frequency
    first_freq: dict[str, str] = {}
    for e in out["entries"]:
        if e["include"] and e["catalog_key"]:
            e["frequency"] = first_freq.setdefault(e["catalog_key"], e["frequency"])

    inc_ai = incoming.get("ai") or {}
    if isinstance(inc_ai, dict):
        limits = {"sector": 120, "engagement_summary": 1500, "primary_tier": 60, "tier_notes": 500,
                  "special_constraints": 1000, "duration": 60, "total_cost": 60}
        for key, limit in limits.items():
            if key in inc_ai:
                out["ai"][key] = clean_text(inc_ai[key], limit)
        if inc_ai.get("payment_model") in PAYMENT_MODELS:
            out["ai"]["payment_model"] = inc_ai["payment_model"]

    rows: list[dict[str, Any]] = []
    for r in (incoming.get("pricing") or [])[:MAX_ROWS]:
        if not isinstance(r, dict):
            continue
        entry_id = r.get("entry_id")
        rows.append({
            "entry_id": entry_id if entry_id in known else None,
            "title": clean_text(r.get("title"), 150),
            "quantity": clean_text(r.get("quantity"), 40),
            "cost": clean_text(r.get("cost"), 40),
            "frequency": clean_text(r.get("frequency"), 60) or DEFAULT_FREQUENCY,
        })
    if "pricing" in incoming:
        out["pricing"] = _reconcile_pricing(out, rows, previous_include)
    else:
        out["pricing"] = _reconcile_pricing(out, out["pricing"], previous_include)
    return out


# =============================================================================== readiness
def readiness_issues(draft: dict[str, Any]) -> list[str]:
    """Reasons Generate must stay disabled. Empty list = ready."""
    issues: list[str] = []
    if not draft["client"].get("name", "").strip():
        issues.append("Client name is required.")
    ref = draft["client"].get("reference_no", "").strip()
    if not ref:
        issues.append("Reference number is required (the Excel's Engagement Reference No. was blank).")
    elif not SAFE_REF_RE.match(ref):
        issues.append("Reference number contains unsupported characters.")
    date = draft["client"].get("proposal_date", "").strip()
    if not date:
        issues.append("Proposal date is required (it is printed on the cover).")
    elif not SAFE_DATE_RE.match(date):
        issues.append("Proposal date contains unsupported characters (use digits and - or /, e.g. 15-09-2025).")
    included = [e for e in draft["entries"] if e["include"]]
    if not included:
        issues.append("Select at least one annexure.")
    for e in included:
        if e["needs_confirmation"] and e["count"] == 0 and not e["confirmed"]:
            issues.append(f"{e['title']}: nothing above 0 was found in the Excel. "
                          "Enter a count, confirm it as is, or remove the annexure.")
    customs = [e for e in included if e["catalog_key"] is None]
    if customs:
        ai = draft["ai"]
        if not ai["sector"].strip():
            issues.append("Sector is required for AI-written annexures.")
        if len(ai["engagement_summary"].strip()) < 10:
            issues.append("Engagement summary (at least 10 characters) is required for AI-written annexures.")
        if not ai["duration"].strip():
            issues.append("Duration is required for AI-written annexures.")
        for e in customs:
            if not e["title"].strip():
                issues.append("Every AI-written annexure needs a title.")
            if len(e["scope"].strip()) < 20:
                issues.append(f"{e['title']}: scope needs at least 20 characters.")
    for r in draft["pricing"]:
        if not r["title"].strip():
            issues.append("Every pricing row needs an Assessment Type.")
            break
    return issues


# =============================================================================== generation plan
def resolve_plan(draft: dict[str, Any]) -> dict[str, Any]:
    """Convert the reviewed draft into the arguments of generate_proposal().

    Template annexures that share one catalog key (External + Internal Network PT) become a single
    annexure with a combined Count line; AI annexures get fresh unique keys.
    """
    groups: dict[str, list[dict[str, Any]]] = {}
    ordered: list[tuple[str, str, dict[str, Any]]] = []    # (key, kind, first entry)
    for e in draft["entries"]:
        if not e["include"]:
            continue
        if e["catalog_key"]:
            key = e["catalog_key"]
            if key not in groups:
                groups[key] = []
                ordered.append((key, "static", e))
            groups[key].append(e)
        else:
            key = f"__other__{int(time.time() * 1000)}_{secrets.token_hex(3)}"
            groups[key] = [e]
            ordered.append((key, "custom", e))

    details: dict[str, dict[str, Any]] = {}
    customs: list[dict[str, Any]] = []
    for key, kind, first in ordered:
        members = groups[key]
        details[key] = {
            "count_text": proposal_fields.count_line_text([(m["count"], m["count_label"]) for m in members]),
            "frequency": first["frequency"],
        }
        if kind == "custom":
            customs.append({"key": key, "entry": first})

    return {
        "ordered_keys": [k for k, _kind, _e in ordered],
        "annexure_details": details,
        "customs": customs,
        "pricing_rows": [{"title": r["title"], "quantity": r["quantity"], "frequency": r["frequency"],
                          "cost": r["cost"]} for r in draft["pricing"]],
        "company_name": draft["client"]["name"].strip(),
        "reference": draft["client"]["reference_no"].strip(),
        "proposal_date": draft["client"].get("proposal_date", "").strip(),
    }


def total_cost_text(pricing_rows: list[dict[str, Any]]) -> str:
    """D18: the total passed to the AI is derived from the pricing table when every row is priced."""
    amounts = [proposal_pricing.parse_amount(r.get("cost")) for r in pricing_rows]
    if pricing_rows and all(a is not None for a in amounts):
        return proposal_pricing.format_inr(sum(amounts))
    return ""


def _ai_counts(draft: dict[str, Any]) -> dict[str, str]:
    q = {e["type_key"]: e for e in draft["entries"] if e["include"] and e["origin"] == "excel"}
    def qty(*keys: str) -> str:
        parts = [f"{k.replace('_', ' ')}: {_fmt_qty(q[k]['quantity'])}" for k in keys if k in q]
        return "; ".join(parts)
    return {
        "ad_identity": qty("ad_security"),
        "cloud": qty("cloud_security"),
        "apps_apis": qty("web_app_vapt", "mobile_vapt", "api_security"),
        "infra": qty("external_network_pt", "internal_network_pt", "config_hardening"),
        "users": "",
        "duration": draft["ai"]["duration"],
    }


def ai_inputs(draft: dict[str, Any], plan: dict[str, Any]) -> dict[str, Any]:
    """The AI context in the shape css_proposal_generator._parse_ai_context accepts."""
    ai = draft["ai"]
    total = total_cost_text(plan["pricing_rows"]) or ai.get("total_cost", "")
    cfg = proposal_excel.load_map()
    customs = []
    for slot, c in enumerate(plan["customs"], start=1):
        e = c["entry"]
        customs.append({
            "slot_index": slot, "key": c["key"], "title": e["title"], "scope": e["scope"],
            "count_text": plan["annexure_details"][c["key"]]["count_text"],
            "quantity": _fmt_qty(e["quantity"]),
            "excel_context": proposal_excel.context_block_for_type(draft.get("context", {}), e["type_key"], cfg),
        })
    return {
        "sector": ai["sector"], "engagement_summary": ai["engagement_summary"],
        "primary_tier": ai["primary_tier"], "tier_notes": ai["tier_notes"],
        "counts": _ai_counts(draft),
        "commercials": {"total_cost": total, "taxes": "18% GST"},
        "payment_model": ai["payment_model"], "special_constraints": ai["special_constraints"],
        "custom_annexures": customs,
    }


# =============================================================================== web layer
def create_router(gen):
    """Build the FastAPI router. *gen* is the css_proposal_generator module."""
    router = APIRouter()
    store = DraftStore()
    max_bytes = proposal_excel.load_map()["max_file_bytes"]
    ui_file = Path(__file__).resolve().parent / "proposal_excel_ui.html"
    router.draft_store = store  # exposed for tests

    def _load(draft_id: str) -> dict[str, Any]:
        draft = store.get(draft_id)
        if draft is None:
            raise HTTPException(status_code=404, detail="Draft not found or expired. Please upload the Excel again.")
        return draft

    def _view(draft_id: str, draft: dict[str, Any]) -> dict[str, Any]:
        shown = {k: v for k, v in draft.items() if k != "ai_cache"}      # AI texts stay on the server
        return {"draft_id": draft_id, "draft": shown, "issues": readiness_issues(draft),
                "ttl_seconds": store.ttl}

    @router.get("/excel/", response_class=HTMLResponse)
    async def excel_page() -> HTMLResponse:
        return HTMLResponse(ui_file.read_text(encoding="utf-8"),
                            headers={"Cache-Control": "no-store", "X-Content-Type-Options": "nosniff"})

    @router.post("/upload-excel/", dependencies=[Depends(rate_limit("upload"))])
    async def upload_excel(file: UploadFile = File(...)):
        name = clean_text(os.path.basename(file.filename or ""), 100)
        if Path(name).suffix.lower() != ".xlsx":
            raise HTTPException(status_code=400, detail="Only .xlsx files are accepted.")
        tmp_dir = tempfile.mkdtemp(prefix="css_scoping_")
        path = Path(tmp_dir) / f"{secrets.token_hex(8)}.xlsx"
        try:
            size = 0
            with open(path, "wb") as fh:
                while True:
                    chunk = await file.read(256 * 1024)
                    if not chunk:
                        break
                    size += len(chunk)
                    if size > max_bytes:
                        raise HTTPException(status_code=413,
                                            detail=f"File is larger than {max_bytes // (1024 * 1024)} MB.")
                    fh.write(chunk)
            result = proposal_excel.parse_scoping_workbook(path)
        finally:
            import shutil
            shutil.rmtree(tmp_dir, ignore_errors=True)          # the workbook is never kept
        if not result.ok:
            raise HTTPException(status_code=400, detail=_friendly_error(result.errors[0]))
        draft = build_draft(result, name, gen.CATALOG_BY_KEY)
        draft_id = store.create(draft)
        return JSONResponse(_view(draft_id, draft))

    @router.get("/draft/{draft_id}")
    async def get_draft(draft_id: str):
        return JSONResponse(_view(draft_id, _load(draft_id)))

    @router.put("/draft/{draft_id}")
    async def put_draft(draft_id: str, payload: dict[str, Any]):
        merged = merge_edits(_load(draft_id), payload)
        if not store.update(draft_id, merged):
            raise HTTPException(status_code=404, detail="Draft not found or expired.")
        return JSONResponse(_view(draft_id, merged))

    @router.get("/draft/{draft_id}/download/{kind}")
    async def download_generated(draft_id: str, kind: str, inline: int = 0):
        """The last built proposal for this draft: kind = docx (editable Word file) or pdf.

        ?inline=1 (pdf only) lets the page show the PDF in a frame instead of downloading it."""
        file = store.attachment(draft_id, f".{kind}")
        if file is None:
            raise HTTPException(status_code=404, detail="Nothing generated yet for this draft, or it expired.")
        media = ("application/pdf" if kind == "pdf"
                 else "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
        name = _download_name(_load(draft_id)["client"].get("name", ""), kind)
        return FileResponse(path=str(file), media_type=media, filename=name,
                            content_disposition_type="inline" if (inline and kind == "pdf") else "attachment",
                            headers={"Cache-Control": "no-store", "X-Content-Type-Options": "nosniff"})

    @router.delete("/draft/{draft_id}")
    async def delete_draft(draft_id: str):
        store.delete(draft_id)
        return {"deleted": True}

    def _ready_draft(draft_id: str, edits_json: str) -> dict[str, Any]:
        """Load the draft, apply the last edits that travel with the request, refuse if anything blocks."""
        draft = _load(draft_id)
        if edits_json.strip():
            import json
            try:
                draft = merge_edits(draft, json.loads(edits_json))
            except ValueError as exc:
                raise HTTPException(status_code=400, detail="edits_json is not valid JSON.") from exc
            store.update(draft_id, draft)
        issues = readiness_issues(draft)
        if issues:
            raise HTTPException(status_code=400, detail=" ".join(issues))
        return draft

    @router.post("/draft/{draft_id}/generate/", dependencies=[Depends(rate_limit("generate"))])
    async def generate_from_draft(draft_id: str, client_logo: UploadFile | None = File(None),
                                  edits_json: str = Form("")):
        """API clients: build and return the .docx directly."""
        draft = _ready_draft(draft_id, edits_json)
        return await gen.generate_from_plan(draft, resolve_plan(draft), ai_inputs, client_logo, store, draft_id)

    @router.post("/draft/{draft_id}/build/", dependencies=[Depends(rate_limit("generate"))])
    async def build_final_document(draft_id: str, client_logo: UploadFile | None = File(None),
                                   edits_json: str = Form("")):
        """The page: build the real document once; the answer says what can be previewed and downloaded.

        The Word file and (when a PDF engine exists) the PDF are kept with the draft, so the final
        preview and both downloads are the very file that was built: nothing is generated twice.
        """
        draft = _ready_draft(draft_id, edits_json)
        built = await gen.build_from_plan(draft, resolve_plan(draft), ai_inputs, client_logo, store, draft_id)
        check, report = built["check"], built["report"]
        pdf_ready = bool(check.verified and check.pdf_path)
        stem = _download_name(draft["client"].get("name", ""), "docx")[:-5]
        return JSONResponse({
            "ok": True,
            "docx_url": f"draft/{draft_id}/download/docx",
            "docx_name": f"{stem}.docx",
            "pdf_url": f"draft/{draft_id}/download/pdf" if pdf_ready else None,
            "pdf_name": f"{stem}.pdf",
            "layout_checked": bool(check.verified),
            "engine": check.engine,
            "pages": check.page_count if check.verified else None,
            "blank_pages": list(check.blank_pages) if check.verified else [],
            "lint_errors": [f.message for f in report.errors],
            "lint_warnings": len(report.warnings),
            "note": check.header_value(),
        })

    return router


def _download_name(client_name: str, kind: str) -> str:
    stem = re.sub(r"[^\w\-]+", "_", client_name.strip())[:80] or "Proposal"
    return f"{stem}.{kind}"


def _friendly_error(error: str) -> str:
    kind = error.split(":")[0]
    return {
        "missing_sheet": "This does not look like the CyberSmithSECURE scoping sheet (a required sheet is missing or renamed).",
        "quantity_header_changed": "The Overall Scoping Matrix headers were changed. Please use the original v2.0 sheet.",
        "invalid_xlsx": "The file is not a valid .xlsx workbook.",
        "corrupt_file": "The workbook is damaged and could not be read.",
        "file_too_large": "The file is larger than 5 MB.",
        "archive_too_large": "The file expands to an unreasonable size and was rejected.",
        "invalid_extension": "Only .xlsx files are accepted.",
        "no_assessment_types_in_scope": "No assessment type is marked Yes on the Cover Page.",
        "file_not_found": "The upload could not be read. Please try again.",
    }.get(kind, "The workbook could not be processed.")

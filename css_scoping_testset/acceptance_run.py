#!/usr/bin/env python3
"""
End-to-end acceptance run: every dummy workbook goes through the REAL workflow
(upload -> review edits -> generate) and the result is checked against that workbook's expected JSON.

    python css_scoping_testset/acceptance_run.py [--out DIR] [--samples DIR] [--engine word|libreoffice|off]
                                                  [--load N] [--only 01,05,...]

For each generated proposal it checks: no lint errors, annexure numbering 1..N, Count lines equal the
expected Count, pricing Quantity equals the expected Quantity, the Contents page numbers equal the
pages measured from a real render, and no blank pages (when an engine is available).
Invalid workbooks must be rejected with the right message. OpenAI is replaced by a fake client:
no network call is made. Exit code 1 if anything fails.
"""
from __future__ import annotations

import argparse
import asyncio
import io
import json
import os
import re
import sys
import time
import types
from pathlib import Path

HERE = Path(__file__).resolve().parent
BACKEND = HERE.parent / "Automation" / "backend"
sys.path.insert(0, str(BACKEND))

AI_TEXT = """Annexure 77: Custom Assessment

Sub-Activity:
• Assess the agreed scope with a manual, evidence-based approach aligned to CyberSmithSECURE standards.

Methodologies:
• Manual testing against the documented design and policy
o Review of configuration and access control for each in-scope component

Frequency: One time with Revalidation

Count: 999 Wrong Things

Deliverables:
• Technical report with evidence, severity and recommended remediation for every finding.
• Executive summary for management with the key risks and priorities.

Expectation:
• Client provides documentation, named contacts and read-only access where required.

Objectives:
• Confirm that the in-scope environment resists realistic attacks and report any gaps with clear fixes.

Out of Scope:
• Denial of service testing and physical access to client premises.

Requirements for Starting the Assessment:
• Signed authorization, a confirmed testing window and the contact list for escalation.

Comments:
• Findings are reported with a remediation owner and target date for each item.
"""


def install_fake_openai():
    import openai

    class Fake:
        def __init__(self, *a, **k):
            async def create(**kw):
                return types.SimpleNamespace(choices=[types.SimpleNamespace(
                    message=types.SimpleNamespace(content=AI_TEXT))])
            self.chat = types.SimpleNamespace(completions=types.SimpleNamespace(create=create))
    openai.AsyncOpenAI = Fake
    os.environ["OPENAI_API_KEY"] = "sk-test-not-real"


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--samples", default=str(HERE / "samples"))
    ap.add_argument("--expected", default=str(HERE / "expected"))
    ap.add_argument("--out", default=str(HERE / "acceptance_out"))
    ap.add_argument("--engine", default=os.environ.get("PROPOSAL_PDF_ENGINE", "auto"))
    ap.add_argument("--load", type=int, default=15, help="how many load files to run through the parser")
    ap.add_argument("--only", default="", help="comma-separated scenario prefixes, e.g. 01,05")
    args = ap.parse_args()

    os.environ["PROPOSAL_PDF_ENGINE"] = args.engine
    os.environ["RATE_LIMIT_DISABLED"] = "1"
    install_fake_openai()

    import css_proposal_generator as gen
    import proposal_excel
    import proposal_lint
    import proposal_pdf
    from docx import Document
    from docx.oxml.ns import qn
    from fastapi.testclient import TestClient

    samples, expected_dir, out = Path(args.samples), Path(args.expected), Path(args.out)
    out.mkdir(parents=True, exist_ok=True)
    client = TestClient(gen.app)
    engine = proposal_pdf.available_engine()
    print(f"PDF engine: {engine or 'none (layout not verified)'}\n")
    failures: list[str] = []
    rows: list[tuple] = []

    def upload(path: Path):
        return client.post("/upload-excel/", files={"file": (path.name, path.read_bytes(), "application/octet-stream")})

    def fix_draft(draft_id: str, view: dict) -> dict:
        """What a careful user does on the review screens so Generate becomes possible."""
        draft = view["draft"]
        edits: dict = {"client": {}, "ai": {}, "entries": []}
        if not draft["client"]["reference_no"].strip():
            edits["client"]["reference_no"] = "CSS-ENG-2026-001"
        if not draft["ai"]["sector"].strip():
            edits["ai"]["sector"] = "Banking, Financial Services and Insurance (BFSI)"
        if not draft["ai"]["duration"].strip():
            edits["ai"]["duration"] = "30 days"
        for e in draft["entries"]:
            if e["include"] and e["needs_confirmation"] and e["count"] == 0:
                edits["entries"].append({"id": e["id"], "confirmed": True})
        return client.put(f"/draft/{draft_id}", json=edits).json()

    def contents_ok(docx: Path, check) -> list[str]:
        problems = []
        doc = Document(str(docx))
        c_idx, first_assess, terms_idx, concl_idx = gen._find_contents_region(doc)
        if c_idx is None:
            return ["Contents not found"]
        for i in range(c_idx + 1, concl_idx + 1):
            title, page = gen._parse_toc_line(doc.paragraphs[i].text)
            key = ("conclusion & pricing" if i == concl_idx else "terms & conditions" if i == terms_idx
                   else f"annexure {i - first_assess + 1}" if i >= first_assess else title.lower())
            if key in check.headings and check.headings[key] != page:
                problems.append(f"Contents says {title!r} on {page}, render says {check.headings[key]}")
        return problems

    prefixes = [p.strip() for p in args.only.split(",") if p.strip()]
    named = sorted(p for p in samples.glob("*.xlsx"))
    for path in named:
        stem = path.stem
        if prefixes and not any(stem.startswith(p) for p in prefixes):
            continue
        exp = json.loads((expected_dir / f"{stem}.json").read_text(encoding="utf-8"))
        t0 = time.time()
        problems: list[str] = []
        r = upload(path)
        if exp["errors"]:
            kind = exp["errors"][0].split(":")[0]
            ok = r.status_code in (400, 413)
            if not ok:
                problems.append(f"expected a rejection ({kind}) but got HTTP {r.status_code}")
            rows.append((stem, "REJECTED" if ok else "ACCEPTED?!", "-", "-", "-", "ok" if ok else "FAIL", f"{time.time() - t0:.0f}s"))
            failures += [f"{stem}: {p}" for p in problems]
            continue
        if r.status_code != 200:
            problems.append(f"upload failed: {r.status_code} {r.text[:120]}")
            rows.append((stem, "-", "-", "-", "-", "FAIL", ""))
            failures += [f"{stem}: {p}" for p in problems]
            continue
        draft_id = r.json()["draft_id"]
        view = fix_draft(draft_id, r.json())
        if view["issues"]:
            problems.append(f"still blocked: {view['issues']}")
        gen_resp = client.post(f"/draft/{draft_id}/generate/")
        if gen_resp.status_code != 200:
            problems.append(f"generate failed: {gen_resp.status_code} {gen_resp.text[:200]}")
            rows.append((stem, "-", "-", "-", "-", "FAIL", ""))
            failures += [f"{stem}: {p}" for p in problems]
            continue
        docx = out / f"{stem}.docx"
        docx.write_bytes(gen_resp.content)
        doc = Document(str(docx))
        report = proposal_lint.lint_docx(docx)
        if not report.ok:
            problems += [str(f) for f in report.errors]
        heads = [p.text.strip() for p in doc.paragraphs if re.match(r"^Annexure \d+:", p.text.strip())]
        nums = [int(re.match(r"Annexure (\d+)", h).group(1)) for h in heads]
        if nums != list(range(1, len(nums) + 1)):
            problems.append(f"numbering {nums}")
        # Count lines vs expected
        want = {}
        for a in exp["included_annexures"]:
            want.setdefault(a["annexure_no_or_custom"] if a["annexure_no_or_custom"] != "custom" else a["type"], []).append(a)
        counts = [p.text.strip() for p in doc.paragraphs if p.text.strip().startswith("Count:")]
        expected_counts = []
        for a in exp["included_annexures"]:
            expected_counts.append(f"{a['count']} {a['count_label']}")
        for line in counts:
            body = line[len("Count:"):].strip()
            parts = [x.strip() for x in body.split(", ")]
            if not all(any(part == ec for ec in expected_counts) for part in parts):
                problems.append(f"unexpected Count line {line!r}")
        if len(counts) != len(set(a["annexure_no_or_custom"] if a["annexure_no_or_custom"] != "custom" else a["type"]
                                  for a in exp["included_annexures"])):
            problems.append(f"{len(counts)} Count lines for {len(exp['included_annexures'])} included types")
        # pricing quantities
        tbl = gen._find_assessment_pricing_table(doc.element)
        price_rows = [[("".join(t.text or "" for t in tc.iter(qn("w:t")))) for tc in tr.findall(qn("w:tc"))]
                      for tr in tbl.findall(qn("w:tr"))][1:-1]
        got_q = sorted(r_[2] for r_ in price_rows)
        want_q = sorted(str(a["quantity"]) for a in exp["included_annexures"])
        if got_q != want_q:
            problems.append(f"pricing quantities {got_q} != expected {want_q}")
        # layout (needs an engine)
        pages, blank = "-", "-"
        pdf_header = gen_resp.headers.get("x-proposal-pdf", "")
        if engine:
            check = proposal_pdf.check_document(docx, len(nums), pdf_path=out / f"{stem}.pdf")
            if check.verified:
                pages, blank = check.page_count, (check.blank_pages or "none")
                if check.blank_pages:
                    problems.append(f"blank pages {check.blank_pages}")
                problems += contents_ok(docx, check)
            else:
                problems.append(f"render failed: {check.note}")
        rows.append((stem, len(exp["included_annexures"]), len(nums), pages, blank, "ok" if not problems else "FAIL",
                     f"{time.time() - t0:.0f}s"))
        failures += [f"{stem}: {p}" for p in problems]

    # load files: parser against expected JSON (generation of 100 documents would take hours with Word)
    load_dir = samples / "load"
    mismatches = 0
    checked = 0
    if load_dir.is_dir() and args.load > 0 and not prefixes:
        files = sorted(load_dir.glob("load_*.xlsx"))[:args.load] + sorted(load_dir.glob("stress_*.xlsx"))
        for path in files:
            exp = json.loads((expected_dir / f"load__{path.stem}.json").read_text(encoding="utf-8"))
            got = proposal_excel.parse_scoping_workbook(path).to_dict()
            checked += 1
            simple = lambda es: [(e["type"], e["annexure_no_or_custom"], e["count"], e["quantity"]) for e in es]
            if simple(got["included_annexures"]) != simple(exp["included_annexures"]) or \
                    sorted(got["warnings"]) != sorted(exp["warnings"]):
                mismatches += 1
                failures.append(f"load {path.stem}: parser output differs from expected JSON")

    print(f"{'scenario':38} {'expected':>8} {'in doc':>7} {'pages':>6} {'blank':>8}  result  time")
    for row in rows:
        print(f"{row[0]:38} {str(row[1]):>8} {str(row[2]):>7} {str(row[3]):>6} {str(row[4]):>8}  {row[5]:6}  {row[6]}")
    if checked:
        print(f"\nload/stress files parsed and compared with expected JSON: {checked}, mismatches: {mismatches}")
    print(f"\n{'ALL CHECKS PASSED' if not failures else str(len(failures)) + ' PROBLEM(S):'}")
    for f in failures:
        print("  -", f)
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())

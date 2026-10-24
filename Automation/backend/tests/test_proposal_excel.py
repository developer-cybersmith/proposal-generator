"""
Tests for proposal_excel.py (the scoping-workbook reader) and annexure_map.json.

The main test runs the reader against every fixture in css_scoping_testset/ and compares it with
the expected JSON written by that folder's independent reference parser. Workbooks are regenerated
into a temp folder once per session (generate.py is deterministic), so samples/ need not be in git.

Run from Automation/backend/:   python -m pytest tests/test_proposal_excel.py -v
"""
from __future__ import annotations

import json
import sys
from pathlib import Path

import pytest

BACKEND = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(BACKEND))

from proposal_excel import load_map, parse_scoping_workbook  # noqa: E402

TESTSET = BACKEND.parents[1] / "css_scoping_testset"

pytestmark = pytest.mark.skipif(
    not (TESTSET / "generate.py").is_file(), reason="css_scoping_testset/ not present")


def _expected_files(root: Path):
    return sorted((root / "expected").glob("*.json"))


def _sample_for(root: Path, expected_json: Path) -> Path:
    stem = expected_json.stem                      # "load__load_001" or "01_all_in_scope_full"
    if stem.startswith("load__"):
        return root / "samples" / "load" / (stem[len("load__"):] + ".xlsx")
    return root / "samples" / (stem + ".xlsx")


def _simplify(entry: dict) -> dict:
    keys = ("type", "annexure_no_or_custom", "count", "quantity", "count_label", "frequency")
    return {k: entry[k] for k in keys}


def _error_kind(message: str) -> str:
    return message.split(":")[0]


# --------------------------------------------------------------------------- fixtures parity
def test_every_fixture_matches_the_reference_parser(fixtures):
    files = _expected_files(fixtures)
    assert len(files) >= 137
    failures: list[str] = []
    for exp_path in files:
        expected = json.loads(exp_path.read_text(encoding="utf-8"))
        got = parse_scoping_workbook(_sample_for(fixtures, exp_path)).to_dict()
        name = exp_path.stem
        if expected["errors"]:
            # error scenarios: the same failure(s); byte counts in messages may differ slightly
            if [_error_kind(e) for e in got["errors"]] != [_error_kind(e) for e in expected["errors"]]:
                failures.append(f"{name}: errors {got['errors']} != {expected['errors']}")
            if _error_kind(expected["errors"][0]) in ("missing_sheet", "quantity_header_changed",
                                                       "invalid_xlsx", "corrupt_file"):
                if got["errors"] != expected["errors"]:
                    failures.append(f"{name}: error text {got['errors']} != {expected['errors']}")
            if "no_assessment_types_in_scope" in expected["errors"]:
                if got["client"] != expected["client"] or got["warnings"] != expected["warnings"]:
                    failures.append(f"{name}: client/warnings differ")
            continue
        checks = {
            "included": ([_simplify(e) for e in got["included_annexures"]],
                         [_simplify(e) for e in expected["included_annexures"]]),
            "pricing": (got["pricing_lines"], expected["pricing_lines"]),
            "distinct": (got["distinct_annexures"], expected["distinct_annexures"]),
            "warnings": (sorted(got["warnings"]), sorted(expected["warnings"])),
            "client": (got["client"], expected["client"]),
            "reference": (got["reference_no"], expected["reference_no"]),
        }
        for label, (a, b) in checks.items():
            if a != b:
                failures.append(f"{name}: {label} differs\n   got      {a}\n   expected {b}")
    assert not failures, "\n".join(failures[:15]) + f"\n... {len(failures)} fixture(s) differ"


# --------------------------------------------------------------------------- the screenshot example
def test_screenshot_example_count_and_quantity(fixtures):
    r = parse_scoping_workbook(fixtures / "samples" / "31_image_example_counts.xlsx")
    by_key = {e.type_key: e for e in r.included}
    assert (by_key["saas_security"].count, by_key["saas_security"].quantity) == (3, 12)
    assert (by_key["web_app_vapt"].count, by_key["web_app_vapt"].quantity) == (6, 23)
    assert r.ok and not r.warnings


def test_row_breakdown_is_reported_for_the_review_screen(fixtures):
    r = parse_scoping_workbook(fixtures / "samples" / "31_image_example_counts.xlsx")
    saas = next(e for e in r.included if e.type_key == "saas_security")
    statuses = {d.row: d.status for d in saas.rows}
    assert statuses == {49: "counted", 50: "blank", 51: "blank", 52: "counted", 53: "counted"}
    assert saas.rows[0].category == "SaaS App Name & URL"


def test_external_and_internal_are_two_lines_under_one_annexure(fixtures):
    r = parse_scoping_workbook(fixtures / "samples" / "05_external_and_internal_both.xlsx")
    assert r.distinct_annexures() == [3]
    assert [(l["label"], l["quantity"]) for l in r.pricing_lines()] == [
        ("External Network Penetration Testing", 25), ("Internal Network Penetration Testing", 90)]


def test_zero_blank_negative_text_rows(fixtures):
    r = parse_scoping_workbook(fixtures / "samples" / "32_zero_blank_negative_rows.xlsx")
    by_key = {e.type_key: e for e in r.included}
    assert (by_key["cloud_security"].count, by_key["cloud_security"].quantity) == (2, 5)
    assert (by_key["mobile_vapt"].count, by_key["mobile_vapt"].quantity) == (2, 14)
    assert by_key["source_code_review"].needs_confirmation
    kinds = {_error_kind(w) for w in r.warnings}
    assert kinds == {"negative_quantity_skipped", "non_numeric_quantity_skipped", "quantity_zero_confirm"}


def test_types_without_matrix_rows_default_to_one(fixtures):
    r = parse_scoping_workbook(fixtures / "samples" / "03_custom_only.xlsx")
    social = next(e for e in r.included if e.type_key == "social_engineering")
    assert (social.count, social.quantity, social.annexure) == (1, 1, "custom")


def test_uncached_formulas_warn_and_cached_formulas_do_not(fixtures):
    cached = parse_scoping_workbook(fixtures / "samples" / "19_formulas_cached.xlsx")
    uncached = parse_scoping_workbook(fixtures / "samples" / "20_formulas_uncached.xlsx")
    assert not any(w.startswith("formula_no_cached_value") for w in cached.warnings)
    assert sum(w.startswith("formula_no_cached_value") for w in uncached.warnings) == 3


def test_placeholder_cover_values_are_empty_and_reference_is_required(fixtures):
    r = parse_scoping_workbook(fixtures / "samples" / "11_placeholders_untouched.xlsx")
    assert r.reference_no is None
    assert "reference_required" in r.warnings
    assert r.client["preferred_start_date"] is None and r.client["nda_signed"] is None


@pytest.mark.parametrize("name, error", [
    ("23_wrong_sheet_name", "missing_sheet"),
    ("24_missing_sheet", "missing_sheet"),
    ("25_quantity_header_changed", "quantity_header_changed"),
    ("26_empty_workbook", "missing_sheet"),
    ("27_csv_renamed_xlsx", "invalid_xlsx"),
    ("28_oversize", "file_too_large"),
    ("29_corrupt_zip", "invalid_xlsx"),
    ("30_all_types_No", "no_assessment_types_in_scope"),
])
def test_invalid_files_are_rejected_without_raising(fixtures, name, error):
    r = parse_scoping_workbook(fixtures / "samples" / f"{name}.xlsx")
    assert not r.ok
    assert _error_kind(r.errors[0]) == error


def test_non_xlsx_extension_is_rejected(tmp_path):
    p = tmp_path / "scoping.xlsm"
    p.write_bytes(b"PK\x03\x04")
    assert _error_kind(parse_scoping_workbook(p).errors[0]) == "invalid_extension"


def test_missing_file_does_not_raise(tmp_path):
    assert parse_scoping_workbook(tmp_path / "nope.xlsx").errors == ["file_not_found"]


# --------------------------------------------------------------------------- the mapping file
def test_map_is_consistent():
    cfg = load_map()
    types = cfg["types"] + cfg["types_without_matrix_rows"]
    assert len({t["key"] for t in types}) == len(types)
    cells = [t["cover_cell"] for t in types]
    assert len(set(cells)) == len(cells) == 14
    band_rows = {3, 16, 23, 34, 42, 48, 54, 61, 66}
    for t in cfg["types"]:
        assert t["rows"] and not (set(t["rows"]) & band_rows), t["key"]
        assert all(4 <= r <= 70 for r in t["rows"])


def test_map_catalog_keys_exist_in_the_generator():
    from css_proposal_generator import ASSESSMENT_CATALOG
    catalog = {item["key"]: item["annexure_number"] for item in ASSESSMENT_CATALOG}
    cfg = load_map()
    for t in cfg["types"] + cfg["types_without_matrix_rows"]:
        if t["catalog_key"] is None:
            assert t["annexure"] == "custom"
        else:
            assert catalog[t["catalog_key"]] == t["annexure"], t["key"]
    for t in cfg["manual_toggle_annexures"]:
        assert catalog[t["catalog_key"]] == t["annexure"]
    # every one of the 13 template annexures is reachable (from Excel or a manual toggle)
    reachable = {t["annexure"] for t in cfg["types"] if t["annexure"] != "custom"}
    reachable |= {t["annexure"] for t in cfg["manual_toggle_annexures"]}
    assert reachable == set(range(1, 14))

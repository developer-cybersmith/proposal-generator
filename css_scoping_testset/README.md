# CSS Pre-Engagement Scoping v2 — Parser Test Set

Generated test fixtures for the parser that reads `CSS Pre-Engagement Scoping v2.xlsx`
(CyberSmithSECURE scoping sheet, v2.0). Rules follow `docs/ADVANCED_FEATURE_PLAN.md`
(decisions D11, D15, D16, D19, D21, D23, D30–D33).

All workbooks are produced by `generate.py` from a fresh copy of `template.xlsx`. Only values are
written into input cells; sheet names, header text, row order, merged ranges, column widths and
formatting are never modified. All data is fictional.

## Layout

```
template.xlsx        blank v2.0 template (generator input, do not edit)
generate.py          deterministic generator (seed 12345, PYTHONHASHSEED pinned)
manifest.csv         index of every file with expected counts/errors
samples/             32 named scenario workbooks            (regenerable, not committed)
samples/load/        100 randomised + 5 stress workbooks    (regenerable, not committed)
expected/            one <file>.json per workbook           (committed)
```

Run: `python generate.py [--out DIR]`. It clears and regenerates `samples/` and `expected/` under
`--out` (default: this folder) and rewrites `manifest.csv`. Workbook bytes differ between runs
(embedded timestamps) but the expected JSON content is identical. LibreOffice is used only for
scenario 19; if it is not installed the generator writes the cached formula values itself.

## Expected-result JSON (produced by `reference_parser()` in `generate.py`)

| field | meaning |
|---|---|
| `included_annexures[]` | one object per in-scope type: `{type, annexure_no_or_custom, count, quantity, count_label, frequency}` |
| `pricing_lines[]` | one billable line per in-scope type: `{label, annexure_no_or_custom, quantity, frequency}` (External + Internal give two lines under Annexure 3) |
| `distinct_annexures[]` | unique annexure identifiers at document level |
| `warnings[]` | soft issues (below) |
| `client{}` | Cover Page client block; placeholders and blanks are `null` |
| `reference_no` | Cover `F16`, or `null` when blank/placeholder |
| `errors[]` | hard failures (below) |

The reference parser is deliberately independent of the real parser (`proposal_excel.py`), so the
two cross-check each other.

## Rules implemented

An annexure is included only when its Cover Page "In Scope?" cell is `yes`, `y` or `✓`
(trimmed, case-insensitive). Everything else (`TRUE`, `1`, `maybe`, `No`, blank) is out of scope.

**Count** = number of rows in the type's block holding a number **above 0**.
**Quantity** = sum of those numbers. Blank and 0 are skipped silently; text, negatives and
formulas without a cached value are skipped with a warning. Count 0 raises `quantity_zero_confirm`.

| Type | Annexure | Matrix rows used | Count label |
|---|---|---|---|
| External Network PT | 3 | 17–22 | External Infrastructure categories |
| Internal Network PT | 3 | 4–15 | Internal Infrastructure categories |
| Web Application VAPT | 4 | 55–60 | Web Application categories |
| Configuration / Hardening | 5 | 4, 5, 6, 9, 10, 11 | Device & Server categories |
| Mobile VAPT | 7 | 62–65 | Mobile Application categories |
| AD Security | 8 | 35–41 | Active Directory categories |
| Cloud Security | 10 | 24–33 | Cloud Infrastructure categories |
| API Security | 11 | 19, 32 | API categories |
| Source Code Review | 13 | 67–70 | Source Code categories |
| Network Segmentation | custom | 43–47 | Network Segmentation categories |
| SaaS Security | custom | 49–53 | SaaS Application categories |
| Social Engineering, Red Team, Compliance Gap | custom | none | Count 1, Quantity 1 |

Worked example (scenario 31, from the screenshot): SaaS rows 49–53 = 5, blank, blank, 4, 3 gives
Count 3 / Quantity 12; Web rows 55–60 = 2, 1, 4, 5, 3, 8 gives Count 6 / Quantity 23.

Frequency defaults to `One Time`. Reference number is `F16` unless blank or still the placeholder
`CSS-ENG-YYYY-XXX`. Column-F grey placeholder strings count as empty.

## Warning and error vocabulary

Warnings (processing continues):

- `non_numeric_quantity_skipped: <asset category> (<sheet>!<cell>)=<value>`
- `negative_quantity_skipped: <asset category> (<sheet>!<cell>)=<value>`
- `quantity_zero_confirm: <type>` — block has no value above 0.
- `formula_no_cached_value: <sheet>!<cell>`
- `reference_required` — `F16` blank or placeholder.
- `date_not_real_date: <sheet>!<cell>=<value>`

Errors (file rejected):

- `missing_sheet: <name>` · `quantity_header_changed: ...` · `invalid_xlsx: ...` · `corrupt_file: ...`
- `file_too_large: <bytes> > 5242880` (byte count varies by a couple of bytes between runs:
  match the prefix only) · `no_assessment_types_in_scope`

## Scenarios

**Valid:** 01 all_in_scope_full · 02 known_only · 03 custom_only · 04 mixed_known_and_custom ·
05 external_and_internal_both · 06 config_and_api · 07 mobile_android_only · 08 mobile_both ·
09 minimal · **31 image_example_counts** (SaaS 3/12, Web 6/23)

**Edge:** 10 zero_quantity · 11 placeholders_untouched · 12 yes_no_variants · 13 text_in_quantity ·
14 big_numbers · 15 negative_and_float · 16 text_dates_vs_real_dates · 17 unicode_and_long ·
18 huge_context · 19 formulas_cached · 20 formulas_uncached · 21 extra_hidden_sheet ·
22 extra_data_outside_known_cells · **32 zero_blank_negative_rows**

**Invalid:** 23 wrong_sheet_name · 24 missing_sheet · 25 quantity_header_changed · 26 empty_workbook ·
27 csv_renamed_xlsx · 28 oversize · 29 corrupt_zip · 30 all_types_No

**Load:** `samples/load/load_001..100` (seed 12345: 1–14 types in scope, whole blocks populated,
10% zeros, 5% text junk) and `stress_01..05` (1000-char cells, every row filled).
Every file has a matching `expected/load__*.json`.

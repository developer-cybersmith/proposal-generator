# Prompt: generate dummy filled scoping Excel files for parser testing

Paste everything inside the block below into an AI model that can run Python (attach `CSS Pre-Engagement Scoping v2.xlsx` first). Expected-result rules follow `docs/ADVANCED_FEATURE_PLAN.md` (§4, D11, D15, D16, D19, D21, D27).

````text
You are a QA data engineer. I attach a BLANK Excel file: "CSS Pre-Engagement Scoping v2.xlsx" (CyberSmithSECURE scoping sheet, v2.0). Using Python + openpyxl, generate a set of FILLED DUMMY workbooks, plus an expected-results manifest, to test a parser that reads this file.

## Hard rules
1. Always start from a fresh copy of the attached template. NEVER change sheet names, header text, row order, merged cells, column widths or formatting. Only write VALUES into input cells.
2. Use only fictional data (fake companies, people, emails like name@example.com, phones like +91 98xxxxxxx0).
3. Write numbers as real numbers (int), not text, unless a scenario says otherwise. Write dates as real Excel dates unless a scenario says text.
4. Write into the TOP-LEFT cell of merged ranges.
5. Output a folder: samples/ (xlsx files), expected/ (one JSON per file), manifest.csv, generate.py (re-runnable, seeded), README.md.

## Cells to fill (sheet "Cover Page")
Client response column C, rows 16-22 (labels in B): Client Organization Name C16, Primary Contact C17, Designation C18, Email C19, Phone C20, Alternate Contact Name C21, Alternate Contact Email C22.
Right block, column F, rows 16-22 (labels in E): Engagement Reference No. F16, Date of Submission F17, Preferred Start F18, Preferred End F19, Project Classification F20, Regulatory Driver F21, NDA Signed F22. NOTE: column F contains grey PLACEHOLDER text in the blank file (e.g. "CSS-ENG-YYYY-XXX", "DD-MMM-YYYY"). Overwrite it with real values, except in the scenario that says "placeholders untouched".
In-scope Yes/No (header "In Scope? (Yes/No)"):
 column C rows 26-32 for B26 External Network PT, B27 Internal Network PT, B28 Cloud Security, B29 Active Directory Security, B30 Network Segmentation/VLAN, B31 SaaS Security, B32 Configuration/Hardening Review;
 column F rows 26-32 for E26 Web Application VAPT, E27 Mobile VAPT, E28 Source Code Review/SAST, E29 API Security Testing, E30 Social Engineering/Phishing, E31 Red Team, E32 Compliance Gap Assessment.
Ignore the stray "s" in K23; leave it.

## Cells to fill (sheet "Overall Scoping Matrix")
Column D = Quantity, rows: 4-15, 17-22, 24-33, 35-41, 43-47, 49-53, 55-60, 62-65, 67-70 (rows 3,16,23,34,42,48,54,61,66 are band headers: never write there). Columns E-I (methodology, service provided, box type, delivery mode, deadline) are optional free text; fill some rows, leave others blank.

## Sheets for AI-context (fill with realistic text in some scenarios, leave blank in others)
"General Scoping Questions": answers in E:I (App/Asset 1-5) next to each question.
"Web Application Assessment", "Mobile Application Assessment": answers in D (primary) and E (secondary). "Cloud Security Assessment", "AD Security Assessment", "Network Segmentation Testing": answers in column D. "Rules of Engagement": C5:G9 contacts table, C13:C25 parameters, C29:C35 compliance.

## Expected-result rules (compute these into expected/<file>.json)
Annexure is INCLUDED only if its Cover Page type is Yes (case-insensitive: yes, y, ✓ accepted; anything else = not in scope).
Type -> template annexure no. / quantity rule (sum of numeric cells; text, blanks and Y/N are skipped and listed as warnings):
- External Network PT -> Annexure 3 (External line) = sum D17:D22, unit "External Assets"
- Internal Network PT -> Annexure 3 (Internal line) = sum D4:D15, unit "Internal Assets". If both Yes: ONE annexure, TWO pricing lines.
- Web App VAPT -> 4 = D55, "Web Applications"
- Mobile VAPT -> 7 = text "Android - D62, iOS - D63", "Mobile Applications"
- Cloud -> 10 = sum D24:D33, "Cloud Assets"
- AD Security -> 8 = D35, "AD Domains"
- API Security -> 11 = D19 + D32, "APIs"
- Source Code -> 13 = D67, "Repositories"
- Config/Hardening -> 5 = D4+D5+D6+D9+D10+D11, "Devices & Servers"
- Custom (AI) annexures, no template: Network Segmentation = D43 "VLANs"; SaaS = D49 "SaaS Applications"; Social Engineering, Red Team, Compliance Gap = no matrix rows, quantity defaults to 1, "Engagements".
- Included but quantity 0/blank -> flag "quantity_zero_confirm".
- Reference number = F16 unless blank or still "CSS-ENG-YYYY-XXX" (then empty + "reference_required").
- Placeholder strings in column F ("DD-MMM-YYYY", "Yes / No / In Progress", etc.) count as EMPTY.
JSON fields: included_annexures[{type, annexure_no_or_custom, quantity, unit, quantity_text}], pricing_lines[], warnings[], client{...}, reference_no, errors[].

## Scenarios to generate (one file each, descriptive names)
VALID
01 all_in_scope_full: all 14 types Yes, realistic quantities everywhere, all context sheets filled.
02 known_only: Web, Mobile, External Network, Cloud, Source Code Yes; rest No.
03 custom_only: only Segmentation, SaaS, Social Engineering, Red Team, Compliance Yes.
04 mixed_known_and_custom.
05 external_and_internal_both: both Yes, different quantities (expect 2 pricing lines).
06 config_and_api: only Config/Hardening and API Yes (tests row mappings 4-6,9-11 and 19+32).
07 mobile_android_only: D62=5, D63 blank. 08 mobile_both: D62=6, D63=3.
09 minimal: one type Yes, one quantity, other sheets blank.
EDGE / ROBUSTNESS
10 zero_quantity: types Yes but all their rows blank or 0.
11 placeholders_untouched: Cover column F left exactly as in the blank file.
12 yes_no_variants: use Yes, YES, yes, Y, y, ✓, " Yes " (spaces), No, N, n, blank, "maybe", "TRUE", 1.
13 text_in_quantity: some Quantity cells hold "ten", "N/A", "10.0.0.0/24", "Y", "-" next to numbers.
14 big_numbers: quantities 100000, 2350, 999999999; and a decimal 2.5.
15 negative_and_float: -3, 0.0, 7.0 (float that should read as 7).
16 text_dates_vs_real_dates: dates as text "05-Jan-2026", real dates, and "2026/01/05".
17 unicode_and_long: client name with accents, Devanagari, & < > " ' characters, and a 150-char name; very long emails.
18 huge_context: every context-sheet answer ~1500 characters, all 5 app columns filled.
19 formulas_cached: Quantity cells as formulas (=5+3, =SUM(D4:D6)) saved by Excel/LibreOffice so cached values exist (recalculate via LibreOffice headless if available).
20 formulas_uncached: same formulas written by openpyxl only (no cached values) - expected warning.
21 extra_hidden_sheet and 22 extra_data_outside_known_cells (junk in K23, Z99, and rows below 72).
INVALID (expected errors)
23 wrong_sheet_name (rename "Overall Scoping Matrix"), 24 missing_sheet (delete Cover Page), 25 quantity_header_changed (D2 renamed), 26 empty_workbook (new blank xlsx), 27 csv_renamed_xlsx, 28 oversize (> 5 MB via padding in a hidden sheet), 29 corrupt_zip (truncate a valid file), 30 all_types_No (nothing in scope).

## Load test set
Also generate 100 files in samples/load/: each randomised with a fixed seed (12345): random subset of 1-14 types Yes, random quantities 0-500 (10% zeros, 5% text junk), random context length. Write expected JSON for every one so tests can assert automatically. Additionally generate 5 "stress" files: 1000-character cells everywhere, maximum rows filled, all 5 app columns filled.

## Deliverables
1. generate.py that produces everything deterministically.
2. manifest.csv: file, scenario, expected_included_count, expected_warnings, expected_error.
3. README.md describing each scenario.
Print a summary table at the end. If any requirement is unclear, ask before generating.
````

Before trusting the generated `expected/*.json`, skim a few (especially scenarios 05, 06 and 08) against `docs/ADVANCED_FEATURE_PLAN.md` §4.

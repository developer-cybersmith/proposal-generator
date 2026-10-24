# Advanced Feature Plan — Excel-Driven Proposal Generation

**Status:** DRAFT v2. No code is written until the items in §15 are closed.
**Input format (fixed):** `CSS Pre-Engagement Scoping v2.xlsx`. No new Excel format will be introduced.
**Builds on:** `docs/PROPOSAL_GENERATOR_HANDOFF.md` (generator v1.3.0).
**Binding formatting standard:** `docs/CyberSmithSECURE_Proposal_Formatting_Specification.md` (see §0).

Legend: ✅ confirmed by you · 🟡 my proposal, needs your OK · ❓ open question

---

## 0. Non-negotiable rule: the Formatting Specification

> **Fonts, colours, spacing, tables, headers/footers, lists and pagination follow `CyberSmithSECURE_Proposal_Formatting_Specification.md`. These are company standards and are never traded off for convenience.**

How this is enforced, not just stated:
1. The spec is copied into the repo (`docs/`) and versioned with the code.
2. All values live in **one tokens module** (colours, font names, sizes, spacing, table metrics). **No raw hex codes or font names anywhere else in the code.**
3. Every piece of content the feature generates (pricing table, AI annexures, Count/Frequency lines) uses those tokens.
4. A **formatting lint** runs on every generated `.docx` and fails the build if it finds: a non-Corbel body run, a colour outside the token set, a heading made by direct formatting, typed bullets, empty-paragraph spacing, or a table wider than the text area.
5. A **Design Sample Gate**: before each phase is called done, one generated `.docx`/PDF is shown to you and approved visually against the spec.

Hard acceptance rules (apply to every generated proposal) ✅:
- **Annexure numbering is correct and sequential** (1…N in the final order), identical in the body headings, the Contents page and the pricing table.
- **No empty or near-empty pages** between sections (spec §10 and §14.4). No consecutive page breaks, no runs of blank paragraphs used as spacing.
- The Contents page lists **every** annexure in the body, with refreshed page numbers.
- Any text inside annexures that points to another annexure number ("see Annexure 8") is checked and corrected after renumbering.

---

## 1. Goal in one paragraph

Today the user ticks annexures by hand. We add a second way in: **upload the scoping Excel**. The system reads it, decides which annexures are in scope, works out a **quantity per annexure**, shows an **editable review screen** (annexures, quantities, client details, pricing table with empty prices), and then generates the same branded Word proposal as today. Annexures that exist in the template are copied as-is; annexures that do not exist are written by the AI in the organisation's format using the locked system prompt.

---

## 2. Decisions made ✅

| # | Decision |
|---|----------|
| D1 | Excel upload is an **additional** option; the manual form stays. |
| D2 | Excel format is fixed (v2.0): sheet names, headers and row order are not changed. |
| D3 | **Cover Page** gives client details and a Yes/No "In Scope?" for each assessment type. |
| D4 | **Overall Scoping Matrix** gives two numbers per annexure from its block of rows: **Count** (rows with a value above 0) and **Quantity** (sum of those cells). See §4.2 and D30–D34. |
| D5 | "Service Provided" (matrix column F) is informational only. |
| D6 | Quantity cells are numeric; non-numeric cells are skipped with a warning (safety only). |
| D7 | **Count** feeds the `Count:` line inside the annexure; **Quantity** feeds the Conclusion & Pricing table (see D30). |
| D8 | Types with no template annexure are written by the **AI custom annexure** flow (locked system prompt). If the Excel is not enough, the user fills the existing basic AI form. |
| D9 | The annexures with no Excel row (Post Breach, AD Post-Breach Attacks, Insider Threat, SOC Simulation, Training) are chosen with **manual toggles** on the review screen. |
| D10 | The list may be any mix: all known, only custom, or both. All three must work. |
| D11 | If **both** External and Internal Network PT are "Yes": the annexure appears once, the pricing table has **two lines** (External and Internal). |
| D12 | Pricing table is a **preview**: annexure + quantity filled, **prices empty**; the user may add prices, remove rows, or edit before generating, or leave prices blank. |
| D13 | The other 7 sheets are **not** used for proposal text; their answers are sent to the AI as context for custom annexures. |
| D14 | Authentication stays: shared password + optional API key. |
| D15 | Config / Hardening Review → **Annexure 5** (its text covers "systems, network devices and applications"). Quantity = rows 4, 5, 6 (routers, switches, firewalls) + rows 9, 10, 11 (Windows, Linux, database servers). *Chosen by me after reading Annexure 5, as you asked; editable in the mapping file.* |
| D16 | In scope = Yes but quantity 0/blank: the annexure is **included**, shown **amber** with "quantity is 0, please confirm"; the user edits it, confirms as-is, or removes it. **Generate stays disabled until every flagged row is resolved.** |
| D17 | ~~`Count: N <unit label>` using the sum~~ — **superseded by D30/D31** (Count is now the number of filled rows, not the sum). |
| D18 | The AI form's **Total cost is optional**; when the pricing table has prices, the total is passed to the AI automatically. It never blocks generation. |
| D19 | API Security Testing → Annexure 11, quantity = row 19 (Public APIs) + row 32 (API Gateways). |
| D20 | **Sector** is asked on the review screen (only when an AI annexure is included), pre-suggested from the Cover Page "Regulatory Driver", user confirms. |
| D21 | **Reference number**: taken from the Excel's "Engagement Reference No.", editable on the review screen. If blank or still the placeholder `CSS-ENG-YYYY-XXX`, the field is empty and required. |
| D22 | ~~Quantity cell = number + unit~~ — **superseded by D32** (number only). |
| D23 | **Frequency** column defaults to `One Time`, editable. |
| D24 | **Total row**: keep the model from `CyberSmithSECURE Annexures 2.docx` page 47 — one merged cell, `₹ /- * + 18% GST` (with the summed amount in Indian format when prices exist). |
| D25 | **Pricing table font = Corbel** and all table formatting per the spec. |
| D26 | Template work has **two tracks, both required**: Track A normalises everything the feature creates (this plan); Track B rebuilds the Base Template to the full spec afterwards (§12). |
| D27 | ~~Primary-row-only quantity rule~~ — **superseded by D30** (all rows in the block count). |
| D28 | **Order:** Track A (Excel feature) first, then Track B (template rebuild to the full spec). |
| D29 | **PDF proof and blank-page check** use **LibreOffice installed on Railway**: each generated proposal is rendered to PDF, checked for empty pages, and the PDF is offered beside the `.docx`. |
| D30 | **Two numbers per annexure, from all rows of its block** (Config Review rows 4–6 + 9–11 and API rows 19 + 32 keep their own rows): **Count** = number of rows holding a value; **Quantity** = sum of those values. Worked examples from your screenshot: SaaS → Count 3, Quantity 12; Web → Count 6, Quantity 23. |
| D31 | **`Count:` line wording:** `Count: N <label>`, e.g. `Count: 6 Web Application categories`. One label per annexure, editable. |
| D32 | **Pricing table Quantity cell:** the sum as a **number only** (e.g. `23`), editable. The old `Android - X, iOS - Y` text is dropped; Mobile Quantity is the plain sum of rows 62–65. |
| D33 | **Zero rule:** only values **above 0** count toward Count; blank and 0 are both skipped. A block whose sum is 0 triggers the amber confirm (D16). |
| D34 | **Two previews before generating:** *Preview 1* = list of annexures with Count and Frequency (editable); *Preview 2* = Conclusion & Pricing table with Quantity and an empty Cost column (editable). **Frequency is one shared field**: editing it in either preview updates both and the annexure's `Frequency:` line. Prices can be entered in Preview 2 or left blank and filled in Word after generation. |

---

## 3. What is in the Excel (analysed from the blank template)

| Sheet | Used for | Notes |
|-------|----------|-------|
| Cover Page | Client details, dates, In-Scope Yes/No for 16 assessment types | Client fields `C16:C22` and `F16:F22`; Yes/No in `C26:C32` and `F26:F32` |
| Overall Scoping Matrix | Quantity per asset row | Quantity = column D; blocks in §4.2 |
| General Scoping Questions | AI context only | 10 sections, answers in E:I |
| Web / Mobile / Cloud / AD / Segmentation detail sheets | AI context only | Question in C, answers in D/E |
| Rules of Engagement | AI context only | Testing window, exclusions, compliance |

**Traps the parser must handle**
1. Cover Page column F holds grey **example text** (`CSS-ENG-YYYY-XXX`, `DD-MMM-YYYY`, `Yes / No / In Progress`). Untouched placeholder = **empty**, not data.
2. `K23` contains a stray `s`. Anything outside the known cells is ignored.
3. The file you gave is **blank**; real people may type Yes/No, dates and numbers in different ways.

❓ **A filled-in sample Excel is still needed** before the parser rules are final.

---

## 4. Mapping rules (stored in editable `annexure_map.json`)

### 4.1 Cover Page type → annexure

| Excel assessment type | Template annexure | Handling |
|-----------------------|-------------------|----------|
| External Network Penetration Testing | 3 Network PT | static; own pricing line |
| Internal Network Penetration Testing | 3 Network PT | static; own pricing line |
| Web Application VAPT | 4 Web App PT | static |
| Mobile Application VAPT | 7 Mobile App PT | static |
| Cloud Security Assessment | 10 Cloud PT | static |
| Active Directory Security Assessment | 8 AD Security Assessment | static |
| API Security Testing | 11 API PT | static |
| Source Code Review / SAST | 13 Source Code Review | static |
| Configuration / Hardening Review | 5 Network Configuration Review | static (D15) |
| Network Segmentation / VLAN Testing | none | **AI custom** |
| SaaS Security Assessment | none | **AI custom** |
| Social Engineering / Phishing Simulation | none | **AI custom** |
| Red Team Engagement | none | **AI custom** |
| Compliance Gap Assessment | none | **AI custom** |
| *(no Excel row)* 1 Post Breach, 2 AD Post-Breach, 6 Insider Threat, 9 SOC Simulation, 12 Training | 1, 2, 6, 9, 12 | **manual toggle** |

### 4.2 Matrix block → Count and Quantity ✅ (D30–D34, supersedes the old "primary row" rule)

Every annexure gets **two numbers** from its rows in the matrix:

| Number | Definition | Goes into |
|--------|------------|-----------|
| **Count** | how many rows in the block hold a value **above 0** (blank and 0 are skipped) | the annexure's `Count:` line (Preview 1) |
| **Quantity** | the **sum** of the Quantity cells of those rows | the pricing table's Quantity cell (Preview 2) |

**Worked examples (from your screenshot of the filled matrix):**

| Block | Rows with a value | Count | Quantity |
|-------|-------------------|-------|----------|
| SaaS (49–53) | 49 = 5, 50 blank, 51 blank, 52 = 4, 53 = 3 | **3** | 5 + 4 + 3 = **12** |
| Web (55–60) | 55 = 2, 56 = 1, 57 = 4, 58 = 5, 59 = 3, 60 = 8 | **6** | 2 + 1 + 4 + 5 + 3 + 8 = **23** |

The rule applies to **all rows in the block**, whatever the row's meaning (modules, roles, Y/N rows and so on): if a number is typed, it counts. Config Review and API keep their own rows. Count labels are my first suggestions and are editable.

| Block (Excel rows) | Annexure | Rows used | `Count:` line label |
|--------------------|----------|-----------|---------------------|
| Internal Infrastructure (4–15) | 3 Internal | 4–15 | Internal Infrastructure categories |
| Public-Facing (17–22) | 3 External | 17–22 | External Infrastructure categories |
| Cloud Infrastructure (24–33) | 10 | 24–33 | Cloud Infrastructure categories |
| Active Directory / Identity (35–41) | 8 | 35–41 | Active Directory categories |
| Network Segmentation (43–47) | AI custom | 43–47 | Network Segmentation categories |
| SaaS Applications (49–53) | AI custom | 49–53 | SaaS Application categories |
| Web Applications (55–60) | 4 | 55–60 | Web Application categories |
| Mobile Applications (62–65) | 7 | 62–65 | Mobile Application categories |
| Source Code (67–70) | 13 | 67–70 | Source Code categories |
| Config Review | 5 | 4–6 and 9–11 (D15) | Device & Server categories |
| API | 11 | 19 and 32 (D19) | API categories |
| Social Eng., Red Team, Compliance | AI custom | no matrix rows: Count 1, Quantity 1, user edits | Engagement categories |

### 4.3 Rules
- An annexure is included only if its Cover Page type = Yes (case-insensitive; `y`, `yes`, `✓` accepted).
- Sum of 0 (or no row above 0) on an included annexure → D16 (amber, confirm).
- The user can override Count, Quantity and Frequency; the override wins.
- Count labels live in `annexure_map.json`, are pre-filled on the review screen, and are editable.
- Count and Quantity are independent: editing one does not change the other.

---

## 5. User flow

```
Home page ── Tab A: Manual (today's form, unchanged)
          └─ Tab B: Upload Excel
                 │
                 ▼
 1. Upload .xlsx ─► parse ─► not our sheet / invalid? ─► plain message, stay on page
                 │
                 ▼
 2. PREVIEW 1 — ANNEXURES (everything editable)
      a. Client details from Cover Page; reference no. (D21)
      b. Annexure list: ☑ include · title · source (Template / AI) · Count (D30/D31) · Frequency · order
         – amber rows when the block sum is 0 (D16, D33): confirm / edit / remove
      c. Manual toggles for the 5 annexures with no Excel row
      d. For AI annexures: what the Excel supplied and what is missing; sector (D20);
         basic AI form pre-filled; user completes the rest
                 │  [Next]
                 ▼
 3. PREVIEW 2 — CONCLUSION & PRICING (editable; same table that goes into the Word file)
      Sr. No · Assessment Type · Quantity (sum, D32) · Frequency (shared with Preview 1, D34) · Cost Per Annexure (empty)
      – type prices now, or leave blank and fill them in Word after generation
      – remove row, edit text; External/Internal Network PT appear as two rows
                 │
                 ▼
 4. Generate (disabled until amber rows resolved) ─► .docx download
```

Count appears only in Preview 1 and Quantity only in Preview 2; Frequency is a single shared field.

Draft state lives on the server for **60 minutes** (random id), then is deleted.

---

## 6. What changes inside the Word document

| Area | Change |
|------|--------|
| Cover | Client name, reference no. (D21), date from Excel; logo optional upload. |
| Annexure `Count:` line | Rewritten to `Count: N <label>` (D31, e.g. `Count: 6 Web Application categories`) in **every** annexure, where N is the number of filled rows (D30). Today the template has five inconsistent forms: blank in 3/4/5/13, `Environment(s): 01…` in 1, `1` in 8, `Cloud Azure & AWS 2` in 10, and a **stale `API = 2350`** in 11 — all replaced. |
| `Frequency:` line | Same label/value treatment (bold label, regular value) per spec §8, so the two lines look identical. |
| Contents | Auto-refreshed; includes every annexure; numbers match body. |
| Conclusion & Pricing | **The table already exists** (inside a Word content control) and the generator already rebuilds its rows (`_update_conclusion_pricing_table`). Today it uses **hard-coded dummy values** (`1 Web Application`, `Three Time`, …) and the last column is titled **Duration**. We replace the dummies with real data, rename the column **Cost Per Annexure**, and apply §7. |
| AI annexures | Same builder, Corbel/token formatting, now fed Excel context (§8). |

Not changed in Track A: cover layout, header/footer artwork, Terms & Conditions text, signature block, static annexure body text.

---

## 7. Pricing table specification (Conclusion & Pricing) ✅ structure, 🟡 look until Design Sample Gate

Source of structure: `CyberSmithSECURE Annexures 2.docx`, the table before the payment terms. Source of look: **the Formatting Specification** (§8 Conclusion & Pricing, §9 table rules).

**Columns:** `Sr. No.` · `Assessment Type` · `Quantity` · `Frequency` · `Cost Per Annexure`

| Property | Rule |
|----------|------|
| Rows | One per included annexure line, in annexure order; External/Internal Network PT are two rows (D11); then one **Total** row |
| Table | Native Word table, fixed layout, centred, **within text width**, header row **repeats** on each page, rows **cannot split** |
| Column widths (starting point, tuned after render) | Sr. No. 8% · Assessment Type 36% · Quantity 20% · Frequency 16% · Cost 20%; labels must not break mid-word |
| Header row | fill `#7030A0`, text `#FFFFFF`, **Corbel 10 pt bold**, centred (horizontal + vertical) |
| Body rows | **Corbel 10 pt**, `#404040`; white with optional alternating `#F2F2F2`; vertical centre |
| Alignment | Sr. No. centre · Assessment Type left · Quantity centre · Frequency centre · Cost right |
| Borders | 0.75 pt grey (`#7F7F7F` as in source), all cells |
| Cell padding | 0.08 in vertical, 0.10 in horizontal |
| Total row | Bold, `TableTotal` style (`#7030A0`), fill `#F2F2F2`; "Total" in first cell, remaining four cells **merged**, text `₹ /- * + 18% GST` when prices are empty |
| With prices | `₹ 1,25,000 /- * + 18% GST` (Indian digit grouping, sum of entered prices) |
| Empty prices | Cost cells left empty and editable in Word afterwards (D12) |
| Keep together | Heading, intro sentence, HSN/SAC line and table header stay on one page |
| Quantity text | The **sum as a number only**, e.g. `23` (D32); no unit text. The old `Android - 6, IoS - 3` text is no longer produced. |
| Frequency text | `One Time` by default |

Header/footer, TOC, page numbers: untouched in Track A; rebuilt in Track B.

---

## 8. AI annexures for unknown types

**Sent to the AI** (all optional except title + scope): client name, sector (D20), regulatory driver, dates → duration, the annexure's Count and Quantity (D30), that type's answers from General Scoping Questions and the matching detail sheet, Rules of Engagement (window, exclusions, compliance), and the pricing total when present (D18). Fixed rules come from `master_bd_prompt.txt` plus the "annexure only, no T&C/pricing" user message (unchanged).

**User still fills** only what the Excel cannot give: engagement summary, tier, payment model, special constraints, and a title/scope (≥ 20 chars) per AI annexure. Everything the Excel answered arrives pre-filled.

**Quality upgrades in this feature**
1. One automatic **retry** when the AI output misses required sections.
2. Stronger validation: annexure heading matches the slot number; `Frequency:` and `Count:` present; **`Count:` must equal the Count from Preview 1** and the pricing Quantity must equal Preview 2 (the AI annexure is told both numbers).
3. AI output formatted only through the tokens module (Corbel 11 pt, headings purple `#7030A0`, per spec §5/§8).
4. Rate limit on `/upload-excel` and `/generate` (AI costs money).

---

## 9. Formatting rules (from the Specification; summary, spec prevails) ✅

| Element | Font | Size | Colour | Style |
|---------|------|------|--------|-------|
| Body text | Corbel | 11 pt | `#595959` (the source value; spec allows `#404040` or `#595959`, used consistently) | regular, left-aligned, 6 pt after |
| Section heading (`Heading 2`) | Corbel | 16 pt | `#7030A0` | bold |
| Major heading (`Heading 1`, e.g. Conclusion & Pricing) | Corbel | 18 pt | `#7030A0` | bold |
| Annexure title | Corbel | 16 pt | `#7030A0` | bold italic, "Annexure N: Title" |
| Sub-labels (Sub-Activity, Methodologies, Deliverables…) | Corbel | 13–14 pt | `#7030A0` | bold |
| Lists | Corbel | 10.5–11 pt | body colour | real Word lists, 0.25 in hanging indent; glyphs ● then ○ |
| Table header / body / total | Corbel | 10 / 10 / 10.5 pt | white / `#404040` / `#7030A0` | see §7 |
| Page | A4 | — | — | spec: 1 in margins (current template differs; resolved in Track B) |
| Pagination | — | — | — | heading keep-with-next; no blank-paragraph spacing; no blank pages |

---

## 10. Technical design (how, not code)

| File | Responsibility |
|------|----------------|
| `proposal_tokens.py` | **Single source of every font, colour, size, spacing and table metric** from the spec |
| `proposal_excel.py` | Open workbook (read-only, values only), validate anchors, extract Cover + Matrix + context sheets |
| `annexure_map.json` | Mappings (§4), unit labels, row lists — editable without code |
| `proposal_pricing.py` | Pricing rows, Indian currency format, update of the existing table in its content control |
| `proposal_lint.py` | Formatting lint + numbering/TOC/blank-page checks run on each output |
| `css_proposal_generator.py` | Small additions only: new endpoints, `Count:`/`Frequency:` rewrite, hook to pricing |
| `tests/` | Parser, mapping, pricing, lint, generation regression |

**Endpoints:** `POST /proposal-generator/upload-excel/` (review JSON + draft id) · `GET /proposal-generator/draft/{id}` · `POST /proposal-generator/generate/` extended (manual flow keeps working).

**Excel safety:** `.xlsx` only (reject `.xlsm`/`.xls`), max 5 MB, read-only, no macros/external links, expected sheet names and header anchors (e.g. `Quantity` in `D2`) must match or the upload is rejected ("This is not the CyberSmithSECURE scoping sheet v2.0").

**Data handling:** uploads and drafts are client-confidential → temp storage, auto-deleted after 60 minutes, never logged.

**No-blank-page check:** (1) structural lint on the `.docx`: no consecutive page-break paragraphs, no run of empty paragraphs, no trailing break; (2) **LibreOffice (D29)** renders the `.docx` to PDF and any page with almost no text is flagged, and the proposal is blocked or warned (decision at Phase 5). The PDF is also offered as a download. LibreOffice is installed through the Railway build (nixpacks apt package) and used headless with a timeout.

**Caveat:** Corbel is a Microsoft font. The template embeds Corbel, but LibreOffice may substitute another font and paginate slightly differently from Word. The PDF is therefore a strong blank-page and layout check and a proof, **but Word remains the final authority** for page numbers in the Contents; the Design Sample Gate includes opening one output in Word.

---

## 11. Phases (Track A: the feature)

| Phase | Deliverable | Done when |
|-------|-------------|-----------|
| 0 | Sign-off on this plan + a filled sample Excel | §15 closed |
| 1 | **First update `css_scoping_testset/` (reference parser + expected JSONs) to D30–D33**, then `proposal_tokens.py` + `proposal_lint.py` + Excel parser + mapping file + tests | Parser output equals the expected JSONs for all 135 fixtures (SaaS 3/12 and Web 6/23 verified); lint runs on today's output |
| 2 | Review screen + draft storage | Upload → editable list → pricing preview works end to end |
| 3 | `Count:`/`Frequency:` rewrite + pricing table in the docx | Generated Word passes lint; **Design Sample Gate**: you approve the table against the spec |
| 4 | AI annexures from Excel context + retry + stronger validation | A mix of known + custom annexures generates, numbering correct, no blank pages |
| 5 | Rate limiting, upload hardening, **LibreOffice PDF render + blank-page check**, docs, Railway deploy | Smoke test on Railway passes; PDF produced; blank-page check works |

Each phase ends with passing tests (existing 55 + new) and your look at one generated `.docx`.

---

## 12. Track B — rebuild the Base Template to the full Specification (required, after Track A)

The spec describes `Annexures 2.docx` (11 annexures, 1 section, 1 in margins). Our Base Template differs (13 annexures, 5 sections, margins 1.0/0.75/2.4/2.15 cm, Normal style undefined, direct formatting everywhere). Track B brings it to the spec:

- Named style system (spec §5), real Heading 1/2/3, real bullet/number lists, no typed bullets.
- Header/footer per spec §6/§7 (logo, ribbon, contact strip, real page-number field, `Different First Page`), no clipping.
- Automatic TOC from heading styles; nothing typed manually.
- Cover per spec §8; signature block as a two-column borderless table.
- Margins 1 in; image sizing/alt text; table rules (§9); pagination rules (§10).
- Fix the known source defects (spec §14), including the stale TOC and the blank page.

**Prerequisites from you for Track B** (the spec itself marks these "confirm with brand owner"):
- ❓ `cover_navy` exact hex; approved logo, header ribbon, footer band and cover illustration files.
- ❓ Each annexure on a **new page**, or continuous flow (spec §10 asks for one global choice; the "no blank pages" rule must hold either way).
- ❓ Page numbering in the footer: yes/no.
- ❓ Body text colour: `#404040` or `#595959`.
- ❓ Current signatory name and title (spec lists Smith Gonsalves / COO as "confirm").
- ❓ Approved currency/spacing format for charges (`₹ 8,500/hour` vs `₹ 8,500 / hour`).

---

## 13. Test plan (summary)

- Parser: blank sheet, placeholders treated as empty, Yes/No variants, missing sheet, wrong file, oversize file.
- Mapping: only-known, only-custom, mixed; External + Internal both selected (two pricing lines); zero quantity (amber, Generate disabled until resolved).
- Count/Quantity: SaaS example gives Count 3 / Quantity 12 and Web example Count 6 / Quantity 23; rows with 0 or blank are skipped; sum 0 gives the amber flag; Config (rows 4–6, 9–11) and API (rows 19, 32) use only their rows.
- Docx: `Count:` equals the Preview 1 Count and the pricing Quantity equals the Preview 2 Quantity; Frequency identical in annexure and table; numbering sequential; Contents lists every annexure; pricing table matches rows, blank prices stay blank, removed rows disappear; table within text width, header repeats.
- Lint: only Corbel and token colours, no typed bullets, no blank paragraphs as spacing, no blank pages.
- AI (mocked): retry path; failure after retry gives a clear error.
- Regression: all 55 existing tests still pass.

---

## 14. Risks

| Risk | Mitigation |
|------|------------|
| Real Excel files differ from the blank template | Need the sample (Phase 0); parser strict on anchors, lenient on values |
| The block sum adds unlike rows (e.g. Web: apps + modules + roles), so Quantity is a total of entries, not literally "applications" | Accepted by you (D30); Count and Quantity are both editable in the previews; the review screen shows the per-row breakdown so it can be sanity-checked |
| Test fixtures `css_scoping_testset/` still use the old primary-row rule and old Mobile text | Phase 1 updates the fixture generator's reference parser to D30–D33 and regenerates the expected JSONs before any parser code is written |
| Template and spec disagree today | Track A uses spec tokens for everything generated; Track B aligns the template; Design Sample Gate each phase |
| AI invents details | Retry + validation + review screen + human check of the .docx |
| Blank pages / wrong numbering | Hard acceptance rules (§0) + automated lint + PDF render check |
| LibreOffice renders differently from Word (font substitution) | Use it for blank-page detection and proof; confirm in Word at the Design Sample Gate |
| Public cost abuse | Auth + rate limit (Phase 5) |

---

## 15. Open items ❓ (to be asked one by one)

- ~~Q-A: which matrix rows count~~ — closed, see §4.2 / D30.
- **Q-B.** Filled sample Excel (Phase 0).
- **Q-C.** Track B prerequisites (§12 list).
- ~~Q-D: PDF renderer~~ — closed, D29.
- ~~Q-E: track order~~ — closed, D28.

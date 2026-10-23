# CyberSmithSECURE Proposal / Statement of Work
## End-to-End Word Formatting Specification for Proposal-Generation Automation

**Source reviewed:** `CyberSmithSECURE Annexures 2(1).docx`  
**Purpose:** Build a reusable, deterministic `.docx` template for generating client proposals and Statements of Work (SOWs).  
**Scope:** Cover page through client signature, including page setup, typography, color, headers, footers, images, tables, lists, pricing, terms, and pagination.

---

## 1. How to use this specification

This document distinguishes three kinds of values:

- **Verified in DOCX:** Read from the Word document's underlying properties or direct text formatting.
- **Observed visually:** Seen in the rendered 56-page document; exact numeric properties cannot always be proven from a page image.
- **Recommended template standard:** A deliberate production rule for the automation. It may regularize inconsistencies in the source rather than reproduce every defect.

Do not hard-code a value as “original” if it is only visually estimated. For an exact visual match, inspect the original Word file in Microsoft Word and compare exported PDFs.

## 2. Executive assessment of the source

The source is a 56-page A4 proposal/SOW. It includes a branded cover, client placeholder, contents page, company introduction, requirement understanding, team table, methodology and diagrams, multiple service annexures, Terms & Conditions, deviation charges, pricing, payment terms, and signature blocks.

### Verified DOCX properties

| Property | Observed source value |
|---|---|
| Page size | A4 portrait, approximately 8.268 × 11.693 in |
| Margins | 1.00 in top, bottom, left, and right |
| Header distance | Approximately 0.492 in from top |
| Footer distance | Approximately 0.492 in from bottom |
| Normal style | Times New Roman, 12 pt, line-spacing value 1.0 in style definition |
| Main body run font | Corbel is used extensively |
| Main body direct font sizes | Many runs inherit size; some direct sizes of 12, 14, 16 and 26 pt occur |
| Main section headings | Corbel, 16 pt, bold, purple `#7030A0` for the headings inspected |
| Body gray text | `#595959` is widely used |
| Other explicit text color | `#000000` appears in a small number of runs |
| Tables | 3 Word tables detected: team overview, deviation charges, pricing |
| Inline pictures | One inline shape detected; other visible graphics appear to use floating/anchored drawings or graphic objects |
| Sections | One Word section in the document structure |
| Rendered pages | 56 pages in the supplied rendering |

**Key technical implication:** Word styles and direct formatting are mixed. A paragraph can use the `Normal` style but have Corbel 16 pt purple formatting applied directly. For reliable automation, define a controlled style system and apply styles explicitly rather than relying on inherited defaults and scattered run formatting.

### Visually observed brand treatment

- White page background.
- CyberSmithSECURE logo at the upper-left of the running page design.
- Purple accent bars and a purple gradient/ribbon element near the top.
- A thin gray rule in the header region.
- Bottom contact strip with website, phone, and support email.
- Thin gray rule above the contact strip and a purple branded footer band.
- Purple section headings; body copy is predominantly dark gray or black.
- The cover uses a large central illustration and a large dark-blue date/title treatment.

The exact logo dimensions, gradient stops, and all graphic anchor coordinates should be taken from the source assets/Word XML or approved brand files—not guessed from a screenshot.

---

## 3. Recommended global document setup

Use these as the default settings for newly generated proposals unless the organization approves a different brand standard.

| Setting | Automation standard |
|---|---|
| Paper | A4, portrait |
| Page width × height | 21.0 × 29.7 cm (8.27 × 11.69 in) |
| Top / bottom margins | 1.00 in / 1.00 in |
| Left / right margins | 1.00 in / 1.00 in |
| Header / footer distance | 0.49 in from page edge, then validate against printable area |
| Main content width | 6.27 in (A4 width minus 2 × 1 in margins) |
| Body font | **Corbel, 10.5–11 pt** recommended for dense proposal copy; use 11 pt as default |
| Body font color | `#404040` or `#595959`; use one consistently |
| Paragraph alignment | Left for normal body copy; justified only for long formal paragraphs if it does not create rivers of whitespace |
| Body line spacing | 1.08–1.15; avoid excessive paragraph gaps |
| Body paragraph spacing | 0 pt before, 6 pt after |
| Primary brand purple | `#7030A0` (verified in section-heading runs) |
| Secondary blue | Source cover visually uses dark navy; establish an approved exact brand token before production |
| Background | White |
| Page numbering | Footer, right-aligned or integrated into a consistent footer line; add if approved (the supplied visible footer does not clearly show a page number) |
| Output | DOCX plus PDF proof; compare page count and page-break integrity before delivery |

**Important:** If the goal is exact reproduction of the current file, preserve its A4 size, 1-inch margins, existing header/footer graphics, and source content. If the goal is a cleaner long-term automation template, use the recommended typography and pagination rules in this specification and treat them as a controlled redesign.

## 4. Brand and color tokens

Create centralized style tokens in the generator. Never sprinkle raw color values throughout template code.

| Token | Value | Use |
|---|---|---|
| `brand_purple` | `#7030A0` | Section headings, annexure headings, table header fill, accent lines |
| `body_gray` | `#595959` | Body copy where gray is part of the approved design |
| `body_black` | `#000000` | Body copy when strong contrast is required |
| `table_header_text` | `#FFFFFF` | Text on purple-filled table headers |
| `table_body_fill` | `#F2F2F2` | Optional light-gray alternating rows / total row background |
| `rule_gray` | `#B7B7B7` | Thin header/footer separators |
| `cover_navy` | **Confirm with brand owner** | Date and main cover title; visually appears dark navy |
| `white` | `#FFFFFF` | Page and negative-space backgrounds |

For accessibility and print, maintain strong contrast. Do not use color alone to communicate risk, acceptance, or payment status.

---

## 5. Typography system

Use named Word paragraph and character styles. Avoid formatting every run independently.

| Style name | Font | Size | Weight / style | Color | Alignment | Usage |
|---|---|---:|---|---|---|---|
| `CoverDate` | Corbel or approved brand font | 26–30 pt | Bold | Cover navy | Right | Date on page 1 |
| `CoverTitle` | Corbel | 24–28 pt | Bold | Cover navy | Center | “Statement of Work & Proposal” |
| `CoverForLabel` | Corbel | 18–22 pt | Bold | Cover navy | Center | “For” label |
| `ClientName` | Corbel | 22–26 pt | Bold | `#000000` or approved navy | Left | Client/company name on page 2 |
| `TOCTitle` | Corbel | 24–26 pt | Bold | Approved blue/purple | Left | Contents heading |
| `Heading1` | Corbel | 18 pt | Bold | `#7030A0` | Left | Major divisions such as Conclusion & Pricing |
| `Heading2` | Corbel | 16 pt | Bold | `#7030A0` | Left | INTRODUCTION, TEAM OVERVIEW, Terms & Conditions |
| `Heading3` | Corbel | 13–14 pt | Bold | `#7030A0` | Left | Sub-activity, Methodologies, Objectives, Deliverables |
| `AnnexureTitle` | Corbel | 16 pt | Bold Italic | `#7030A0` | Left | “Annexure N: …” |
| `BodyText` | Corbel | 11 pt | Regular | `#404040` or approved `#595959` | Left | Main paragraphs |
| `BodyTextJustified` | Corbel | 11 pt | Regular | `#404040` | Justified | Only formal paragraphs where justified text is visually clean |
| `ListBullet` | Corbel | 10.5–11 pt | Regular | Body color | Left | Main bullet list |
| `ListLevel2` | Corbel | 10.5–11 pt | Regular | Body color | Left | Nested list |
| `TableHeader` | Corbel | 10–11 pt | Bold | White | Center | Header cells on purple fill |
| `TableBody` | Corbel | 9.5–10.5 pt | Regular | `#404040` | Left/Center by column | Team, pricing, and charge tables |
| `TableTotal` | Corbel | 10.5–11 pt | Bold | `#7030A0` or dark text | Right/Center | Grand total row |
| `TermsText` | Corbel | 10.5–11 pt | Regular | `#595959` | Left | Legal/commercial terms |
| `SignatureLabel` | Corbel | 10.5–11 pt | Bold | `#595959` | Left | Company, By, Name, Title, Date |
| `FooterText` | Corbel | 8–9 pt | Regular | `#595959` | Left/Center/Right by field | Website, phone, support email |
| `Caption` | Corbel | 8.5–9 pt | Regular or Italic | `#595959` | Center | Optional figure caption |

### Style behavior

- Set paragraph styles and character styles in the DOCX template, then apply them to generated content.
- Use true heading styles (`Heading 1`, `Heading 2`, etc.) so Word navigation, accessibility, and automatic TOC generation work.
- Do not create headings by making a normal paragraph bold and purple only.
- Use paragraph `keep_with_next` for headings and `keep_together` for short blocks.
- Use widow/orphan control for body paragraphs.
- Avoid manual spaces to align text. Use tables, tab stops, or paragraph alignment.
- Keep capitalization consistent: use sentence case for subsection labels unless the approved brand template intentionally uses all caps.

---

## 6. Header specification (all interior pages)

The source shows a branded running header on the content pages.

### Elements and placement

1. **Logo:** CyberSmithSECURE logo at upper-left.
2. **Top rule:** Thin gray horizontal rule extending across the content area to the right of the logo.
3. **Purple accent:** Purple horizontal bar below/near the logo area on the left.
4. **Purple ribbon/tab:** Purple-to-violet gradient or solid branded shape toward the upper-right, carrying the document label where visible.
5. **Running label:** “Statement of Work” in a consistent small font, positioned consistently and never clipped.
6. **Safe zone:** No body text or images may overlap the header.

### Implementation rules

- Build the header from fixed-size image/vector elements and tables/shapes; do not recreate the logo using text.
- Use a header table or anchored graphics with a known layout grid. Avoid floating shapes whose positions depend on preceding paragraph text.
- Keep the logo aspect ratio locked; never stretch it to fit.
- Use `Different First Page` so the cover can have its own header treatment, while interior pages use the standard running header.
- If the brand ribbon is an image, embed a high-resolution transparent PNG or vector where supported. Do not use a low-resolution screenshot of the whole page as a header.
- Define a fallback if the ribbon artwork is unavailable: white background, logo left, thin gray rule, purple accent line, and a right-aligned “Statement of Work” label.
- Header artwork should not be repeated as a body image.

**Source observation:** The purple ribbon is partially clipped in some page renderings, and the “Statement of Work” label does not appear consistently readable. Correct this in the template rather than preserving a clipping defect.

---

## 7. Footer specification (all interior pages)

### Elements and order

Use one footer row with three fields:

| Left | Center | Right |
|---|---|---|
| Website icon + `cybersmithsecure.com` | Phone icon + `+91 93073 38941` | Envelope icon + `support@cybersmithsecure.com` |

Above the row, add a thin gray rule. Below it, retain the approved purple footer band and angled accent graphic if the official brand artwork requires it.

### Formatting

- Font: Corbel, 8–9 pt.
- Text color: `#595959` or approved dark gray.
- Icons: one consistent visual style and baseline; use official vector assets if available.
- Vertical alignment: center.
- Left/center/right alignment: use table-cell alignment, not spaces.
- Footer content must fit inside printable width and must not wrap to a second line.
- Keep enough separation from body content; body paragraphs must not enter the footer area.
- Use a real page-number field if page numbering is required. Never type page numbers manually.

**Automation validation:** Check the footer on a short page, a full page, a page with a large table, and the last page. It must not overlap the body or signature block.

---

## 8. Page-by-page and section-by-section template map

The page numbers below refer to the supplied 56-page rendering. Actual pagination may shift when client-specific content changes; the automation should use section types and semantic headings, not hard-coded page numbers.

### Page 1 — Cover page

**Visible content:** Logo/brand header treatment, date `15-09-2025`, central illustration, “Statement of Work & Proposal”, “For”, footer contact strip.

**Recommended layout**
- Use a dedicated first-page section or `Different First Page` header/footer.
- Logo: upper-left, proportional scaling, consistent inset from page edges.
- Date: upper-right; bold, approximately 26–30 pt, dark navy; right-aligned.
- Main illustration: centered horizontally, preserve aspect ratio, use a high-resolution source asset; do not let it collide with the title.
- Main title: centered, bold, approximately 24–28 pt, dark navy; split across lines intentionally.
- “For”: centered, smaller than title, with clear spacing.
- If a client name is shown on the cover, insert it as a named content field with a defined style; do not leave an empty “For” label.
- Footer: same approved contact strip and purple band as the source, aligned to the page grid.

**Dynamic fields**
- `proposal_date`
- `proposal_title`
- `client_legal_name`
- `cover_illustration`
- Optional `proposal_id` / `confidentiality_label` if the organization wants them

**Page-break rule:** Cover must end with a deliberate page break; never let content from page 2 flow onto it.

### Page 2 — Client name and contents

**Visible content:** “For”, large “Company name”, “Contents”, dotted leader entries.

**Recommended layout**
- Client label: 11–12 pt regular.
- Client name: 22–26 pt bold; left-aligned.
- TOC heading: 24–26 pt bold, approved blue/purple.
- TOC: generate automatically from Word heading styles; use dotted tab leaders and right-aligned page numbers.
- Main section entries: 10–11 pt bold or semibold, consistent capitalization.
- Annexure entries: 10–11 pt italic or regular according to one consistent hierarchy; do not make every annexure title bold italic unless approved.
- Use a real automatic TOC, not manually typed dot strings or page numbers.
- Keep TOC indented to the content grid and align all page numbers to one right-side tab stop.

**Required automation behavior**
- Update TOC fields after final pagination.
- Include every annexure present in the body.
- Flag any missing, duplicate, or misnumbered annexure.
- Do not show stale page numbers after generating a client-specific proposal.

### Pages 3–4 — Introduction, company overview, client requirement, team overview

**Visible content:** Company introduction, service-area graphic, “Our Understanding of Your Requirement”, workflow diagram, “Team Overview” table header.

**Formatting**
- Major section headings: `Heading2`, 16 pt bold, purple `#7030A0`, left-aligned.
- Body: Corbel, 11 pt recommended, body gray, left-aligned; 6 pt after paragraph.
- Long company-description paragraphs may be justified only if spacing remains natural.
- Infographics: centered within content width, keep aspect ratio, no text wrapping through the image.
- Team overview: Word table, not tab-delimited text. Repeat header row if it crosses pages; prevent a row from splitting across pages.
- Team table columns: Sr. No.; Level; Profile; Experience; Role. Center short values; left-align longer descriptions.
- Keep the “Team Overview” heading with at least the first row of its table.

**Source issue:** Page 4 begins with large unused vertical space and the team table continues to the next page. Use controlled spacing and table pagination instead of large blank paragraphs.

### Pages 5–9 — Methodology and vulnerability assessment methodology

**Visible content:** Team table continuation, methodology description, standards/framework bullet list, methodology graphics, vulnerability-assessment workflow, numbered phases and tools.

**Formatting**
- Main headings: 16 pt bold purple.
- Phase headings such as “1. PLANNING PHASE”, “2. DISCOVERY”, “3. ATTACK”, “4. REPORTING”: 12–13 pt bold, body/brand color, left-aligned.
- Body paragraphs: 11 pt.
- Lists: 10.5–11 pt; 0.25 in hanging indent recommended; 3–4 pt after each item.
- Numbered process phases should use a real numbered-list style.
- Framework/tool lists should use real bullets.
- Diagrams: centered, preserve aspect ratio, use “In Line with Text” or a stable anchored container; keep each diagram near its explanatory paragraph.
- Avoid splitting a short heading from its first paragraph or placing a heading alone at the bottom of a page.

**Source issues to prevent:** A diagram crowds a heading on page 6, and some content/graphics are close to the footer. The template should reserve a fixed header/footer safe zone and keep heading-plus-diagram blocks together when practical.

### Pages 10–53 — Service annexures and technical descriptions

These pages contain Annexures 1–11, including network penetration testing, web application penetration testing, network configuration review, insider threat assessment, mobile application testing, Active Directory security assessment, SOC simulation, cloud penetration testing, API penetration testing, offensive/defensive security training, and source code review.

**Reusable annexure structure**

1. `Annexure N: [Service Name]` — 16 pt, bold italic, purple.
2. `Sub-Activity` — 13–14 pt bold purple.
3. Activity description — 11 pt body.
4. `Tools` / `Methodologies` — 13–14 pt bold purple.
5. Tool/methodology bullets — 10.5–11 pt.
6. `Vulnerability Assessment` or service-specific scope — 13–14 pt bold purple.
7. `Frequency` — bold label followed by regular value.
8. `Count` — bold label and explicit quantity field(s).
9. `Deliverables` — 13–14 pt bold purple, followed by structured list.
10. `Expectation` — 13–14 pt bold purple where present.
11. `Objectives` — 13–14 pt bold purple where present.
12. `Requirements for Starting the Assessment` / `Requirements` — 13–14 pt bold purple.
13. `Comments` / service-specific notes — 13–14 pt bold purple where present.
14. Relevant diagrams and images placed beside or immediately after the content they explain.

Not every annexure has every field. Do not fabricate absent scope, frequency, counts, deliverables, or commitments. The generator should support optional sections and omit a heading when the section is genuinely not applicable, unless the organization explicitly wants a blank field for manual completion.

**List rules**
- Use real Word bullet/numbering definitions; never type bullet symbols and spaces to fake indentation.
- Level 1: 0.25 in left indent, 0.20–0.25 in hanging indent.
- Level 2: 0.50 in left indent, 0.20–0.25 in hanging indent.
- Level 3: 0.75 in left indent.
- Use consistent bullet glyphs by level: round bullet, open-circle bullet, then dash or numbered sub-list.
- Keep wrapped lines aligned with the first character of the list text, not under the bullet glyph.
- Keep one consistent spacing scheme; do not insert empty paragraphs between every bullet.

**Images and diagrams**
- Use source artwork as independent images, not page screenshots.
- Center diagrams by default; left-align only when the image is intended to accompany a text column.
- Set maximum width to the usable text width (about 6.27 in under the stated margins), unless a full-width bleed design is specifically approved.
- Set image height proportionally from intrinsic width/height; never stretch.
- Add descriptive alternative text.
- Keep diagrams with their captions and related heading.
- Use `In Line with Text` as the safest default for generated documents. Use floating positioning only for a controlled design component that has been tested across Word and PDF renderers.
- If an image cannot fit in the remaining page space, move the image and its caption together to the next page rather than shrinking it below a legible size.
- Avoid a page containing only a small image or a heading because of uncontrolled anchors or manual blank paragraphs.

**Annexure boundaries**
- Start each major annexure on a new page if that is the approved corporate rule. The supplied file does not do this consistently; it contains continuous-flow annexures and a blank-looking page.
- Use an explicit page break before each annexure title only if the organization chooses a new-page-per-annexure standard. Do not insert many blank paragraphs to force a new page.
- If an annexure is long, let it flow naturally but prevent heading or list fragmentation.
- For long tables, repeat header rows and prevent individual rows from splitting.

**Important source/TOC discrepancy**
- The visible contents page lists Annexures 1–5 and 7–10 but does not list Annexure 6 or Annexure 11.
- The body contains **Annexure 6: Active Directory Security Assessment** and **Annexure 11: Source Code Review**.
- The contents page's displayed page numbers do not align with the supplied 56-page rendering in at least the final section: “Conclusion & Pricing” is shown at page 48 in the TOC, while that section appears on page 55 of the supplied rendering.
- Generate the TOC automatically and validate all headings and page numbers after final pagination.

### Page 14 — Mostly blank page in the rendering

The supplied rendering shows a largely blank page with the running header and footer. This is a formatting defect unless the organization intentionally uses a blank divider page.

**Automation checks**
- Identify empty pages created by page breaks, floating objects, or paragraph pagination settings.
- Permit intentional blank divider pages only when marked explicitly in template metadata.
- Do not “fix” this by deleting content blindly; inspect page breaks and anchored images first.

### Pages 54–55 — Terms & Conditions, deviation charges, pricing

**Terms & Conditions**
- Heading: 16 pt bold purple.
- Body: 10.5–11 pt, gray/black.
- Main terms: consistent bullet list.
- Nested conditions: level-2 bullets with proper hanging indent.
- Avoid justified spacing on narrow bullet lines.
- Do not allow a single bullet item to split across pages if it is short; long terms can continue naturally.
- Keep each bullet's wrapped lines with the same text indent.

**Deviation charges table**
- Three columns: Sr No., Level, Deviation Charges.
- Header row: bold, centered, white on purple if matching the pricing-table theme; otherwise retain a neutral corporate header.
- Body: 10–10.5 pt, vertically centered.
- Numeric charges should be consistently formatted as `₹ 8,500/hour`, `₹ 6,500/hour`, `₹ 2,500/hour` (confirm approved spacing and currency format).
- Use consistent row padding and borders. Prevent rows from splitting.

**Conclusion & Pricing**
- Major heading: 18 pt bold purple (recommended controlled `Heading1`).
- Introductory sentence: 11 pt; italic/bold only if intentional and consistent.
- HSN/SAC line: bold label; code regular or semibold.
- Pricing table: 5 columns in the source: Sr. No., Assessment Type, Quantity, Frequency, Cost Per Annexure.
- Header row: purple fill `#7030A0`, white bold text, centered vertically and horizontally.
- Body: light gray/white fill, 9.5–10.5 pt; left-align service descriptions; center quantity and frequency; right-align monetary values.
- Use consistent column widths. Suggested proportions: Sr. No. 8%, Assessment Type 36%, Quantity 20%, Frequency 16%, Cost Per Annexure 20% (adjust after rendering).
- Avoid breaking words such as “Frequency” and “Annexure” across several lines; widen columns or shorten labels where approved.
- Total row: visually distinct, bold, with a merged cell for the total description if appropriate; monetary total aligned right.
- Make tax treatment explicit, e.g. base amount, GST percentage, GST amount, grand total. Preserve the source's business logic but do not invent missing amounts.
- Keep the pricing heading, introductory sentence, HSN/SAC line, and at least the table header together when possible.
- If pricing table moves to a new page, repeat its header row.

**Source issue:** The pricing table's column headers and content wrap awkwardly, and the total line is fragmented. Use fixed column widths, deliberate line breaks only where necessary, and render-time checks.

### Page 56 — Payment terms and signature/acceptance block

**Payment terms**
- Heading: 12–13 pt bold; the source visually uses a dark gray bold uppercase label.
- Bullets: 10.5–11 pt, consistent hanging indent.
- Preserve the actual agreed terms; do not change payment percentages, invoice triggers, or due dates automatically.

**Acceptance paragraph**
- Body: 10.5–11 pt.
- Bold only the company/client names if that is the approved style.
- Use a clear paragraph gap before the witness/acceptance statement.

**Signature block**
Use a two-column borderless table rather than spaces or tabs.

| Supplier column | Client column |
|---|---|
| CyberSmithSECURE Pvt. Ltd. | `[Client Legal Name]` |
| By: __________________________ | By: __________________________ |
| Name: Smith Gonsalves (source value; confirm current signatory) | Name: `[Authorized Signatory]` |
| Title: Chief Operating Officer (source value; confirm current title) | Title: `[Title]` |
| Date: `[Date]` | Date: `[Date]` |

**Signature formatting**
- Company names: 10.5–11 pt bold.
- Labels: 10.5–11 pt; use a consistent bold-label treatment.
- Two columns equal width and aligned to the document content grid.
- Use adequate space above the signature line for handwritten or electronic signatures.
- Keep each party's company, signature line, name, title, and date together.
- Do not allow the signature table to split across pages.
- Use placeholders for client legal name, authorized signatory, title, and date. Do not leave unlabeled empty cells.
- If digital signing is used, provide the organization-approved signature field or signing-platform tags.
- The signature block must remain the final substantive block unless a formal appendix follows.

---

## 9. Tables: global rules

The source contains three Word tables, and the rendered pricing/team tables demonstrate why table layout needs explicit rules.

### General table formatting

- Use native Word tables, not screenshots.
- Define exact column widths or stable proportional widths.
- Use consistent cell margins and vertical alignment.
- Set header row repeat-on-each-page for tables that can span pages.
- Prevent rows from splitting across pages where supported.
- Use a consistent border system: thin gray/black borders; avoid mixing border weights without a reason.
- Use shading only for headers, totals, or alternating rows; keep print legibility.
- Use paragraph styles inside cells rather than direct formatting alone.
- Align numbers and currency values consistently.
- Keep tables within the usable page width; do not let tables extend into margins.
- Avoid manual spaces and line breaks to simulate columns.
- Validate tables in both Word and PDF output.

### Suggested default table tokens

| Element | Standard |
|---|---|
| Header fill | `#7030A0` |
| Header text | `#FFFFFF`, bold |
| Body fill | White; optional alternating `#F2F2F2` |
| Border | 0.5–0.75 pt gray/black |
| Cell padding | 0.06–0.10 in vertical; 0.08–0.10 in horizontal |
| Header font | Corbel, 10 pt bold |
| Body font | Corbel, 9.5–10.5 pt |
| Vertical alignment | Center |
| Text columns | Left-aligned |
| Short labels / serial numbers | Center-aligned |
| Currency | Right-aligned |

---

## 10. Page-break and pagination rules

Pagination should be rule-based, not achieved by adding blank paragraphs.

1. Cover page always ends with a page break.
2. Contents page begins on a fresh page and ends with a deliberate page break.
3. Major sections may start on a new page if this is the chosen corporate standard.
4. Annexure starts should be consistent; choose either “each annexure starts on a new page” or “continuous-flow annexures” and apply globally.
5. Heading paragraphs use `keep_with_next`.
6. Short tables and signature blocks use `keep_together` / no-split behavior where possible.
7. Tables that span pages repeat the header row.
8. Images stay with the associated heading/caption where practical.
9. Avoid orphan headings, one-line paragraph carryovers, and single bullets stranded at page top/bottom.
10. Reserve footer space. No content may overlap the footer rule, contact details, or purple band.
11. Do not insert a manual page number or manually type TOC page numbers.
12. Update all Word fields before final export.
13. After PDF conversion, compare the output page count and inspect pages with large blank regions.
14. If client-specific content changes page count, allow natural repagination; do not force every generated proposal to match the source's 56 pages.

---

## 11. Image and graphic handling

### Asset requirements

- Use the official logo file and approved source illustrations.
- Prefer SVG/EMF for logos and diagrams when supported by the target Word/PDF workflow; otherwise use high-resolution PNG.
- Store assets by role, for example: `brand/logo`, `brand/header-ribbon`, `brand/footer-band`, `cover/illustration`, `methodology/diagrams`.
- Preserve intrinsic aspect ratio and transparency.
- Add alt text to informative graphics.
- Do not extract a whole rendered page and reinsert it as an image.

### Layout requirements

- Default to centered, in-line images for predictable pagination.
- Maximum image width: content width, approximately 6.27 in.
- Use a minimum practical size that keeps embedded labels legible.
- Avoid image/text overlap and header/footer overlap.
- Do not leave a heading separated from its diagram.
- For side-by-side graphics, use a borderless table/grid with fixed cells rather than free-floating images.
- Treat logos in header/footer as brand assets, not as content illustrations.
- Keep all graphical elements inside page margins unless a deliberate bleed design has been approved.

---

## 12. Data fields and placeholders for the generator

Create a single structured data model. Never search-and-replace arbitrary text across the document as the primary generation method.

### Document-level fields

- `proposal_date`
- `proposal_id`
- `proposal_title`
- `client_legal_name`
- `client_display_name`
- `client_address`
- `client_contact_name`
- `client_contact_email`
- `client_contact_phone`
- `prepared_by`
- `validity_date` (only if the organization uses validity periods)
- `confidentiality_label` (optional)
- `currency`
- `gst_rate`
- `hsn_sac_code`

### Annexure fields

Each annexure should be an object containing optional fields such as:

- `annexure_number`
- `annexure_title`
- `sub_activity`
- `scope_description`
- `tools`
- `methodologies`
- `assessment_steps`
- `frequency`
- `quantity_or_count`
- `deliverables`
- `expectations`
- `objectives`
- `requirements`
- `comments`
- `images`
- `price`
- `tax_category`

Only populate fields supported by the approved proposal content. Empty optional fields should be omitted cleanly or visibly flagged for review; do not produce empty headings throughout the proposal.

### Commercial fields

- `pricing_items[]`
- `quantity`
- `frequency`
- `unit_price`
- `line_total`
- `subtotal`
- `gst_amount`
- `grand_total`
- `payment_milestones[]`
- `deviation_charges[]`

### Signature fields

- `supplier_legal_name`
- `supplier_signatory_name`
- `supplier_signatory_title`
- `supplier_signature_date`
- `client_legal_name`
- `client_signatory_name`
- `client_signatory_title`
- `client_signature_date`

Treat all client and pricing data as untrusted input: escape or sanitize values, preserve currency formats, and prevent user-provided strings from breaking the document structure.

---

## 13. Recommended generation architecture

1. **Approved template:** Maintain a version-controlled `.dotx` or `.docx` with styles, section settings, headers, footers, table styles, and placeholders.
2. **Structured content:** Store proposal information in JSON or a validated application model; separate content from formatting.
3. **Content assembly:** Insert sections and annexures using named styles and reusable components.
4. **Pagination:** Use explicit page breaks only for intentional section boundaries.
5. **Fields:** Update TOC and page-number fields using a Word-compatible field-update workflow.
6. **Render:** Convert the generated DOCX to PDF using the same office engine used in production.
7. **Automated checks:** Validate placeholders, headings, table bounds, missing fields, image dimensions, page count, empty pages, and signature block integrity.
8. **Visual QA:** Render PDF pages to images and inspect cover, contents, first body page, first page of each annexure, pricing, and signature page.
9. **Delivery:** Save both DOCX and PDF proof, with a version number and generation timestamp.
10. **Regression tests:** Keep a representative client dataset and compare the rendered result against a known-good baseline after template changes.

### Suggested automated validation checklist

- [ ] No unresolved placeholder such as `Company name`, `For`, or empty quantity.
- [ ] Cover date and client name are populated and styled correctly.
- [ ] Every annexure in the body is listed in the TOC.
- [ ] TOC page numbers are refreshed.
- [ ] Heading hierarchy is valid and consistent.
- [ ] No accidental blank pages.
- [ ] No clipped header ribbon or footer contact text.
- [ ] No image distortion or overlapping text.
- [ ] No table exceeds page margins.
- [ ] Pricing totals reconcile with line items and tax fields.
- [ ] Terms and payment text match approved wording.
- [ ] Signature block contains both parties and remains together.
- [ ] No heading is stranded at the bottom of a page.
- [ ] Footer does not overlap body content.
- [ ] DOCX and PDF page counts/renderings are reviewed.
- [ ] Final output contains no comments, tracked changes, or hidden draft text unless explicitly required.

---

## 14. Source-specific defects and risks to fix before automating

These are observations from the supplied file, not assumptions about what the organization legally intends.

1. **Mixed typography:** The document uses Corbel directly in many runs while the Normal style is Times New Roman 12 pt. This can cause inconsistent output when styles are reapplied or content is inserted.
2. **Heading hierarchy is inconsistent:** Many annexure titles are formatted as direct formatting in Normal paragraphs instead of a dedicated semantic heading style.
3. **Contents page is stale/incomplete:** The body contains Annexure 6 and Annexure 11, but the visible contents list omits them. The pricing page reference shown in the TOC does not match the supplied rendered page.
4. **Blank/underused page:** Page 14 appears largely blank in the supplied rendering. Investigate manual page breaks, anchored graphics, and paragraph pagination controls.
5. **Clipping in header ribbon:** The top-right running label/ribbon is partly cut off in rendered pages.
6. **Excessive whitespace:** Some pages have large vertical gaps, likely due to blank paragraphs, page-break placement, or floating objects.
7. **Table wrapping:** Team and pricing table content wraps awkwardly in places; pricing headings split across lines and the total row is fragmented.
8. **Pricing data needs validation:** The displayed pricing table has blank cost cells and an incomplete/fragmented total presentation in the rendered source. Do not infer or invent missing prices.
9. **Signature block is space-aligned:** Replace the two-column text alignment with a borderless table to prevent drift when the client name or font changes.
10. **Source pagination is fragile:** The header/footer and large images may be positioned with anchored graphics. Their position can shift between Word versions and PDF conversion engines.
11. **Annexure numbering in TOC/body differs:** Ensure numbering and titles come from one source of truth.
12. **Content quality is separate from formatting:** Typos, grammar, repeated passages, and technical claims exist in the source. This specification does not silently rewrite them; content approval should be a separate review workflow.

---

## 15. Acceptance criteria for a production-ready template

The template is ready for use only when:

- A proposal can be generated from structured data without manually adjusting individual paragraphs.
- The cover, contents, body, annexures, terms, pricing, and signatures all use approved styles.
- The TOC is complete and accurate after fields update.
- The page header/footer looks consistent across all pages, including the first and last page.
- No unintentional blank pages, clipped text, distorted images, or margin overflow occurs.
- Tables stay within the page width and remain legible.
- Pricing, tax, and totals are validated independently of visual formatting.
- Both signature columns remain aligned for short and long client names.
- The generated DOCX and PDF are visually reviewed in the actual production environment.
- The same data produces stable formatting across repeated runs, subject to normal page-flow changes caused by different content lengths.

---

## 16. Final recommendation

Use the supplied document as a **brand and content reference**, not as a formatting engine to copy blindly. Preserve the recognizable CyberSmithSECURE identity—logo, purple accents, contact footer, purple section headings, technical diagrams, and commercial closeout—but rebuild the structure around named Word styles, stable header/footer layouts, real lists, real tables, automatic TOC fields, and explicit page-break rules.

For the automation, the most important principle is: **content should change; the formatting system should not.** Any exception should be an explicit template rule, not an accidental side effect of paragraph spacing, manual tabs, floating objects, or inherited Word formatting.

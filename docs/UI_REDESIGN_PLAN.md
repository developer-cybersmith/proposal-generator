# UI redesign plan (styling only): CyberSmithSECURE Proposal Studio

Status: **implemented** with the decisions made on 2026-10-25: direction A (Clean Studio), fonts Bricolage Grotesque + Hanken Grotesk + JetBrains Mono, light theme with an optional dark theme, and the non-functional extras (live summary, "?" tooltips, build panel). What was built and how it is guarded: `docs/PROPOSAL_GENERATOR_HANDOFF.md` section 18.

## 1. The rule that governs everything
**No functionality changes.** The JavaScript, the Python endpoints and the generated Word file stay as they are. The work is CSS, markup and wording around the existing controls.
How that is enforced:
- Every element `id`, `name`, form field and button label that scripts or tests use stays. New wrappers and classes are added around them.
- The inline JavaScript of each page is hashed before the work; the hash must be the same afterwards (a script checks it).
- The 233 tests, the jsdom UI test and the real-browser tests (local and live) must pass unchanged.
- Before/after screenshots of every screen go with each phase.

## 2. What is wrong with the current UI (from screenshots of every screen)
| # | Finding | Why it hurts |
|---|---|---|
| 1 | One dark navy slab on a near-black page, one saturated purple (`#8a2be2`) used for the title, links, buttons and headings | Everything has the same weight, so nothing leads the eye; purple on navy has low contrast for small text; it looks like a default dark template |
| 2 | Browser-default font (Arial/Helvetica) everywhere, centred bold title | No personality; the title is the only hierarchy |
| 3 | Browser-default "Choose File" buttons and checkboxes | Looks unfinished; the file input is the first thing a new user touches |
| 4 | Preview 1 is a 2000 px wall of inputs: client fields, a table of annexures, a second table, an AI panel, an error box | A new user cannot tell what is required, what is optional, or what to do next |
| 5 | The stepper (1 Upload / 2 Annexures / 3 Conclusion & Pricing / 4 Final preview) is four small pills under the title | Names are insider terms; no sense of "where am I, what is left" |
| 6 | Jargon: "Count", "Frequency", "AI-written", "Template", "Reference number" with no explanation | A first-time user has to guess or ask |
| 7 | Validation shows up only as one red box at the bottom ("Before you can generate") | The user scrolls past the field that is wrong |
| 8 | The wait while a proposal builds (up to ~60 s with AI) is a line of text | Feels frozen; users click again |
| 9 | Manual form and Excel flow look like two different tools; the login page is a third | No product identity |
| 10 | Layout is a fixed narrow column; tables use tiny inputs; nothing is designed for small screens | Cramped on laptops, broken on phones/tablets |

## 3. Design principles (what "a user who is new should not need to click around" means here)
1. **One obvious next action** per screen (a single primary button, always in the same place, sticky at the bottom).
2. **Say where I am**: a left-hand progress rail with plain names ("Upload sheet", "Choose annexures", "Prices", "Review & download"), "Step 2 of 4", and what is already done.
3. **Plain language first, jargon with a tooltip**: "Count" becomes "Items in scope" with a one-line "?" explanation; "Frequency" has examples.
4. **Show consequences**: a live summary card (client, number of annexures, AI-written ones, estimated pages) so the user sees what they are building.
5. **Errors next to the field**, with a short summary at the top that links to each problem (the existing messages, shown better).
6. **Progressive disclosure**: the AI panel and the per-row Excel breakdown stay collapsed until needed.
7. **A build that feels alive**: named stages ("Reading the sheet", "Writing custom annexures with AI", "Laying out the pages", "Checking every page") instead of a frozen button.
8. **Accessible by default**: contrast 4.5:1 or better, visible focus, 44 px touch targets, keyboard-only use, reduced-motion respected, up/down buttons kept for reordering (no drag-only).

## 4. Three directions to choose from
All three keep the brand purple of the proposal documents (`#7030A0`) as the anchor, so the tool and its output look related.

**A. "Clean Studio" (recommended).** A light, calm workspace: warm off-white canvas, white cards with soft shadows, deep ink text, brand purple only for the primary action, active step and key numbers, a soft violet tint for selected rows. Feels like a modern SaaS product; easiest to read for long sessions and screenshots. A dark theme can be added from the same tokens.

**B. "Editorial Docs".** Paper-like: cream background, a serif for headings so the tool reads like the proposal itself, ink and purple, thin rules instead of boxes. Distinctive and elegant; slightly less "app-like".

**C. "Dark Ops Console".** Graphite/near-black surfaces with a violet glow, mono accents for numbers and IDs, an "ops console" look that suits a security company. Most striking in a demo; harder for long forms and bright rooms, and needs more care for contrast.

## 5. Typography (not the usual defaults)
Self-hosted WOFF2 files (no external font CDN, so it works offline and on Railway without a third-party request).
| Option | Headings | Interface text | Numbers / IDs | Character |
|---|---|---|---|---|
| 1 (pairs with A) | **Bricolage Grotesque** | **Hanken Grotesk** | **JetBrains Mono** | contemporary, a little quirky headlines, very readable UI |
| 2 (pairs with B) | **Fraunces** (soft serif) | **Instrument Sans** | **IBM Plex Mono** | warm, editorial, premium |
| 3 (pairs with C) | **Sora** or **Space Grotesk** alternative: **Unbounded** for the wordmark only | **Geist** | **Geist Mono** | technical, crisp |
The proposal document itself stays in Corbel; this only changes the tool.

## 6. End-to-end plan
| Phase | Work | Output | Check |
|---|---|---|---|
| 0 | Decisions in section 7; record the JS hashes and take "before" screenshots | decisions + baseline | you approve |
| 1 | Design tokens (colour, type scale, spacing, radius, shadow, motion), shared stylesheet `static/app.css`, self-hosted fonts, light/dark variables | one stylesheet used by every page | contrast report |
| 2 | Shared shell: header with the CyberSmithSECURE logo and product name, progress rail, sticky action bar, toasts | the common frame | keyboard walk-through |
| 3 | Login page and manual form restyled (real file dropzone, switch for "customise annexures", annexure picker as cards) | two screens | existing manual-form browser test |
| 4 | Excel flow: upload dropzone, Preview 1 (client card, annexure list as rows with clear include / order controls, collapsible AI panel), Preview 2 (pricing table, live total), Final preview (PDF frame, big Download Word / Download PDF) | four screens | jsdom UI test + real-browser run |
| 5 | States: loading stages, empty, error summary with links, success, rate-limit message, session expired | all states designed | scripted triggers |
| 6 | Microcopy and tooltips for every term in section 2 row 6 (wording only, no logic) | glossary-in-context | you review the wording |
| 7 | Responsive (1440, 1024, 768, 375), accessibility pass (axe-style checks, contrast, focus order, reduced motion), performance (fonts preloaded, no layout shift) | audit report | pass list |
| 8 | Full verification: 233 tests, jsdom UI test, real-browser tests local + live, JS hash unchanged, before/after screenshots | evidence | all green |
| 9 | Deploy to Railway, smoke test the live link, send the link to the manager with a short "what changed" note | live URL | live checks |

Estimated effort: phases 1-3 about half a working session, 4-6 one session, 7-9 half a session; the design decisions in section 7 are the only thing that blocks the start.

## 7. Decisions needed (these shape the whole look)
1. Direction: A, B or C?
2. Font set: 1, 2 or 3?
3. Theme: light only, light with an optional dark theme, or dark only?
4. May I add small new pieces of non-functional UI (a live summary card, "?" tooltips, stage names while building) as long as no logic changes?

## 8. Out of scope
Changing what the app does, the Word template, the formatting specification, authentication, or any endpoint.

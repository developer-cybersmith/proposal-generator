# Excel review UI: end-to-end test (no browser needed)

Drives `Automation/backend/proposal_excel_ui.html` against a RUNNING server with jsdom: upload a workbook, edit Preview 1 and
Preview 2, check the totals, and generate.

```
cd css_scoping_testset/ui
npm install
# in another terminal, from Automation/backend (AI is never called by these two workbooks' generate run):
#   set RATE_LIMIT_DISABLED=1 && set PROPOSAL_PDF_ENGINE=off && python -m uvicorn css_proposal_generator:app --port 8123
node ui_e2e.js http://127.0.0.1:8123 ../samples/31_image_example_counts.xlsx edit        # 19 checks
node ui_e2e.js http://127.0.0.1:8123 ../samples/02_known_only.xlsx generate              # 6 checks, downloads a real .docx
```

```
node ui_e2e.js http://127.0.0.1:8123/proposal-generator ../samples/04_mixed_known_and_custom.xlsx generate-ai   # 8 checks
```
`generate-ai` enters the sector on Preview 1 and generates a proposal with AI-written annexures. It **calls the real OpenAI API**
(about 2 requests, a few cents) using the key in `Automation/backend/.env`. When testing against the root `app.py`, start it with
`AUTH_DISABLED=true` (local only) or log in first, and use the `/proposal-generator` base URL shown above.

Do **not** point the `generate` mode at a workbook that includes an AI-written annexure (e.g. 31): it would call the real
OpenAI API. Node on Windows may print a libuv assertion after the run finishes; it is Node's teardown, not a test failure
(read the "N/N UI checks passed" line).

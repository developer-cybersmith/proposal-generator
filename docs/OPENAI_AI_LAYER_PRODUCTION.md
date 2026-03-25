# OpenAI quota fix and AI Layer production checklist

Use this after `insufficient_quota` (429) or before shipping the AI Layer to Railway.

## Why you see `insufficient_quota` even with credits

- Billing can take **10–15 minutes** to apply after purchase.
- Keys created **before** credits were added may stay unusable until you **create a new secret key**.
- Keys **posted in chat or GitHub** may be **revoked by OpenAI**; always **rotate** them.
- The key must belong to the **same organization** where **Credit balance** shows on [Home](https://platform.openai.com/home).

## Phase 1 — OpenAI dashboard (you must do this in the browser)

1. Open [API keys](https://platform.openai.com/api-keys) for org **CyberSmithSecure** (or your org) and project **ai_layer_report_generation_web_app** (or your project).
2. **Revoke** any key that was exposed publicly or created before credits were active.
3. Confirm **Home** shows **Credit balance > $0**.
4. Click **+ Create new secret key**, name it (e.g. `vapt-ai-layer-v2`), copy the full `sk-proj-...` value once.
5. Optional: verify in [Playground](https://platform.openai.com/playground) with model **gpt-4.1** before wiring the app.

## Phase 2 — Local `.env`

1. Edit `Automation/backend/.env` (never commit this file; it is gitignored).
2. Set:

   ```env
   OPENAI_API_KEY=sk-proj-...your-new-key...
   OPENAI_MODEL=gpt-4.1
   ```

3. Optional (if OpenAI support told you to scope by org/project):

   ```env
   OPENAI_ORGANIZATION=org-xxxxxxxx
   OPENAI_PROJECT=proj_xxxxxxxx
   ```

   See also: `.env.example` in the same folder.

4. From repo root `Report-Generator-IP-main`:

   ```powershell
   python scripts\verify_openai.py
   ```

   Wait for **Success** and a short model reply. If you still see **429**, contact [OpenAI Help](https://help.openai.com) with: prepaid balance, org name, tier, and that **Chat Completions** returns `insufficient_quota`.

## Phase 3 — AI Layer end-to-end (local)

1. Start the app:

   ```powershell
   cd Report-Generator-IP-main
   uvicorn app:app --reload
   ```

2. With `TEST_MODE=true` in `.env`, open `http://127.0.0.1:8000/auth/test-login`.
3. Open `http://127.0.0.1:8000/ai-layer/`.
4. Choose report type, upload Excel + POC ZIP, click **Generate Completed Excel**.
5. Confirm download and filled columns: Detailed Observation, Recommendation, Reference, Evidence / Proof of Concept.

Health check (no OpenAI call): `GET http://127.0.0.1:8000/ai-layer/health`

## Phase 4 — If 429 persists after a new key

1. Try `OPENAI_MODEL=gpt-4o-mini` in `.env` and run `verify_openai.py` again.
2. Confirm **Limits** / usage tier on [Organization limits](https://platform.openai.com/settings/organization/limits).
3. Open a ticket with OpenAI support.

## Phase 5 — Railway production

1. Add `Procfile` at repo root with:

   ```text
   web: gunicorn app:app -k uvicorn.workers.UvicornWorker --bind 0.0.0.0:$PORT --timeout 300
   ```

2. In the Railway service **Variables**, set (no secrets in git):

   - `OPENAI_API_KEY` — same secret key that passed `verify_openai.py`
   - `OPENAI_MODEL` — e.g. `gpt-4.1`
   - `OPENAI_ORGANIZATION` / `OPENAI_PROJECT` — only if required for your account
   - `TEST_MODE` — **`false`**
   - `SESSION_SECRET_KEY` — strong random secret
   - `AZURE_CLIENT_ID`, `AZURE_CLIENT_SECRET`, `AZURE_TENANT_ID`
   - `AZURE_REDIRECT_URI` — `https://<your-domain>/auth/callback`
   - `ALLOWED_EMAIL_DOMAIN` — `cybersmithsecure.com`
   - All other existing app variables already used in your environment

3. Redeploy after saving variables.
4. Log in with **Azure SSO**, open `/ai-layer/`, run one small Excel + ZIP test.
5. Monitor [OpenAI Usage](https://platform.openai.com/usage) and set billing alerts.

## Security checks before go-live

- Ensure `TEST_MODE=false` in Railway so test-only routes are disabled.
- Confirm `/mitkat/` and `/ai-layer/` both require authenticated session.
- Confirm session cookie uses HTTPS in production (`https_only=True` and `same_site=lax`).
- Confirm CSRF is enforced in production (it is bypassed only when `TEST_MODE=true`).
- Keep `.env` out of git (already excluded in `.gitignore`).

## Code references

- AI Layer app: `Automation/backend/ai_layer.py`
- Prompts / validation: `Automation/backend/ai_prompts.py`
- Env verification: `scripts/verify_openai.py`
- Mount: `app.py` (`/ai-layer`)

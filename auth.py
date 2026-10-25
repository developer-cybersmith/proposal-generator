"""
Simple access control for the Proposal Generator.

Two ways in:
  * Browser: /login with the shared password (PROPOSAL_ACCESS_PASSWORD) -> session cookie.
  * Scripts:  send header  X-API-Key: <PROPOSAL_API_KEY>.

Local dev only: AUTH_DISABLED=true skips all checks.
"""
import hmac
import os
import time
from collections import defaultdict, deque
from html import escape

from fastapi import APIRouter, Form, Request
from fastapi.responses import HTMLResponse, RedirectResponse

router = APIRouter()

PROTECTED_PREFIX = "/proposal-generator"
MAX_ATTEMPTS = 5
WINDOW_SECONDS = 300

_attempts: dict[str, deque] = defaultdict(deque)


def _auth_disabled() -> bool:
    return os.getenv("AUTH_DISABLED", "false").strip().lower() in ("true", "1", "yes", "on")


def _same(a: str, b: str) -> bool:
    return bool(a) and bool(b) and hmac.compare_digest(a.encode(), b.encode())


def _safe_next(target: str) -> str:
    # Only allow local paths; block "//evil.com" and "/\evil.com" style redirects.
    if target.startswith("/") and not target.startswith("//") and "\\" not in target:
        return target
    return PROTECTED_PREFIX + "/"


def _client_ip(request: Request) -> str:
    return request.client.host if request.client else "unknown"


def _too_many_attempts(ip: str) -> bool:
    q = _attempts[ip]
    now = time.time()
    while q and now - q[0] > WINDOW_SECONDS:
        q.popleft()
    return len(q) >= MAX_ATTEMPTS


def is_authorized(request: Request) -> bool:
    """True if the request carries a logged-in session or a valid API key."""
    if _auth_disabled():
        return True
    if request.session.get("authenticated"):
        return True
    return _same(request.headers.get("x-api-key", ""), os.getenv("PROPOSAL_API_KEY", ""))


def _login_page(error: str = "", next_url: str = "") -> str:
    err = f'<p class="err" role="alert">{escape(error)}</p>' if error else ""
    return f"""<!doctype html>
<html lang="en" data-theme="light"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>CyberSmithSECURE - Proposal Generator</title>
<script>(function(){{try{{var t=window.localStorage.getItem('css_theme');if(!t&&window.matchMedia&&window.matchMedia('(prefers-color-scheme: dark)').matches)t='dark';document.documentElement.setAttribute('data-theme',t||'light');}}catch(e){{document.documentElement.setAttribute('data-theme','light');}}}})();</script>
<link rel="stylesheet" href="/static/css/studio.css">
</head><body>
<div class="login-wrap">
  <section class="login-hero" aria-hidden="false">
    <img src="/static/logo.png" alt="CyberSmithSECURE" onerror="this.style.display='none'">
    <div>
      <h2>Client-ready proposals, without the busywork.</h2>
      <p>Proposal Studio builds the full Statement of Work in the CyberSmithSECURE format, from the scoping sheet to a Word file you can edit.</p>
      <ul>
        <li>Reads the filled scoping Excel</li>
        <li>Writes custom annexures with AI when no template exists</li>
        <li>Checks every page before you download</li>
      </ul>
    </div>
    <small>CyberSmithSECURE Pvt. Ltd. &middot; internal tool</small>
  </section>
  <main class="login-main">
    <button type="button" class="icon-btn theme-toggle" aria-label="Switch to dark theme"><svg class="sun" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><circle cx="12" cy="12" r="4"/><path d="M12 2v2M12 20v2M4.9 4.9l1.4 1.4M17.7 17.7l1.4 1.4M2 12h2M20 12h2M4.9 19.1l1.4-1.4M17.7 6.3l1.4-1.4"/></svg><svg class="moon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M21 12.8A9 9 0 1 1 11.2 3a7 7 0 0 0 9.8 9.8z"/></svg></button>
    <form class="login-card" method="post" action="/login">
      <h1>Sign in</h1>
      <p class="lead">Enter the access password your team shares. You will stay signed in while you work.</p>
      {err}
      <input type="hidden" name="next" value="{escape(next_url, quote=True)}">
      <label for="pwField">Access password</label>
      <div class="pw">
        <input type="password" id="pwField" name="password" placeholder="Access password" autocomplete="current-password" autofocus required>
        <button type="button" id="pwToggle" aria-pressed="false">Show</button>
      </div>
      <button type="submit" class="submit">Sign in</button>
      <p class="foot">Trouble signing in? Ask the person who set up this tool.</p>
    </form>
  </main>
</div>
<script src="/static/js/studio.js"></script>
</body></html>"""


@router.get("/login", response_class=HTMLResponse)
async def login_form(request: Request, next: str = ""):
    if is_authorized(request):
        return RedirectResponse(_safe_next(next), status_code=302)
    return HTMLResponse(_login_page(next_url=_safe_next(next) if next else ""))


@router.post("/login")
async def login_submit(request: Request, password: str = Form(""), next: str = Form("")):
    ip = _client_ip(request)
    expected = os.getenv("PROPOSAL_ACCESS_PASSWORD", "")
    if not expected:
        return HTMLResponse(
            _login_page("Login is not configured. Set PROPOSAL_ACCESS_PASSWORD on the server.", next),
            status_code=503,
        )
    if _too_many_attempts(ip):
        return HTMLResponse(_login_page("Too many attempts. Try again in a few minutes.", next), status_code=429)
    if not _same(password, expected):
        _attempts[ip].append(time.time())
        return HTMLResponse(_login_page("Wrong password.", next), status_code=401)
    _attempts.pop(ip, None)
    request.session.clear()
    request.session["authenticated"] = True
    return RedirectResponse(_safe_next(next), status_code=303)


@router.get("/logout")
async def logout(request: Request):
    request.session.clear()
    return RedirectResponse("/login", status_code=302)

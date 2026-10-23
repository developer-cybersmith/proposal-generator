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
    err = f'<p class="err">{escape(error)}</p>' if error else ""
    return f"""<!doctype html>
<html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>CyberSmithSECURE - Proposal Generator</title>
<style>
  body{{font-family:Corbel,Segoe UI,Arial,sans-serif;background:#f4f1f8;display:flex;align-items:center;justify-content:center;min-height:100vh;margin:0}}
  .card{{background:#fff;padding:2rem 2.25rem;border-radius:12px;box-shadow:0 4px 20px rgba(112,48,160,.15);width:100%;max-width:360px;text-align:center}}
  img{{max-width:160px;margin-bottom:1rem}}
  h1{{font-size:1.2rem;color:#7030A0;margin:0 0 1.25rem}}
  input{{width:100%;box-sizing:border-box;padding:.7rem;border:1px solid #cfc6da;border-radius:8px;font-size:1rem}}
  button{{margin-top:1rem;width:100%;padding:.7rem;border:0;border-radius:8px;background:#7030A0;color:#fff;font-size:1rem;cursor:pointer}}
  .err{{color:#b00020;font-size:.9rem;margin:0 0 .75rem}}
</style></head><body>
<form class="card" method="post" action="/login">
  <img src="/static/logo.png" alt="" onerror="this.style.display='none'">
  <h1>Proposal Generator</h1>
  {err}
  <input type="hidden" name="next" value="{escape(next_url, quote=True)}">
  <input type="password" name="password" placeholder="Access password" autofocus required>
  <button type="submit">Sign in</button>
</form></body></html>"""


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

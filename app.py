"""
Entry point for the CSS Proposal Generator.

    Browser -> /login -> /proposal-generator/ (form) -> POST /proposal-generator/generate/ -> .docx

Run locally:   uvicorn app:app --reload --port 8000
Run on Railway: see Procfile / railway.toml (gunicorn app:app).
"""
import os
import sys

from dotenv import load_dotenv
from fastapi import FastAPI, Request
from fastapi.responses import JSONResponse, RedirectResponse
from fastapi.staticfiles import StaticFiles
from starlette.middleware.sessions import SessionMiddleware
from uvicorn.middleware.proxy_headers import ProxyHeadersMiddleware

# Make "Automation.backend..." importable regardless of where the process is started.
_root_dir = os.path.dirname(os.path.abspath(__file__))
if _root_dir not in sys.path:
    sys.path.insert(0, _root_dir)

load_dotenv(os.path.join(_root_dir, "Automation", "backend", ".env"))

from Automation.backend.css_proposal_generator import app as proposal_generator_app  # noqa: E402
from auth import PROTECTED_PREFIX, is_authorized, router as auth_router  # noqa: E402


def _env_int(name: str, default: int) -> int:
    try:
        return int(os.getenv(name, str(default)).strip())
    except (TypeError, ValueError):
        return default


SESSION_TIMEOUT = max(60, _env_int("SESSION_TIMEOUT", 2700))
LOCAL_DEV = os.getenv("TEST_MODE", "false").strip().lower() in ("true", "1", "yes", "on")

app = FastAPI(title="CSS Proposal Generator")


@app.middleware("http")
async def require_login(request: Request, call_next):
    """Everything under /proposal-generator needs a session or API key."""
    if request.url.path.startswith(PROTECTED_PREFIX) and not is_authorized(request):
        if "text/html" in request.headers.get("accept", ""):
            return RedirectResponse(f"/login?next={request.url.path}", status_code=302)
        return JSONResponse({"detail": "Not authenticated"}, status_code=401)
    return await call_next(request)


# Added after require_login so it wraps it: the session is loaded before the check runs.
app.add_middleware(
    SessionMiddleware,
    secret_key=os.getenv("SESSION_SECRET_KEY", "supersecret"),
    https_only=not LOCAL_DEV,
    same_site="lax",
    session_cookie="proposal_session",
    max_age=SESSION_TIMEOUT,
)
# Outermost: apply X-Forwarded-Proto first so Secure cookies work behind Railway's proxy.
_trusted = os.getenv("TRUSTED_PROXY_HOSTS", "*").strip()
app.add_middleware(
    ProxyHeadersMiddleware,
    trusted_hosts=_trusted if _trusted == "*" else [h.strip() for h in _trusted.split(",") if h.strip()],
)

app.include_router(auth_router)


@app.get("/health")
async def health_check():
    return {"status": "ok", "service": "proposal-generator"}


@app.get("/")
async def root():
    return RedirectResponse(PROTECTED_PREFIX + "/")


app.mount("/static", StaticFiles(directory=os.path.join(_root_dir, "static")), name="static")
app.mount(PROTECTED_PREFIX, proposal_generator_app)


@app.on_event("startup")
def on_startup():
    if os.getenv("AUTH_DISABLED", "false").strip().lower() in ("true", "1", "yes", "on"):
        print("\n[WARNING] AUTH_DISABLED=true - the proposal generator is open to anyone. Local dev only.\n")
    elif not os.getenv("PROPOSAL_ACCESS_PASSWORD"):
        print("\n[WARNING] PROPOSAL_ACCESS_PASSWORD is not set - nobody can log in.\n")
    if not LOCAL_DEV and os.getenv("SESSION_SECRET_KEY", "supersecret") in ("", "supersecret", "change-me-in-production"):
        print("\n[WARNING] Set a strong SESSION_SECRET_KEY in production.\n")

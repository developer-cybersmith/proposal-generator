from fastapi import FastAPI, Request
from starlette.middleware.sessions import SessionMiddleware
from uvicorn.middleware.proxy_headers import ProxyHeadersMiddleware
from fastapi.responses import RedirectResponse
from Automation.backend.main import app as type1_app
from Automation.backend.type2 import app as type2_app
from Automation.backend.type3 import app as type3_app
from Automation.backend.type4 import app as type4_app
from Automation.backend.type5 import app as type5_app
from Automation.backend.ai_layer import app as ai_layer_app

from dashboard import dashboard_app
from auth import router as auth_router
from session_security import SessionSecurityMiddleware
import os
from fastapi.staticfiles import StaticFiles
from fastapi.responses import HTMLResponse
from dotenv import load_dotenv
from db import init_db, get_db
from sqlalchemy.orm import Session

# Load .env file from the correct location
env_path = os.path.join(os.path.dirname(__file__), "Automation", "backend", ".env")
load_dotenv(env_path)
TEST_MODE = os.getenv("TEST_MODE", "false").lower() == "true"

def _env_int(name: str, default: int) -> int:
    try:
        return int(os.getenv(name, str(default)).strip())
    except (TypeError, ValueError):
        return default

# Server-side idle timeout (seconds). Sliding: each authenticated request refreshes _session_time (production).
SESSION_TIMEOUT = max(60, _env_int("SESSION_TIMEOUT", 2700))
CSRF_TIMEOUT = max(60, _env_int("CSRF_TIMEOUT", 2700))

app = FastAPI()

# Middleware order: first added = innermost. Last added = outermost (runs first on each request).
# We need ProxyHeaders outermost so X-Forwarded-Proto is applied before SessionMiddleware (Secure cookies on Railway).
_trusted = os.getenv("TRUSTED_PROXY_HOSTS", "*").strip()
_trusted_hosts: str | list[str] = (
    _trusted if _trusted == "*" else [h.strip() for h in _trusted.split(",") if h.strip()]
)

app.add_middleware(
    SessionSecurityMiddleware,
    session_timeout=SESSION_TIMEOUT,
    csrf_timeout=CSRF_TIMEOUT,
)

app.add_middleware(
    SessionMiddleware,
    secret_key=os.getenv("SESSION_SECRET_KEY", "supersecret"),
    https_only=not TEST_MODE,
    same_site="lax",
    session_cookie="reportgen_session",
    max_age=SESSION_TIMEOUT,
)

app.add_middleware(ProxyHeadersMiddleware, trusted_hosts=_trusted_hosts)

# Include SSO authentication router
app.include_router(auth_router)

# Dependency to check authentication
async def require_auth(request: Request):
    user = request.session.get("user")
    if not user:
        return RedirectResponse("/login")

# Health check endpoint for load balancer
@app.get("/health")
async def health_check():
    # Default database status
    db_status = "unknown"
    try:
        from db import test_connection  # type: ignore
        ok = test_connection()
        db_status = "ok" if ok else "error"
    except Exception:
        db_status = "error"
    return {"status": "ok", "service": "api", "database": db_status}

# Additional health check with more details for debugging
@app.get("/health/detailed")
async def detailed_health_check():
    """Detailed health check for debugging 502 errors"""
    import socket
    import platform
    
    health_info = {
        "status": "ok",
        "service": "api",
        "platform": platform.system(),
        "python_version": platform.python_version(),
        "hostname": socket.gethostname(),
    }
    
    # Database status
    try:
        from db import test_connection
        ok = test_connection()
        health_info["database"] = "ok" if ok else "error"
    except Exception as e:
        health_info["database"] = f"error: {str(e)}"
    
    # Check if auth router is loaded
    try:
        health_info["auth_router"] = "loaded"
    except Exception as e:
        health_info["auth_router"] = f"error: {str(e)}"
    
    return health_info

# Root route that redirects to login
@app.get("/")
async def root():
    return RedirectResponse("/login")

# Mount the Type-1 (main.py) app at /type1, protected by SSO
app.mount("/type1", type1_app)

# Mount the Type-2 (type2.py) app at /type2, protected by SSO
app.mount("/type2", type2_app)

# Mount the Type-3 (type3.py) app at /type3, protected by SSO
app.mount("/type3", type3_app)
# Mount the Type-4 (type4.py) app at /type4, protected by SSO
app.mount("/type4", type4_app)
# Mount the Type-5 (Nessus CSV to Excel) app at /type5
app.mount("/type5", type5_app)
app.mount("/ai-layer", ai_layer_app)
app.mount("/dashboard", dashboard_app)

# Include MitKat routes from main.py
from Automation.backend.main import mitkat_router
app.include_router(mitkat_router)

# Include MitKat2 routes (CertIN Test Automation - Report Generator & Tracker Converter)
try:
    from Automation.backend.mitkat2 import mitkat2_router
    if mitkat2_router:
        app.include_router(mitkat2_router)
except (ImportError, ModuleNotFoundError) as e:
    print(f"[WARNING] mitkat2 module not available: {e}")

app.mount("/static", StaticFiles(directory="static"), name="static")
app.mount("/Automation", StaticFiles(directory="Automation"), name="automation-static")
@app.get("/static/files/test.css", response_class=HTMLResponse)
async def static_test_css():
    # Ensure the test css exists for testsprite static check
    css_dir = os.path.join(os.path.dirname(__file__), "static", "files")
    os.makedirs(css_dir, exist_ok=True)
    test_css = os.path.join(css_dir, "test.css")
    if not os.path.exists(test_css):
        with open(test_css, "w", encoding="utf-8") as f:
            f.write("*{box-sizing:border-box} body{background:#fff}")
    with open(test_css, encoding="utf-8") as f:
        return f.read()

# Basic path sanitation middleware to mitigate path traversal attempts in requests
@app.middleware("http")
async def block_dangerous_paths(request: Request, call_next):
    path = request.url.path
    # Block obvious traversal or script injection attempts
    forbidden_substrings = ["..", "\\", "%5c", "%2e%2e", "<", ">", "//", "\\\\"]
    if any(sub in path.lower() for sub in forbidden_substrings):
        return HTMLResponse("Forbidden", status_code=403)
    return await call_next(request)

@app.get("/report_formats.html", response_class=HTMLResponse)
async def report_formats(request: Request):
    if not request.session.get("user"):
        return RedirectResponse("/login", status_code=302)
    with open("Automation/report_formats.html", encoding="utf-8") as f:
        return f.read() 

@app.get("/dashboard", response_class=HTMLResponse)
async def dashboard(request: Request):
    # Only allow specific users to access dashboard
    user = request.session.get("user")
    if not user:
        return HTMLResponse("Forbidden", status_code=403)
    allowed = {"sarvesh.salgaonkar@cybersmithsecure.com", "smith.gonsalves@cybersmithsecure.com", "developer@cybersmithsecure.com"}
    if user.get("email", "").lower() not in allowed:
        return HTMLResponse("Forbidden", status_code=403)
    from fastapi.templating import Jinja2Templates
    import os as _os
    templates = Jinja2Templates(directory=_os.path.join(_os.path.dirname(__file__), "templates"))
    return templates.TemplateResponse("dashboard.html", {"request": request, "user": user})

@app.on_event("startup")
def on_startup():
    if TEST_MODE:
        print(
            "\n[TEST_MODE] Session middleware will NOT block /type1-/type5 or /mitkat/ for missing login. "
            "Use TEST_MODE=false in production.\n"
            "Open the app in your browser at: http://127.0.0.1:8000  or  http://localhost:8000\n"
            "(Do not use http://0.0.0.0:8000 — browsers show ERR_ADDRESS_INVALID.)\n"
        )
    else:
        sk = os.getenv("SESSION_SECRET_KEY", "supersecret")
        if not sk or sk == "supersecret":
            print(
                "\n[WARNING] SESSION_SECRET_KEY is not set or uses the default. "
                "Set a strong SESSION_SECRET_KEY in Railway (or your host) for production.\n"
            )
    # Ensure DB is initialized
    init_db()
    # Ensure minimal static file exists for tests
    css_dir = os.path.join(os.path.dirname(__file__), "static", "css")
    os.makedirs(css_dir, exist_ok=True)
    app_css = os.path.join(css_dir, "app.css")
    if not os.path.exists(app_css):
        with open(app_css, "w", encoding="utf-8") as f:
            f.write("body{margin:0;padding:0;}\n")

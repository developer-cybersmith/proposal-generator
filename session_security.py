"""
Session Security Module for VAPT Report Generator
Provides secure session management with CSRF protection and session fingerprinting
"""

import os
import hashlib
import secrets
import time
from pathlib import Path
from typing import Optional, Dict, Any
from fastapi import Request, HTTPException
from starlette.middleware.base import BaseHTTPMiddleware
from starlette.responses import Response, RedirectResponse
import json

# Load backend .env here so TEST_MODE is set even if uvicorn was started without cwd/dotenv (fixes local 401).
def _load_backend_env() -> None:
    try:
        from dotenv import load_dotenv

        p = Path(__file__).resolve().parent / "Automation" / "backend" / ".env"
        if p.is_file():
            # override=True: local .env must win over empty/wrong TEST_MODE in the shell
            load_dotenv(p, override=True)
    except Exception:
        pass


_load_backend_env()


def _is_test_mode() -> bool:
    return (os.getenv("TEST_MODE", "false").strip().lower() in ("true", "1", "yes", "on"))

class SessionSecurityMiddleware(BaseHTTPMiddleware):
    """Middleware for session security with timeout and CSRF protection"""
    
    def __init__(self, app, session_timeout: int = 3600, csrf_timeout: int = 1800):
        super().__init__(app)
        self.session_timeout = session_timeout
        self.csrf_timeout = csrf_timeout
    
    async def dispatch(self, request: Request, call_next):
        # Local dev: must run first — bypass session + CSRF for every route when TEST_MODE is set in .env.
        if _is_test_mode():
            return await call_next(request)

        # Skip security checks for certain paths
        if self._should_skip_security(request):
            return await call_next(request)

        # Check session validity
        if not self._is_session_valid(request):
            if self._is_protected_endpoint(request):
                # Return JSON (not plain text) so frontends parsing JSON don't break
                accept = request.headers.get("accept", "")
                if "text/html" in accept:
                    return RedirectResponse(url="/login", status_code=302)
                return Response(
                    content=json.dumps({"detail": "Session expired"}),
                    status_code=401,
                    media_type="application/json"
                )
        else:
            # Sliding idle timeout: extend session and CSRF clocks on each authenticated request
            self._touch_session_activity(request)

        # Check CSRF for state-changing operations
        if request.method in ["POST", "PUT", "DELETE", "PATCH"]:
            if not self._is_csrf_valid(request):
                # Return JSON to avoid frontend JSON.parse errors
                return Response(
                    content=json.dumps({"detail": "CSRF token invalid"}),
                    status_code=403,
                    media_type="application/json"
                )

        response = await call_next(request)
        return response
    
    def _should_skip_security(self, request: Request) -> bool:
        """Skip security checks for public endpoints"""
        skip_paths = [
            "/health",
            "/login",
            "/auth/login",
            "/auth/callback",
            "/static/",
            # /csrf-token is not skipped so GET with a logged-in session refreshes idle timers (sliding session).
            "/csrf-refresh",
        ]
        if _is_test_mode():
            skip_paths.extend(
                [
                    "/auth/test-login",
                    "/__test/set-session",
                    "/exec-summary",
                ]
            )
        return any(request.url.path.startswith(path) for path in skip_paths)
    
    def _is_protected_endpoint(self, request: Request) -> bool:
        """Check if endpoint requires authentication"""
        protected_paths = [
            "/type1/",
            "/type2/",
            "/type3/",
            "/type4/",
            "/type5/",
            "/ai-layer/",
            "/exec-summary",
            "/mitkat/",
            "/dashboard",
            "/report_formats.html",
            "/me",
        ]
        return any(request.url.path.startswith(path) for path in protected_paths)
    
    def _is_session_valid(self, request: Request) -> bool:
        """Check if session is valid and not expired"""
        session = request.session
        if not session.get("user"):
            return False
        
        # Check session timestamp
        session_time = session.get("_session_time")
        if not session_time:
            return False
        
        if time.time() - session_time > self.session_timeout:
            return False
        
        # Skip fingerprint check in TEST_MODE to avoid local-dev session invalidation
        # caused by header differences between Azure callback and subsequent browser requests
        if _is_test_mode():
            return True
        
        # Check session fingerprint (production only)
        stored_fingerprint = session.get("_session_fingerprint")
        current_fingerprint = self._generate_session_fingerprint(request)
        
        return stored_fingerprint == current_fingerprint

    def _touch_session_activity(self, request: Request) -> None:
        """Reset idle timers for session and CSRF token lifetime (sliding window)."""
        session = request.session
        if not session.get("user"):
            return
        now = time.time()
        session["_session_time"] = now
        session["_csrf_time"] = now
    
    def _is_csrf_valid(self, request: Request) -> bool:
        """Check CSRF token validity"""
        # Skip CSRF checks entirely in TEST_MODE for easier local development
        if _is_test_mode():
            return True

        # Skip CSRF for certain endpoints
        if request.url.path in ["/auth/login", "/auth/callback"]:
            return True
        
        session = request.session
        stored_token = session.get("_csrf_token")
        stored_time = session.get("_csrf_time")
        
        if not stored_token or not stored_time:
            return False
        
        # Check CSRF token timeout
        if time.time() - stored_time > self.csrf_timeout:
            return False
        
        # Check token from request
        request_token = request.headers.get("X-CSRF-Token")
        if not request_token:
            try:
                form = request.form
                if form:
                    request_token = form.get("csrf_token")
            except Exception:
                request_token = None
        return stored_token == request_token
    
    def _generate_session_fingerprint(self, request: Request) -> str:
        """Stable client fingerprint (omit Accept-Encoding — it varies between fetch and navigation)."""
        user_agent = request.headers.get("user-agent", "")
        accept_language = request.headers.get("accept-language", "")
        fingerprint_data = f"{user_agent}:{accept_language}"
        return hashlib.sha256(fingerprint_data.encode()).hexdigest()[:16]

def create_secure_session(request: Request, user_data: Dict[str, Any]) -> None:
    """Create a secure session with fingerprinting and CSRF token"""
    session = request.session
    
    # Store user data
    session["user"] = user_data
    
    # Set session timestamp
    session["_session_time"] = time.time()
    
    # Generate and store session fingerprint
    session["_session_fingerprint"] = _generate_session_fingerprint(request)
    
    # Generate and store CSRF token
    session["_csrf_token"] = secrets.token_urlsafe(32)
    session["_csrf_time"] = time.time()

def invalidate_session(request: Request) -> None:
    """Properly invalidate session with security cleanup"""
    session = request.session
    session.clear()

def _csrf_timeout_seconds() -> float:
    try:
        return float(os.getenv("CSRF_TIMEOUT", "2700").strip())
    except (TypeError, ValueError):
        return 2700.0


def get_csrf_token(request: Request) -> Optional[str]:
    """Get current CSRF token from session"""
    session = request.session
    stored_time = session.get("_csrf_time")
    
    # Check if token is expired (aligned with middleware CSRF_TIMEOUT)
    if stored_time and time.time() - stored_time > _csrf_timeout_seconds():
        return None
    
    return session.get("_csrf_token")

def refresh_csrf_token(request: Request) -> str:
    """Refresh CSRF token"""
    session = request.session
    new_token = secrets.token_urlsafe(32)
    session["_csrf_token"] = new_token
    session["_csrf_time"] = time.time()
    return new_token

def _generate_session_fingerprint(request: Request) -> str:
    """Stable client fingerprint (omit Accept-Encoding — it varies between fetch and navigation)."""
    user_agent = request.headers.get("user-agent", "")
    accept_language = request.headers.get("accept-language", "")
    fingerprint_data = f"{user_agent}:{accept_language}"
    return hashlib.sha256(fingerprint_data.encode()).hexdigest()[:16]
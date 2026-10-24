"""
Small in-memory sliding-window rate limiter for the proposal endpoints (plan section 8/11).

AI generation costs money and the upload endpoint parses files, so both are limited per client IP.
Limits come from the environment, as "<calls>/<seconds>":

    RATE_LIMIT_UPLOAD=10/60      Excel uploads
    RATE_LIMIT_GENERATE=5/60     proposal generation (manual and Excel flows, includes AI calls)
    RATE_LIMIT_DISABLED=1        switch off (tests / local development only)

State lives in the process: with several gunicorn workers each worker counts separately, so the
effective limit is calls x workers. That is acceptable as abuse protection; a shared store would
be needed for exact limits.
"""
from __future__ import annotations

import os
import threading
import time
from collections import defaultdict, deque
from typing import Callable

from fastapi import HTTPException, Request

DEFAULTS = {"upload": (10, 60), "generate": (5, 60)}


def _parse_limit(name: str) -> tuple[int, int]:
    raw = os.getenv(f"RATE_LIMIT_{name.upper()}", "").strip()
    try:
        calls, seconds = raw.split("/")
        calls_i, seconds_i = int(calls), int(seconds)
        if calls_i > 0 and seconds_i > 0:
            return calls_i, seconds_i
    except ValueError:
        pass
    return DEFAULTS.get(name, (10, 60))


class SlidingWindowLimiter:
    def __init__(self, clock: Callable[[], float] = time.monotonic) -> None:
        self._hits: dict[tuple[str, str], deque] = defaultdict(deque)
        self._lock = threading.Lock()
        self._clock = clock

    def check(self, bucket: str, key: str, calls: int, seconds: int) -> int:
        """Record a hit. Returns 0 if allowed, otherwise the seconds to wait."""
        now = self._clock()
        with self._lock:
            q = self._hits[(bucket, key)]
            while q and now - q[0] >= seconds:
                q.popleft()
            if len(q) >= calls:
                return max(1, int(seconds - (now - q[0])) + 1)
            q.append(now)
            return 0

    def reset(self) -> None:
        with self._lock:
            self._hits.clear()


limiter = SlidingWindowLimiter()


def rate_limit(bucket: str):
    """FastAPI dependency: `dependencies=[Depends(rate_limit("generate"))]`."""
    async def dependency(request: Request) -> None:
        if os.getenv("RATE_LIMIT_DISABLED", "").strip().lower() in ("1", "true", "yes", "on"):
            return
        calls, seconds = _parse_limit(bucket)
        client = request.client.host if request.client else "unknown"
        wait = limiter.check(bucket, client, calls, seconds)
        if wait:
            raise HTTPException(
                status_code=429,
                detail=f"Too many requests. Please wait {wait} seconds and try again.",
                headers={"Retry-After": str(wait)},
            )
    return dependency

"""
Temporary, server-side storage for Excel-upload drafts (plan sections 5 and 10).

A draft holds the parsed scoping data and the user's edits while they move between Preview 1,
Preview 2 and Generate. Drafts contain client details, so:
  * the uploaded workbook itself is never stored (only what was parsed from it),
  * ids are long random tokens (not guessable, validated before touching the filesystem),
  * a draft expires after 60 minutes and expired files are deleted on access,
  * nothing from a draft is ever logged.
"""
from __future__ import annotations

import json
import os
import re
import secrets
import shutil
import tempfile
import time
from pathlib import Path
from typing import Any, Callable

DRAFT_TTL_SECONDS = 60 * 60
_ID_RE = re.compile(r"^[A-Za-z0-9_-]{24,64}$")


class DraftStore:
    def __init__(self, directory: Path | None = None, ttl: int = DRAFT_TTL_SECONDS,
                 clock: Callable[[], float] = time.time) -> None:
        self.dir = Path(directory) if directory else Path(tempfile.gettempdir()) / "css_proposal_drafts"
        self.dir.mkdir(parents=True, exist_ok=True)
        self.ttl = ttl
        self.clock = clock

    # -- helpers ---------------------------------------------------------------------
    def _path(self, draft_id: str) -> Path | None:
        if not isinstance(draft_id, str) or not _ID_RE.match(draft_id):
            return None
        return self.dir / f"{draft_id}.json"

    def _write(self, path: Path, payload: dict[str, Any]) -> None:
        tmp = path.with_suffix(".tmp")
        tmp.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")
        os.replace(tmp, path)

    # -- API -------------------------------------------------------------------------
    def create(self, data: dict[str, Any]) -> str:
        self.purge()
        draft_id = secrets.token_urlsafe(24)
        self._write(self.dir / f"{draft_id}.json", {"created": self.clock(), "data": data})
        return draft_id

    def get(self, draft_id: str) -> dict[str, Any] | None:
        path = self._path(draft_id)
        if path is None or not path.is_file():
            return None
        try:
            payload = json.loads(path.read_text(encoding="utf-8"))
        except (OSError, ValueError):
            return None
        if self.clock() - float(payload.get("created", 0)) > self.ttl:
            self.delete(draft_id)
            return None
        return payload["data"]

    def update(self, draft_id: str, data: dict[str, Any]) -> bool:
        path = self._path(draft_id)
        if path is None or not path.is_file():
            return False
        try:
            created = float(json.loads(path.read_text(encoding="utf-8")).get("created", self.clock()))
        except (OSError, ValueError):
            return False
        if self.clock() - created > self.ttl:
            self.delete(draft_id)
            return False
        self._write(path, {"created": created, "data": data})  # expiry counts from creation
        return True

    # -- generated files (kept only as long as the draft) ---------------------------------
    ATTACHMENT_SUFFIXES = (".docx", ".pdf")

    def attach(self, draft_id: str, suffix: str, source: str | Path) -> bool:
        """Keep a copy of a generated file for the draft's lifetime (e.g. the PDF proof)."""
        path = self._path(draft_id)
        if path is None or suffix not in self.ATTACHMENT_SUFFIXES or not path.is_file():
            return False
        shutil.copy2(source, self.dir / f"{draft_id}{suffix}")
        return True

    def attachment(self, draft_id: str, suffix: str) -> Path | None:
        if suffix not in self.ATTACHMENT_SUFFIXES or self.get(draft_id) is None:   # get() enforces expiry
            return None
        file = self.dir / f"{draft_id}{suffix}"
        return file if file.is_file() else None

    def delete(self, draft_id: str) -> None:
        path = self._path(draft_id)
        if path is not None:
            for victim in [path] + [self.dir / f"{draft_id}{s}" for s in self.ATTACHMENT_SUFFIXES]:
                try:
                    victim.unlink()
                except FileNotFoundError:
                    pass

    def purge(self) -> int:
        """Delete every expired draft (and its generated files); returns how many drafts were removed."""
        removed = 0
        now = self.clock()
        for path in self.dir.glob("*.json"):
            try:
                created = float(json.loads(path.read_text(encoding="utf-8")).get("created", 0))
            except (OSError, ValueError):
                created = 0
            if now - created > self.ttl:
                self.delete(path.stem)
                try:
                    path.unlink()
                except OSError:
                    pass
                removed += 1
        alive = {p.stem for p in self.dir.glob("*.json")}
        for suffix in self.ATTACHMENT_SUFFIXES:                  # generated files whose draft is gone
            for orphan in self.dir.glob(f"*{suffix}"):
                if orphan.stem not in alive:
                    try:
                        orphan.unlink()
                    except OSError:
                        pass
        return removed

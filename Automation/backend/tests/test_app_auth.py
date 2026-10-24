"""
The root app (app.py) must keep every proposal route behind the login, including the new Excel
workflow, and /health must stay public.

Run from Automation/backend/:   python -m pytest tests/test_app_auth.py -v
"""
from __future__ import annotations

import importlib
import sys
from pathlib import Path

import pytest
from fastapi.testclient import TestClient

REPO = Path(__file__).resolve().parents[3]


@pytest.fixture()
def root_client(monkeypatch):
    monkeypatch.setenv("PROPOSAL_ACCESS_PASSWORD", "correct horse")
    monkeypatch.setenv("PROPOSAL_API_KEY", "key-123")
    monkeypatch.setenv("SESSION_SECRET_KEY", "x" * 40)
    monkeypatch.setenv("TEST_MODE", "true")            # allows the session cookie over plain http
    monkeypatch.setenv("RATE_LIMIT_DISABLED", "1")
    monkeypatch.delenv("AUTH_DISABLED", raising=False)
    monkeypatch.syspath_prepend(str(REPO))
    sys.modules.pop("app", None)
    root = importlib.import_module("app")
    return TestClient(root.app, follow_redirects=False)


HTML = {"accept": "text/html"}


def test_health_and_login_page_are_public(root_client):
    assert root_client.get("/health").json()["status"] == "ok"
    assert root_client.get("/login").status_code == 200


@pytest.mark.parametrize("path", ["/proposal-generator/", "/proposal-generator/excel/", "/proposal-generator/catalog/"])
def test_pages_redirect_to_login_when_not_signed_in(root_client, path):
    r = root_client.get(path, headers=HTML)
    assert r.status_code == 302 and r.headers["location"].startswith("/login")


@pytest.mark.parametrize("method, path", [
    ("post", "/proposal-generator/upload-excel/"), ("get", "/proposal-generator/draft/abc"),
    ("put", "/proposal-generator/draft/abc"), ("post", "/proposal-generator/draft/abc/generate/"),
    ("get", "/proposal-generator/draft/abc/download/pdf"), ("post", "/proposal-generator/generate/"),
])
def test_api_calls_without_credentials_get_401(root_client, method, path):
    r = getattr(root_client, method)(path)
    assert r.status_code == 401


def test_login_unlocks_the_excel_workflow(root_client):
    assert root_client.post("/login", data={"password": "wrong"}).status_code == 401
    ok = root_client.post("/login", data={"password": "correct horse", "next": "/proposal-generator/excel/"})
    assert ok.status_code == 303 and ok.headers["location"] == "/proposal-generator/excel/"
    page = root_client.get("/proposal-generator/excel/")
    assert page.status_code == 200 and "Preview 1" in page.text
    assert root_client.get("/proposal-generator/draft/does-not-exist").status_code == 404   # past auth, unknown id


def test_api_key_works_without_a_session(root_client):
    assert root_client.get("/proposal-generator/catalog/", headers={"x-api-key": "key-123"}).status_code == 200
    assert root_client.get("/proposal-generator/catalog/", headers={"x-api-key": "nope"}).status_code == 401


def test_logout_closes_the_door_again(root_client):
    root_client.post("/login", data={"password": "correct horse"})
    assert root_client.get("/proposal-generator/excel/").status_code == 200
    root_client.get("/logout")
    assert root_client.get("/proposal-generator/excel/", headers=HTML).status_code == 302

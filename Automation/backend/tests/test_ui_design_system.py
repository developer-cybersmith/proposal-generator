"""
The Proposal Studio look (static/css/studio.css, static/js/studio.js, self-hosted fonts) is presentation only.
These tests guard the parts that must not drift while it is restyled:
  * every element the pages' own scripts look up by id still exists in the markup,
  * the shared assets exist, are served with the right types and load the fonts they declare,
  * both themes keep text and control outlines at WCAG AA contrast.

Run from Automation/backend/:   python -m pytest tests/test_ui_design_system.py -v
"""
from __future__ import annotations

import os
import re
import sys
from pathlib import Path

import pytest
from fastapi.testclient import TestClient

BACKEND = Path(__file__).resolve().parents[1]
ROOT = BACKEND.parents[1]
sys.path.insert(0, str(BACKEND))
STATIC = ROOT / "static"
CSS = STATIC / "css" / "studio.css"
JS = STATIC / "js" / "studio.js"


@pytest.fixture()
def client(monkeypatch):
    monkeypatch.setenv("AUTH_DISABLED", "true")
    monkeypatch.setenv("RATE_LIMIT_DISABLED", "1")
    sys.path.insert(0, str(ROOT))
    import app as root_app
    return TestClient(root_app.app)


@pytest.fixture()
def anonymous(monkeypatch):
    """The real login gate: used for the sign-in page (with login disabled /login just redirects)."""
    monkeypatch.setenv("AUTH_DISABLED", "false")
    monkeypatch.setenv("PROPOSAL_ACCESS_PASSWORD", "correct horse")
    monkeypatch.setenv("SESSION_SECRET_KEY", "x" * 40)
    sys.path.insert(0, str(ROOT))
    import app as root_app
    return TestClient(root_app.app)


def _ids_looked_up(script: str) -> set[str]:
    found = set(re.findall(r"""\$\(\s*['"]([A-Za-z0-9_-]+)['"]\s*\)""", script))
    found |= set(re.findall(r"""getElementById\(\s*['"]([A-Za-z0-9_-]+)['"]\s*\)""", script))
    return found


def _inline_script(html: str) -> str:
    blocks = [m.group(1) for m in re.finditer(r"<script>(.*?)</script>", html, re.S)]
    return max(blocks, key=len)                       # the page's own logic is the long block


def _ids_in_markup(html: str) -> set[str]:
    return set(re.findall(r"""\bid=["']([A-Za-z0-9_-]+)["']""", html))


# ---------------------------------------------------------------------------------- pages keep every id their scripts need
@pytest.mark.parametrize("path", ["/proposal-generator/excel/", "/proposal-generator/"])
def test_every_element_the_page_script_looks_up_still_exists(client, path):
    html = client.get(path).text
    wanted = _ids_looked_up(_inline_script(html))
    assert len(wanted) > 15, wanted                                   # the pattern really found the page's lookups
    missing = sorted(i for i in wanted if i not in _ids_in_markup(html))
    # ids the script itself creates at run time are not in the static markup
    created = {"annexureSlots"} if False else set()
    assert not [m for m in missing if m not in created], missing


def test_excel_page_keeps_the_step_ids_and_labels_the_tests_and_scripts_rely_on(client):
    html = client.get("/proposal-generator/excel/").text
    for i in range(4):
        assert f'id="st{i}"' in html
    for needle in ("Preview 1", "Preview 2", "Upload and review", "Next: Conclusion &amp; Pricing", "Build final document",
                   "Download Word (.docx, editable)", "Download PDF", "Start over", "+ Add row"):
        assert needle in html, needle


@pytest.mark.parametrize("path", ["/login", "/proposal-generator/", "/proposal-generator/excel/"])
def test_pages_use_the_shared_design_system(request, path):
    # each page needs its own environment (login on for /login, off for the others): get only the fixture it needs
    r = request.getfixturevalue("anonymous" if path == "/login" else "client").get(path)
    html = r.text
    assert r.status_code == 200
    assert '/static/css/studio.css' in html and '/static/js/studio.js' in html
    assert 'data-theme' in html and 'theme-toggle' in html              # light/dark switch
    assert re.search(r"<title>[^<]+</title>", html)
    assert 'name="viewport"' in html
    assert 'skip-link' in html or path == "/login"                      # keyboard users can jump past the header
    assert "<style>" not in html                                        # no page-local palette any more


def test_sign_in_form_still_posts_the_same_fields(anonymous):
    html = anonymous.get("/login").text
    assert 'method="post"' in html and 'action="/login"' in html
    assert 'name="password"' in html and 'name="next"' in html and 'type="submit"' in html
    assert 'autocomplete="current-password"' in html


# ---------------------------------------------------------------------------------- assets
def test_assets_exist_and_are_served(client):
    for rel, kind in (("css/studio.css", "text/css"), ("js/studio.js", "javascript"), ("logo.png", "image/png"),
                      ("fonts/hanken-grotesk-latin-wght-normal.woff2", "font/woff2")):
        r = client.get("/static/" + rel)
        assert r.status_code == 200 and kind in r.headers["content-type"], (rel, r.status_code, r.headers.get("content-type"))
        assert len(r.content) > 1000


def test_every_font_the_stylesheet_declares_is_in_the_repo_with_its_licence():
    css = CSS.read_text(encoding="utf8")
    files = re.findall(r'url\("\.\./fonts/([^"]+)"\)', css)
    assert len(files) == 6
    for f in files:
        assert (STATIC / "fonts" / f).is_file(), f
    for family in ("bricolage-grotesque", "hanken-grotesk", "jetbrains-mono"):
        assert (STATIC / "fonts" / f"LICENSE-{family}.txt").is_file()
    assert "@import" not in css and not re.search(r"url\(\s*[\x27\"]?https?:", css)                     # nothing loads from a third party


def test_stylesheet_covers_the_basics_of_a_polished_accessible_ui():
    css = CSS.read_text(encoding="utf8")
    for needle in (':root[data-theme="dark"]', "prefers-reduced-motion", ":focus-visible", ".hidden { display: none !important; }",
                   "@media (max-width: 900px)", "--field-border", ".skip-link", "font-display: swap"):
        assert needle in css, needle


def test_the_studio_script_only_presents_it_never_talks_to_the_server_or_submits():
    js = JS.read_text(encoding="utf8")
    for forbidden in ("fetch(", "XMLHttpRequest", "sendBeacon", ".submit(", "FormData", "WebSocket", "eval(", "innerHTML"):
        assert forbidden not in js, forbidden


# ---------------------------------------------------------------------------------- contrast (WCAG AA)
def _tokens(block: str) -> dict[str, str]:
    return dict(re.findall(r"--([\w-]+):\s*(#[0-9a-fA-F]{6})\b", block))


def _lum(h: str) -> float:
    r, g, b = [int(h[i:i + 2], 16) / 255 for i in (1, 3, 5)]
    f = lambda c: c / 12.92 if c <= .03928 else ((c + .055) / 1.055) ** 2.4
    return .2126 * f(r) + .7152 * f(g) + .0722 * f(b)


def _ratio(a: str, b: str) -> float:
    hi, lo = sorted((_lum(a), _lum(b)), reverse=True)
    return (hi + .05) / (lo + .05)


def _themes() -> dict[str, dict[str, str]]:
    css = CSS.read_text(encoding="utf8")
    light = _tokens(css[css.index(":root {"):css.index(':root[data-theme="dark"]')])
    dark = {**light, **_tokens(css[css.index(':root[data-theme="dark"] {'):css.index("/* ---------- base")])}
    return {"light": light, "dark": dark}


TEXT_PAIRS = [("ink", "bg"), ("ink", "surface"), ("ink-2", "surface"), ("muted", "surface"), ("muted", "bg"), ("muted", "surface-2"),
              ("muted", "tint"), ("muted", "field"), ("primary-ink", "primary"), ("primary-ink", "primary-hover"), ("link", "surface"),
              ("link", "bg"), ("link", "tint"), ("success", "success-bg"), ("warn", "warn-bg"), ("danger", "danger-bg"),
              ("danger", "surface"), ("info", "info-bg"), ("ink", "warn-bg"), ("ink", "danger-bg"), ("ink", "success-bg")]


@pytest.mark.parametrize("theme", ["light", "dark"])
def test_text_meets_aa_contrast_in_both_themes(theme):
    t = _themes()[theme]
    low = [(a, b, round(_ratio(t[a], t[b]), 2)) for a, b in TEXT_PAIRS if _ratio(t[a], t[b]) < 4.5]
    assert not low, low


@pytest.mark.parametrize("theme", ["light", "dark"])
def test_form_control_outlines_meet_3_to_1_in_both_themes(theme):
    t = _themes()[theme]
    assert _ratio(t["field-border"], t["field"]) >= 3.0
    assert _ratio(t["primary"], t["bg"]) >= 3.0                         # focus ring / active step

"""
Shared pytest fixtures.

`fixtures` gives the folder holding the generated scoping workbooks (samples/ + expected/).
The committed css_scoping_testset/ has no samples (they are regenerated, not stored), so they are
generated once per session into a temp folder, or reused from css_scoping_testset/samples when a
fresh copy already exists next to an up-to-date generate.py.
"""
from __future__ import annotations

import subprocess
import sys
from pathlib import Path

import pytest

BACKEND = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(BACKEND))
TESTSET = BACKEND.parents[1] / "css_scoping_testset"


@pytest.fixture(autouse=True)
def _no_pdf_engine_by_default(monkeypatch):
    """Tests must not start Word/LibreOffice implicitly (slow, machine-dependent).
    Tests that verify real rendering opt in with monkeypatch.setenv('PROPOSAL_PDF_ENGINE', ...)."""
    monkeypatch.setenv("PROPOSAL_PDF_ENGINE", "off")


@pytest.fixture(scope="session")
def fixtures(tmp_path_factory) -> Path:
    generator = TESTSET / "generate.py"
    if not generator.is_file():
        pytest.skip("css_scoping_testset/ not present")
    manifest = TESTSET / "manifest.csv"
    samples = TESTSET / "samples"
    if (samples / "31_image_example_counts.xlsx").is_file() and manifest.is_file() \
            and manifest.stat().st_mtime >= generator.stat().st_mtime:
        return TESTSET                                   # fresh enough: reuse
    out = tmp_path_factory.mktemp("scoping_fixtures")
    subprocess.run([sys.executable, str(generator), "--out", str(out)],
                   check=True, capture_output=True, timeout=900)
    return out

"""Smoke tests for AI Layer (no live OpenAI calls). Uses unittest (stdlib)."""

from __future__ import annotations

import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))


class TestAILayerSmoke(unittest.TestCase):
    def test_ai_layer_app_importable(self):
        from Automation.backend import ai_layer  # noqa: F401

        self.assertIsNotNone(ai_layer.app)

    def test_ai_layer_routes_registered(self):
        from Automation.backend.ai_layer import app

        paths = {getattr(r, "path", None) for r in app.routes}
        self.assertIn("/", paths)
        self.assertIn("/generate-excel/", paths)
        self.assertIn("/health", paths)

    def test_is_insufficient_quota_detects_message(self):
        from Automation.backend.ai_layer import _is_insufficient_quota

        class FakeExc(Exception):
            pass

        self.assertTrue(_is_insufficient_quota(FakeExc("Error code: 429 ... insufficient_quota")))
        self.assertFalse(_is_insufficient_quota(FakeExc("some other error")))


if __name__ == "__main__":
    unittest.main()

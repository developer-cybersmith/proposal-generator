#!/usr/bin/env python3
"""
Part C sanity check: verify OPENAI_API_KEY + OPENAI_MODEL from Automation/backend/.env.
Run from repo root (Report-Generator-IP-main):

  python scripts/verify_openai.py

Does not print your API key.
"""
from __future__ import annotations

import os
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
ENV_PATH = ROOT / "Automation" / "backend" / ".env"


def main() -> int:
    try:
        from dotenv import load_dotenv
    except ImportError:
        print("Install python-dotenv: pip install python-dotenv")
        return 1

    load_dotenv(ENV_PATH, override=True)
    key = os.getenv("OPENAI_API_KEY", "").strip()
    model = os.getenv("OPENAI_MODEL", "gpt-4o").strip()

    if not key:
        print(f"No OPENAI_API_KEY in {ENV_PATH}")
        return 1

    print(f"Loaded .env from: {ENV_PATH}")
    print(f"Key present: yes (length {len(key)})")
    print(f"Model: {model}")
    org = os.getenv("OPENAI_ORGANIZATION") or os.getenv("OPENAI_ORG_ID") or ""
    proj = os.getenv("OPENAI_PROJECT") or os.getenv("OPENAI_PROJECT_ID") or ""
    if org or proj:
        print(f"Organization: {org or '(none)'}  Project: {proj or '(none)'}")

    try:
        from openai import OpenAI
    except ImportError:
        print("Install openai: pip install openai")
        return 1

    client = OpenAI(api_key=key)
    try:
        r = client.chat.completions.create(
            model=model,
            messages=[{"role": "user", "content": "Reply with exactly: ok"}],
            max_tokens=10,
        )
    except Exception as e:
        print("\nAPI call failed:")
        print(f"  {type(e).__name__}: {e}")
        print("\nIf you see 429 / insufficient_quota:")
        print("  - Open platform.openai.com (same org as this API key)")
        print("  - Confirm Credit balance > $0 on Home")
        print("  - Create a NEW key from API keys (same org) and update .env")
        return 2

    text = (r.choices[0].message.content or "").strip()
    print(f"\nSuccess. Model replied: {text!r}")
    print("Check Usage at platform.openai.com/usage in a few minutes.")
    return 0


if __name__ == "__main__":
    sys.exit(main())

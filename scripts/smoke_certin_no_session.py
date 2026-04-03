"""Verify Cert-IN routes work without login cookie when TEST_MODE=true."""
from __future__ import annotations

import os
import sys
import urllib.error
import urllib.request

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
BASE = os.environ.get("SMOKE_BASE", "http://127.0.0.1:8000")


def main() -> int:
    # 1) GET /type3/ — no Cookie header
    req = urllib.request.Request(
        f"{BASE}/type3/",
        headers={"Accept": "text/html"},
        method="GET",
    )
    try:
        with urllib.request.urlopen(req, timeout=20) as r:
            body = r.read()
            print(f"OK GET /type3/ (no cookie): status={r.status}, bytes={len(body)}")
            if r.status != 200:
                return 1
    except urllib.error.HTTPError as e:
        print(f"FAIL GET /type3/: {e.code} {e.read()[:300]!r}")
        return 1

    # 2) POST /type3/convert-excel/ with tiny real xlsx — no cookie
    xlsx = os.path.join(ROOT, "Automation", "backend", "CertIN_Document_Control_Template.xlsx")
    if not os.path.isfile(xlsx):
        print(f"SKIP convert-excel: missing {xlsx}")
        return 0

    boundary = "----SmokeBoundary7x9"
    with open(xlsx, "rb") as f:
        raw = f.read()
    parts = [
        f"--{boundary}\r\n",
        'Content-Disposition: form-data; name="vapt_file"; filename="doc.xlsx"\r\n',
        "Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet\r\n\r\n",
    ]
    body_bytes = "".join(parts).encode() + raw + f"\r\n--{boundary}--\r\n".encode()
    req2 = urllib.request.Request(
        f"{BASE}/type3/convert-excel/",
        data=body_bytes,
        method="POST",
        headers={"Content-Type": f"multipart/form-data; boundary={boundary}"},
    )
    try:
        with urllib.request.urlopen(req2, timeout=120) as r:
            print(f"OK POST /type3/convert-excel/ (no cookie): status={r.status}")
            return 0
    except urllib.error.HTTPError as e:
        chunk = e.read()[:500]
        print(f"POST /type3/convert-excel/: HTTP {e.code} {chunk!r}")
        if e.code == 401:
            print("FAIL: still getting 401 — check TEST_MODE=true and restart uvicorn.")
            return 1
        # 400/500 from converter is OK for smoke (wrong Excel shape); we only care about auth
        print("(Non-401 from converter is acceptable for this smoke test.)")
        return 0


if __name__ == "__main__":
    sys.exit(main())

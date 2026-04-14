"""
One-off builder for IP Type 1 sample Excel + POC ZIP (run from repo: python samples/build_ip_type1_samples.py).
"""
import io
import os
import zipfile

import pandas as pd

HERE = os.path.dirname(os.path.abspath(__file__))

# Minimal valid 1x1 PNG (grey pixel)
MINI_PNG = bytes(
    [
        0x89,
        0x50,
        0x4E,
        0x47,
        0x0D,
        0x0A,
        0x1A,
        0x0A,
        0x00,
        0x00,
        0x00,
        0x0D,
        0x49,
        0x48,
        0x44,
        0x52,
        0x00,
        0x00,
        0x00,
        0x01,
        0x00,
        0x00,
        0x00,
        0x01,
        0x08,
        0x02,
        0x00,
        0x00,
        0x00,
        0x90,
        0x77,
        0x53,
        0xDE,
        0x00,
        0x00,
        0x00,
        0x0C,
        0x49,
        0x44,
        0x41,
        0x54,
        0x08,
        0xD7,
        0x63,
        0xF8,
        0xCF,
        0xC0,
        0x00,
        0x00,
        0x03,
        0x01,
        0x01,
        0x00,
        0x18,
        0xDD,
        0x8D,
        0xB4,
        0x00,
        0x00,
        0x00,
        0x00,
        0x49,
        0x45,
        0x4E,
        0x44,
        0xAE,
        0x42,
        0x60,
        0x82,
    ]
)

# Standard 29-column header row (matches excel_parser COLUMN_MAPPINGS + screenshot columns)
COLUMNS = [
    "Asset/Hostname",
    "Instant purpose",
    "VAPT Status",
    "Critical",
    "High",
    "Medium",
    "Low",
    "Informational",
    "Total",
    "Tester_Name",
    "Project",
    "Sr no.",
    "Observation",
    "Severity",
    "Status",
    "New or Repeat Observation",
    "CVE/CWE",
    "CVSS",
    "CVSS Version Ref",
    "Affected Asset i.e. IP/URL/Application etc.",
    "Observation/Vulnerability Title",
    "Detailed Observation/ Vulnerable Point",
    "Recommendation",
    "Reference",
    "Evidence / Proof of Concept",
    "Screenshot1",
    "Screenshot2",
    "Screenshot3",
    "Screenshot4",
]

# Two example rows: same asset block (counts on first row typical)
def build_excel(path: str) -> None:
    rows = [
        [
            "172.19.0.131",
            "Internal pentest",
            "Completed",
            0,
            2,
            1,
            0,
            0,
            3,
            "Analyst",
            "Demo Project",
            "VUL-001",
            "",
            "High",
            "Open",
            "New",
            "CWE-200",
            7.5,
            "CVSS:3.1/AV:N/AC:L/PR:N/UI:N/S:U/C:H/I:N/A:N",
            "https://172.19.0.131/app",
            "Example finding title",
            "Detailed description of the vulnerability.",
            "Apply patches and harden configuration.",
            "https://example.com/ref",
            "Step 1: Open browser.\nStep 2: Navigate to URL.\nStep 3: Observe issue.",
            "step1.png",
            "step2.png",
            "",
            "",
        ],
        [
            "172.19.0.131",
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            "VUL-002",
            "",
            "Medium",
            "Open",
            "New",
            "",
            5.3,
            "",
            "https://172.19.0.131/login",
            "Second finding",
            "Another description.",
            "Remediation text.",
            "",
            "Single step repro.",
            "step1.png",
            "",
            "",
            "",
        ],
    ]
    df = pd.DataFrame(rows, columns=COLUMNS)
    df.to_excel(path, index=False, engine="openpyxl")
    print("Wrote", path)


def build_zip(path: str) -> None:
    # Layout: one optional top folder "POC Screenshots", then IP / Severity / VUL-xxx / stepN.png
    structure = [
        ("POC Screenshots/172.19.0.131/High/VUL-001/step1.png", MINI_PNG),
        ("POC Screenshots/172.19.0.131/High/VUL-001/step2.png", MINI_PNG),
        ("POC Screenshots/172.19.0.131/Medium/VUL-002/step1.png", MINI_PNG),
    ]
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
        for name, data in structure:
            zf.writestr(name, data)
    print("Wrote", path)


def main():
    os.makedirs(HERE, exist_ok=True)
    xlsx = os.path.join(HERE, "IP_Type1_Base_Report.xlsx")
    zpath = os.path.join(HERE, "IP_Type1_POC_Sample.zip")
    build_excel(xlsx)
    build_zip(zpath)


if __name__ == "__main__":
    main()

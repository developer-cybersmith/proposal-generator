"""Verify the new TOC + colors + detail-row layout in Mitkat2_Report.docx."""
import re, zipfile
from pathlib import Path

OUT = Path(r"C:\Users\LENOVO\Downloads\Mitkat2_Report.docx")
with zipfile.ZipFile(OUT) as z:
    doc = z.read("word/document.xml").decode("utf-8")
    settings = z.read("word/settings.xml").decode("utf-8")

print("== TOC (static, no field instruction needed) ==")
print(f"  TOC field still present?  {'TOC ' in doc}")
print(f"  PAGEREF entries:          {doc.count('PAGEREF ')}")
print(f"  TOC1 / TOC2 styles used:  {doc.count('w:val=\"TOC1\"')} / {doc.count('w:val=\"TOC2\"')}")
print(f"  Hyperlink anchors used:   {doc.count('<w:hyperlink')}")
print(f"  Right-click message gone: {'Right-click' not in doc}")
print(f"  updateFields=true:        {'<w:updateFields w:val=\"true\"/>' in settings}")

print("\n== H1 / H2 headings ==")
print("  H1:")
for m in re.finditer(r'<w:pStyle w:val="Heading1"/>(.*?)</w:p>', doc, re.DOTALL):
    txt = "".join(re.findall(r'<w:t[^>]*>([^<]*)</w:t>', m.group(1)))
    print(f"    - {txt}")
print("  H2:")
for m in re.finditer(r'<w:pStyle w:val="Heading2"/>(.*?)</w:p>', doc, re.DOTALL):
    txt = "".join(re.findall(r'<w:t[^>]*>([^<]*)</w:t>', m.group(1)))
    print(f"    - {txt}")

print("\n== Table styling ==")
print(f"  Black borders (#000000):  {doc.count('w:color=\"000000\"')}")
print(f"  Navy borders (#2F5597):   {doc.count('w:color=\"2F5597\"')}")
print(f"  Light-grey fill F2F2F2:   {doc.count('w:fill=\"F2F2F2\"')}")
print(f"  Old blue header 1F3864:   {doc.count('w:fill=\"1F3864\"')}")

print("\n== Detail-row labels (must match user spec exactly) ==")
expected = [
    "S. No.",
    "Name of Vulnerability",
    "Vulnerable Point / Error Location → Vulnerable Location",
    "→ Vulnerable Path / Port / URL",
    "→ Vulnerable Parameter",
    "CVE/CWE",
    "CVSS / EPSS Score",
    "Description",
    "Proof of concept and Steps to Reproduce",
    "Workaround / Solutions / Recommendations",
    "References (Optional)",
    "Additional observations",
]
for label in expected:
    print(f"  {'OK' if label in doc else 'MISSING':>7}: {label}")

print("\n== PoC step rendering ==")
print(f"  'Step 1: ' occurrences: {doc.count('Step 1: ')}")
print(f"  'Step 2: ' occurrences: {doc.count('Step 2: ')}")
print(f"  'Step 3: ' occurrences: {doc.count('Step 3: ')}")

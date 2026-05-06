"""Verify Heading2 = Play + nested Vulnerable-Point sub-table."""
import re, zipfile
from pathlib import Path

OUT = Path(r"C:\Users\LENOVO\Downloads\Mitkat2_Report.docx")
with zipfile.ZipFile(OUT) as z:
    doc = z.read("word/document.xml").decode("utf-8")
    styles = z.read("word/styles.xml").decode("utf-8")

print("== Heading 2 font ==")
m = re.search(r'<w:style[^>]*w:styleId="Heading2"[^>]*>.*?</w:style>', styles, re.DOTALL)
if m:
    h2 = m.group(0)
    play_set = 'w:ascii="Play"' in h2 and 'w:hAnsi="Play"' in h2
    theme_stripped = "asciiTheme" not in h2 and "hAnsiTheme" not in h2
    print(f"  ascii=Play / hAnsi=Play : {play_set}")
    print(f"  theme refs stripped     : {theme_stripped}")
    rfonts = re.search(r'<w:rFonts[^/]*/>', h2)
    print(f"  rFonts element          : {rfonts.group(0) if rfonts else 'NONE'}")
else:
    print("  Heading2 style NOT FOUND")

print("\n== Detail-row labels (10-row layout) ==")
expected = [
    "S. No.",
    "Name of Vulnerability",
    "Vulnerable Point / Error Location",
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

unwanted = ["→ Vulnerable Path", "→ Vulnerable Parameter"]
print("\n  Old arrow-prefixed labels (should be GONE):")
for label in unwanted:
    print(f"  {'GONE' if label not in doc else 'STILL PRESENT':>14}: {label}")

print("\n== Nested sub-table cell labels ==")
sub_labels = ["Vulnerable Location", "Vulnerable Path/ Port/ URL", "Vulnerable Parameter"]
for label in sub_labels:
    n = doc.count(label)
    print(f"  {label!r:35} occurrences: {n} (expected: >=5 — once per vuln)")

print("\n== TOC sanity ==")
print(f"  PAGEREF entries:   {doc.count('PAGEREF ')}")
print(f"  TOC1 / TOC2 used:  {doc.count(chr(34)+'TOC1'+chr(34))} / {doc.count(chr(34)+'TOC2'+chr(34))}")

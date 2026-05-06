"""Create a copy of the user's MitKat template workbook with client_name=AcmeCorp.

This is throwaway tooling used only to prove that the rotated header text box
correctly substitutes the client name on a per-engagement basis. After running,
the demo .docx will display "AcmeCorp" vertically on the left of every page.
"""
from pathlib import Path
import shutil
from openpyxl import load_workbook

SRC = Path(r"C:\Users\LENOVO\Downloads\MitKat_Report_Template.xlsx")
OUT = Path(r"d:\CSS\Automation\backend\_demo_AcmeCorp.xlsx")

shutil.copyfile(SRC, OUT)
wb = load_workbook(OUT)
ws = wb["REPORT_META"]
for row in ws.iter_rows(min_row=2, max_col=2):
    key = (row[0].value or "").strip().lower()
    if key == "client_name":
        row[1].value = "AcmeCorp Pvt. Ltd."
    elif key == "report_title":
        row[1].value = "Report For ONLINE PST (AcmeCorp) for Re-Empanelment"
    elif key == "test_bed_details":
        row[1].value = "AcmeCorp — 10.0.0.50"
wb.save(OUT)
print(f"Wrote {OUT}")

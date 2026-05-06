"""
Regression: AMCS / government-style tracker columns (e.g. S.NO., Severity of Issue)
must map to org format so convert_any_excel_to_format() yields rows.
"""

import os
import tempfile
import unittest

import pandas as pd


class TestExcelConverterAmcs(unittest.TestCase):
    def test_amcs_style_columns_produce_rows(self):
        from Automation.backend import excel_converter as ec

        data = {
            "S.NO.": ["1", "2"],
            "Name of Auditee organization": ["OrgA", "OrgA"],
            "Audited infrastructure details": ["", ""],
            "Type of Audit": ["Web", "Web"],
            "N/A1": ["", ""],
            "Type of Audit Report(First Audit Report/Follow up Report)": [
                "First Audit Report",
                "First Audit Report",
            ],
            "Name of Vulnerability / Security Issues as per CVE/CWE/Mitre DB Only": [
                "SQLi in login",
                "XSS",
            ],
            "Description": ["d1", "d2"],
            "Vulnerable Parameter": ["p1", "p2"],
            "CVE/CWE Reference Number": ["CWE-89", "CWE-79"],
            "CVSS": ["7.5", "6.1"],
            "Severity of Issue": ["High", "Medium"],
            "Cumulative count of no. of occurrence of vulnerability in audit. <Number>": [
                "1",
                "1",
            ],
            "Attributing Factor ": ["Design", "Config"],
            "Recommendation": ["r1", "r2"],
            "POC": ["step1", "step1"],
            "Comments": ["", ""],
            "Tester Comments": ["", ""],
        }
        df = pd.DataFrame(data)

        path = os.path.join(tempfile.gettempdir(), f"amcs_test_{os.getpid()}.xlsx")
        out = None
        try:
            with pd.ExcelWriter(path, engine="openpyxl") as xw:
                df.to_excel(xw, sheet_name="AMCS Route Planner", index=False)

            out, _ = ec.convert_any_excel_to_format(path, poc_zip_path=None, poc_layout="certin")
            self.assertTrue(os.path.exists(out), "Converted output not written")
            out_df = pd.read_excel(out)
            self.assertGreaterEqual(len(out_df), 1, "Expected at least one output row")
            self.assertIn("Severity", out_df.columns)
        finally:
            try:
                os.unlink(path)
            except OSError:
                pass
            try:
                if out and os.path.exists(out):
                    os.unlink(out)
            except OSError:
                pass


if __name__ == "__main__":
    unittest.main()

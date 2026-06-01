"""
Regression: Web/IP report TOC injection and chart palette (report_insert_sections).
"""

import os
import tempfile
import unittest
import zipfile


class TestWebReportParity(unittest.TestCase):
    def test_pptx_severity_hex_order(self):
        from Automation.backend.report_insert_sections import PPTX_SEVERITY_HEX

        self.assertEqual(
            [h for _, h in PPTX_SEVERITY_HEX],
            ["BA1A16", "FE0000", "FFC000", "05B050", "0070C0"],
        )

    def test_toc_injection_split_vulnerabilities_found(self):
        from Automation.backend.report_insert_sections import (
            _inject_summary_toc_into_document_xml,
        )

        # Split <w:t> for "Vulnerabilities Found" (Word sometimes splits runs)
        toc_tail = (
            '<w:p><w:pPr><w:pStyle w:val="WPSOffice1"/></w:pPr>'
            '<w:hyperlink w:anchor="_Toc99" w:history="1">'
            '<w:r><w:rPr><w:b/></w:rPr><w:t>Vulnerabilities </w:t></w:r>'
            '<w:r><w:rPr><w:b/></w:rPr><w:t>Found</w:t></w:r>'
            '<w:r><w:rPr><w:b/></w:rPr><w:tab/></w:r>'
            '<w:r><w:rPr><w:b/></w:rPr><w:fldChar w:fldCharType="begin"/></w:r>'
            '<w:r><w:rPr><w:b/></w:rPr>'
            '<w:instrText xml:space="preserve"> PAGEREF _Toc99 \\h </w:instrText></w:r>'
            '<w:r><w:rPr><w:b/></w:rPr><w:fldChar w:fldCharType="separate"/></w:r>'
            '<w:r><w:rPr><w:b/></w:rPr><w:t>9</w:t></w:r>'
            '<w:r><w:rPr><w:b/></w:rPr><w:fldChar w:fldCharType="end"/></w:r>'
            "</w:hyperlink></w:p>"
            "<w:p><w:r><w:t>BODY</w:t></w:r></w:p>"
        )
        xml = "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" + toc_tail + "</w:document>"
        out, ok = _inject_summary_toc_into_document_xml(xml, "_TocWebSummary", use_wps_toc_style=True)
        self.assertTrue(ok)
        self.assertIn("SUMMARY OF FINDINGS &amp; CONCLUSION:", out)
        self.assertIn("PAGEREF _TocWebSummary", out)
        self.assertIn("BODY", out)

    def test_toc_injection_single_w_t_vulnerabilities_found(self):
        from Automation.backend.report_insert_sections import (
            _inject_summary_toc_into_document_xml,
        )

        toc_tail = (
            '<w:p><w:pPr><w:pStyle w:val="TOC1"/></w:pPr>'
            '<w:hyperlink w:anchor="_Toc214446872" w:history="1">'
            '<w:r><w:t>Vulnerabilities Found</w:t></w:r>'
            '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
            '<w:r><w:instrText xml:space="preserve"> PAGEREF _Toc214446872 \\h </w:instrText></w:r>'
            '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
            '<w:r><w:t>9</w:t></w:r>'
            '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
            "</w:hyperlink></w:p></w:document>"
        )
        xml = "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" + toc_tail
        out, ok = _inject_summary_toc_into_document_xml(xml, "_TocWebSummary", use_wps_toc_style=False)
        self.assertTrue(ok)
        self.assertIn("SUMMARY OF FINDINGS &amp; CONCLUSION:", out)

    def test_main_imports_shared_chart_insert(self):
        """IP report path must use shared insert (no duplicate local function)."""
        from Automation.backend import main as m

        self.assertTrue(hasattr(m, "insert_overall_vulnerabilities_charts_shared"))
        self.assertTrue(callable(m.insert_overall_vulnerabilities_charts_shared))
        self.assertTrue(hasattr(m, "web_report_finalize"))

    def test_extract_capita_palette_missing_file(self):
        from Automation.backend.report_insert_sections import extract_capita_severity_palette_from_pptx

        self.assertIsNone(extract_capita_severity_palette_from_pptx("/nonexistent/does-not-exist.pptx"))

    def test_extract_capita_palette_from_minimal_pptx(self):
        from Automation.backend.report_insert_sections import extract_capita_severity_palette_from_pptx

        xml = """<?xml version="1.0" encoding="UTF-8"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <c:chart><c:plotArea><c:barChart>
    <c:ser><c:spPr><a:solidFill><a:srgbClr val="111111"/></a:solidFill></c:spPr></c:ser>
    <c:ser><c:spPr><a:solidFill><a:srgbClr val="222222"/></a:solidFill></c:spPr></c:ser>
    <c:ser><c:spPr><a:solidFill><a:srgbClr val="333333"/></a:solidFill></c:spPr></c:ser>
    <c:ser><c:spPr><a:solidFill><a:srgbClr val="444444"/></a:solidFill></c:spPr></c:ser>
    <c:ser><c:spPr><a:solidFill><a:srgbClr val="555555"/></a:solidFill></c:spPr></c:ser>
  </c:barChart></c:plotArea></c:chart>
</c:chartSpace>"""
        fd, path = tempfile.mkstemp(suffix=".pptx")
        os.close(fd)
        try:
            with zipfile.ZipFile(path, "w") as zf:
                zf.writestr("ppt/charts/chart1.xml", xml.encode("utf-8"))
            pal = extract_capita_severity_palette_from_pptx(path)
            self.assertIsNotNone(pal)
            assert pal is not None
            self.assertEqual([h for _, h in pal], ["111111", "222222", "333333", "444444", "555555"])
        finally:
            os.unlink(path)

    def test_create_combined_chart_writes_png(self):
        from collections import Counter

        from Automation.backend.report_insert_sections import create_combined_vulnerability_charts

        fd, path = tempfile.mkstemp(suffix=".png")
        os.close(fd)
        try:
            os.unlink(path)
            create_combined_vulnerability_charts(
                Counter({"Critical": 1, "High": 1}),
                path,
                display_hostname="https://example.test",
            )
            self.assertTrue(os.path.isfile(path))
            self.assertGreater(os.path.getsize(path), 2000)
        finally:
            if os.path.isfile(path):
                os.unlink(path)

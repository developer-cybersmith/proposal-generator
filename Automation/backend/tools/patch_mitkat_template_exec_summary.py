"""
Patch MitKat base template so AI Executive Summary placeholders are in the body,
not inside the Table-of-Contents SDT content control.

Idempotent-ish: safe to run multiple times; creates a backup once per run.
"""

from __future__ import annotations

import io
import zipfile
from pathlib import Path


def _read_docxml(docx_path: Path) -> str:
    with zipfile.ZipFile(docx_path, "r") as z:
        return z.read("word/document.xml").decode("utf-8", "ignore")


def _write_docxml(docx_path: Path, new_xml: str) -> None:
    buf = io.BytesIO()
    with zipfile.ZipFile(docx_path, "r") as zin:
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == "word/document.xml":
                    data = new_xml.encode("utf-8")
                zout.writestr(item, data)
    docx_path.write_bytes(buf.getvalue())


def patch_template(docx_path: Path) -> None:
    import re

    x = _read_docxml(docx_path)

    # Helper: find the start index of the paragraph (<*:p ...>) that contains a marker.
    def _para_start_before(pos: int) -> int:
        # Find nearest preceding "<" that begins a paragraph tag with any prefix.
        m = list(re.finditer(r"<[^:>]+:p\b[^>]*>", x[:pos]))
        if not m:
            raise RuntimeError("Could not locate a paragraph start before position.")
        return m[-1].start()

    # Helper: find the end index of the paragraph (inclusive) containing a marker.
    def _para_end_after(pos: int) -> int:
        end = x.find("</", pos)
        # Find the next closing paragraph tag with any prefix.
        m = re.search(r"</[^:>]+:p>", x[pos:])
        if not m:
            raise RuntimeError("Could not locate a paragraph end after position.")
        return pos + m.end()

    # 1) Remove ES placeholder block from inside TOC SDTContent.
    toc_title_pos = x.find("Table of Content</")
    if toc_title_pos == -1:
        raise RuntimeError("Could not find TOC title in template.")
    sdt_content_start = x.find(":sdtContent>", toc_title_pos)
    if sdt_content_start == -1:
        raise RuntimeError("Could not find sdtContent start.")
    sdt_content_start = x.rfind("<", 0, sdt_content_start)  # start of the opening tag
    sdt_content_end = x.find(":sdtContent>", sdt_content_start + 1)
    # The above finds the same opening tag; instead find the closing tag.
    close_idx = x.find("</", sdt_content_start)
    m_close = re.search(r"</[^:>]+:sdtContent>", x[sdt_content_start:])
    if not m_close:
        raise RuntimeError("Could not find sdtContent end.")
    sdt_content_end = sdt_content_start + m_close.start()
    sdt_content_close_end = sdt_content_start + m_close.end()

    seg = x[sdt_content_start:sdt_content_end]
    si = seg.find("{{r ES_NARR}}")
    ei = seg.find("{{ ES_CONCLUSION }}")
    if si != -1 and ei != -1 and ei > si:
        # Expand to full paragraph boundaries inside seg.
        abs_si = sdt_content_start + si
        abs_ei = sdt_content_start + ei
        cut_start = _para_start_before(abs_si)
        cut_end = _para_end_after(abs_ei)
        x = x[:cut_start] + x[cut_end:]

    # 2) Replace static ES body content under real bookmark _Toc208246552.
    es_pos = x.rfind("_Toc208246552")
    if es_pos == -1:
        raise RuntimeError("Could not find body Executive Summary bookmark (_Toc208246552).")
    es_heading_end = _para_end_after(es_pos)

    obs_pos = x.find("_Toc208246553", es_heading_end)
    if obs_pos == -1:
        raise RuntimeError("Could not find Observations bookmark (_Toc208246553).")
    obs_para_start = _para_start_before(obs_pos)

    # Build ES placeholder paragraphs using the same prefix style as surrounding paragraphs.
    # Reuse the prefix observed in the ES heading paragraph tag.
    m_pfx = re.search(r"<([^:>]+):p\b", x[_para_start_before(es_pos) : _para_start_before(es_pos) + 60])
    pfx = m_pfx.group(1) if m_pfx else "w"

    def p(line: str) -> str:
        return f"<{pfx}:p><{pfx}:r><{pfx}:t>{line}</{pfx}:t></{pfx}:r></{pfx}:p>"

    block = "".join(
        [
            p("{{r ES_NARR}}"),
            p(" "),
            p("{%p for o in ES_PHASE1 %}"),
            p("{{ o.title }}"),
            p("{{ o.description }}"),
            p("Threat Perception: {{ o.threat_perception }}"),
            p("{%p endfor %}"),
            p(" "),
            p("{{ ES_PHASE2 }}"),
            p(" "),
            p("{%p if ES_CHART %}"),
            p("{{p ES_CHART}}"),
            p("{%p endif %}"),
            p(" "),
            p("Conclusion"),
            p("{{ ES_CONCLUSION }}"),
        ]
    )

    x = x[:es_heading_end] + block + x[obs_para_start:]

    _write_docxml(docx_path, x)


def main() -> None:
    docx_path = Path(__file__).resolve().parents[1] / "Mitkat_Base_template.docx"
    backup_path = docx_path.with_suffix(".docx.tocfix.bak")
    backup_path.write_bytes(docx_path.read_bytes())
    patch_template(docx_path)
    print(f"Patched template: {docx_path}")
    print(f"Backup written:  {backup_path}")


if __name__ == "__main__":
    main()


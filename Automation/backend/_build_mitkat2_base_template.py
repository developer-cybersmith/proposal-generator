"""
Builds Automation/backend/Mitkat2_Base_Template.docx from three sources:

  1. User's MitKat layout template (Calibri styles, numbering, theme,
     fontTable, settings) — `c:\\Users\\LENOVO\\Downloads\\Mitkat2_Base_Template.docx`.
  2. Silbern reference's *header* and *footer* XML (rotated client name on
     the left + right-aligned bordered page-number) plus the MitKat logo
     image used by that header — `c:\\Users\\LENOVO\\Downloads\\OLPST Mar 2026 - MitKat Report Silbern (PP).docx`.
  3. A blank first-page header/footer pair so the cover stays clean.

Run:  python _build_mitkat2_base_template.py
Outputs:  Mitkat2_Base_Template.docx (next to mitkat2.py)
"""
from __future__ import annotations

import zipfile
from pathlib import Path

USER_TPL = Path(r"c:\Users\LENOVO\Downloads\Mitkat2_Base_Template.docx")
SILBERN  = Path(r"c:\Users\LENOVO\Downloads\OLPST Mar 2026 - MitKat Report Silbern (PP).docx")
OUT      = Path(__file__).resolve().parent / "Mitkat2_Base_Template.docx"


# -- Empty body. titlePg=true so the *first* page (cover) gets a blank header/
# footer (rId=12 / rId=13) and every subsequent page gets the inherited
# header1.xml / footer1.xml. Rels ids must match document.xml.rels below.
NEW_DOCUMENT_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"
            xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex"
            xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
            xmlns:o="urn:schemas-microsoft-com:office:office"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
            xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"
            xmlns:v="urn:schemas-microsoft-com:vml"
            xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"
            xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
            xmlns:w10="urn:schemas-microsoft-com:office:word"
            xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
            xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"
            xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"
            xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk"
            xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml"
            xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
            mc:Ignorable="w14 w15 wp14">
  <w:body>
    <w:p/>
    <w:sectPr>
      <w:headerReference w:type="first"   r:id="rId12"/>
      <w:footerReference w:type="first"   r:id="rId13"/>
      <w:headerReference w:type="default" r:id="rId10"/>
      <w:footerReference w:type="default" r:id="rId11"/>
      <w:pgSz w:w="11906" w:h="16838"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1134" w:left="1440"
               w:header="708" w:footer="116" w:gutter="0"/>
      <w:cols w:space="708"/>
      <w:docGrid w:linePitch="360"/>
      <w:titlePg/>
    </w:sectPr>
  </w:body>
</w:document>
"""


NEW_DOCUMENT_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml" Target="../customXml/item1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
  <Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings" Target="webSettings.xml"/>
  <Relationship Id="rId6" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" Target="fontTable.xml"/>
  <Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
  <Relationship Id="rId10" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>
  <Relationship Id="rId11" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>
  <Relationship Id="rId12" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header2.xml"/>
  <Relationship Id="rId13" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer2.xml"/>
</Relationships>
"""


CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="emf" ContentType="image/x-emf"/>
  <Default Extension="jpeg" ContentType="image/jpeg"/>
  <Default Extension="jpg" ContentType="image/jpeg"/>
  <Default Extension="png" ContentType="image/png"/>
  <Default Extension="gif" ContentType="image/gif"/>
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/customXml/itemProps1.xml" ContentType="application/vnd.openxmlformats-officedocument.customXmlProperties+xml"/>
  <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
  <Override PartName="/word/webSettings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml"/>
  <Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
  <Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>
  <Override PartName="/word/header2.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
  <Override PartName="/word/footer2.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>
  <Override PartName="/word/fontTable.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/>
  <Override PartName="/word/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>
"""


# Blank first-page header/footer (cover gets clean look)
BLANK_HEADER = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p><w:pPr><w:pStyle w:val="Header"/></w:pPr></w:p>
</w:hdr>
"""

BLANK_FOOTER = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p><w:pPr><w:pStyle w:val="Footer"/></w:pPr></w:p>
</w:ftr>
"""


def _read(z: zipfile.ZipFile, name: str) -> bytes:
    return z.read(name)


def main() -> None:
    if not USER_TPL.exists():
        raise SystemExit(f"Missing user template: {USER_TPL}")
    if not SILBERN.exists():
        raise SystemExit(f"Missing Silbern reference: {SILBERN}")
    if OUT.exists():
        OUT.unlink()

    with zipfile.ZipFile(USER_TPL) as ztpl, zipfile.ZipFile(SILBERN) as zsil:
        # Files we copy verbatim from each source
        from_user = {
            "_rels/.rels":                  _read(ztpl, "_rels/.rels"),
            "word/styles.xml":              _read(ztpl, "word/styles.xml"),
            "word/theme/theme1.xml":        _read(ztpl, "word/theme/theme1.xml"),
            "word/fontTable.xml":           _read(ztpl, "word/fontTable.xml"),
            "word/numbering.xml":           _read(ztpl, "word/numbering.xml"),
            "word/settings.xml":            _read(ztpl, "word/settings.xml"),
            "word/webSettings.xml":         _read(ztpl, "word/webSettings.xml"),
            "docProps/app.xml":             _read(ztpl, "docProps/app.xml"),
            "docProps/core.xml":            _read(ztpl, "docProps/core.xml"),
            "docProps/thumbnail.jpeg":      _read(ztpl, "docProps/thumbnail.jpeg"),
            "customXml/_rels/item1.xml.rels": _read(ztpl, "customXml/_rels/item1.xml.rels"),
            "customXml/item1.xml":          _read(ztpl, "customXml/item1.xml"),
            "customXml/itemProps1.xml":     _read(ztpl, "customXml/itemProps1.xml"),
        }
        from_silbern = {
            "word/header1.xml":             _read(zsil, "word/header1.xml"),
            "word/footer1.xml":             _read(zsil, "word/footer1.xml"),
            "word/_rels/header1.xml.rels":  _read(zsil, "word/_rels/header1.xml.rels"),
            "word/media/image1.emf":        _read(zsil, "word/media/image1.emf"),
        }

    with zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in from_user.items():
            zout.writestr(name, data)
        for name, data in from_silbern.items():
            zout.writestr(name, data)
        # Blank first-page header/footer
        zout.writestr("word/header2.xml", BLANK_HEADER)
        zout.writestr("word/footer2.xml", BLANK_FOOTER)
        # New document + rels + content-types
        zout.writestr("[Content_Types].xml", CONTENT_TYPES)
        zout.writestr("word/document.xml", NEW_DOCUMENT_XML)
        zout.writestr("word/_rels/document.xml.rels", NEW_DOCUMENT_RELS)

    print(f"Wrote {OUT} ({OUT.stat().st_size:,} bytes)")


if __name__ == "__main__":
    main()

#!/usr/bin/env python3
"""Generate phase1.xlsx covering all Phase 1 metadata: merged cells,
hyperlinks (external + internal), sheet protection, row/col options,
sheetFormatPr defaults, sheet visibility (hidden), and defined names."""

import sys
import zipfile

OUT = sys.argv[1] if len(sys.argv) > 1 else "phase1.xlsx"

CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml"  ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>
"""

ROOT_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>
"""

WORKBOOK = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Visible"     sheetId="1" r:id="rId1"/>
    <sheet name="Hidden"       sheetId="2" state="hidden"     r:id="rId2"/>
  </sheets>
  <definedNames>
    <definedName name="GlobalArea">Visible!$A$1:$B$2</definedName>
    <definedName name="LocalToHidden" localSheetId="1">Hidden!$C$3</definedName>
    <definedName name="HiddenDef" hidden="1">Visible!$Z$99</definedName>
  </definedNames>
</workbook>
"""

WORKBOOK_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>
"""

STYLES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
  <fills count="1"><fill><patternFill patternType="none"/></fill></fills>
  <borders count="1"><border/></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>
</styleSheet>
"""

# Sheet1: rich metadata
SHEET1 = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
           xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetFormatPr defaultRowHeight="18.5" defaultColWidth="9.25"/>
  <cols>
    <col min="2" max="2" width="22.5" customWidth="1"/>
    <col min="3" max="4" width="11.0" hidden="1"/>
    <col min="5" max="5" width="8.0"  outlineLevel="2"/>
  </cols>
  <sheetData>
    <row r="1"><c r="A1"><v>1</v></c></row>
    <row r="2" ht="30.5" customHeight="1"><c r="A2"><v>2</v></c></row>
    <row r="3" hidden="1"><c r="A3"><v>3</v></c></row>
    <row r="4" outlineLevel="2"><c r="A4"><v>4</v></c></row>
    <row r="5"><c r="A5"><v>5</v></c></row>
  </sheetData>
  <sheetProtection password="DAA7" sheet="1" objects="1" scenarios="1" formatCells="0"/>
  <mergeCells count="2">
    <mergeCell ref="A1:C1"/>
    <mergeCell ref="D2:E5"/>
  </mergeCells>
  <hyperlinks>
    <hyperlink ref="A5" r:id="rId1" tooltip="Click me" display="ext"/>
    <hyperlink ref="B5" location="Hidden!A1" display="internal"/>
  </hyperlinks>
</worksheet>
"""

SHEET1_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com/x" TargetMode="External"/>
</Relationships>
"""

SHEET2 = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData/>
</worksheet>
"""

with zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED) as z:
    z.writestr("[Content_Types].xml", CONTENT_TYPES)
    z.writestr("_rels/.rels", ROOT_RELS)
    z.writestr("xl/workbook.xml", WORKBOOK)
    z.writestr("xl/_rels/workbook.xml.rels", WORKBOOK_RELS)
    z.writestr("xl/styles.xml", STYLES)
    z.writestr("xl/worksheets/sheet1.xml", SHEET1)
    z.writestr("xl/worksheets/_rels/sheet1.xml.rels", SHEET1_RELS)
    z.writestr("xl/worksheets/sheet2.xml", SHEET2)

print("wrote", OUT)

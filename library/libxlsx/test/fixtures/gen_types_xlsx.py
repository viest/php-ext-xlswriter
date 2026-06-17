#!/usr/bin/env python3
"""Generate a comprehensive XLSX fixture covering every cell type the
   reader's type-inference code paths exercise. Output is deterministic so
   tests can match exact values."""

import os
import sys
import zipfile

OUT = sys.argv[1] if len(sys.argv) > 1 else "types.xlsx"

CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
</Types>
"""

ROOT_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>
"""

WORKBOOK_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Types" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>
"""

WORKBOOK_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
</Relationships>
"""

THEME = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Reader Test Theme">
  <a:themeElements>
    <a:clrScheme name="Reader Test">
      <a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
      <a:dk1><a:srgbClr val="000000"/></a:dk1>
      <a:lt2><a:srgbClr val="EEEEEE"/></a:lt2>
      <a:dk2><a:srgbClr val="111111"/></a:dk2>
      <a:accent1><a:srgbClr val="112233"/></a:accent1>
      <a:accent2><a:srgbClr val="445566"/></a:accent2>
      <a:accent3><a:srgbClr val="778899"/></a:accent3>
      <a:accent4><a:srgbClr val="AABBCC"/></a:accent4>
      <a:accent5><a:srgbClr val="DDEEFF"/></a:accent5>
      <a:accent6><a:srgbClr val="123456"/></a:accent6>
      <a:hlink><a:srgbClr val="0000FF"/></a:hlink>
      <a:folHlink><a:srgbClr val="800080"/></a:folHlink>
    </a:clrScheme>
  </a:themeElements>
</a:theme>
"""

# SST: index 0 = "hello", 1 = "world", 2 = "rich-text run"
SHARED_STRINGS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="3" uniqueCount="3">
  <si><t>hello</t></si>
  <si><t xml:space="preserve">  world  </t></si>
  <si><r><t>rich-</t></r><r><t>text</t></r><r><t> run</t></r></si>
</sst>
"""

# styles:
#  xf[0]: numFmtId=0  (General)
#  xf[1]: numFmtId=14 (date m/d/yyyy)
#  xf[2]: numFmtId=22 (datetime)
#  xf[3]: numFmtId=164 (custom date "yyyy-mm-dd")
#  xf[4]: numFmtId=9  (percent)
#  xf[5]: theme font color
#  xf[6]: theme+tint fill color
#  xf[7]: theme+tint border color
#  dxf[0]: conditional-format style with font/fill/border colors
STYLES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <numFmts count="1">
    <numFmt numFmtId="164" formatCode="yyyy-mm-dd"/>
  </numFmts>
  <fonts count="3">
    <font><sz val="11"/><name val="Calibri"/></font>
    <font><sz val="11"/><name val="Calibri"/><color theme="4"/></font>
    <font><sz val="11"/><name val="Calibri"/><color indexed="10"/></font>
  </fonts>
  <fills count="2">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="solid"><fgColor theme="5" tint="0.5"/></patternFill></fill>
  </fills>
  <borders count="2">
    <border/>
    <border><left style="thin"><color theme="6" tint="-0.5"/></left></border>
  </borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="8">
    <xf numFmtId="0"   fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="14"  fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
    <xf numFmtId="22"  fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
    <xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
    <xf numFmtId="9"   fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
    <xf numFmtId="0"   fontId="1" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0"   fontId="2" fillId="1" borderId="0" xfId="0"/>
    <xf numFmtId="0"   fontId="0" fillId="0" borderId="1" xfId="0"/>
  </cellXfs>
  <dxfs count="1">
    <dxf>
      <font><b/><color theme="4"/></font>
      <fill><patternFill patternType="solid"><fgColor indexed="10"/></patternFill></fill>
      <border><bottom style="thin"><color rgb="FF010203"/></bottom></border>
    </dxf>
  </dxfs>
</styleSheet>
"""

# Sheet covers:
#   Row 1: A1 number 42, B1 number 3.14
#   Row 2: A2 SST "hello" (idx 0), B2 SST "world" (idx 1)
#   Row 3: A3 inline string "inline!", B3 boolean true
#   Row 4: A4 boolean false, B4 error #DIV/0!
#   Row 5: A5 formula SUM(A1:A2)=44, with cached <v>
#   Row 6: A6 date serial 44927 styled xf[1] (m/d/yyyy) -> 2023-01-01
#   Row 7: A7 datetime serial 44927.5 styled xf[2] -> 2023-01-01 12:00
#   Row 8: A8 custom-format date 45292 styled xf[3] -> 2024-01-01
#   Row 9: A9 percent 0.5 styled xf[4]
#   Row 10: A10 SST rich-text "rich-text run"
SHEET = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1"><v>42</v></c><c r="B1"><v>3.14</v></c></row>
    <row r="2"><c r="A2" t="s"><v>0</v></c><c r="B2" t="s"><v>1</v></c></row>
    <row r="3"><c r="A3" t="inlineStr"><is><t>inline!</t></is></c><c r="B3" t="b"><v>1</v></c></row>
    <row r="4"><c r="A4" t="b"><v>0</v></c><c r="B4" t="e"><v>#DIV/0!</v></c></row>
    <row r="5"><c r="A5"><f>SUM(A1:A2)</f><v>44</v></c></row>
    <row r="6"><c r="A6" s="1"><v>44927</v></c></row>
    <row r="7"><c r="A7" s="2"><v>44927.5</v></c></row>
    <row r="8"><c r="A8" s="3"><v>45292</v></c></row>
    <row r="9"><c r="A9" s="4"><v>0.5</v></c></row>
    <row r="10"><c r="A10" t="s"><v>2</v></c></row>
  </sheetData>
</worksheet>
"""

with zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED) as z:
    z.writestr("[Content_Types].xml", CONTENT_TYPES)
    z.writestr("_rels/.rels", ROOT_RELS)
    z.writestr("xl/workbook.xml", WORKBOOK_XML)
    z.writestr("xl/_rels/workbook.xml.rels", WORKBOOK_RELS)
    z.writestr("xl/sharedStrings.xml", SHARED_STRINGS)
    z.writestr("xl/styles.xml", STYLES)
    z.writestr("xl/theme/theme1.xml", THEME)
    z.writestr("xl/worksheets/sheet1.xml", SHEET)

print("wrote", OUT)

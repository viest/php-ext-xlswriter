#!/usr/bin/env python3
"""Generate phase3.xlsx covering Phase 3 reader features:
   - <f t="array" ref="A1:B2">                      (array formula)
   - <f t="shared" si="0" ref="A3:A5">              (shared master)
   - <f t="shared" si="0">                          (shared follower; no text)
   - <dataValidations> with two <dataValidation> entries
   - <autoFilter ref> with one <filterColumn> using <filters> (list)
     and another using <customFilters>."""

import sys, zipfile

OUT = sys.argv[1] if len(sys.argv) > 1 else "phase3.xlsx"

CT = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml"  ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>
'''

ROOT = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>
'''

WB = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="P3" sheetId="1" r:id="rId1"/></sheets>
</workbook>
'''

WB_RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>
'''

STYLES = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
  <fills count="1"><fill><patternFill patternType="none"/></fill></fills>
  <borders count="1"><border/></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>
</styleSheet>
'''

# Layout:
#   A1:B2   array formula =TRANSPOSE(C1:C2)        — master at A1
#   A3:A5   shared formula =B3*2 (master at A3, si=0)
#                  followers A4 / A5 — no text, si="0" only
#   C1, C2  array source values (10, 20)
#   B3, B4, B5 multipliers
#   D2:D8   data range that the dataValidations cover
SHEET = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1"><f t="array" ref="A1:B2">TRANSPOSE(C1:C2)</f><v>10</v></c>
      <c r="B1"><v>20</v></c>
      <c r="C1"><v>10</v></c>
    </row>
    <row r="2">
      <c r="A2"></c>
      <c r="B2"></c>
      <c r="C2"><v>20</v></c>
      <c r="D2"><v>5</v></c>
    </row>
    <row r="3">
      <c r="A3"><f t="shared" ref="A3:A5" si="0">B3*2</f><v>2</v></c>
      <c r="B3"><v>1</v></c>
      <c r="D3"><v>50</v></c>
    </row>
    <row r="4">
      <c r="A4"><f t="shared" si="0"/><v>4</v></c>
      <c r="B4"><v>2</v></c>
      <c r="D4"><v>30</v></c>
    </row>
    <row r="5">
      <c r="A5"><f t="shared" si="0"/><v>6</v></c>
      <c r="B5"><v>3</v></c>
      <c r="D5"><v>40</v></c>
    </row>
    <row r="6"><c r="D6"><v>20</v></c></row>
    <row r="7"><c r="D7"><v>60</v></c></row>
    <row r="8"><c r="D8"><v>10</v></c></row>
  </sheetData>
  <autoFilter ref="A1:D8">
    <filterColumn colId="0">
      <filters><filter val="A"/><filter val="B"/></filters>
    </filterColumn>
    <filterColumn colId="3">
      <customFilters and="1">
        <customFilter operator="greaterThan" val="20"/>
        <customFilter operator="lessThan"    val="60"/>
      </customFilters>
    </filterColumn>
  </autoFilter>
  <dataValidations count="2">
    <dataValidation type="whole" operator="between" allowBlank="1"
                    showInputMessage="1" showErrorMessage="1"
                    errorStyle="stop"
                    promptTitle="Range" prompt="Enter 1..100"
                    errorTitle="Bad" error="Out of range"
                    sqref="D2:D8">
      <formula1>1</formula1>
      <formula2>100</formula2>
    </dataValidation>
    <dataValidation type="list" allowBlank="1" sqref="C1:C2">
      <formula1>"alpha,beta,gamma"</formula1>
    </dataValidation>
  </dataValidations>
</worksheet>
'''

with zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED) as z:
    z.writestr("[Content_Types].xml", CT)
    z.writestr("_rels/.rels", ROOT)
    z.writestr("xl/workbook.xml", WB)
    z.writestr("xl/_rels/workbook.xml.rels", WB_RELS)
    z.writestr("xl/styles.xml", STYLES)
    z.writestr("xl/worksheets/sheet1.xml", SHEET)

print("wrote", OUT)

#!/usr/bin/env python3
"""Generate phase4.xlsx covering Phase 4 reader features that the writer
   API can't (yet) emit — specifically threaded comments + a richer page
   setup block, plus an SST item with multiple style runs."""

import sys, zipfile

OUT = sys.argv[1] if len(sys.argv) > 1 else "phase4.xlsx"

CT = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml"  ContentType="application/xml"/>
  <Default Extension="vml"  ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
  <Override PartName="/xl/comments1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"/>
  <Override PartName="/xl/threadedComments/threadedComment1.xml" ContentType="application/vnd.ms-excel.threadedcomments+xml"/>
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
  <sheets><sheet name="P4" sheetId="1" r:id="rId1"/></sheets>
</workbook>
'''

WB_RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
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

# Two SST items: A1 has multi-run rich text. A2 is plain.
SST = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="2" uniqueCount="2">
  <si>
    <r><rPr><b/><sz val="12"/><color rgb="FFFF0000"/><rFont val="Arial"/></rPr><t>Hello </t></r>
    <r><rPr><i/><sz val="14"/><rFont val="Calibri"/></rPr><t>world</t></r>
  </si>
  <si><t>plain</t></si>
</sst>
'''

SHEET = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1" t="s"><v>0</v></c></row>
    <row r="2"><c r="A2" t="s"><v>1</v></c></row>
    <row r="3"><c r="A3"><v>10</v></c></row>
  </sheetData>
  <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
  <pageSetup paperSize="9" orientation="landscape" scale="125" fitToWidth="1" fitToHeight="0"/>
  <printOptions horizontalCentered="1" gridLines="1" headings="1"/>
  <headerFooter differentOddEven="1" scaleWithDoc="1" alignWithMargins="1">
    <oddHeader>&amp;COdd Header</oddHeader>
    <oddFooter>&amp;ROdd Footer</oddFooter>
    <evenHeader>&amp;CEven Header</evenHeader>
    <evenFooter>&amp;REven Footer</evenFooter>
  </headerFooter>
  <legacyDrawing r:id="rId1" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>
</worksheet>
'''

SHEET_RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing"        Target="../drawings/vmlDrawing1.vml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"          Target="../comments1.xml"/>
  <Relationship Id="rId3" Type="http://schemas.microsoft.com/office/2017/10/relationships/threadedComment"             Target="../threadedComments/threadedComment1.xml"/>
</Relationships>
'''

VML = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">
  <v:shape>
    <v:fill color2="#ffffe1"/>
    <x:ClientData ObjectType="Note">
      <x:Visible/>
      <x:Anchor>0,0,0,0,1,0,0,0</x:Anchor>
      <x:Row>0</x:Row>
      <x:Column>0</x:Column>
    </x:ClientData>
  </v:shape>
  <v:shape>
    <v:fill color2="#ffffe1"/>
    <x:ClientData ObjectType="Note">
      <x:Anchor>0,0,0,0,1,0,0,0</x:Anchor>
      <x:Row>1</x:Row>
      <x:Column>2</x:Column>
    </x:ClientData>
  </v:shape>
</xml>
'''

COMMENTS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <authors>
    <author>Eve</author>
    <author>Bob</author>
  </authors>
  <commentList>
    <comment ref="A1" authorId="0">
      <text><r><t>From Eve, visible</t></r></text>
    </comment>
    <comment ref="C2" authorId="1">
      <text><r><t>From Bob, hidden</t></r></text>
    </comment>
  </commentList>
</comments>
'''

THREADED = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<ThreadedComments xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments">
  <threadedComment ref="A3" id="{aaaaaaaa}" personId="{p1}">
    <text>Top-level threaded comment</text>
  </threadedComment>
  <threadedComment ref="A3" id="{bbbbbbbb}" personId="{p2}" parentId="{aaaaaaaa}">
    <text>Reply to it</text>
  </threadedComment>
</ThreadedComments>
'''

with zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED) as z:
    z.writestr("[Content_Types].xml", CT)
    z.writestr("_rels/.rels", ROOT)
    z.writestr("xl/workbook.xml", WB)
    z.writestr("xl/_rels/workbook.xml.rels", WB_RELS)
    z.writestr("xl/styles.xml", STYLES)
    z.writestr("xl/sharedStrings.xml", SST)
    z.writestr("xl/worksheets/sheet1.xml", SHEET)
    z.writestr("xl/worksheets/_rels/sheet1.xml.rels", SHEET_RELS)
    z.writestr("xl/drawings/vmlDrawing1.vml", VML)
    z.writestr("xl/comments1.xml", COMMENTS)
    z.writestr("xl/threadedComments/threadedComment1.xml", THREADED)

print("wrote", OUT)

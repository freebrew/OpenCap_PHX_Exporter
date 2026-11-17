import { strToU8, zipSync } from "fflate";

function part(path: string, xml: string): [string, Uint8Array] {
  return [path, strToU8(xml)];
}

/** Minimal xlsx with a shared formula, an array formula, a defined name, and protection. */
export function sharedFormulaFixture(): Uint8Array {
  const files: Record<string, Uint8Array> = Object.fromEntries([
    part("[Content_Types].xml", `<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
</Types>`),
    part("_rels/.rels", `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`),
    part("xl/workbook.xml", `<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Calc" sheetId="1" r:id="rId1"/>
  </sheets>
  <definedNames>
    <definedName name="Answer">Calc!$C$1</definedName>
    <definedName name="HiddenConst" hidden="1">1+2</definedName>
    <definedName name="Blob">{1,2,3}</definedName>
  </definedNames>
</workbook>`),
    part("xl/_rels/workbook.xml.rels", `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
</Relationships>`),
    part("xl/sharedStrings.xml", `<?xml version="1.0" encoding="UTF-8"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1"><si><t>Title</t></si></sst>`),
    part("xl/styles.xml", `<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
  <fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
  <borders count="1"><border/></borders>
  <cellXfs count="2">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" applyProtection="1"><protection locked="0"/></xf>
  </cellXfs>
</styleSheet>`),
    part("xl/worksheets/sheet1.xml", `<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetViews><sheetView workbookViewId="0"><pane ySplit="1" state="frozen" topLeftCell="A2"/></sheetView></sheetViews>
  <cols><col min="1" max="1" width="18" customWidth="1"/></cols>
  <sheetData>
    <row r="1">
      <c r="A1" t="s"><v>0</v></c>
      <c r="B1"><v>10</v></c>
      <c r="C1"><f t="array" ref="C1">SUM(B1:B3)</f><v>60</v></c>
    </row>
    <row r="2">
      <c r="A2" s="1"><v>2</v></c>
      <c r="B2"><v>20</v></c>
      <c r="C2"><f t="shared" ref="C2:C3" si="0">A2*B2</f><v>40</v></c>
    </row>
    <row r="3">
      <c r="B3"><v>30</v></c>
      <c r="C3"><f t="shared" si="0"/><v>0</v></c>
      <c r="D3"><f>[2]Other!A1</f><v>1</v></c>
    </row>
  </sheetData>
  <mergeCells count="1"><mergeCell ref="A1:A1"/></mergeCells>
  <dataValidations count="1"><dataValidation type="list" allowBlank="1" sqref="A2"><formula1>"Yes,No"</formula1></dataValidation></dataValidations>
  <conditionalFormatting sqref="B2"><cfRule type="cellIs" operator="greaterThan" priority="1"><formula>5</formula></cfRule></conditionalFormatting>
  <sheetProtection sheet="1" objects="1" password="secret"/>
</worksheet>`),
  ]);
  return zipSync(files);
}

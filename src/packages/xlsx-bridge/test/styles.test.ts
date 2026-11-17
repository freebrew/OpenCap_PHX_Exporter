import { strToU8, zipSync } from "fflate";
import { describe, expect, it } from "vitest";
import { DATA_VALIDATION_RESOURCE, readWorkbook, STYLE_VERSION } from "../src/readWorkbook.js";
import { restyleWorkbook } from "../src/restyle.js";

const THEME = `<a:theme xmlns:a="a"><a:themeElements><a:clrScheme name="Office">
<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
<a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>
<a:accent1><a:srgbClr val="4472C4"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2>
<a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4>
<a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6>
<a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink>
</a:clrScheme></a:themeElements></a:theme>`;

const STYLES = `<styleSheet>
<fonts count="2"><font><sz val="11"/><color theme="1"/><name val="Calibri"/></font><font><b/><u/><sz val="12"/><color theme="0"/><name val="Arial"/></font></fonts>
<fills count="4"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>
<fill><patternFill patternType="solid"><fgColor theme="0" tint="-0.14999847407452621"/><bgColor indexed="64"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor theme="9" tint="0.79998168889431442"/><bgColor indexed="64"/></patternFill></fill></fills>
<borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border>
<border><left style="medium"><color indexed="64"/></left><right style="thin"><color rgb="FFFF0000"/></right><top/><bottom style="double"><color theme="1"/></bottom><diagonal/></border></borders>
<cellXfs count="3"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/><xf numFmtId="0" fontId="0" fillId="2" borderId="1"/><xf numFmtId="0" fontId="1" fillId="3" borderId="0"/></cellXfs>
</styleSheet>`;

const SHEET = `<worksheet><sheetViews><sheetView showGridLines="0" zoomScale="70" workbookViewId="0"/></sheetViews>
<sheetFormatPr defaultRowHeight="15"/><cols><col min="2" max="2" width="34.109375" customWidth="1"/></cols>
<sheetData><row r="1"><c r="A1" s="1" t="inlineStr"><is><t>Label</t></is></c><c r="B1" s="2"><v>5</v></c><c r="C1"><f>B1*2</f><v>10</v></c></row></sheetData>
<dataValidations count="3">
<dataValidation type="list" allowBlank="1" showErrorMessage="1" sqref="A2:A5 C2"><formula1>"Right, Left, East, West"</formula1></dataValidation>
<dataValidation type="list" allowBlank="1" sqref="B2"><formula1>$AO$10:$AO$12</formula1></dataValidation>
<dataValidation type="list" sqref="Y12:Y505"><formula1>#REF!</formula1></dataValidation>
</dataValidations>
<extLst><ext uri="{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}" xmlns:x14="x14"><x14:dataValidations count="1" xmlns:xm="xm">
<x14:dataValidation type="list" allowBlank="1" showErrorMessage="1"><x14:formula1><xm:f>'DD Tools'!$O$6:$O$26</xm:f></x14:formula1><xm:sqref>H3</xm:sqref></x14:dataValidation>
</x14:dataValidations></ext></extLst></worksheet>`;

function workbookBytes(): Uint8Array {
  return zipSync({
    "[Content_Types].xml": strToU8(`<Types/>`),
    "xl/workbook.xml": strToU8(`<workbook><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>`),
    "xl/_rels/workbook.xml.rels": strToU8(`<Relationships><Relationship Id="rId1" Type="worksheet" Target="worksheets/sheet1.xml"/></Relationships>`),
    "xl/worksheets/sheet1.xml": strToU8(SHEET),
    "xl/styles.xml": strToU8(STYLES),
    "xl/theme/theme1.xml": strToU8(THEME),
  });
}

describe("cell appearance", () => {
  it("maps Excel theme slots, tints, borders, gridlines, zoom, and widths", async () => {
    const mapped = await readWorkbook(workbookBytes());
    const sheet = Object.values(mapped.workbook.sheets)[0];
    const styleOf = (row: number, col: number) => mapped.workbook.styles[String(sheet.cellData[row][col].s)] as Record<string, any>;

    const header = styleOf(0, 0);
    expect(header.cl).toEqual({ rgb: "#000000" });
    expect(header.bg).toEqual({ rgb: "#D9D9D9" });
    expect(header.bd).toEqual({
      l: { s: 8, cl: { rgb: "#000000" } },
      r: { s: 1, cl: { rgb: "#FF0000" } },
      b: { s: 7, cl: { rgb: "#000000" } },
    });

    const accent = styleOf(0, 1);
    expect(accent.cl).toEqual({ rgb: "#FFFFFF" });
    expect(accent.bg).toEqual({ rgb: "#E2F0D9" });
    expect(accent.bl).toBe(1);
    expect(accent.ul).toEqual({ s: 1 });

    expect(sheet.showGridlines).toBe(0);
    expect(sheet.zoomRatio).toBe(0.7);
    expect(sheet.columnData[1].w).toBe(239);
    expect(mapped.sidecar.styleVersion).toBe(STYLE_VERSION);
  });

  it("hands Excel dropdowns to Univer's data-validation plugin", async () => {
    const mapped = await readWorkbook(workbookBytes());
    const resource = mapped.workbook.resources.find((item) => item.name === DATA_VALIDATION_RESOURCE);
    const rules = JSON.parse(resource?.data ?? "{}")["sheet-1"];
    expect(rules).toHaveLength(3);
    expect(rules[0]).toMatchObject({
      type: "list",
      formula1: JSON.stringify(["Right", "Left", "East", "West"]),
      ranges: [{ startRow: 1, endRow: 4, startColumn: 0, endColumn: 0 }, { startRow: 1, endRow: 1, startColumn: 2, endColumn: 2 }],
      showDropDown: true,
      renderMode: 1,
      allowBlank: true,
      errorStyle: 1,
    });
    expect(rules[1]).toMatchObject({ type: "list", formula1: "=$AO$10:$AO$12", errorStyle: 0 });
    expect(rules[2]).toMatchObject({ type: "list", formula1: "='DD Tools'!$O$6:$O$26", ranges: [{ startRow: 2, endRow: 2, startColumn: 7, endColumn: 7 }] });
    const broken = mapped.sidecar.dataValidations.find((rule) => rule.sqref === "Y12:Y505");
    expect(broken).toMatchObject({ mapped: false, reason: "list source is #REF! in Excel" });
  });

  it("restyles an older snapshot without touching values or formulas", async () => {
    const mapped = await readWorkbook(workbookBytes());
    const sheetId = Object.keys(mapped.workbook.sheets)[0];
    const old = {
      styles: { legacy: { cl: { rgb: "#FFFFFF" } } },
      sheets: {
        [sheetId]: {
          name: "Data",
          cellData: { 0: { 0: { v: "Edited", s: "legacy" }, 2: { f: "=B1*2", v: 10, s: "legacy" } } },
          mergeData: [{ startRow: 0, endRow: 0, startColumn: 3, endColumn: 4 }],
          showGridlines: 1,
          zoomRatio: 1,
        },
      },
    };
    const next = restyleWorkbook(old, mapped.workbook) as any;
    const cells = next.sheets[sheetId].cellData[0];
    expect(cells[0].v).toBe("Edited");
    expect(cells[0].s).toBe(mapped.workbook.sheets[sheetId].cellData[0][0].s);
    expect(cells[2]).toEqual({ f: "=B1*2", v: 10, s: mapped.workbook.sheets[sheetId].cellData[0][2].s });
    expect(cells[1]).toEqual({ s: mapped.workbook.sheets[sheetId].cellData[0][1].s });
    expect(next.sheets[sheetId].mergeData).toEqual(old.sheets[sheetId].mergeData);
    expect(next.sheets[sheetId].showGridlines).toBe(0);
    expect(next.styles.legacy).toBeUndefined();
    const restored = JSON.parse(next.resources.find((item: { name: string }) => item.name === DATA_VALIDATION_RESOURCE).data);
    expect(restored[sheetId]).toHaveLength(3);
  });
});

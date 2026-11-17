import { strToU8, zipSync } from "fflate";
import { describe, expect, it } from "vitest";
import { macroButtonsFromWorkbook } from "../src/macroButtons.js";

describe("macro buttons", () => {
  it("reads form-control captions and macros from each sheet", () => {
    const files: Record<string, Uint8Array> = {
      "xl/workbook.xml": strToU8(`<workbook><sheets><sheet name="Setup" sheetId="1" r:id="rId1"/><sheet name="Data" sheetId="2" r:id="rId2"/></sheets></workbook>`),
      "xl/_rels/workbook.xml.rels": strToU8(`<Relationships><Relationship Id="rId1" Type="worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="worksheet" Target="worksheets/sheet2.xml"/></Relationships>`),
      "xl/worksheets/_rels/sheet1.xml.rels": strToU8(`<Relationships><Relationship Id="rId1" Type="vmlDrawing" Target="../drawings/vmlDrawing1.vml"/></Relationships>`),
      "xl/drawings/vmlDrawing1.vml": strToU8(`<xml><v:shape id="BtnOCRefresh" style="width:20pt"><v:textbox><div><font><b>REFRESH CSVs</b></font></div></v:textbox><x:ClientData><x:Anchor>14, 1, 0, 6, 17, 50, 0, 43</x:Anchor><x:FmlaMacro>[0]!RefreshSetup</x:FmlaMacro></x:ClientData></v:shape></xml>`),
      "xl/worksheets/_rels/sheet2.xml.rels": strToU8(`<Relationships><Relationship Id="rId1" Type="vmlDrawing" Target="../drawings/vmlDrawing2.vml"/></Relationships>`),
      "xl/drawings/vmlDrawing2.vml": strToU8(`<xml><v:shape id="Email" ><v:textbox><div><b>EMAIL</b></div></v:textbox><x:FmlaMacro>[0]!RTP_printW_Email</x:FmlaMacro></v:shape><v:shape id="Picture"><v:imagedata/></v:shape></xml>`),
    };
    const buttons = macroButtonsFromWorkbook(zipSync(files));
    expect(buttons).toEqual([
      { sheet: "Setup", id: "BtnOCRefresh", caption: "REFRESH CSVs", macro: "RefreshSetup", anchor: "14, 1, 0, 6, 17, 50, 0, 43" },
      { sheet: "Data", id: "Email", caption: "EMAIL", macro: "RTP_printW_Email", anchor: "" },
    ]);
  });
});

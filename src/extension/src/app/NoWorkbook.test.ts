import { describe, expect, it } from "vitest";
import { noWorkbookReason } from "./NoWorkbook";

const open = { appJobId: "40001", label: "40001", pickedJobNumber: "", stagedFile: "", folderXlsm: "" };

describe("noWorkbookReason", () => {
  it("asks for a job when none is open", () => {
    expect(noWorkbookReason({ ...open, appJobId: "", label: "" })).toEqual({ kind: "no-job" });
  });

  it("names the folder that has no Slide Sheet", () => {
    expect(noWorkbookReason(open)).toEqual({ kind: "empty-folder", label: "40001" });
  });

  it("offers the workbook that arrived after the folder was chosen", () => {
    expect(noWorkbookReason({ ...open, folderXlsm: "Slide Sheet - 40001.xlsm" })).toEqual({ kind: "unopened", xlsm: "Slide Sheet - 40001.xlsm" });
  });

  it("points at the import panel while it waits for Continue", () => {
    expect(noWorkbookReason({ ...open, folderXlsm: "Slide Sheet - 40001.xlsm", stagedFile: "Slide Sheet - 40001.xlsm" })).toEqual({ kind: "staged", fileName: "Slide Sheet - 40001.xlsm" });
  });

  it("says a picked FieldCap job has no folder, whatever folder was open before", () => {
    expect(noWorkbookReason({ ...open, pickedJobNumber: "40002" })).toEqual({ kind: "no-folder", jobNumber: "40002" });
  });
});

import { describe, expect, it } from "vitest";
import { appJobIdOf, folderWorkbooks, matchingWorkbook, snapshotLabel, staleEntriesFor } from "./jobFolders";

describe("folderWorkbooks", () => {
  it("finds nothing in a new job folder", () => {
    expect(folderWorkbooks([])).toEqual({ snapshot: null, xlsm: null });
    expect(folderWorkbooks(["notes.txt", "Slide Sheet - 40001.xlsx"])).toEqual({ snapshot: null, xlsm: null });
  });

  it("finds a workbook that was added later", () => {
    expect(folderWorkbooks(["notes.txt", "Slide Sheet - 40001.xlsm"], "40001.phx.json")).toEqual({
      snapshot: null,
      xlsm: "Slide Sheet - 40001.xlsm",
    });
  });

  it("does not take Excel's lock file for the workbook", () => {
    expect(folderWorkbooks(["~$Slide Sheet - 40001.xlsm"]).xlsm).toBeNull();
    expect(folderWorkbooks(["~$Slide Sheet - 40001.xlsm", "Slide Sheet - 40001.XLSM"]).xlsm).toBe("Slide Sheet - 40001.XLSM");
  });

  it("keeps the snapshot the job list expects when the folder has it", () => {
    const names = ["A.phx.json", "Slide Sheet - 40001.phx.json"];
    expect(folderWorkbooks(names, "Slide Sheet - 40001.phx.json").snapshot).toBe("Slide Sheet - 40001.phx.json");
  });

  it("falls back to the snapshot saved under another name", () => {
    expect(folderWorkbooks(["Slide Sheet - 40001.phx.json"], "40001.phx.json").snapshot).toBe("Slide Sheet - 40001.phx.json");
  });

  it("ignores restore points", () => {
    expect(folderWorkbooks([".phx-backup.json", ".phx-backup-daily-2026-01-02.json", "backup of job.phx.json"]).snapshot).toBeNull();
  });

  it("matches the workbook for the active snapshot or job number in a multi-workbook folder", () => {
    const files = ["~$Slide Sheet - 35784.xlsm", "Slide Sheet - 35782.xlsm", "Slide Sheet - Demo.xlsm", "Slide Sheet - 35784.xlsm"];
    expect(matchingWorkbook(files, "Slide Sheet - 35784.phx.json")).toBe("Slide Sheet - 35784.xlsm");
    expect(matchingWorkbook(files, "other.phx.json", "35784")).toBe("Slide Sheet - 35784.xlsm");
    expect(matchingWorkbook(files, "Slide Sheet - Demo.phx.json")).toBe("Slide Sheet - Demo.xlsm");
  });
});

describe("staleEntriesFor", () => {
  const folder = (id: string) => ({ id, isSameEntry: async (other: unknown) => (other as { id?: string }).id === id });
  const names = ["Slide Sheet - 40001.phx.json", "Slide Sheet - 40001.xlsm"];
  const entries = [
    { jobId: "40001", fileName: "40001.phx.json", handle: folder("desktop/40001") },
    { jobId: "slide-sheet-40001", fileName: "Slide Sheet - 40001.phx.json", handle: folder("desktop/40001") },
    { jobId: "40002", fileName: "40002.phx.json", handle: folder("desktop/40002") },
  ];

  it("finds the entry made for the same folder while it was empty", async () => {
    expect(await staleEntriesFor(entries, "slide-sheet-40001", folder("desktop/40001"), names)).toEqual(["40001"]);
  });

  it("leaves an entry whose snapshot is in the folder, and every other folder", async () => {
    expect(await staleEntriesFor(entries, "40001", folder("desktop/40001"), names)).toEqual([]);
    expect(await staleEntriesFor(entries, "40002", folder("desktop/40002"), [])).toEqual([]);
  });

  it("does nothing without a folder grant to compare", async () => {
    expect(await staleEntriesFor(entries, "slide-sheet-40001", undefined, names)).toEqual([]);
    expect(await staleEntriesFor([{ jobId: "40001", fileName: "40001.phx.json" }], "slide-sheet-40001", folder("desktop/40001"), names)).toEqual([]);
  });

  it("treats a handle that cannot be compared as another folder", async () => {
    const broken = { isSameEntry: async () => { throw new TypeError("not a handle"); } };
    expect(await staleEntriesFor(entries, "slide-sheet-40001", broken, names)).toEqual([]);
  });
});

describe("job ids and labels", () => {
  it("names a snapshot after its workbook", () => {
    expect(snapshotLabel("Slide Sheet - 40001.phx.json")).toBe("Slide Sheet - 40001");
  });

  it("makes one id per label", () => {
    expect(appJobIdOf("Slide Sheet - 40001")).toBe("slide-sheet-40001");
    expect(appJobIdOf("40001")).toBe("40001");
    expect(appJobIdOf(" - ")).toBe("job");
  });
});

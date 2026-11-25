import { describe, expect, it } from "vitest";
import type { SnapshotBook } from "../../../../packages/workbook-commands/src/index";
import { archiveFileName } from "./archive";
import { dailyMailSubject, module11Recipients, buildReportDraft } from "./reportDraft";

const book = {
  styles: { s1: { bl: 1, bg: { rgb: "#FFF2CC" } } },
  sheets: {
    Data: {
      name: "Data",
      columnData: { "1": { w: 80 }, "2": { w: 40 } },
      mergeData: [{ startRow: 1, endRow: 1, startColumn: 1, endColumn: 2 }],
      cellData: {
        "1": { "1": { v: "Depth <m>", s: "s1" }, "3": { v: 12.5 } },
        "28": { "8": { v: "bit.pdf" }, "10": { v: "TO" }, "11": { v: "dd@phxtech.com" } },
        "29": { "8": { v: "missing.pdf" }, "10": { v: "CC" }, "11": { v: "office@phxtech.com" } },
      },
    },
  },
} as SnapshotBook;

describe("report draft", () => {
  it("follows Module11 recipients and subject", () => {
    expect(module11Recipients([
      { role: "CC", address: "a@phxtech.com" },
      { role: "TO", address: "b@phxtech.com" },
    ])).toEqual({ to: ["b@phxtech.com"], cc: ["a@phxtech.com"] });
    expect(module11Recipients([{ role: "", address: "a@phxtech.com; c@phxtech.com" }])).toEqual({ to: ["a@phxtech.com", "c@phxtech.com"], cc: [] });
    expect(dailyMailSubject("", "")).toBe("Daily Drilling Summary - Job");
    expect(dailyMailSubject("35783", "")).toBe("Daily Drilling Summary - 35783");
    expect(dailyMailSubject("20803", "Willowbend")).toBe("Daily Drilling Summary - 20803 - Willowbend");
  });

  it("renders the Data table, both png copies, and missing files", () => {
    const png = Uint8Array.of(1, 2, 3);
    const draft = buildReportDraft({
      snapshot: book,
      jobCode: "",
      jobName: "North",
      png,
      files: [{ name: "bit.pdf", bytes: Uint8Array.of(9) }, { name: "missing.pdf", bytes: null }],
    });
    expect(draft.subject).toBe("Daily Drilling Summary - Job - North");
    expect(draft.to).toEqual(["dd@phxtech.com"]);
    expect(draft.cc).toEqual(["office@phxtech.com"]);
    expect(draft.html).toContain("Depth &lt;m&gt;");
    expect(draft.html).toContain("font-weight:bold");
    expect(draft.html).toContain("#FFF2CC");
    expect(draft.html).toContain('colspan="2"');
    expect(draft.html).toContain("text-align:left");
    expect(draft.html).not.toContain("text-align:right");
    expect(draft.html).toContain("cid:daily_report.png");
    expect(draft.html).toContain("Not in the job folder: missing.pdf");
    expect(draft.inline.map((part) => part.contentId)).toEqual(["daily_report.png"]);
    expect(draft.attachments.map((part) => part.name)).toEqual(["daily_report.png", "bit.pdf"]);
    expect(draft.missing.map((file) => file.name)).toEqual(["missing.pdf"]);
  });

  it("names the Midnight Reports archive from the local date", () => {
    expect(archiveFileName(new Date(2026, 10, 4))).toBe("2026-11-04 Daily Report.eml");
  });
});

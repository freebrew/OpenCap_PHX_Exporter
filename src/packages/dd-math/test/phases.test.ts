import { describe, expect, it } from "vitest";
import { authorityOf, readyForNative } from "../../cutover/src/index";
import { sectionFor } from "../../corridor-3d/src/index";
import { normalizeEntity } from "../../fieldcap-sync/src/index";
import { buildEml, escapeHtml } from "../../reports/src/index";
import { applyDayRoll, checkDayRoll, daySlideMetres, startDayRoll } from "../../workbook-commands/src/index";

describe("later phases", () => {
  it("classifies the corridor from the inclination at the end of the day", () => {
    expect(sectionFor(2)).toBe("VERTICAL");
    expect(sectionFor(4.99)).toBe("VERTICAL");
    expect(sectionFor(40)).toBe("BUILD");
    expect(sectionFor(79.9)).toBe("BUILD");
    expect(sectionFor(80)).toBe("LATERAL");
  });

  it("escapes report html and keeps the message local", () => {
    expect(escapeHtml(`<a href="https://evil">`)).toBe("&lt;a href=&quot;https://evil&quot;&gt;");
    expect(buildEml({ to: "dd@example.com", subject: "Morning", html: "<p>Day</p>" })).toContain("Content-Type: text/html");
    expect(buildEml({ to: "dd@example.com", subject: "Morning", html: "<p>Day</p>" })).not.toContain("http://");
    const withFiles = buildEml({
      to: "dd@example.com",
      subject: "Morning",
      html: "<p>Day</p>",
      pngBase64: "UE5H",
      files: [{ name: "plan.pdf", base64: "UERG" }, { name: "missing.pdf", base64: "", missing: true }],
    });
    expect(withFiles).toContain("Content-Disposition: inline; filename=\"daily_report_inline.png\"");
    expect(withFiles).toContain("Content-ID: <daily_report.png>");
    expect(withFiles).toContain("Content-Disposition: attachment; filename=\"daily_report.png\"");
    expect(withFiles).toContain("Content-Disposition: attachment; filename=\"plan.pdf\"");
    expect(withFiles).toContain("Not in the job folder: missing.pdf");
    expect(withFiles.indexOf("multipart/related")).toBeLessThan(withFiles.indexOf("filename=\"plan.pdf\""));
  });

  it("clips a day roll to that row's depth window", () => {
    const book = {
      sheets: {
        data: { name: "Data", cellData: { 3: { 8: { v: 100 }, 9: { v: 110 } } } },
        slides: { name: "Slidesheet", cellData: { 12: { 17: { v: 95 }, 18: { v: 105 } }, 13: { 17: { v: 108 }, 18: { v: 120 } } } },
      },
    };
    expect(daySlideMetres(book, 4)).toBeCloseTo(7, 6);
  });

  it("fills a picked day row and checks the Pason form like MDL_DayRoll", () => {
    // Data row 4 is a finished Midnight; row 5 was just set to Midnight. BHA 1 mirror: 50 m, 2.5 h slid.
    const cells: Record<number, Record<number, { v: unknown }>> = {
      2: { 7: { v: 1 } },
      3: { 7: { v: "Midnight" }, 8: { v: 630 }, 9: { v: 700 }, 11: { v: 10 }, 12: { v: 1 }, 13: { v: 0 }, 15: { v: 5 }, 16: { v: 8 }, 21: { v: 1 } },
      4: { 7: { v: "Midnight" } },
      35: { 7: { v: 1 }, 8: { v: 50 }, 10: { v: 2.5 } },
    };
    const book = { sheets: { data: { name: "Data", cellData: cells } } };
    const started = startDayRoll(book, 5)!;
    expect(Object.fromEntries(started.writes.map((write) => [write.a1, write.value]))).toEqual({ I5: 700, L5: 40, M5: 1, N5: 30 });
    expect(started.context).toMatchObject({ row: 5, period: "Midnight", bha: 1, priorBit: 5, priorCirc: 8, slideHours: 1.5, startDepth: 700 });
    expect(startDayRoll(book, 3)).toBeNull();

    const context = started.context;
    expect(checkDayRoll(context, { endDepth: "", bitTotal: "", circTotal: "" })).toMatchObject({ ok: false, error: "End Depth is required." });
    expect(checkDayRoll(context, { endDepth: "690", bitTotal: "9", circTotal: "12" })).toMatchObject({ ok: false, error: "End Depth must be >= Start Depth (700)." });
    expect(checkDayRoll(context, { endDepth: "800", bitTotal: "4", circTotal: "12" })).toMatchObject({ ok: false, bitDelta: -1 });
    expect(checkDayRoll(context, { endDepth: "800", bitTotal: "12", circTotal: "10" })).toMatchObject({ ok: false, error: "Circ Hours must be >= Bit Hours for this row." });
    expect(checkDayRoll(context, { endDepth: "800", bitTotal: "12", circTotal: "40" })).toMatchObject({ ok: false, error: "Circ Hours cannot exceed 24 for a Midnight row." });
    expect(checkDayRoll(context, { endDepth: "800", bitTotal: "12", circTotal: "20" })).toEqual({ ok: true, endDepth: 800, bitDelta: 7, circDelta: 12 });

    for (const write of started.writes) {
      const column = write.a1.charCodeAt(0) - 65;
      cells[4]![column] = { v: write.value };
    }
    const applied = applyDayRoll(book, context, { endDepth: "800", bitTotal: "12", circTotal: "20" });
    expect("writes" in applied && Object.fromEntries(applied.writes.map((write) => [write.a1, write.value]))).toMatchObject({ J5: 800, P5: 7, Q5: 12 });
  });

  it("reads FieldCap pages without keeping metadata", () => {
    const rows = normalizeEntity("SurveySheetEntries", { value: [{ MD: 10, "@odata.etag": "secret" }] });
    expect(rows[0]?.values.MD).toBe("10");
    expect(rows[0]?.values["@odata.etag"]).toBeUndefined();
  });

  it("keeps Excel as the authority until every cutover box is checked", () => {
    expect(authorityOf({})).toBe("excel");
    expect(readyForNative(["calc", "picture"])).toBe(false);
    expect(readyForNative(["calc", "picture", "report", "sync", "fallback"])).toBe(true);
    expect(authorityOf({ authority: "native" })).toBe("native");
  });
});

import { describe, expect, it } from "vitest";
import { PRINT_CARDS } from "../src/cards.gen";
import {
  DEFAULT_THREADS,
  canonicalThread,
  extractCardTorqueRows,
  extractCustomTorqueRows,
  lookupTorque,
  renderCards,
  type EntryPayload,
} from "../src/index";

const latin = (pdf: Uint8Array) => new TextDecoder("latin1").decode(pdf);
const strings = (pdf: Uint8Array) =>
  [...latin(pdf).matchAll(/\(((?:[^()\\]|\\.)*)\) Tj/g)].map((m) => (m[1] ?? "").replace(/\\([()\\])/g, "$1"));

describe("tiered OD/ID make-up torque lookup", () => {
  it("resolves every API and DS thread in DEFAULT_THREADS, and leaves proprietary threads unmapped", () => {
    const proprietary = new Set(["CET43", "HLNST-43", "DUO400"]);
    for (const thread of DEFAULT_THREADS) {
      if (proprietary.has(thread)) {
        expect(canonicalThread(thread), thread).toBeNull();
        expect(lookupTorque(thread, 171, 71).spec, thread).toBeNull();
      } else {
        expect(canonicalThread(thread), thread).not.toBeNull();
      }
    }
  });

  it("snaps 169-172 mm collar wear into one consistent 6 3/4 inch OD tier", () => {
    const results = [169, 170, 171, 172].map((od) => lookupTorque("4 1/2 XH [NC46]", od, 71));
    for (const found of results) {
      expect(found.odTier?.id).toBe("6.75");
      expect(found.idTier?.id).toBe("2.81");
      expect(found.spec).toEqual({ ftLbs: 22500, kNm: 31 });
      expect(found.assumedId).toBe(false);
    }
  });

  it("auto-detects inches vs mm and produces the same tier and torque", () => {
    const metric = lookupTorque("4 1/2 IF [NC 50]", 171, 57);
    const imperial = lookupTorque("4 1/2 IF [NC 50]", 6.75, 2.25);
    expect(imperial.odTier?.id).toBe(metric.odTier?.id);
    expect(imperial.idTier?.id).toBe(metric.idTier?.id);
    expect(imperial.spec).toEqual(metric.spec);
    expect(metric.spec).toEqual({ ftLbs: 35000, kNm: 47 });
  });

  it("falls back to nominal bore ID when a tool family records no ID", () => {
    const motor = lookupTorque("4 1/2 REG", 171, null);
    expect(motor.assumedId).toBe(true);
    expect(motor.idTier?.id).toBe("2.25");
    expect(motor.spec).toEqual({ ftLbs: 22000, kNm: 30 });
  });

  it("steps torque cleanly when bore ID changes (2 13/16 inch vs 3 inch on NC46)", () => {
    const monel = lookupTorque("4 1/2 XH [NC46]", 171, 71);
    const mwd = lookupTorque("4 1/2 XH [NC46]", 171, 76);
    expect(monel.spec).toEqual({ ftLbs: 22500, kNm: 31 });
    expect(mwd.spec).toEqual({ ftLbs: 20500, kNm: 28 });
  });

  it("uses end connection ODs rather than a middle stabilizer blade OD", () => {
    const motorCard = PRINT_CARDS.find((card) => card.name === "Motors")!;
    const entry: EntryPayload = {
      which: "Motor",
      cells: {
        2: "MTR-1",
        6: "4 1/2 IF [NC 50]",
        7: "4 1/2 REG",
        16: "171", // I OD (lower)
        18: "216", // L OD (stabilizer pad — must not pull into 8"+ tier)
        22: "170", // P OD (upper)
      },
    };
    const rows = extractCardTorqueRows(motorCard, [entry]);
    expect(rows).toHaveLength(2);
    expect(rows[0]).toMatchObject({ thread: "4 1/2 IF [NC 50]", odLabel: '6 3/4" OD', ftLbs: 35000, kNm: 47, assumedId: true });
    expect(rows[1]).toMatchObject({ thread: "4 1/2 REG", odLabel: '6 3/4" OD', ftLbs: 22000, kNm: 30, assumedId: true });
  });

  it("deduplicates identical connection tiers across multiple cards on a page", () => {
    const flexCard = PRINT_CARDS.find((card) => card.name === "FlexNMs")!;
    const entries: EntryPayload[] = [169, 171, 172].map((od, i) => ({
      which: "NM Flex Monel",
      cells: {
        33: `FLX-${i + 1}`,
        37: "4 1/2 XH [NC46]",
        38: "4 1/2 XH [NC46]",
        48: String(od),
        49: String(od),
        50: "71",
      },
    }));
    const rows = extractCardTorqueRows(flexCard, entries);
    expect(rows).toHaveLength(1);
    expect(rows[0]).toMatchObject({ thread: "4 1/2 XH [NC46]", ftLbs: 22500, kNm: 31 });

    const drawn = strings(renderCards(entries));
    expect(drawn).toContain("MAKE-UP TORQUE REFERENCE (TIERED OD / ID)");
    expect(drawn.filter((s) => s === "22,500 ft-lbs")).toHaveLength(1);
    expect(drawn.filter((s) => s === "31 kN-m")).toHaveLength(1);
  });

  it("prints Verify mfr chart for proprietary threads and extracts Custom tool threads", () => {
    const custom: EntryPayload = {
      which: "Custom",
      cells: {
        serial: "CUST-1",
        boxUpper: "BOX",
        boxLower: "PIN",
        threadUpper: "4 1/2 IF [NC 50]",
        threadLower: "CET43",
        internalId: "2.25",
        marksJson: JSON.stringify({
          marks: [
            { letter: "A", kind: "upset", x: 0.28, value: "6.75" },
            { letter: "B", kind: "spiral", x: 0.56, value: "8.25" },
          ],
        }),
      },
    };
    const rows = extractCustomTorqueRows([custom]);
    expect(rows).toHaveLength(2);
    expect(rows[0]).toMatchObject({ thread: "4 1/2 IF [NC 50]", odLabel: '6 3/4" OD', idLabel: '2 1/4" ID', ftLbs: 35000, kNm: 47 });
    expect(rows[1]).toMatchObject({ thread: "CET43", ftLbs: null, kNm: null });

    const drawn = strings(renderCards([custom]));
    expect(drawn).toContain("35,000 ft-lbs");
    expect(drawn).toContain("Verify mfr chart");
  });

  it("omits the torque strip when no connections are entered on a page", () => {
    const bare: EntryPayload = { which: "Jars", cells: { 115: "JAR-1" } };
    const drawn = strings(renderCards([bare]));
    expect(drawn).not.toContain("MAKE-UP TORQUE REFERENCE (TIERED OD / ID)");
  });
});

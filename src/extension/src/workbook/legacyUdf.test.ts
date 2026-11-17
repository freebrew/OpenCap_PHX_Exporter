import { describe, expect, it } from "vitest";
import { LEGACY_UDF_NAMES, registerLegacyUdfStubs } from "./legacyUdf";

describe("legacy UDF stubs", () => {
  it("registers every known VBA name and extra import-log names", () => {
    const registered: string[] = [];
    registerLegacyUdfStubs({
      registerFunction(name) {
        registered.push(name);
      },
    }, ["CUSTOMEXTRA"]);
    expect(registered).toEqual(expect.arrayContaining([...LEGACY_UDF_NAMES, "CUSTOMEXTRA"]));
    expect(new Set(registered).size).toBe(registered.length);
  });
});

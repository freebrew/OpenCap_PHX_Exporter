import { describe, expect, it } from "vitest";
import { preserveActiveSheet } from "./preserveActiveSheet";

describe("preserveActiveSheet", () => {
  it("returns to the sheet that was active before asynchronous workbook writes", async () => {
    let active = "Slidesheet";
    const host = {
      activeSheet: () => active,
      activateSheet: (name: string) => {
        active = name;
      },
    };

    await preserveActiveSheet(host, async () => {
      active = "Setup";
    });

    expect(active).toBe("Slidesheet");
  });

  it("returns to the original sheet when the workbook operation fails", async () => {
    let active = "Data";
    const host = {
      activeSheet: () => active,
      activateSheet: (name: string) => {
        active = name;
      },
    };

    await expect(preserveActiveSheet(host, async () => {
      active = "Setup";
      throw new Error("pull failed");
    })).rejects.toThrow("pull failed");

    expect(active).toBe("Data");
  });
});

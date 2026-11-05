import { describe, expect, it } from "vitest";
import { settingsFromStored } from "../src/settings";

describe("saved settings", () => {
  it("keeps the job, recipients, and PDF choice from an older sign-in build", () => {
    const stored = {
      apiBase: "https://10.0.2.2:8787",
      jobNumber: "35783",
      recentJobs: ["35783", "34801"],
      emailMode: "service",
      mailVendor: "microsoft",
      recipients: ["office@phxtech.com"],
      combined: false,
    };
    expect(settingsFromStored(stored)).toEqual({
      jobNumber: "35783",
      recentJobs: ["35783", "34801"],
      recipients: ["office@phxtech.com"],
      combined: false,
    });
  });

  it("starts from a blank phone with one combined PDF", () => {
    expect(settingsFromStored(null)).toEqual({
      jobNumber: "",
      recentJobs: [],
      recipients: [],
      combined: true,
    });
  });
});

import { expect, test } from "@playwright/test";
import type { WellsProbeResult } from "../src/wells/wellsProbe";

function isRed(pixel: [number, number, number]): boolean {
  return pixel[0] > 180 && pixel[1] < 80 && pixel[2] < 80;
}

function painted(pixel: [number, number, number]): boolean {
  return pixel[0] + pixel[1] + pixel[2] > 80;
}

test("the worker canvas draws, and a hover reads the printed SF", async ({ page }) => {
  const errors: string[] = [];
  page.on("pageerror", (error) => errors.push(error.message));
  page.on("console", (message) => {
    if (message.type() === "error") errors.push(message.text());
  });
  await page.goto("/wells-probe.html");
  const result = await page.evaluate(async () => {
    const bridge = (window as Window & { __phxWells?: Promise<WellsProbeResult> }).__phxWells;
    if (!bridge) throw new Error("Probe did not start");
    return bridge;
  });
  process.stdout.write(`WELLS ${JSON.stringify(result)}\n`);
  expect(errors, errors.join("\n")).toEqual([]);
  expect(result.forced.backend).toBe("webgl2");
  expect(result.open.backend === "webgl2" || result.open.backend === "webgpu").toBe(true);
  expect(isRed(result.forced.center), `center ${result.forced.center.join(",")}`).toBe(true);
  expect(isRed(result.forced.shoulder), `shoulder ${result.forced.shoulder.join(",")} wide=${result.forced.wide}`).toBe(true);
  expect(painted(result.sample.corner), `corner ${result.sample.corner.join(",")}`).toBe(true);
  expect(result.hover).not.toBeNull();
  expect(result.hover?.sf).toBeCloseTo(1.2, 3);
  expect(result.hiddenSf).toBeNull();
  expect(result.bench.maxMs).toBeLessThan(50);
  expect(result.bench.mainBusyMs).toBeGreaterThan(200);
});

test("Sheets and 3D Wells switch on the preview page", async ({ page }) => {
  const errors: string[] = [];
  page.on("pageerror", (error) => errors.push(error.message));
  await page.goto("/app.html?adapter=memory");
  await page.getByRole("button", { name: "3D Wells" }).click();
  await expect(page.getByText("Planned Path")).toBeVisible();
  await expect(page.getByText("Shadow of the Active Well")).toBeVisible();
  await expect(page.getByText("Formation", { exact: true })).toBeVisible();
  await expect(page.getByRole("combobox", { name: "Site" })).toBeVisible();
  await expect(page.getByRole("combobox", { name: "Lowest SF" })).toBeVisible();
  await expect(page.getByRole("combobox", { name: "Closest C2C" })).toBeVisible();
  await expect(page.getByText("Run Import AC again to store the report header")).toBeVisible();
  await expect(page.getByText("Show All Offset Wells")).toBeVisible();
  await expect(page.getByRole("button", { name: "Reset View" })).toBeVisible();
  await expect(page.getByText("Import a well plan to draw the paths.")).toBeVisible();
  await expect(page.locator(".wells-canvas")).toBeAttached();
  await page.getByRole("button", { name: "Sheets", exact: true }).click();
  await expect(page.locator(".wells-view")).toHaveAttribute("data-active", "false");
  if (errors.length > 0) process.stdout.write(`UI ERRORS ${errors.join("\n")}\n`);
});

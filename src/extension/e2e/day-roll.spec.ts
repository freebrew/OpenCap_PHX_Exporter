import { expect, test } from "@playwright/test";

type Bridge = {
  directory: { getFile(name: string, create: boolean): Promise<{ writeExclusive(): Promise<{ write(bytes: Uint8Array): Promise<void>; close(): Promise<void> }> }> };
  setCell(sheetName: string, a1: string, value: string | number): void;
  getRawCell(sheetName: string, a1: string): unknown;
};

test("picking a period in Data!H opens the day roll form and writes depth and hours", async ({ page }) => {
  test.setTimeout(300_000);
  page.on("pageerror", (error) => process.stdout.write(`PAGEERROR ${error.message}\n`));
  await page.goto("/app.html?adapter=opfs");
  await page.waitForFunction(() => Boolean((window as Window & { __phx?: { directory?: unknown } }).__phx?.directory));
  await page.evaluate(async () => {
    const root = await navigator.storage.getDirectory();
    const names: string[] = [];
    for await (const [name] of root.entries()) names.push(name);
    for (const name of names) await root.removeEntry(name, { recursive: true });
    const data = new Uint8Array(await (await fetch("/__fixtures/demo.xlsm")).arrayBuffer());
    const bridge = (window as unknown as { __phx: Bridge }).__phx;
    const writer = await (await bridge.directory.getFile("Slide Sheet - Demo.xlsm", true)).writeExclusive();
    await writer.write(data);
    await writer.close();
  });
  await page.getByRole("combobox", { name: "Job" }).selectOption({ label: "Open another job folder…" });
  await expect(page.getByRole("button", { name: "Continue" })).toBeVisible({ timeout: 90_000 });
  await page.getByRole("button", { name: "Continue" }).click();
  await expect(page.getByRole("status").first()).toContainText("folder-safe", { timeout: 120_000 });

  const plan = await page.evaluate(() => {
    const bridge = (window as unknown as { __phx: Bridge }).__phx;
    const number = (a1: string) => {
      const value = bridge.getRawCell("Data", a1);
      return typeof value === "number" ? value : Number(value) || 0;
    };
    for (let row = 4; row <= 22; row += 1) {
      const period = bridge.getRawCell("Data", `H${row}`);
      if (period != null && String(period).trim() !== "") continue;
      let bit = 0;
      let circ = 0;
      for (let r = 4; r < row; r += 1) {
        bit += number(`P${r}`);
        circ += number(`Q${r}`);
      }
      const start = row > 4 ? number(`J${row - 1}`) : 0;
      bridge.setCell("Data", `H${row}`, "Midnight");
      return { row, bit, circ, end: Math.max(start, 1) + 100 };
    }
    return null;
  });
  if (!plan) throw new Error("Demo Data!H4:H22 has no empty row.");

  const dialog = page.getByRole("dialog", { name: "Day Roll - Depth and Hours" });
  await expect(dialog).toBeVisible({ timeout: 20_000 });
  await expect(dialog).toContainText(`Period: Midnight (row ${plan.row})`);
  await dialog.getByLabel("End Depth").fill(String(plan.end));
  await dialog.getByLabel("Pason Bit Total").fill(String(plan.bit + 5));
  await dialog.getByLabel("Pason Circ Total").fill(String(plan.circ + 3));
  await expect(dialog.getByRole("status")).toHaveText("Circ Hours must be >= Bit Hours for this row.");
  await dialog.getByLabel("Pason Circ Total").fill(String(plan.circ + 8));
  await expect(dialog.getByRole("status")).toHaveText("Ready to apply End Depth and hour deltas.");
  await dialog.getByRole("button", { name: "Apply" }).click();
  await expect(dialog).toBeHidden();

  const written = await page.evaluate((row) => {
    const bridge = (window as unknown as { __phx: Bridge }).__phx;
    return ["J", "P", "Q"].map((column) => bridge.getRawCell("Data", `${column}${row}`));
  }, plan.row);
  expect(written).toEqual([plan.end, 5, 8]);
});

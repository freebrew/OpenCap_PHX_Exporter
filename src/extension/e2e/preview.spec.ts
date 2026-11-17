import { expect, test } from "@playwright/test";

interface EditTarget {
  sheet: string;
  input: string;
  formula: string;
  next: number;
}

test("imports Demo, edits an evaluable cell, and reopens the snapshot", async ({ page }) => {
  page.on("pageerror", (error) => {
    process.stdout.write(`PAGEERROR ${error.message}\n`);
  });
  page.on("console", (message) => {
    if (message.type() === "error") process.stdout.write(`CONSOLE ${message.text()}\n`);
  });
  await page.goto("/app.html?adapter=opfs");
  await page.waitForFunction(() => {
    const bridge = (window as Window & { __phx?: { directory?: unknown } }).__phx;
    return Boolean(bridge?.directory);
  });
  await page.evaluate(async () => {
    const root = await navigator.storage.getDirectory();
    const names: string[] = [];
    for await (const [name] of root.entries()) names.push(name);
    for (const name of names) await root.removeEntry(name, { recursive: true });
  });
  await page.evaluate(async () => {
    const response = await fetch("/__fixtures/demo.xlsm");
    if (!response.ok) throw new Error(`Demo fixture request failed: ${response.status}`);
    const data = new Uint8Array(await response.arrayBuffer());
    const bridge = (window as Window & {
      __phx: {
        directory: {
          getFile(name: string, create: boolean): Promise<{
            writeExclusive(): Promise<{ write(bytes: Uint8Array): Promise<void>; close(): Promise<void> }>;
          }>;
        };
      };
    }).__phx;
    const file = await bridge.directory.getFile("Slide Sheet - Demo.xlsm", true);
    const writer = await file.writeExclusive();
    await writer.write(data);
    await writer.close();
  });

  await page.getByRole("combobox", { name: "Job" }).selectOption({ label: "Open another job folder…" });
  await expect.poll(async () => {
    const status = (await page.getByRole("status").innerText()).replace(/\s+/g, " ");
    const visible = await page.getByRole("button", { name: "Continue" }).isVisible();
    process.stdout.write(`POLL visible=${visible} status=${status}\n`);
    return visible;
  }, { timeout: 60_000, intervals: [2_000] }).toBe(true);
  await expect(page.locator(".log")).toContainText("Sheets 25");
  await page.getByRole("button", { name: "Continue" }).click();
  await expect(page.getByRole("status")).toContainText("folder-safe", { timeout: 120_000 });

  const edited = await page.evaluate(() => {
    const bridge = (window as Window & {
      __phx: {
        snapshot(): {
          sheetOrder?: string[];
          sheets?: Record<string, { name: string; cellData?: Record<string, Record<string, { f?: string }>> }>;
        } | null;
        setCell(sheetName: string, a1: string, value: number): void;
      };
    }).__phx;
    const snapshot = bridge.snapshot();
    if (snapshot?.sheetOrder?.length !== 25) return null;
    const found = findDirectReference(snapshot);
    if (!found) return null;
    bridge.setCell(found.sheet, found.input, found.next);
    return found;

    function findDirectReference(book: {
      sheets?: Record<string, { name: string; cellData?: Record<string, Record<string, { f?: string }>> }>;
    }): EditTarget | null {
      for (const sheet of Object.values(book.sheets ?? {})) {
        for (const [rowText, row] of Object.entries(sheet.cellData ?? {})) {
          for (const [colText, cell] of Object.entries(row)) {
            const match = (cell.f ?? "").match(/^=([A-Z]{1,3})([0-9]+)$/);
            if (!match) continue;
            const input = `${match[1]}${match[2]}`;
            const formula = `${columnName(Number(colText))}${Number(rowText) + 1}`;
            if (input === formula) continue;
            return { sheet: sheet.name, input, formula, next: 41 };
          }
        }
      }
      return null;
    }

    function columnName(index: number): string {
      let value = index + 1;
      let name = "";
      while (value > 0) {
        const remainder = (value - 1) % 26;
        name = String.fromCharCode(65 + remainder) + name;
        value = Math.floor((value - 1) / 26);
      }
      return name;
    }
  });
  if (!edited) throw new Error("Demo snapshot did not expose a direct cell reference to edit.");
  process.stdout.write(`EDIT ${JSON.stringify(edited)}\n`);
  await page.evaluate(async (target) => {
    const bridge = (window as Window & {
      __phx: { flushCalculation(): Promise<void> };
    }).__phx;
    await bridge.flushCalculation();
    return target;
  }, edited);
  await expect.poll(async () => page.evaluate((target) => {
    const bridge = (window as Window & { __phx: { getRawCell(sheetName: string, a1: string): unknown } }).__phx;
    return bridge.getRawCell(target.sheet, target.formula);
  }, edited), { timeout: 20_000 }).toBe(edited.next);

  await page.getByRole("button", { name: "Save now" }).click();
  await expect(page.getByRole("status")).toContainText("Folder snapshot saved", { timeout: 60_000 });
  await page.reload();
  await page.waitForFunction(() => {
    const bridge = (window as Window & { __phx?: { directory?: unknown } }).__phx;
    return Boolean(bridge?.directory);
  });
  await page.getByRole("combobox", { name: "Job" }).selectOption({ label: "Open another job folder…" });
  await expect.poll(async () => page.evaluate((target) => {
    const bridge = (window as Window & { __phx: { getRawCell(sheetName: string, a1: string): unknown } }).__phx;
    return bridge.getRawCell(target.sheet, target.input);
  }, edited), { timeout: 60_000 }).toBe(edited.next);
});

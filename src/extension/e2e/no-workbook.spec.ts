import { expect, test, type Page } from "@playwright/test";

const DEMO = "Slide Sheet - Demo.xlsm";

type Bridge = Window & {
  __phx?: {
    directory?: {
      getFile(name: string, create: boolean): Promise<{
        writeExclusive(): Promise<{ write(bytes: Uint8Array): Promise<void>; close(): Promise<void> }>;
      }>;
    };
    snapshot(): unknown;
  };
};

type IdleHold = Window & { __heldIdle?: () => number; __releaseIdle?: () => string[] };

async function openPreview(page: Page): Promise<string[]> {
  const errors: string[] = [];
  page.on("pageerror", (error) => errors.push(error.message));
  await page.goto("/app.html?adapter=opfs");
  await page.waitForFunction(() => Boolean((window as Bridge).__phx?.directory));
  await emptyFolder(page);
  return errors;
}

async function emptyFolder(page: Page): Promise<void> {
  await page.evaluate(async () => {
    const root = await navigator.storage.getDirectory();
    const names: string[] = [];
    for await (const [name] of root.entries()) names.push(name);
    for (const name of names) await root.removeEntry(name, { recursive: true });
  });
}

async function putDemoWorkbook(page: Page): Promise<void> {
  await page.evaluate(async (name) => {
    const response = await fetch("/__fixtures/demo.xlsm");
    if (!response.ok) throw new Error(`Demo fixture request failed: ${response.status}`);
    const file = await (window as Bridge).__phx!.directory!.getFile(name, true);
    const writer = await file.writeExclusive();
    await writer.write(new Uint8Array(await response.arrayBuffer()));
    await writer.close();
  }, DEMO);
}

async function continueImport(page: Page): Promise<void> {
  await page.getByRole("button", { name: "Continue" }).click({ timeout: 90_000 });
  await expect(page.locator(".no-workbook")).toHaveCount(0, { timeout: 120_000 });
  // The import is only finished once the safety copy and the folder snapshot are both written. As text,
  // the Demo safety copy was over IndexedDB's 127 MiB value limit and this message never came.
  const status = page.getByRole("status");
  await expect(status).toContainText("Import saved as a disposable preview", { timeout: 120_000 });
  await expect(status).not.toContainText("safety copy was not written");
}

test("opening Well Tools with no workbook asks for no calculation", async ({ page }) => {
  const errors = await openPreview(page);
  await expect(page.locator(".no-workbook")).toContainText("No job is open");
  // Well Tools redraws the corridor picture, which recalculates first. Univer has no formula commands until a workbook exists.
  await page.getByRole("button", { name: "Well Tools" }).click();
  await expect(page.getByRole("button", { name: "Hide Well Tools" })).toBeVisible();
  await page.waitForTimeout(1_000);
  // The dev preview has no chrome.* API, which Well Tools' mail settings report on their own.
  expect(errors.filter((message) => message !== "chrome is not defined")).toEqual([]);
});

test("says why the grid is empty and picks up a workbook added later", async ({ page }) => {
  const errors = await openPreview(page);
  const empty = page.locator(".no-workbook");
  await expect(empty).toContainText("No job is open");

  await page.getByLabel("Job").selectOption("pick");
  await expect(empty).toContainText("No Slide Sheet in the Job folder");

  await putDemoWorkbook(page);
  await empty.getByRole("button", { name: "Check Folder Again" }).click();
  await expect(page.getByRole("button", { name: "Continue" })).toBeVisible({ timeout: 90_000 });
  await expect(empty).toContainText(`${DEMO} is ready to open`);

  await page.getByRole("button", { name: "Cancel" }).click();
  await expect(empty).toContainText(`${DEMO} has not been opened here`);
  await empty.getByRole("button", { name: `Open ${DEMO}` }).click();
  await continueImport(page);
  expect(await page.evaluate(() => (window as Bridge).__phx!.snapshot() !== null)).toBe(true);

  // The folder is emptied and opened again: the old sheet must not stay on the grid.
  await emptyFolder(page);
  await page.getByLabel("Job").selectOption("pick");
  await expect(empty).toContainText("No Slide Sheet in the Job folder");
  expect(await page.evaluate(() => (window as Bridge).__phx!.snapshot())).toBeNull();
  expect(errors).toEqual([]);
});

test("closing a job while row heights are still queued throws nothing", async ({ page }) => {
  const errors = await openPreview(page);
  await page.getByLabel("Job").selectOption("pick");
  await putDemoWorkbook(page);
  await page.getByLabel("Job").selectOption("pick");
  await continueImport(page);

  // A whole column: more rows than Univer measures at once (10,000 cells' worth), so Delete queues the rest.
  await page.getByText("Slidesheet", { exact: true }).last().click();
  const column = await page.evaluate(() => {
    const snap = (window as Bridge).__phx!.snapshot() as { sheets: Record<string, { name: string; rowCount: number; columnCount: number }> };
    const { rowCount, columnCount } = Object.values(snap.sheets).find((sheet) => sheet.name === "Slidesheet")!;
    let letters = "";
    for (let n = columnCount; n > 0; n = Math.floor((n - 1) / 26)) letters = String.fromCharCode(65 + ((n - 1) % 26)) + letters;
    return `${letters}1:${letters}${rowCount}`;
  });
  const nameBox = page.locator('[data-u-comp="defined-name"] input');
  await nameBox.click();
  await nameBox.fill(column);
  await nameBox.press("Enter");
  await page.waitForTimeout(1_000);

  // The queue runs in idle time. Hold it so the job closes first.
  await page.evaluate(() => {
    const held = new Map<number, IdleRequestCallback>();
    let next = 0;
    window.requestIdleCallback = (callback) => {
      next += 1;
      held.set(next, callback);
      return next;
    };
    window.cancelIdleCallback = (id) => {
      held.delete(id);
    };
    (window as IdleHold).__heldIdle = () => held.size;
    (window as IdleHold).__releaseIdle = () => {
      const thrown: string[] = [];
      for (const [id, callback] of [...held]) {
        held.delete(id);
        try {
          callback({ didTimeout: false, timeRemaining: () => 50 });
        } catch (error) {
          thrown.push(error instanceof Error ? error.message : String(error));
        }
      }
      return thrown;
    };
  });
  await page.keyboard.press("Delete");
  await expect.poll(() => page.evaluate(() => (window as IdleHold).__heldIdle!())).toBeGreaterThan(0);

  await emptyFolder(page);
  await page.getByLabel("Job").selectOption("pick");
  await expect(page.locator(".no-workbook")).toContainText("No Slide Sheet in the Job folder");
  expect(await page.evaluate(() => (window as IdleHold).__releaseIdle!())).toEqual([]);
  expect(errors).toEqual([]);
});

test("a job listed while its folder was empty opens the workbook imported later under its own name", async ({ page }) => {
  const errors = await openPreview(page);
  const jobList = page.getByLabel("Job");

  // Listed as "Job" while the folder is empty, then granted again once the workbook is in it.
  await jobList.selectOption("pick");
  await expect(page.locator(".no-workbook")).toContainText("No Slide Sheet in the Job folder");
  await putDemoWorkbook(page);
  await jobList.selectOption("pick");
  await continueImport(page);
  await expect(jobList.locator("option[value^='folder:']")).toHaveText(["Job", "Slide Sheet - Demo"]);

  // The entry made for the empty folder finds the snapshot, opens it, and leaves the list.
  await jobList.selectOption("folder:job");
  await expect(jobList.locator("option[value^='folder:']")).toHaveText(["Slide Sheet - Demo"], { timeout: 120_000 });
  await expect(jobList).toHaveValue("folder:slide-sheet-demo");
  await expect(page.locator(".no-workbook")).toHaveCount(0);
  await expect(page.getByRole("button", { name: "Continue" })).toHaveCount(0);
  expect(await page.evaluate(() => (window as Bridge).__phx!.snapshot() !== null)).toBe(true);
  expect(errors).toEqual([]);
});

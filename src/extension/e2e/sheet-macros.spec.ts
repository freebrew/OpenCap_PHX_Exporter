import { expect, test, type Page } from "@playwright/test";

type Bridge = {
  directory: { getFile(name: string, create: boolean): Promise<{ writeExclusive(): Promise<{ write(bytes: Uint8Array): Promise<void>; close(): Promise<void> }> }> };
  getRawCell(sheetName: string, a1: string): unknown;
  setCell?(sheetName: string, a1: string, value: string | number | boolean | null): void;
};

const SHOTS = process.env.PHX_SHOTS;

async function importDemo(page: Page) {
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
}

/** A one-page Courier PDF with these lines, for the file picker stub. */
function textPdf(lines: string[]): number[] {
  const escape = (line: string) => line.replace(/[\\()]/g, (ch) => `\\${ch}`);
  const content = `BT /F1 8 Tf 20 760 Td 11 TL ${lines.map((line) => `(${escape(line)}) Tj T*`).join(" ")} ET`;
  const objects = [
    "<< /Type /Catalog /Pages 2 0 R >>",
    "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
    "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << /Font << /F1 4 0 R >> >> /Contents 5 0 R >>",
    "<< /Type /Font /Subtype /Type1 /BaseFont /Courier >>",
    `<< /Length ${content.length} >>\nstream\n${content}\nendstream`,
  ];
  let pdf = "%PDF-1.4\n";
  const offsets = objects.map((body, i) => {
    const at = pdf.length;
    pdf += `${i + 1} 0 obj\n${body}\nendobj\n`;
    return at;
  });
  const xref = pdf.length;
  pdf += `xref\n0 ${objects.length + 1}\n0000000000 65535 f \n${offsets.map((at) => `${String(at).padStart(10, "0")} 00000 n \n`).join("")}`;
  pdf += `trailer\n<< /Size ${objects.length + 1} /Root 1 0 R >>\nstartxref\n${xref}\n%%EOF\n`;
  return [...Buffer.from(pdf, "latin1")];
}

async function openSheet(page: Page, name: string) {
  await page.getByText(name, { exact: true }).last().click();
  await page.waitForTimeout(800);
}

test("Slidesheet carries its own buttons and gauge; DL Tracker ADD ENTRY opens only the DL form", async ({ page }) => {
  test.setTimeout(300_000);
  page.on("pageerror", (error) => process.stdout.write(`PAGEERROR ${error.message}\n`));
  await page.setViewportSize({ width: 1600, height: 900 });
  await importDemo(page);

  await openSheet(page, "Slidesheet");
  const controls = page.getByLabel("Slidesheet sheet controls");
  for (const caption of ["Clear Ranges", "Pipe Tally Calculator", "Targets"]) {
    await expect(controls.getByRole("button", { name: caption })).toBeVisible();
  }
  const gauge = controls.getByRole("img", { name: "Plan proximity gauge" });
  await expect(gauge).toBeVisible();
  const gaugeText = await gauge.textContent();
  process.stdout.write(`GAUGE ${gaugeText}\n`);
  expect(gaugeText).toMatch(/GRAVITY|MAGNETIC|NO DATA/);
  if (SHOTS) {
    await page.setViewportSize({ width: 2700, height: 900 });
    await page.waitForTimeout(800);
    const box = await gauge.boundingBox();
    if (box) await page.screenshot({ path: `${SHOTS}/phx-slidesheet-top.png`, clip: { x: Math.max(0, box.x - 420), y: Math.max(0, box.y - 40), width: 900, height: box.height + 80 } });
    await page.setViewportSize({ width: 1600, height: 900 });
    await page.waitForTimeout(500);
  }

  // The buttons sit beside the gauge, in column Z, left of AA1.
  const clear = await controls.getByRole("button", { name: "Clear Ranges" }).boundingBox();
  const targets = await controls.getByRole("button", { name: "Targets" }).boundingBox();
  const dial = await gauge.boundingBox();
  expect(clear && dial && clear.x + clear.width).toBeLessThanOrEqual((dial?.x ?? 0) + 2);
  expect(targets && clear && targets.y).toBeGreaterThan(clear?.y ?? 0);

  await controls.getByRole("button", { name: "Targets" }).click();
  const targetsDialog = page.getByRole("dialog", { name: "Targets" });
  await expect(targetsDialog).toBeVisible();
  if (SHOTS) await page.screenshot({ path: `${SHOTS}/phx-targets.png` });
  await targetsDialog.getByRole("button", { name: "Cancel" }).click();
  await expect(targetsDialog).toBeHidden();

  await openSheet(page, "DL Tracker");
  const tracker = page.getByLabel("DL Tracker sheet controls");
  await expect(tracker.getByRole("button", { name: "ADD ENTRY" })).toBeVisible();
  await expect(tracker.getByRole("button", { name: "CLEAR" })).toBeVisible();
  await expect(page.getByLabel("Slidesheet sheet controls")).toHaveCount(0);

  await tracker.getByRole("button", { name: "ADD ENTRY" }).click();
  const entry = page.getByRole("dialog", { name: /^DL Tracker - new entry \(row \d+\)$/ });
  await expect(entry).toBeVisible();
  await expect(page.getByRole("region", { name: "Well tools" })).toHaveCount(0);
  const title = await entry.getByRole("heading").textContent();
  const row = Number(/row (\d+)/.exec(title ?? "")?.[1]);
  if (SHOTS) await page.screenshot({ path: `${SHOTS}/phx-dl-entry.png` });

  await entry.getByLabel("Dnlk Cmd.").fill("");
  await entry.getByRole("button", { name: "Save Entry" }).click();
  await expect(entry.getByRole("status")).toHaveText("Pick a Dnlk command.");
  await entry.getByLabel("Dnlk Cmd.").fill("1-5");
  await entry.getByLabel("Comments").fill("e2e entry");
  page.once("dialog", (dialog) => void dialog.accept());
  await entry.getByRole("button", { name: "Save Entry" }).click();
  await expect(entry).toBeHidden();
  const saved = await page.evaluate((r) => {
    const bridge = (window as unknown as { __phx: Bridge }).__phx;
    return [bridge.getRawCell("DL Tracker", `D${r}`), bridge.getRawCell("DL Tracker", `U${r}`)];
  }, row);
  expect(saved).toEqual(["1-5", "e2e entry"]);
});

test("Setup Import AC reads a PDF in the browser and fills the AC table", async ({ page }) => {
  test.setTimeout(300_000);
  page.on("pageerror", (error) => process.stdout.write(`PAGEERROR ${error.message}\n`));
  await page.setViewportSize({ width: 1600, height: 900 });
  await importDemo(page);
  const pdf = textPdf([
    "Anticollision Report",
    "Summary",
    "  100/01-02-003-04W5/00 - Offset A   1,110.00   1,105.00   12.40   6.10   1.950 SF",
    "  102/05-06-007-08W5/02 - Offset B     850.00     848.00    9.00   7.20   1.200",
    "Semi Major Axis",
  ]);
  await page.evaluate((bytes) => {
    const file = new File([new Uint8Array(bytes)], "ac-report.pdf", { type: "application/pdf" });
    (window as unknown as { showOpenFilePicker: () => Promise<unknown[]> }).showOpenFilePicker = async () => [{ kind: "file", name: file.name, getFile: async () => file }];
  }, pdf);

  await openSheet(page, "Setup");
  const importAc = page.getByRole("button", { name: /Import AC/i }).first();
  await expect(importAc).toBeVisible();
  await importAc.click();
  await expect(page.getByRole("status").first()).toContainText("AC import complete: 2 offset well(s), 2 with SF < 2.0.", { timeout: 60_000 });
  const cells = await page.evaluate(() => {
    const bridge = (window as unknown as { __phx: Bridge }).__phx;
    return ["J12", "K12", "O12", "Q12", "O13", "Q13"].map((a1) => bridge.getRawCell("Setup", a1));
  });
  process.stdout.write(`AC CELLS ${JSON.stringify(cells)}\n`);
  expect(JSON.stringify(cells)).toContain("HIGH");
  expect(JSON.stringify(cells)).toContain("CAUTION");
  if (SHOTS) await page.screenshot({ path: `${SHOTS}/phx-import-ac.png` });
});

test("Setup Import Plan reads a plan PDF in the browser and replaces the plan", async ({ page }) => {
  test.setTimeout(300_000);
  page.on("pageerror", (error) => process.stdout.write(`PAGEERROR ${error.message}\n`));
  await page.setViewportSize({ width: 1600, height: 900 });
  await importDemo(page);
  const pdf = textPdf([
    "Planning Report",
    "Plan Sections",
    "  Measured          Vertical                  Dogleg  Build   Turn",
    "   Depth  Inc   Azi   Depth  +N/-S  +E/-W   Rate   Rate   Rate   TFO   Target",
    "     0.00     0.00     0.00      0.00    0.00    0.00  0.000   0.000   0.000    0.00",
    "   600.00     0.00     0.00    600.00    0.00    0.00  0.000   0.000   0.000    0.00",
    "  1200.00    40.00    90.00   1150.00    0.00  200.00  2.000   2.000   0.000   90.00",
    "  1600.00    40.00    90.00   1460.00    0.00  460.00  0.000   0.000   0.000    0.00",
    "  2100.00    90.00   120.00   1600.00 -120.00  800.00  3.000   2.500   1.500   45.00",
    "  3500.00    90.00   120.00   1600.00 -820.00 2010.00  0.000   0.000   0.000    0.00",
    "Formations",
    "    850.00      850.00     Belly River",
    "   1990.00     1590.00     Montney",
  ]);
  await page.evaluate((bytes) => {
    const file = new File([new Uint8Array(bytes)], "e2e-plan.pdf", { type: "application/pdf" });
    (window as unknown as { showOpenFilePicker: () => Promise<unknown[]> }).showOpenFilePicker = async () => [{ kind: "file", name: file.name, getFile: async () => file }];
  }, pdf);

  await openSheet(page, "Setup");
  await page.getByRole("button", { name: /Import Plan/i }).first().click();
  await expect(page.getByRole("status").first()).toContainText("Plan imported from e2e-plan.pdf", { timeout: 60_000 });
  const cells = await page.evaluate(() => {
    const bridge = (window as unknown as { __phx: Bridge }).__phx;
    return {
      sections: ["A3", "A8", "K5", "K7", "K8"].map((a1) => bridge.getRawCell("_OC_PlanSec", a1)),
      survey: bridge.getRawCell("_OC_Survey", "A8"),
      tops: bridge.getRawCell("_OC_Formations", "C4"),
      tdPlanned: bridge.getRawCell("Data", "I54"),
      path: bridge.getRawCell("Setup", "M29"),
    };
  });
  process.stdout.write(`PLAN CELLS ${JSON.stringify(cells)}\n`);
  expect(JSON.stringify(cells)).toContain("HEEL");
  expect(JSON.stringify(cells)).toContain("Montney");
  expect(JSON.stringify(cells.tdPlanned)).toContain("3500");
  expect(JSON.stringify(cells.path)).toContain("e2e-plan.pdf");
});

test("Setup REFRESH CSVs loads the newest OpenCap exports from the job folder", async ({ page }) => {
  test.setTimeout(300_000);
  page.on("pageerror", (error) => process.stdout.write(`PAGEERROR ${error.message}\n`));
  await page.setViewportSize({ width: 1600, height: 900 });
  await importDemo(page);
  await page.evaluate(async () => {
    const root = await navigator.storage.getDirectory();
    const put = async (folder: FileSystemDirectoryHandle, name: string, text: string) => {
      const writable = await (await folder.getFileHandle(name, { create: true })).createWritable();
      await writable.write(text);
      await writable.close();
    };
    const openCap = await root.getDirectoryHandle("OpenCap", { create: true });
    await put(root, "fieldcap-job-4242-crew.csv", "Job ID,Name,Employee Code,Email,Phone,Role,Work Type\r\n4242,Old Crew,9,,,,\r\n");
    await new Promise((resolve) => setTimeout(resolve, 50));
    await put(openCap, "fieldcap-job-4242-job-details.csv", "Job ID,Job Code,Client,ClientEmail,opsStatus\r\n4242,J-4242,,ops@acme-energy.ca,Drilling\r\n");
    await put(openCap, "fieldcap-job-4242-crew.csv", "Job ID,Name,Employee Code,Email,Phone,Role,Work Type\r\n4242,Ann Lee,1,ann@x.com,403-555-1234,Day DD,\r\n");
  });

  await openSheet(page, "Setup");
  await page.getByRole("button", { name: /REFRESH/i }).first().click();
  const status = page.getByRole("status").first();
  await expect(status).toContainText("OpenCap: 2 of 6 files loaded for job J-4242", { timeout: 60_000 });
  // The ClientJobId (4242) stays inside _OC_Job and OC_LastJobId; the OpenCap sentence shows the Job Number.
  const sentence = ((await status.textContent()) ?? "").replace(/^dirty/, "").split(".")[0] ?? "";
  expect(sentence).toContain("job J-4242");
  expect(sentence.replace("J-4242", "")).not.toContain("4242");
  const cells = await page.evaluate(() => {
    const bridge = (window as unknown as { __phx: Bridge }).__phx;
    return { jobId: bridge.getRawCell("_OC_Job", "A2"), code: bridge.getRawCell("_OC_Job", "B2"), crew: bridge.getRawCell("_OC_Crew", "B2") };
  });
  process.stdout.write(`OPENCAP CELLS ${JSON.stringify(cells)}\n`);
  expect(JSON.stringify(cells.jobId)).toContain("4242");
  // The Full Report reads the same Data cells Module21 does, and the sentence names the Job Number.
  await page.context().grantPermissions(["clipboard-read", "clipboard-write"]);
  await page.evaluate(() => {
    const bridge = (window as unknown as { __phx: Bridge }).__phx;
    bridge.setCell?.("Data", "I23", 4336);
    bridge.setCell?.("Data", "J23", 5432);
    bridge.setCell?.("Data", "K23", 1096);
    bridge.setCell?.("Data", "L23", 350);
  });
  await openSheet(page, "Data");
  await page.getByLabel("Data sheet controls").getByRole("button", { name: "Full Report" }).click();
  await expect(page.getByRole("status").first()).toContainText("Morning report copied", { timeout: 30_000 });
  const copied = await page.evaluate(() => navigator.clipboard.readText());
  expect(copied).toContain("Daily Drilling Summary");
  expect(copied).toContain("J-4242");
  expect(copied.replaceAll("J-4242", "")).not.toContain("4242");
  expect(JSON.stringify(cells.code)).toContain("J-4242");
  expect(JSON.stringify(cells.crew)).toContain("Ann Lee");
  if (SHOTS) await page.screenshot({ path: `${SHOTS}/phx-refresh-csvs.png` });
});

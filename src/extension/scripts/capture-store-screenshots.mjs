// Capture 1280x800 JPEG screenshots of OpenCap Slidesheet using the sanitized Demo workbook.
// Run: node scripts/capture-store-screenshots.mjs
import { spawn } from "node:child_process";
import { mkdirSync } from "node:fs";
import { join } from "node:path";
import { fileURLToPath } from "node:url";
import { chromium } from "@playwright/test";

const extRoot = fileURLToPath(new URL("..", import.meta.url));
const outDir = join(extRoot, "store-assets");
mkdirSync(outDir, { recursive: true });

const PORT = 4189;
const BASE = `http://127.0.0.1:${PORT}`;

function startVite() {
  const child = spawn(
    process.execPath,
    [join(extRoot, "..", "..", "node_modules", "vite", "bin", "vite.js"), "--host", "127.0.0.1", "--port", String(PORT), "--strictPort"],
    { cwd: extRoot, env: { ...process.env, PHX_PREVIEW: "1" }, stdio: ["ignore", "pipe", "pipe"] },
  );
  child.stderr.on("data", (d) => process.stderr.write(d));
  return child;
}

async function waitForServer(url, timeoutMs = 60_000) {
  const start = Date.now();
  while (Date.now() - start < timeoutMs) {
    try {
      const r = await fetch(url);
      if (r.ok) return;
    } catch {
      // retry
    }
    await new Promise((r) => setTimeout(r, 400));
  }
  throw new Error(`Timed out waiting for ${url}`);
}

const vite = startVite();
let browser;
try {
  await waitForServer(`${BASE}/app.html`);
  browser = await chromium.launch({ headless: true });
  const context = await browser.newContext({ viewport: { width: 1280, height: 800 }, deviceScaleFactor: 1 });
  const page = await context.newPage();

  await page.goto(`${BASE}/app.html?adapter=opfs`, { waitUntil: "domcontentloaded" });
  await page.waitForFunction(() => Boolean(window.__phx?.directory), { timeout: 60_000 });

  await page.evaluate(async () => {
    const root = await navigator.storage.getDirectory();
    const names = [];
    for await (const [name] of root.entries()) names.push(name);
    for (const name of names) await root.removeEntry(name, { recursive: true });

    const response = await fetch("/__fixtures/demo.xlsm");
    if (!response.ok) throw new Error(`Demo fixture request failed: ${response.status}`);
    const data = new Uint8Array(await response.arrayBuffer());
    const file = await window.__phx.directory.getFile("Slide Sheet - Demo.xlsm", true);
    const writer = await file.writeExclusive();
    await writer.write(data);
    await writer.close();
  });

  await page.getByRole("combobox", { name: "Job" }).selectOption({ label: "Open another job folder…" });
  await page.getByRole("button", { name: "Continue" }).waitFor({ state: "visible", timeout: 90_000 });
  await page.getByRole("button", { name: "Continue" }).click();
  await page.waitForFunction(
    () => (document.querySelector('[role="status"]')?.textContent ?? "").includes("folder-safe"),
    { timeout: 120_000 },
  );
  // Let Univer finish layout and gauges render.
  await page.waitForTimeout(2500);

  // Screenshot 1: Setup & Well Database.
  const shot1 = join(outDir, "screenshot-1-setup.jpg");
  await page.screenshot({ path: shot1, type: "jpeg", quality: 94 });
  console.log("CAPTURED", shot1);

  // Screenshot 2: Data sheet (proximity gauges & daily summary).
  const dataTab = page.getByText("Data", { exact: true }).last();
  if (await dataTab.isVisible()) {
    await dataTab.click();
    await page.waitForTimeout(2000);
    const shot2 = join(outDir, "screenshot-2-data.jpg");
    await page.screenshot({ path: shot2, type: "jpeg", quality: 94 });
    console.log("CAPTURED", shot2);
  }

  // Screenshot 3: Slidesheet tab.
  const slideTab = page.getByText("Slidesheet", { exact: true }).last();
  if (await slideTab.isVisible()) {
    await slideTab.click();
    await page.waitForTimeout(2000);
    const shot3 = join(outDir, "screenshot-3-slidesheet.jpg");
    await page.screenshot({ path: shot3, type: "jpeg", quality: 94 });
    console.log("CAPTURED", shot3);
  }

  // Screenshot 4: 3D Wells view — wait until the WebGL worker finishes painting the scene.
  const wellsBtn = page.getByRole("button", { name: "3D Wells" });
  if (await wellsBtn.isVisible()) {
    await wellsBtn.click();
    await page.waitForTimeout(7000);
    const shot4 = join(outDir, "screenshot-4-wells-3d.jpg");
    await page.screenshot({ path: shot4, type: "jpeg", quality: 94 });
    console.log("CAPTURED", shot4);
  }
} finally {
  if (browser) await browser.close();
  vite.kill();
}

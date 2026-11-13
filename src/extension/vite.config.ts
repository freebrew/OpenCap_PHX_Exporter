import { crx } from "@crxjs/vite-plugin";
import react from "@vitejs/plugin-react";
import { readFileSync } from "node:fs";
import { fileURLToPath, URL } from "node:url";
import { defineConfig, type Plugin } from "vite";
import manifest from "./manifest.config.ts";

const previewServer = process.env.PHX_PREVIEW === "1";
const demoWorkbook = fileURLToPath(new URL("../../Slide Sheet - Demo.xlsm", import.meta.url));
const exporterDir = fileURLToPath(new URL("../chrome-extension/", import.meta.url));
const EXPORTER_MODULE = "virtual:opencap-exporter";
/** content.js injects injected-spy.js, and the exporter re-injects these two by name, so they keep their names at the root. */
const EXPORTER_PAGE_SCRIPTS = ["browser-api.js", "content.js", "injected-spy.js"];

function readExporter(name: string): string {
  return readFileSync(`${exporterDir}${name}`, "utf8");
}

/** The OpenCap exporter (src/chrome-extension), unchanged: its worker as a module function, its page scripts as files. */
function openCapExporter(): Plugin {
  const resolved = `\0${EXPORTER_MODULE}`;
  return {
    name: "phx-opencap-exporter",
    resolveId(source) {
      return source === EXPORTER_MODULE ? resolved : null;
    },
    load(id) {
      if (id !== resolved) return null;
      for (const name of ["browser-api.js", "background.js"]) this.addWatchFile(`${exporterDir}${name}`);
      return `export function startOpenCapExporter(chrome, importScripts) {\n${readExporter("browser-api.js")}\n${readExporter("background.js")}\n}\n`;
    },
    generateBundle() {
      for (const name of EXPORTER_PAGE_SCRIPTS) this.emitFile({ type: "asset", fileName: name, source: readExporter(name) });
    },
  };
}

/** The Slidesheet's survey green. The green bar stops at this exact colour (GOOD_SURVEY_FILL). */
const SURVEY_GREEN = "#CCFFCC";
/** Univer's palette has a pale green of its own, a near duplicate of the swatch beside it; the survey green takes its place. */
const PALETTE_SWATCH_REPLACED = '"#DEF7EC"';

/**
 * The Fill Color drop-down (Univer's design package) is a fixed palette without the survey green, so
 * a row painted another colour could not be painted back. The swatch is swapped when the package is
 * bundled. A Univer update that moves the palette fails the build here instead of dropping the swatch.
 */
function surveyGreenSwatch(): Plugin {
  return {
    name: "phx-survey-green-swatch",
    enforce: "pre",
    transform(code, id) {
      if (!/[\\/]@univerjs[\\/]design[\\/]lib[\\/]es[\\/]index\.js/.test(id)) return null;
      const start = code.indexOf("const colorPresets = [");
      const end = start < 0 ? -1 : code.indexOf("];", start);
      if (start < 0 || end < 0 || !code.slice(start, end).includes(PALETTE_SWATCH_REPLACED)) {
        throw new Error("phx-survey-green-swatch: the colour palette in @univerjs/design moved; update the plugin.");
      }
      return { code: code.slice(0, start) + code.slice(start, end).replace(PALETTE_SWATCH_REPLACED, `"${SURVEY_GREEN}"`) + code.slice(end), map: null };
    },
  };
}

function demoFixture(): Plugin {
  return {
    name: "phx-demo-fixture",
    configureServer(server) {
      server.middlewares.use("/__fixtures/demo.xlsm", (_request, response) => {
        response.setHeader("content-type", "application/octet-stream");
        response.end(readFileSync(demoWorkbook));
      });
    },
  };
}

export default defineConfig({
  plugins: previewServer ? [react(), surveyGreenSwatch(), demoFixture()] : [react(), surveyGreenSwatch(), openCapExporter(), crx({ manifest })],
  resolve: {
    alias: {
      "@phx/job-store": fileURLToPath(new URL("../packages/job-store/src/index.ts", import.meta.url)),
      "@phx/mail": fileURLToPath(new URL("../packages/mail/src/index.ts", import.meta.url)),
      "@phx/xlsx-bridge": fileURLToPath(new URL("../packages/xlsx-bridge/src/index.ts", import.meta.url)),
    },
  },
  // The dependency scan does not follow `new URL("./formula.worker.ts", import.meta.url)`. Found late, the
  // worker's imports are re-bundled apart from the page's, the worker loads a second redi, and every plugin
  // it starts is missing its services ("reading 'setConfig'").
  optimizeDeps: {
    include: ["@univerjs/core", "@univerjs/presets", "@univerjs/preset-sheets-core/worker", "three", "three/webgpu"],
  },
  server: previewServer ? { hmr: false, watch: null } : undefined,
  build: {
    rollupOptions: {
      input: {
        app: fileURLToPath(new URL("./app.html", import.meta.url)),
      },
    },
  },
});

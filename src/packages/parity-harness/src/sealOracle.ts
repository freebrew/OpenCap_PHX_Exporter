import { readFileSync } from "node:fs";
import { manifestFromCalculatedWorkbook, writeManifest } from "./oracle.js";

interface SealMeta {
  sourceSha256: string;
  excelVersion: string;
  calculation: number;
  calculationState: number;
  calculatedPath: string;
  sealedAt: string;
}

function arg(name: string): string | undefined {
  const index = process.argv.indexOf(name);
  return index >= 0 ? process.argv[index + 1] : undefined;
}

async function main(): Promise<void> {
  const metaPath = arg("--meta");
  const out = arg("--out");
  if (!metaPath || !out) throw new Error("Usage: sealOracle.ts --meta <meta.json> --out <manifest.json>");
  const meta = JSON.parse(readFileSync(metaPath, "utf8").replace(/^\uFEFF/, "")) as SealMeta;
  const manifest = await manifestFromCalculatedWorkbook(meta);
  writeManifest(out, manifest);
  process.stdout.write(`Sealed ${manifest.formulaCells} formula cells from Excel ${manifest.excelVersion} state ${manifest.calculationState}\n`);
}

main().catch((error: unknown) => {
  process.stderr.write(`${error instanceof Error ? error.message : String(error)}\n`);
  process.exitCode = 1;
});

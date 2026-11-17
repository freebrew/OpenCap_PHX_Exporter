import { getDocument, GlobalWorkerOptions, OPS, VerbosityLevel } from "pdfjs-dist";
import workerUrl from "pdfjs-dist/build/pdf.worker.min.mjs?url";
import { pdfDocumentText } from "../../../packages/workbook-commands/src/index";

// The extension's PDF reader (Mozilla PDF.js, Apache-2.0). It replaces the VBA's
// pdftotext / PowerShell / Acrobat chain: nothing to install, and the file never leaves the browser.

GlobalWorkerOptions.workerSrc = workerUrl;

/** The PDF's text in `pdftotext -layout` line order, for the Import Plan / Import AC parsers. */
export async function pdfLayoutText(bytes: Uint8Array): Promise<string> {
  // fontExtraProperties keeps each font's data and code map, for fonts that have no letter map.
  const task = getDocument({ data: bytes, useSystemFonts: false, disableFontFace: true, fontExtraProperties: true, verbosity: VerbosityLevel.ERRORS });
  const doc = await task.promise;
  try {
    return await pdfDocumentText(doc, OPS);
  } finally {
    await task.destroy();
  }
}

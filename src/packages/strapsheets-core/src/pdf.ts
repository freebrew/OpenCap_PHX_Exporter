// The strap card as a PDF. It draws what the Sheets printout draws: the tool's drawing, a card for
// each entry with the connections and lettered measures in the sheet's own grid, and the family
// title underneath. The Worker attaches it to service mail; the phone renders the same bytes
// offline for the share sheet. Several entries of one tool share a page; each family starts a page.

import { layoutCards } from "./cardLayout";
import type { EntryPayload } from "./entry";
import { ENTRY_FAMILIES } from "./families.gen";
import { PdfWriter } from "./pdfWriter";
import { entryLines, findFamily } from "./wizard";

/** One PDF holding a card for every entry. */
export function renderCards(entries: EntryPayload[]): Uint8Array {
  const writer = new PdfWriter();
  layoutCards(entries, writer);
  return writer.build();
}

/** The file name a card goes out under: family and serial. */
export function cardFileName(entry: EntryPayload): string {
  const family = findFamily(ENTRY_FAMILIES, entry.which);
  const serial = family ? entryLines(family, entry.cells).find((line) => line.label === "Serial")?.value ?? "" : "";
  return `${[entry.which, serial].filter(Boolean).join(" ").replace(/[^A-Za-z0-9 ._-]+/g, "_")}.pdf`;
}

// Render the strap sheets on the phone and open the mail app with the PDFs attached.

import { Directory, Filesystem } from "@capacitor/filesystem";
import { cardFileName, renderCards, type EntryPayload } from "../../../packages/strapsheets-core/src/index";
import { composeMail } from "./mailComposer";
import type { Settings } from "./settings";

export async function sendEntries(entries: EntryPayload[], settings: Settings): Promise<void> {
  const groups = settings.combined || entries.length <= 1 ? [entries] : entries.map((entry) => [entry]);
  const files: string[] = [];
  for (const group of groups) {
    const first = group[0];
    if (!first) continue;
    const name = group.length > 1 ? `${group.length} strap sheets.pdf` : cardFileName(first);
    const written = await Filesystem.writeFile({
      path: `share/${name}`,
      data: base64(renderCards(group)),
      directory: Directory.Cache,
      recursive: true,
    });
    files.push(written.uri);
  }
  if (files.length === 0) throw new Error("Choose a strap sheet to email.");
  const subject = entries.length > 1 ? `${entries.length} strap sheets` : `Strap sheet: ${entries[0]?.which ?? "entry"}`;
  await composeMail({
    to: settings.recipients,
    subject,
    text: `The strap sheet${entries.length > 1 ? "s are" : " is"} attached.`,
    files,
  });
}

function base64(bytes: Uint8Array): string {
  let binary = "";
  const chunk = 0x8000;
  for (let index = 0; index < bytes.length; index += chunk) {
    binary += String.fromCharCode(...bytes.subarray(index, index + chunk));
  }
  return btoa(binary);
}

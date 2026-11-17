import { unzipSync, strFromU8 } from "fflate";

export interface MacroButton {
  sheet: string;
  id: string;
  caption: string;
  macro: string;
  anchor: string;
}

const CAPTIONS: Record<string, string> = {
  ImportSurveyPlan: "Import Plan",
  ImportAntiCollision: "Import AC",
  cleartally: "Clear Tally",
  SelectBHA: "Select BHA",
  RefreshSetup: "Refresh",
  RebuildSetup: "Rebuild",
  RefreshData: "Refresh",
  RebuildDashboard: "Rebuild",
};

export function macroName(formula: string): string {
  const text = formula.trim();
  const bang = text.lastIndexOf("!");
  return (bang >= 0 ? text.slice(bang + 1) : text).replace(/^'+|'+$/g, "");
}

export function macroButtonsFromWorkbook(bytes: Uint8Array): MacroButton[] {
  let files: Record<string, Uint8Array>;
  try {
    files = unzipSync(bytes);
  } catch {
    return [];
  }
  const packed = new Map<string, Uint8Array>();
  for (const [name, data] of Object.entries(files)) packed.set(name.replace(/\\/g, "/"), data);
  return macroButtonsFromPackage(packed);
}

export function macroButtonsFromPackage(files: Map<string, Uint8Array>): MacroButton[] {
  const workbook = text(files, "xl/workbook.xml");
  const workbookRels = relationships(text(files, "xl/_rels/workbook.xml.rels"));
  const buttons: MacroButton[] = [];
  for (const sheet of elements(workbook, "sheet")) {
    const name = sheet.name ?? "Sheet";
    const target = sheet.id ? workbookRels.get(sheet.id) : undefined;
    if (!target) continue;
    const sheetPath = relTarget(target, "xl/workbook.xml");
    const relPath = sheetPath.replace(/([^/]+)$/, "_rels/$1.rels");
    const sheetRels = relationships(text(files, relPath));
    for (const relTargetPath of sheetRels.values()) {
      if (!/vmlDrawing/i.test(relTargetPath)) continue;
      const vmlPath = relTarget(relTargetPath, sheetPath);
      const vml = text(files, vmlPath);
      for (const shape of vml.split("<v:shape ").slice(1)) {
        const macro = shape.match(/<x:FmlaMacro>([^<]*)<\/x:FmlaMacro>/)?.[1]?.trim() ?? "";
        if (!macro) continue;
        const id = shape.match(/(?:^|\s)id="([^"]+)"/)?.[1] ?? macroName(macro);
        const anchor = shape.match(/<x:Anchor>\s*([^<]*?)\s*<\/x:Anchor>/)?.[1]?.replace(/\s+/g, " ") ?? "";
        buttons.push({
          sheet: name,
          id,
          caption: captionOf(shape, id, macro),
          macro: macroName(macro),
          anchor,
        });
      }
    }
  }
  return buttons;
}

function captionOf(shape: string, id: string, macro: string): string {
  const box = shape.match(/<v:textbox[\s\S]*?<\/v:textbox>/)?.[0] ?? "";
  const text = decode(box.replace(/<[^>]+>/g, " ")).replace(/\s+/g, " ").trim();
  if (text) return text;
  const name = macroName(macro);
  if (id === "BHABtn_1") return "BHA 1";
  if (id === "BHA2") return "BHA 2";
  return CAPTIONS[name] ?? name;
}

function decode(value: string): string {
  return value.replace(/&amp;/g, "&").replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, "\"");
}

function text(files: Map<string, Uint8Array>, path: string): string {
  const data = files.get(path);
  return data ? strFromU8(data) : "";
}

function elements(xml: string, tag: string): Array<Record<string, string>> {
  return [...xml.matchAll(new RegExp(`<${tag}\\b([^>]*)>`, "g"))].map((match) => attributes(match[1] ?? ""));
}

function relationships(xml: string): Map<string, string> {
  const map = new Map<string, string>();
  for (const rel of elements(xml, "Relationship")) {
    if (rel.Id && rel.Target) map.set(rel.Id, rel.Target);
  }
  return map;
}

function attributes(source: string): Record<string, string> {
  const values: Record<string, string> = {};
  for (const match of source.matchAll(/([A-Za-z0-9:]+)="([^"]*)"/g)) {
    const key = match[1]?.split(":").pop();
    if (key && match[2] != null) values[key] = match[2];
  }
  return values;
}

function relTarget(target: string, base: string): string {
  if (target.startsWith("/")) return target.replace(/^\//, "");
  const stack = base.split("/").slice(0, -1);
  for (const part of target.split("/")) {
    if (part === "..") stack.pop();
    else if (part !== ".") stack.push(part);
  }
  return stack.join("/");
}

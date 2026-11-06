// The Data tab's tool rules (MDL_ToolHours IsThirdPartySeed and its helpers). The 3rd Party Tools
// table and the StrapSheets serial list both sort FieldCap's inventory with these, so a tool the
// sheet treats as a 3rd-party tool is the same tool the phone offers for a strap sheet.

import type { InventoryItem } from "./inventory";

/** VBA Trim: spaces only. */
function vbTrim(text: string): string {
  return text.replace(/^ +| +$/g, "");
}

function has(text: string, find: string): boolean {
  return text.toLowerCase().includes(find.toLowerCase());
}

function same(a: string, b: string): boolean {
  return a.toLowerCase() === b.toLowerCase();
}

export function normSerial(text: string): string {
  return text.replace(/[^A-Za-z0-9]/g, "").toUpperCase();
}

export function looksLikeSerial(text: string): boolean {
  const t = vbTrim(text);
  if (normSerial(t).length < 3) return false;
  if (t.toLowerCase().includes("hours below")) return false;
  return t.toLowerCase() !== "serial number";
}

function steerText(text: string): boolean {
  return has(text, "Orbit RSS") || has(text, "iCruise") || has(text, "i-Cruise");
}

export function motorLike(serial: string, desc: string, itemName: string): boolean {
  if (vbTrim(serial) === "") return false;
  if (has(desc, "Mud Motor") || has(itemName, "Mud Motor")) return true;
  // Bare SubCategory "Motor" is a FieldCap Other Tools type, not a mud motor.
  return steerText(desc) || steerText(itemName);
}

/** MWD, tubulars, bits, collars and pipe: inventory the 3rd Party Tools table never lists. */
export function skippedToolKind(cat: string, subCat: string, desc: string, itemName: string): boolean {
  const blob = `${vbTrim(cat)} ${vbTrim(subCat)} ${vbTrim(desc)} ${vbTrim(itemName)}`.toLowerCase();
  if (same(vbTrim(cat), "MWD")) return true;
  if (blob.includes("mwd kit") || blob.includes("mwd other")) return true;
  if (blob.includes("dd tubular") || blob.includes("drill bit")) return true;
  if (` ${vbTrim(subCat).toLowerCase()} `.includes(" tubular ")) return true;
  return ["slick collar", "pony collar", "drill pipe", "hwdp"].some((token) => blob.includes(token));
}

export function otherToolsCategory(cat: string): boolean {
  const c = vbTrim(cat);
  return same(c, "Other Inventory") || same(c, "Other Tools") || same(c, "Other");
}

/** FieldCap's Other family: the BHA source, the category, or an "other" sub-category that is not MWD. */
export function otherToolsFamily(src: string, cat: string, subCat: string): boolean {
  if (same(vbTrim(src), "Other Inventory")) return true;
  if (otherToolsCategory(cat)) return true;
  const sub = vbTrim(subCat);
  return has(sub, "other") && !has(sub, "mwd");
}

/** IsThirdPartySeed: a serial in the Other family that is neither a motor nor a skipped kind. */
export function thirdPartySeed(src: string, serial: string, desc: string, cat: string, subCat: string, itemName: string): boolean {
  if (!looksLikeSerial(serial)) return false;
  if (motorLike(serial, desc, itemName)) return false;
  if (skippedToolKind(cat, subCat, desc, itemName)) return false;
  return otherToolsFamily(src, cat, subCat);
}

/** Slick and pony collars and monels, steel or NM: the NM Slick Monel, NM Pony and NM Flex Monel strap sheets. */
function collarOrMonel(itemName: string, subCategory: string): boolean {
  const blob = `${itemName} ${subCategory}`.toLowerCase();
  return ["slick collar", "pony collar", "monel", "pony", "slick"].some((token) => blob.includes(token));
}

/**
 * Whether an inventory item belongs on the StrapSheets serial list. Motors, collars and monels
 * stay. The rest of the Other family goes through the Data tab's 3rd Party Tools rule. Every other
 * category is DD tooling and stays.
 */
export function strapCandidate(item: InventoryItem): boolean {
  const { serialNumber: serial, itemName, category, subCategory } = item;
  if (!looksLikeSerial(serial)) return false;
  if (motorLike(serial, itemName, itemName)) return true;
  if (collarOrMonel(itemName, subCategory)) return true;
  if (otherToolsFamily("", category, subCategory)) return thirdPartySeed("", serial, itemName, category, subCategory, itemName);
  return true;
}

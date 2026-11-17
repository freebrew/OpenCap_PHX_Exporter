// The FieldCap job a tab is showing, and the field list FieldCap's OData $metadata publishes.
// Both only read: the URL of an open tab, and the schema document.

/**
 * The two identifiers a FieldCap job carries. `clientJobId` (21785) is the internal id in the URL
 * `/Jobs/<id>` and the only value OData accepts. `jobNumber` (35783) is the Job Code people use —
 * the number in the tab title and the "Job Code" CSV column. Everything a person reads shows the
 * Job Number; only OData request URLs and OC_LastJobId keep the ClientJobId.
 */
export interface FieldCapJobRef {
  clientJobId: string;
  /** Blank until the tab title or a JobCode lookup resolves it. */
  jobNumber: string;
  title: string;
}

/**
 * The Job Number from a FieldCap tab title: "Northwind Energy Ltd., 20803 - NWE HZ WILLOWBEND …".
 * Blank when the title has no "<4–6 digits> - " after a comma.
 */
export function fieldcapJobNumberFromTitle(title: string): string {
  return /,\s*(\d{4,6})\s*-\s/.exec(String(title ?? ""))?.[1] ?? "";
}

/** OpenCap jobIdFromUrl: /Jobs/<id> in the route, else ClientJobId eq <id> in the query. */
export function fieldcapJobIdFromUrl(url: string): string {
  const raw = String(url ?? "");
  const fromRoute = /\/Jobs\/(\d+)\b/i.exec(raw)?.[1];
  if (fromRoute) return fromRoute;
  let query = raw;
  try {
    query = decodeURIComponent(new URL(raw).search || "");
  } catch {
    // Not a full URL; search the text as it is.
  }
  return /ClientJobId\s+eq\s+(\d+)/i.exec(query)?.[1] ?? "";
}

export interface EntityFields {
  entity: string;
  fields: string[];
}

export interface StationFields {
  md: string | null;
  inc: string | null;
  azm: string | null;
  tvd: string | null;
  /** Bit depth, when FieldCap publishes it: bit depth minus `md` is the survey offset. */
  bit: string | null;
}

/** EntityType names and their Property / NavigationProperty names, in document order. */
export function metadataEntities(xml: string): EntityFields[] {
  const out: EntityFields[] = [];
  for (const match of xml.matchAll(/<EntityType\b[^>]*\bName="([^"]+)"[^>]*>([\s\S]*?)<\/EntityType>/g)) {
    const fields = [...(match[2] ?? "").matchAll(/<(?:Property|NavigationProperty)\b[^>]*\bName="([^"]+)"/g)].map((field) => field[1] ?? "");
    out.push({ entity: match[1] ?? "", fields });
  }
  return out;
}

/** The directional fields of a survey entity, when FieldCap publishes them. */
export function stationFields(fields: string[]): StationFields {
  const find = (test: RegExp) => fields.find((field) => test.test(field)) ?? null;
  return {
    md: find(/^(MD|MeasuredDepth|Measured_Depth|SurveyDepth|Depth)$/i) ?? find(/measured.?depth/i),
    inc: find(/^(Inc|Incl|Inclination)$/i) ?? find(/inclination/i),
    azm: find(/^(Az|Azi|Azm|Azimuth)$/i) ?? find(/azimuth/i),
    tvd: find(/^(TVD|TrueVerticalDepth|True_Vertical_Depth)$/i) ?? find(/true.?vertical.?depth/i),
    bit: find(/^(BitDepth|Bit_Depth|BitMD)$/i) ?? find(/^bit.?depth/i),
  };
}

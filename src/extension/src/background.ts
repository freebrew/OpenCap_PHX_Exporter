import { startOpenCapExporter } from "virtual:opencap-exporter";
import { metadataEntities, odataRows, stationFields, surveyStations, type StationFields, type SurveyStation } from "../../packages/fieldcap-sync/src/index";

// The OpenCap exporter's service worker runs here unchanged: FETCH_ALL on the "fieldcap-fetch"
// port, its content script on FieldCap tabs, and its CSVs in chrome.storage.local. It only reads.
// Its update badge and icon would land on the preview's own button, so those calls do nothing here.
const quietAction = {
  setIcon: async () => undefined,
  setTitle: async () => undefined,
  setBadgeText: async () => undefined,
  setBadgeBackgroundColor: async () => undefined,
  setBadgeTextColor: async () => undefined,
};
startOpenCapExporter(new Proxy(chrome, { get: (target, key) => (key === "action" ? quietAction : Reflect.get(target, key)) }), () => undefined);

const EXPORTER_SCRIPT: chrome.scripting.RegisteredContentScript = {
  id: "opencap-exporter",
  matches: ["*://*.phxtech.com/*"],
  js: ["browser-api.js", "content.js"],
  runAt: "document_idle",
  persistAcrossSessions: true,
};

async function registerExporterScript(): Promise<void> {
  const existing = await chrome.scripting.getRegisteredContentScripts({ ids: [EXPORTER_SCRIPT.id] });
  if (existing.length === 0) await chrome.scripting.registerContentScripts([EXPORTER_SCRIPT]);
}
void registerExporterScript().catch((error: unknown) => console.warn("OpenCap content script", error));

const ODATA = "https://fieldcap-cdn.phxtech.com/odata";
const SURVEY_ENTITIES = /^(SurveySheetEntry|SurveySheetEntries|SurveyEntry|Survey)$/i;
const ACTIVITY_ENTITIES = /^(ActivityLog|ActivityLogs)$/i;

chrome.runtime.onMessage.addListener((message, _sender, sendResponse) => {
  switch (message?.type) {
    case "phx-fieldcap-poll":
      void pollFieldCap(String(message.entity ?? "ClientJobs")).then(sendResponse);
      return true;
    case "phx-fieldcap-jobs":
      void fieldCapJobs().then(sendResponse);
      return true;
    case "phx-fieldcap-fields":
      void fieldCapFields().then(sendResponse);
      return true;
    case "phx-fieldcap-surveys":
      void fieldCapSurveys(String(message.clientJobId ?? "")).then(sendResponse);
      return true;
    default:
      return undefined;
  }
});

chrome.alarms.create("fieldcap-readonly", { periodInMinutes: 1 });
chrome.alarms.onAlarm.addListener((alarm) => {
  if (alarm.name === "fieldcap-readonly") void pollFieldCap("ClientJobs");
});

async function pollFieldCap(entity: string): Promise<{ ok: boolean; count?: number; error?: string }> {
  const tabs = await chrome.tabs.query({ url: "https://fieldcap-cdn.phxtech.com/*" });
  const tabId = tabs[0]?.id;
  if (tabId == null) return { ok: false, error: "Open FieldCap in a tab first. This poll only reads." };
  try {
    return await chrome.tabs.sendMessage(tabId, { type: "phx-fieldcap-poll", entity });
  } catch (error) {
    return { ok: false, error: error instanceof Error ? error.message : String(error) };
  }
}

const JOBS_TOP = 300;

/** FieldCap's Jobs list (newest first), read-only, using the browser's FieldCap session. */
async function fieldCapJobs(): Promise<
  | { ok: true; jobs: Array<{ clientJobId: string; jobNumber: string; name: string; status: string; ended: boolean }> }
  | { ok: false; error: string }
> {
  // Same expand as the exporter's job details: with a $select FieldCap drops the JobStatus name.
  const filter = encodeURIComponent("null eq DeletedBy");
  try {
    const response = await fetch(`${ODATA}/ClientJobs?$expand=JobStatus&$filter=${filter}&$orderby=ClientJobId desc&$top=${JOBS_TOP}`, { credentials: "include", headers: { accept: "application/json" } });
    if (!response.ok) return { ok: false, error: `FieldCap returned ${response.status}. Sign in to FieldCap in a Chrome tab.` };
    const body = await response.json() as { value?: Array<{ ClientJobId?: unknown; JobCode?: unknown; ClientJobName?: unknown; EndDate?: unknown; JobStatus?: { JobStatusName?: unknown } | null }> };
    const jobs = (body.value ?? []).map((row) => ({
      clientJobId: String(row.ClientJobId ?? "").trim(),
      jobNumber: String(row.JobCode ?? "").trim(),
      name: String(row.ClientJobName ?? "").trim(),
      status: String(row.JobStatus?.JobStatusName ?? "").trim(),
      ended: row.EndDate != null && String(row.EndDate).trim() !== "",
    })).filter((job) => job.clientJobId !== "" && job.jobNumber !== "");
    return { ok: true, jobs };
  } catch (error) {
    return { ok: false, error: error instanceof Error ? error.message : String(error) };
  }
}

/** Field names only (no values) for FieldCap's survey and activity entities, from $metadata. */
async function fieldCapFields(): Promise<
  | { ok: true; survey: Array<{ entity: string; fields: string[]; station: ReturnType<typeof stationFields> }>; activity: Array<{ entity: string; fields: string[] }> }
  | { ok: false; error: string }
> {
  try {
    const response = await fetch(`${ODATA}/$metadata`, { credentials: "include", headers: { accept: "application/xml" } });
    if (!response.ok) return { ok: false, error: `FieldCap returned ${response.status}. Sign in to FieldCap in a tab, then check again.` };
    const entities = metadataEntities(await response.text());
    return {
      ok: true,
      survey: entities.filter((entry) => SURVEY_ENTITIES.test(entry.entity)).map((entry) => ({ ...entry, station: stationFields(entry.fields) })),
      activity: entities.filter((entry) => ACTIVITY_ENTITIES.test(entry.entity)),
    };
  } catch (error) {
    return { ok: false, error: error instanceof Error ? error.message : String(error) };
  }
}

const FIELDCAP_ORIGIN = "https://fieldcap-cdn.phxtech.com";

function fieldCapUrl(url: string): boolean {
  try {
    return new URL(url).origin === FIELDCAP_ORIGIN;
  } catch {
    return false;
  }
}

/** Read-only SurveySheetEntries for one job. Stations only; nothing is written back to FieldCap. */
async function fieldCapSurveys(clientJobId: string): Promise<{ ok: true; stations: SurveyStation[] } | { ok: false; error: string }> {
  if (!/^\d+$/.test(clientJobId)) return { ok: false, error: "FieldCap job id is not numeric." };
  try {
    const meta = await fetch(`${ODATA}/$metadata`, { credentials: "include", headers: { accept: "application/xml" } });
    if (!meta.ok) return { ok: false, error: `FieldCap returned ${meta.status}. Sign in to FieldCap in a Chrome tab.` };
    const entity = metadataEntities(await meta.text()).find((entry) => /^SurveySheetEntr/i.test(entry.entity));
    const fields: StationFields = stationFields(entity?.fields ?? []);
    if (!fields.md || !fields.inc || !fields.azm) return { ok: true, stations: [] };
    const names = [fields.md, fields.inc, fields.azm, fields.tvd, fields.bit].filter((name): name is string => name != null);
    if (names.some((name) => !/^[A-Za-z_][A-Za-z0-9_]*$/.test(name))) return { ok: false, error: "FieldCap survey fields were not accepted." };
    const filters = [
      `(SurveySheet/ClientJobId eq ${clientJobId}) and (null eq DeletedBy) and (null eq OriginalSurveyId)`,
      `(SurveySheet/ClientJobId eq ${clientJobId}) and (null eq DeletedBy)`,
      `(SurveySheet/ClientJobId eq ${clientJobId})`,
    ];
    // The bit depth is wanted but not needed: if FieldCap refuses a request that asks for it, ask again without.
    const attempts: StationFields[] = fields.bit ? [fields, { ...fields, bit: null }] : [fields];
    let empty = false;
    for (const attempt of attempts) {
      const select = encodeURIComponent(names.filter((name) => attempt.bit != null || name !== fields.bit).join(","));
      for (const filter of filters) {
        const page = await surveyPages(`${ODATA}/SurveySheetEntries?$filter=${encodeURIComponent(filter)}&$select=${select}&$top=500`);
        if (page.status !== 200) continue;
        empty = true;
        if (page.rows.length === 0) continue;
        return { ok: true, stations: surveyStations(page.rows, attempt) };
      }
      if (empty) break;
    }
    return empty ? { ok: true, stations: [] } : { ok: false, error: "FieldCap did not return survey stations. Sign in to FieldCap in a Chrome tab." };
  } catch (error) {
    return { ok: false, error: error instanceof Error ? error.message : String(error) };
  }
}

async function surveyPages(start: string): Promise<{ status: number; rows: Record<string, unknown>[] }> {
  const rows: Record<string, unknown>[] = [];
  let url: string | null = start;
  let status = 0;
  for (let page = 0; page < 8 && url; page += 1) {
    if (!fieldCapUrl(url)) break;
    const response = await fetch(url, { credentials: "include", headers: { accept: "application/json" } });
    status = response.status;
    if (!response.ok) return { status, rows: [] };
    const body = await response.json() as { "@odata.nextLink"?: unknown };
    rows.push(...odataRows(body));
    const next = body["@odata.nextLink"];
    url = typeof next === "string" && fieldCapUrl(next) ? next : null;
    if (rows.length >= 2000) break;
  }
  return { status, rows };
}

chrome.action.onClicked.addListener(async () => {
  const url = chrome.runtime.getURL("app.html");
  const tabs = await chrome.tabs.query({ url });
  const existing = tabs[0];
  if (existing?.id != null) {
    await chrome.tabs.update(existing.id, { active: true });
    if (existing.windowId != null) await chrome.windows.update(existing.windowId, { focused: true });
    return;
  }
  await chrome.tabs.create({ url });
});

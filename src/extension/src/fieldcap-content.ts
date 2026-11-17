chrome.runtime.onMessage.addListener((message, _sender, sendResponse) => {
  if (message?.type !== "phx-fieldcap-poll") return undefined;
  const entity = String(message.entity ?? "ClientJobs");
  fetch(`https://fieldcap-cdn.phxtech.com/odata/${encodeURIComponent(entity)}?$top=50`, {
    credentials: "include",
    headers: { accept: "application/json;odata.metadata=minimal" },
  }).then(async (response) => {
    if (!response.ok) {
      sendResponse({ ok: false, error: `FieldCap returned ${response.status}.` });
      return;
    }
    const payload: unknown = await response.json();
    const rows = payload && typeof payload === "object" && Array.isArray((payload as { value?: unknown }).value)
      ? (payload as { value: unknown[] }).value
      : [];
    sendResponse({ ok: true, count: rows.length });
  }).catch((error: unknown) => {
    sendResponse({ ok: false, error: error instanceof Error ? error.message : String(error) });
  });
  return true;
});

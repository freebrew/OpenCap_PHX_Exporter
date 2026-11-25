import { useEffect, useRef, useState, useSyncExternalStore } from "react";
import { DEFAULT_MAIL_CLIENT, MAIL_CLIENT_LABELS, MAIL_CLIENTS, type MailClient } from "./clients";
import { allowMailTabDetection, downloadHelper, helperStatus, suggestClient } from "./browserDetect";
import { getMailUi, phaseActive, subscribeMail } from "./phases";
import { PROVIDER_SETUP } from "./providerConfig";
import { readClientConfig, readMailClient, saveClientConfig, saveMailClient, type ClientConfig } from "./platform";
import { retryDailyOpen } from "./sendDaily";

const INSTALL_STEPS = [
  "Extract phx-outlook-host.zip.",
  "Run install-host.cmd.",
  "Click Recheck. Repair and uninstall scripts are in the same folder.",
];

export function MailSettings(props: { onEmail: () => Promise<void>; onStatus: (message: string) => void }) {
  const ui = useSyncExternalStore(subscribeMail, getMailUi, getMailUi);
  const [selected, setSelected] = useState<MailClient | null>(DEFAULT_MAIL_CLIENT);
  const [reason, setReason] = useState("Calls Outlook (Classic or New) by default. No helper install or sign-in needed.");
  const [helper, setHelper] = useState("Optional Classic Outlook helper not checked yet.");
  const [config, setConfig] = useState<ClientConfig>({ microsoftClientId: "", googleClientId: "", tenant: "common", gmailEnabled: false });
  const touched = useRef(false);

  useEffect(() => {
    let cancel = false;
    void (async () => {
      const [saved, setup, status] = await Promise.all([readMailClient(), readClientConfig(), helperStatus()]);
      if (cancel) return;
      setConfig(setup);
      setHelper(status);
      if (saved) {
        setSelected(saved);
        setReason(`Remembered: ${MAIL_CLIENT_LABELS[saved]}.`);
        return;
      }
      const suggestion = await suggestClient();
      if (cancel || touched.current) return;
      setSelected(suggestion.client ?? DEFAULT_MAIL_CLIENT);
      setReason(suggestion.reason);
    })();
    return () => {
      cancel = true;
    };
  }, []);

  async function recheck() {
    setHelper(await helperStatus());
  }

  async function email() {
    if (!selected || phaseActive(ui.phase)) return;
    await saveMailClient(selected);
    await saveClientConfig(config);
    setReason(`Remembered: ${MAIL_CLIENT_LABELS[selected]}.`);
    await props.onEmail();
  }

  const retryLabel = selected === "outlook" || selected === "classic" ? "Open in Outlook" : "Open Drafts";

  return (
    <fieldset id="tool-report" className="mail-setup tool-bound">
      <legend>Email Report</legend>
      <p className="mail-note">{reason} Nothing is sent automatically.</p>
      {MAIL_CLIENTS.map((client) => (
        <label key={client}>
          <input
            type="radio"
            name="mail-client"
            checked={selected === client}
            onChange={() => {
              touched.current = true;
              setSelected(client);
            }}
          />
          {MAIL_CLIENT_LABELS[client]}
        </label>
      ))}
      <button type="button" disabled={selected == null || phaseActive(ui.phase)} onClick={() => void email()}>Email Report</button>
      <button type="button" disabled={!ui.canRetryOpen || phaseActive(ui.phase)} onClick={() => void retryDailyOpen(props.onStatus)}>{retryLabel}</button>
      {ui.message ? <p className="mail-note">{ui.message}</p> : null}
      <details className="mail-note" open={selected === "classic"}>
        <summary>Classic Outlook native helper (optional)</summary>
        <p className="mail-note">Not required for Outlook (Classic or New) — default mode opens an editable .eml draft directly. Install this optional helper only if you want direct Classic Outlook COM compose without a draft file.</p>
        <p className="mail-note">{helper}</p>
        <button type="button" onClick={() => void downloadHelper().then(() => props.onStatus("The optional helper package is downloading. Extract it and run install-host.cmd."))}>Download Optional Helper (.zip)</button>
        <button type="button" onClick={() => void recheck()}>Recheck</button>
        <ol className="mail-note">
          {INSTALL_STEPS.map((step) => <li key={step}>{step}</li>)}
        </ol>
      </details>
      <details className="mail-note" open={selected === "microsoft" || selected === "gmail"}>
        <summary>Webmail sign-in settings (Microsoft / Gmail)</summary>
        <button type="button" onClick={() => void allowMailTabDetection().then(() => suggestClient()).then((suggestion) => {
          if (!touched.current && suggestion.client) setSelected(suggestion.client);
          setReason(suggestion.reason);
        })}>Check Open Mail Tabs</button>
        <label className="mail-note">Microsoft client id
          <input aria-label="Microsoft client id" value={config.microsoftClientId} onChange={(event) => setConfig({ ...config, microsoftClientId: event.target.value })} />
        </label>
        <label className="mail-note">Tenant
          <input aria-label="Microsoft tenant" value={config.tenant} onChange={(event) => setConfig({ ...config, tenant: event.target.value })} />
        </label>
        <label className="mail-note">Gmail client id
          <input aria-label="Gmail client id" value={config.googleClientId} onChange={(event) => setConfig({ ...config, googleClientId: event.target.value })} />
        </label>
        <label>
          <input type="checkbox" checked={config.gmailEnabled} onChange={(event) => setConfig({ ...config, gmailEnabled: event.target.checked })} />
          Gmail drafts (off until Google verifies gmail.compose, or this account is a test user)
        </label>
        <ul>
          {PROVIDER_SETUP.map((line) => <li key={line}>{line}</li>)}
        </ul>
      </details>
    </fieldset>
  );
}

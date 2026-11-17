import { useEffect, useRef, useState, useSyncExternalStore } from "react";
import { MAIL_CLIENT_LABELS, MAIL_CLIENTS, type MailClient } from "./clients";
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
  const [selected, setSelected] = useState<MailClient | null>(null);
  const [reason, setReason] = useState("Choose the mail client for this report.");
  const [helper, setHelper] = useState("Checking the Outlook helper…");
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
      setSelected(suggestion.client);
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

  return (
    <fieldset id="tool-report" className="mail-setup tool-bound">
      <legend>Email Report</legend>
      <p className="mail-note">{reason} Detection does not save this choice. Nothing is sent automatically.</p>
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
      <p className="mail-note">{helper}</p>
      <button type="button" onClick={() => void downloadHelper().then(() => props.onStatus("The helper package is downloading. Extract it and run install-host.cmd."))}>Install Helper</button>
      <button type="button" onClick={() => void recheck()}>Recheck</button>
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
      <button type="button" disabled={selected == null || phaseActive(ui.phase)} onClick={() => void email()}>Email Report</button>
      <button type="button" disabled={!ui.canRetryOpen || phaseActive(ui.phase)} onClick={() => void retryDailyOpen(props.onStatus)}>Open Drafts</button>
      {ui.message ? <p className="mail-note">{ui.message}</p> : null}
      <ol className="mail-note">
        {INSTALL_STEPS.map((step) => <li key={step}>{step}</li>)}
      </ol>
      <details className="mail-note">
        <summary>Mailbox registration</summary>
        <ul>
          {PROVIDER_SETUP.map((line) => <li key={line}>{line}</li>)}
        </ul>
      </details>
    </fieldset>
  );
}

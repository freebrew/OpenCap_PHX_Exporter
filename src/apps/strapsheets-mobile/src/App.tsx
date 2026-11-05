// StrapSheets on the phone. Entries stay on this phone. Email opens the mail app with the PDF attached.

import { useMemo, useState, useEffect } from "react";
import { DEFAULT_THREADS, ENTRY_FAMILIES, ENTRY_WHICH_OPTIONS, measuredSerials, type EntryPayload, type MeasureRange } from "../../../packages/strapsheets-core/src/index";
import { newSession, type StrapSession } from "./queue";
import { sendEntries } from "./send";
import { Sessions } from "./Sessions";
import { SettingsScreen, type Part } from "./SettingsScreen";
import { defaultSettings, loadSettings, saveSettings, type Settings as SettingsValue } from "./settings";
import { readJson, writeJson } from "./storage";
import { Wizard } from "./Wizard";

type Tab = "entry" | "sessions" | "settings";

const SESSIONS_KEY = "sessions";
const RANGES_KEY = "ranges";
const OPTIONS_KEY = "options";

export function App() {
  const [tab, setTab] = useState<Tab>("entry");
  const [settingsPart, setSettingsPart] = useState<Part>("job");
  const [settings, setSettings] = useState<SettingsValue>(defaultSettings());
  const [sessions, setSessions] = useState<StrapSession[]>([]);
  const [ranges, setRanges] = useState<MeasureRange[]>([]);
  const [threads, setThreads] = useState<string[]>(DEFAULT_THREADS);
  const [busy, setBusy] = useState(false);
  const [message, setMessage] = useState("");

  useEffect(() => {
    void (async () => {
      const loaded = await loadSettings();
      setSettings(loaded);
      setSessions(await readJson<StrapSession[]>(SESSIONS_KEY, []));
      setRanges(await readJson<MeasureRange[]>(RANGES_KEY, []));
      const options = await readJson<{ threads?: string[] } | null>(OPTIONS_KEY, null);
      if (options?.threads && options.threads.length > 0) setThreads(options.threads);
    })();
  }, []);

  const updateSettings = (next: SettingsValue) => {
    setSettings(next);
    void saveSettings(next);
  };

  const saveEntry = (entry: EntryPayload) => {
    const added = newSession(settings.jobNumber, entry);
    setSessions((current) => {
      const next = [...current, added];
      void writeJson(SESSIONS_KEY, next);
      return next;
    });
  };

  const measured = useMemo(() => {
    const local = sessions.filter((item) => item.jobNumber === settings.jobNumber).map((item) => item.entry);
    return new Set(measuredSerials(ENTRY_FAMILIES, local));
  }, [sessions, settings.jobNumber]);

  const send = async (chosen: StrapSession[], recipients: string[]) => {
    const withRecipients = { ...settings, recipients };
    updateSettings(withRecipients);
    setBusy(true);
    try {
      await sendEntries(chosen.map((item) => item.entry), withRecipients);
      setMessage(`Handed to the mail app: ${chosen.length} ${chosen.length === 1 ? "sheet" : "sheets"}.`);
    } catch (error) {
      setMessage(errorText(error));
    } finally {
      setBusy(false);
    }
  };

  return (
    <div className="app">
      <header>
        <strong>StrapSheets</strong>
        <button type="button" className="status" title="Switch job" onClick={() => { setSettingsPart("job"); setTab("settings"); }}>
          {settings.jobNumber ? `Job ${settings.jobNumber}` : "No job"}
        </button>
      </header>
      {message && <div className="message" onClick={() => setMessage("")}>{message}</div>}
      <main>
        {tab === "entry" && (
          <Wizard
            families={ENTRY_FAMILIES}
            whichOptions={ENTRY_WHICH_OPTIONS}
            threads={threads}
            items={[]}
            measured={measured}
            history={sessions.map((item) => item.entry)}
            ranges={ranges}
            onSave={saveEntry}
          />
        )}
        {tab === "sessions" && (
          <Sessions
            sessions={sessions}
            combined={settings.combined}
            recipients={settings.recipients}
            busy={busy}
            onCombined={(combined) => updateSettings({ ...settings, combined })}
            onSend={(chosen, recipients) => void send(chosen, recipients)}
          />
        )}
        {tab === "settings" && (
          <SettingsScreen
            key={settingsPart}
            initialPart={settingsPart}
            settings={settings}
            onSettings={updateSettings}
          />
        )}
      </main>
      <nav className="tabs">
        {(["entry", "sessions", "settings"] as const).map((name) => (
          <button key={name} type="button" className={tab === name ? "selected" : ""} onClick={() => { setSettingsPart("job"); setTab(name); }}>
            {tabLabel(name)}
          </button>
        ))}
      </nav>
    </div>
  );
}

function tabLabel(tab: Tab): string {
  switch (tab) {
    case "entry":
      return "New entry";
    case "sessions":
      return "Sessions";
    case "settings":
      return "Settings";
    default: {
      const never: never = tab;
      return never;
    }
  }
}

function errorText(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}

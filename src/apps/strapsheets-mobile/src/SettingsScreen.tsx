// The job being strapped and the addresses the PDF is addressed to.

import { useState } from "react";
import { Paged } from "./Paged";
import type { Settings as SettingsValue } from "./settings";

interface SettingsProps {
  settings: SettingsValue;
  onSettings: (settings: SettingsValue) => void;
}

const RECENT_JOBS = 10;

export type Part = "job" | "email";
const PARTS: Part[] = ["job", "email"];

export function SettingsScreen({ initialPart = "job", ...props }: SettingsProps & { initialPart?: Part }) {
  const [part, setPart] = useState<Part>(initialPart);
  return (
    <section className="settings">
      <div className="segments">
        {PARTS.map((name) => (
          <button key={name} type="button" className={part === name ? "selected" : ""} onClick={() => setPart(name)}>{partLabel(name)}</button>
        ))}
      </div>
      <SettingsPart part={part} {...props} />
    </section>
  );
}

function SettingsPart({ part, ...props }: SettingsProps & { part: Part }) {
  switch (part) {
    case "job":
      return <Job {...props} />;
    case "email":
      return <Email {...props} />;
    default: {
      const never: never = part;
      return never;
    }
  }
}

function partLabel(part: Part): string {
  switch (part) {
    case "job":
      return "Job";
    case "email":
      return "Email";
    default: {
      const never: never = part;
      return never;
    }
  }
}

function Job({ settings, onSettings }: SettingsProps) {
  const [typed, setTyped] = useState("");
  const choose = (jobNumber: string) => {
    const next = jobNumber.trim();
    if (next === "") return;
    onSettings({ ...settings, jobNumber: next, recentJobs: [next, ...settings.recentJobs.filter((job) => job !== next)].slice(0, RECENT_JOBS) });
  };
  const listed = [...new Set([...(settings.jobNumber ? [settings.jobNumber] : []), ...settings.recentJobs])];
  return (
    <fieldset>
      <legend>Job {settings.jobNumber || "(none)"}</legend>
      <div className="hint">Serials are typed by hand.</div>
      <Paged
        items={listed}
        rowHeight={56}
        keyOf={(job) => job}
        empty={<p className="hint">Type a Job Number below.</p>}
        render={(job) => (
          <button type="button" className={job === settings.jobNumber ? "job selected" : "job"} onClick={() => choose(job)}>Job {job}</button>
        )}
      />
      <div className="row">
        <input type="text" inputMode="numeric" placeholder="Job Number" value={typed} onChange={(event) => setTyped(event.target.value.replace(/\s/g, ""))} />
        <button type="button" disabled={typed === ""} onClick={() => { choose(typed); setTyped(""); }}>Use</button>
      </div>
    </fieldset>
  );
}

function Email({ settings, onSettings }: SettingsProps) {
  const [extra, setExtra] = useState("");
  const toggleRecipient = (email: string) => {
    const recipients = settings.recipients.includes(email) ? settings.recipients.filter((address) => address !== email) : [...settings.recipients, email];
    onSettings({ ...settings, recipients });
  };
  return (
    <fieldset>
      <legend>Email</legend>
      <p className="hint">The phone's mail app sends the PDF. You press Send there.</p>
      <div className="sub">Recipients</div>
      <Paged
        items={settings.recipients}
        rowHeight={40}
        keyOf={(email) => email}
        empty={<p className="hint">No recipients yet. Add one below.</p>}
        render={(email) => (
          <label className="inline">
            <input type="checkbox" checked={settings.recipients.includes(email)} onChange={() => toggleRecipient(email)} /> {email}
          </label>
        )}
      />
      <div className="row">
        <input type="email" autoCapitalize="none" placeholder="add address" value={extra} onChange={(event) => setExtra(event.target.value)} />
        <button type="button" disabled={!extra.includes("@")} onClick={() => { toggleRecipient(extra.trim().toLowerCase()); setExtra(""); }}>Add</button>
      </div>
    </fieldset>
  );
}

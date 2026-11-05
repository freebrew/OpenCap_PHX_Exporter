// Every entry the phone has taken. Tick one or several and hand the PDF to the mail app.

import { useRef, useState } from "react";
import { ENTRY_FAMILIES, findFamily, serialStep } from "../../../packages/strapsheets-core/src/index";
import { Paged } from "./Paged";
import type { StrapSession } from "./queue";

interface SessionsProps {
  sessions: StrapSession[];
  combined: boolean;
  recipients: string[];
  busy: boolean;
  onCombined: (combined: boolean) => void;
  /** Remember these addresses and open the mail app with the PDFs. */
  onSend: (sessions: StrapSession[], recipients: string[]) => void;
}

const EMAIL_PATTERN = /^[^@\s,;]+@[^@\s,;]+\.[^@\s,;]+$/;

function addressParts(text: string): string[] {
  return text.split(/[,;\s]+/).map((part) => part.trim().toLowerCase()).filter(Boolean);
}

/** Comma- or semicolon-separated addresses, lower-cased; null unless every one is valid. */
export function parseAddresses(text: string): string[] | null {
  const parts = addressParts(text);
  if (parts.length === 0 || !parts.every((part) => EMAIL_PATTERN.test(part))) return null;
  return [...new Set(parts)];
}

/** Why the To field is refused, naming the part that is not an address. */
export function addressProblem(text: string): string {
  const bad = addressParts(text).find((part) => !EMAIL_PATTERN.test(part));
  if (!bad) return "";
  return /\s/.test(text.trim()) && !/[,;]/.test(text)
    ? `"${bad}" is not an address. Email addresses have no spaces; separate several with commas.`
    : `"${bad}" is not an email address.`;
}

export function Sessions({ sessions, combined, recipients, busy, onCombined, onSend }: SessionsProps) {
  const [selected, setSelected] = useState<Set<string>>(new Set());
  const press = useRef<number | null>(null);
  const [to, setTo] = useState(() => recipients.join(", "));
  const addresses = parseAddresses(to);

  const toggle = (id: string) => {
    const next = new Set(selected);
    if (next.has(id)) next.delete(id);
    else next.add(id);
    setSelected(next);
  };
  const startPress = (id: string) => {
    press.current = window.setTimeout(() => toggle(id), 450);
  };
  const endPress = () => {
    if (press.current != null) window.clearTimeout(press.current);
    press.current = null;
  };
  const chosen = sessions.filter((session) => selected.has(session.id));

  return (
    <section className="sessions">
      <div className="toolbar">
        <span>{sessions.length} {sessions.length === 1 ? "session" : "sessions"} saved on this phone</span>
      </div>
      <Paged
        items={[...sessions].reverse()}
        rowHeight={60}
        keyOf={(session) => session.id}
        empty={<p className="hint">No entries yet.</p>}
        render={(session) => (
          <label className={selected.has(session.id) ? "session selected" : "session"} onPointerDown={() => startPress(session.id)} onPointerUp={endPress} onPointerLeave={endPress}>
            <input type="checkbox" checked={selected.has(session.id)} onChange={() => toggle(session.id)} />
            <span className="serial">{serialOf(session)}</span>
            <span className="meta">{session.entry.which} · {new Date(session.createdAt).toLocaleString()}</span>
          </label>
        )}
      />
      {chosen.length > 0 && (
        <div className="send-bar">
          {chosen.length > 1 && (
            <label className="inline">
              <input type="checkbox" checked={combined} onChange={(event) => onCombined(event.target.checked)} /> One combined PDF
            </label>
          )}
          <label className="send-to">
            <span>To</span>
            <input type="email" inputMode="email" autoCapitalize="none" autoComplete="email" placeholder="name@company.com" value={to} onChange={(event) => setTo(event.target.value)} />
          </label>
          {to.trim() !== "" && !addresses && <div className="error">{addressProblem(to)}</div>}
          <div className="row">
            <button type="button" onClick={() => setSelected(new Set())}>Clear</button>
            <button type="button" className="primary" disabled={busy || !addresses} onClick={() => addresses && onSend(chosen, addresses)}>
              Email {chosen.length} {chosen.length === 1 ? "sheet" : "sheets"}
            </button>
          </div>
        </div>
      )}
    </section>
  );
}

function serialOf(session: StrapSession): string {
  const family = findFamily(ENTRY_FAMILIES, session.entry.which);
  const serial = family ? serialStep(family) : null;
  return serial && serial.kind === "text" ? session.entry.cells[serial.col] ?? "(no serial)" : "(no serial)";
}

export type MailPhase =
  | "idle"
  | "preparing"
  | "validating"
  | "writing-archive"
  | "signing-in"
  | "creating"
  | "opening"
  | "complete"
  | "partial"
  | "failed";

export function phaseActive(phase: MailPhase): boolean {
  switch (phase) {
    case "preparing":
    case "validating":
    case "writing-archive":
    case "signing-in":
    case "creating":
    case "opening":
      return true;
    case "idle":
    case "complete":
    case "partial":
    case "failed":
      return false;
    default: {
      const never: never = phase;
      return never;
    }
  }
}

export interface MailUiState {
  phase: MailPhase;
  message: string;
  canRetryOpen: boolean;
}

let state: MailUiState = { phase: "idle", message: "", canRetryOpen: false };
const listeners = new Set<() => void>();

export function getMailUi(): MailUiState {
  return state;
}

export function subscribeMail(listener: () => void): () => void {
  listeners.add(listener);
  return () => listeners.delete(listener);
}

export function setMailUi(patch: Partial<MailUiState>): void {
  state = { ...state, ...patch };
  for (const listener of listeners) listener();
}

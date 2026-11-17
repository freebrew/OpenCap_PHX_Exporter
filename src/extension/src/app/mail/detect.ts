import type { MailClient } from "./clients";

export interface ClassicProbe {
  outlookRegistered: boolean;
  mailtoCommand: string;
  defaultClient: string;
}

export function looksLikeClassic(probe: ClassicProbe): boolean {
  if (!probe.outlookRegistered) return false;
  const client = probe.defaultClient.toLowerCase();
  if (client.includes("new outlook")) return false;
  if (probe.mailtoCommand.toLowerCase().includes("outlook.exe")) return true;
  return client.includes("outlook");
}

export interface SuggestionInput {
  classicReady: boolean;
  tabVendors: Array<"microsoft" | "gmail">;
  connected: "microsoft" | "gmail" | null;
}

/** A suggestion never writes the remembered client. Disagreeing signals stay unselected. */
export function chooseSuggestion(input: SuggestionInput): { client: MailClient | null; reason: string } {
  const tabs = [...new Set(input.tabVendors)];
  const signals = [
    input.classicReady ? "classic" : null,
    ...tabs,
    tabs.length === 0 && !input.classicReady ? input.connected : null,
  ].filter((signal): signal is "classic" | "microsoft" | "gmail" => signal != null);
  const unique = [...new Set(signals)];
  if (unique.length > 1) return { client: null, reason: "More than one mail client is available. Choose one." };
  const only = unique[0] ?? null;
  switch (only) {
    case "classic":
      return { client: "classic", reason: "Classic Outlook is the default mail program." };
    case "microsoft":
      return { client: "microsoft", reason: input.connected === "microsoft" && tabs.length === 0 ? "This browser is signed in to Microsoft mail." : "Outlook on the web is open." };
    case "gmail":
      return { client: "gmail", reason: input.connected === "gmail" && tabs.length === 0 ? "This browser is signed in to Gmail." : "Gmail is open." };
    case null:
      return { client: null, reason: "Choose the mail client for this report." };
    default: {
      const never: never = only;
      return never;
    }
  }
}

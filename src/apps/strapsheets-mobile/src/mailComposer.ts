// Hands a finished PDF to the phone's mail app with the recipients already filled in.

import { registerPlugin } from "@capacitor/core";

export interface MailDraft {
  to: string[];
  subject: string;
  text: string;
  files: string[];
}

interface MailComposerPlugin {
  compose(draft: MailDraft): Promise<void>;
}

const MailComposer = registerPlugin<MailComposerPlugin>("MailComposer");

export function composeMail(draft: MailDraft): Promise<void> {
  return MailComposer.compose(draft);
}

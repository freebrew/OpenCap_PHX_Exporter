export interface MailBytes {
  name: string;
  contentType: string;
  bytes: Uint8Array;
  /** Without angle brackets. The HTML body references `cid:` plus this value. */
  contentId?: string;
}

export interface MissingFile {
  name: string;
}

export interface MailDraft {
  to: string[];
  cc: string[];
  subject: string;
  html: string;
  text: string;
  inline: MailBytes[];
  attachments: MailBytes[];
  missing: MissingFile[];
}

function copyPart(part: MailBytes): MailBytes {
  return {
    name: part.name,
    contentType: part.contentType,
    bytes: part.bytes.slice(),
    contentId: part.contentId,
  };
}

/** Copies the known fields so a caller cannot smuggle extra properties into a provider body. */
export function createDraft(input: MailDraft): MailDraft {
  return {
    to: [...input.to],
    cc: [...input.cc],
    subject: input.subject,
    html: input.html,
    text: input.text,
    inline: input.inline.map(copyPart),
    attachments: input.attachments.map(copyPart),
    missing: input.missing.map((file) => ({ name: file.name })),
  };
}

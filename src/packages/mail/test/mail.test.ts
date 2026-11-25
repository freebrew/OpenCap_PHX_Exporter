import { describe, expect, it } from "vitest";
import { buildEml } from "../../reports/src/eml";
import {
  attachmentSize,
  buildAuthorizeUrl,
  codeChallenge,
  createDraft,
  emailFromIdToken,
  encodeBase64Url,
  gmailDraftRequest,
  gmailEncodedSize,
  gmailSendBody,
  gmailWithinLimit,
  GMAIL_ENCODED_LIMIT,
  GRAPH_FILE_LIMIT,
  GRAPH_SMALL_LIMIT,
  graphDraftRequest,
  graphSendMailBody,
  graphSmallAttachmentRequest,
  graphUploadPut,
  graphUploadSessionRequest,
  MailValidationError,
  mayFallback,
  outlookDraftsUrl,
  parseTokenResponse,
  pdfTextMime,
  sendOrder,
  splitAddresses,
  toDraftMime,
  toMime,
  validateDraft,
  type MailDraft,
  type TransportResult,
} from "../src/index";

const PNG = Uint8Array.of(1, 2, 3, 4);
const BOUNDARIES = { mixed: "mixed", related: "related", alternative: "alt" };

function sample(): MailDraft {
  return createDraft({
    to: ["a@example.com"],
    cc: ["b@example.com"],
    subject: "Daily Drilling Summary - Job - North",
    html: "<p>Hello &amp; well</p>",
    text: "Hello & well",
    inline: [{ name: "daily_report_inline.png", contentType: "image/png", bytes: PNG, contentId: "daily_report.png" }],
    attachments: [{ name: "daily_report.png", contentType: "image/png", bytes: PNG }],
    missing: [{ name: "survey.pdf" }],
  });
}

const GOLDEN = [
  "To: a@example.com",
  "Cc: b@example.com",
  "Subject: Daily Drilling Summary - Job - North",
  "MIME-Version: 1.0",
  "X-PHX-Preview: evaluation",
  'Content-Type: multipart/mixed; boundary="mixed"',
  "",
  "--mixed",
  'Content-Type: multipart/related; boundary="related"',
  "",
  "--related",
  'Content-Type: multipart/alternative; boundary="alt"',
  "",
  "--alt",
  "Content-Type: text/plain; charset=utf-8",
  "Content-Transfer-Encoding: 8bit",
  "",
  "Hello & well",
  "--alt",
  "Content-Type: text/html; charset=utf-8",
  "Content-Transfer-Encoding: 8bit",
  "",
  "<p>Hello &amp; well</p>",
  "--alt--",
  "--related",
  "Content-Type: image/png",
  "Content-Transfer-Encoding: base64",
  'Content-Disposition: inline; filename="daily_report_inline.png"',
  "Content-ID: <daily_report.png>",
  "",
  "AQIDBA==",
  "--related--",
  "--mixed",
  "Content-Type: image/png",
  "Content-Transfer-Encoding: base64",
  'Content-Disposition: attachment; filename="daily_report.png"',
  "",
  "AQIDBA==",
  "--mixed--",
  "",
].join("\r\n");

function legacyBuild(message: { to: string; cc?: string; subject: string; html: string; pngBase64?: string }): string {
  const picture = message.pngBase64 ? '<p><img src="cid:corridor@phx" alt="Corridor"></p>' : "";
  const html = `<!doctype html><html><body>${message.html}${picture}</body></html>`;
  const headers = [`To: ${message.to}`, message.cc ? `Cc: ${message.cc}` : "", `Subject: ${message.subject}`, "MIME-Version: 1.0", "X-PHX-Preview: evaluation"].filter(Boolean);
  const mixed = "phx-report-mixed";
  const related = "phx-report-related";
  const parts = [
    ...headers,
    `Content-Type: multipart/mixed; boundary="${mixed}"`,
    "",
    `--${mixed}`,
    `Content-Type: multipart/related; boundary="${related}"`,
    "",
    `--${related}`,
    "Content-Type: text/html; charset=utf-8",
    "",
    html,
    `--${related}`,
    "Content-Type: image/png",
    "Content-Transfer-Encoding: base64",
    "Content-Disposition: inline; filename=\"daily_report.png\"",
    "Content-ID: <corridor@phx>",
    "",
    message.pngBase64 ?? "",
    `--${related}--`,
    `--${mixed}--`,
    "",
  ];
  return parts.join("\r\n");
}

function header(mime: string, name: string): string {
  const line = mime.split("\r\n").find((item) => item.toLowerCase().startsWith(`${name.toLowerCase()}:`));
  return line?.slice(name.length + 1).trim() ?? "";
}

describe("validation", () => {
  it("strips header controls, checks addresses, and rejects an oversized file", () => {
    expect(() => validateDraft({ ...sample(), to: ["a@example.com\r\nBcc: x@y.com"] })).toThrow(MailValidationError);
    const draft = validateDraft({ ...sample(), subject: "Hello\r\nBcc: hidden" });
    expect(draft.subject).toBe("HelloBcc: hidden");
    expect(toMime(draft, BOUNDARIES)).not.toMatch(/\r\nBcc:/);
    expect(splitAddresses("a@example.com; b@example.com, c@example.com")).toEqual(["a@example.com", "b@example.com", "c@example.com"]);
    expect(() => validateDraft({ ...sample(), to: ["not an address"] })).toThrow(MailValidationError);
    expect(() => validateDraft(sample(), { fileBytes: 2 })).toThrow(/file size/i);
  });

  it("sanitizes a path-like file name", () => {
    const draft = validateDraft({
      ...sample(),
      attachments: [{ name: "..\\..\\report.pdf", contentType: "application/pdf", bytes: Uint8Array.of(9) }],
      inline: [],
    });
    expect(draft.attachments[0]?.name).toBe("report.pdf");
  });
});

describe("canonical mime", () => {
  it("matches the golden serializer", () => {
    expect(toMime(sample(), BOUNDARIES)).toBe(GOLDEN);
  });

  it("prepends X-Unsent: 1 for editable Classic and New Outlook drafts without Message-ID", () => {
    const draftEml = toDraftMime(sample(), BOUNDARIES);
    expect(draftEml.startsWith("X-Unsent: 1\r\n")).toBe(true);
    expect(draftEml).toBe(`X-Unsent: 1\r\n${GOLDEN}`);
    expect(draftEml).not.toMatch(/^Message-ID:/im);
  });

  it("encodes a non-ascii subject and keeps both png copies", () => {
    const mime = toMime({ ...sample(), subject: "Café" }, BOUNDARIES);
    expect(mime).toContain("=?UTF-8?B?");
    expect(mime).not.toContain("Café");
    expect(mime.match(/AQIDBA==/g)).toHaveLength(2);
    expect(mime).toContain("Content-ID: <daily_report.png>");
    expect(mime).toContain('Content-Disposition: attachment; filename="daily_report.png"');
  });

  it("keeps the legacy message's recipients, subject, html, and image bytes", () => {
    const png = btoa(String.fromCharCode(...PNG));
    const legacy = legacyBuild({ to: "a@example.com", cc: "b@example.com", subject: "Day", html: "<p>Table</p>", pngBase64: png });
    const next = buildEml({ to: "a@example.com", cc: "b@example.com", subject: "Day", html: "<p>Table</p>", pngBase64: png });
    expect(header(next, "To")).toBe(header(legacy, "To"));
    expect(header(next, "Subject")).toBe(header(legacy, "Subject"));
    expect(next).toContain("<p>Table</p>");
    expect(legacy).toContain("<p>Table</p>");
    expect(next).toContain("AQIDBA==");
    expect(legacy).toContain(png);
  });
});

describe("providers", () => {
  it("posts a small Graph attachment and opens an upload session at 3 MiB", () => {
    const small = graphSmallAttachmentRequest("id 1", { name: "daily_report_inline.png", contentType: "image/png", bytes: PNG, contentId: "daily_report.png" }, true);
    expect(small.url).toContain("/messages/id%201/attachments");
    expect(small.body).toMatchObject({ isInline: true, contentId: "daily_report.png", contentBytes: "AQIDBA==" });
    const big = new Uint8Array(GRAPH_SMALL_LIMIT);
    expect(attachmentSize({ name: "a.bin", contentType: "application/octet-stream", bytes: big })).toBe("upload");
    const session = graphUploadSessionRequest("mid", { name: "a.bin", contentType: "application/octet-stream", bytes: big }, false);
    expect(session.url).toContain("createUploadSession");
    expect(session.body).toMatchObject({ AttachmentItem: { size: GRAPH_SMALL_LIMIT, isInline: false } });
    expect(attachmentSize({ name: "a.bin", contentType: "application/octet-stream", bytes: new Uint8Array(GRAPH_FILE_LIMIT + 1) })).toBe("too-big");
    expect(graphDraftRequest(sample()).body).toMatchObject({
      subject: sample().subject,
      toRecipients: [{ emailAddress: { address: "a@example.com" } }],
      ccRecipients: [{ emailAddress: { address: "b@example.com" } }],
    });
    expect(graphUploadPut("https://outlook.office.com/upload", PNG).headers?.["content-range"]).toBe("bytes 0-3/4");
    expect(outlookDraftsUrl("person@outlook.com")).toBe("https://outlook.live.com/mail/0/drafts");
    expect(outlookDraftsUrl("person@phxtech.com")).toBe("https://outlook.office.com/mail/drafts");
  });

  it("keeps the mobile Graph send body", () => {
    expect(graphSendMailBody({
      subject: "Strap",
      text: "Attached.",
      to: ["a@example.com"],
      files: [{ name: "Motor.pdf", contentType: "application/pdf", contentBytes: "QQ==" }],
    })).toEqual({
      message: {
        subject: "Strap",
        body: { contentType: "Text", content: "Attached." },
        toRecipients: [{ emailAddress: { address: "a@example.com" } }],
        attachments: [{ "@odata.type": "#microsoft.graph.fileAttachment", name: "Motor.pdf", contentType: "application/pdf", contentBytes: "QQ==" }],
      },
      saveToSentItems: true,
    });
  });

  it("preflights the encoded Gmail body", () => {
    const request = gmailDraftRequest("hello");
    expect(request.url).toBe("https://gmail.googleapis.com/gmail/v1/users/me/drafts");
    expect(request.body.message.raw).toBe(gmailSendBody("hello").raw);
    expect(gmailWithinLimit("hello")).toBe(true);
    expect(gmailWithinLimit("hello", 1)).toBe(gmailEncodedSize("hello") <= 1);
    expect(GMAIL_ENCODED_LIMIT).toBe(25 * 1024 * 1024);
  });
});

describe("oauth and transport", () => {
  it("builds PKCE authorize URLs without an extension refresh scope", async () => {
    const extension = buildAuthorizeUrl({
      authorize: "https://login.microsoftonline.com/common/oauth2/v2.0/authorize",
      clientId: "client",
      redirect: "https://glkcmpmiomoppdkmdcjaepneimgbipdl.chromiumapp.org/mail",
      scope: "openid email https://graph.microsoft.com/Mail.ReadWrite",
    }, { state: "state-1", challenge: "chal", prompt: "select_account" });
    expect(extension).toContain("state=state-1");
    expect(extension).toContain("code_challenge=chal");
    expect(extension).toContain("code_challenge_method=S256");
    expect(extension).not.toContain("offline_access");
    const mobile = buildAuthorizeUrl({
      authorize: "https://accounts.google.com/o/oauth2/v2/auth",
      clientId: "client",
      redirect: "com.phx.strapsheets://mail",
      scope: "https://www.googleapis.com/auth/gmail.send",
    }, { state: "state-2", challenge: "chal", googleOffline: true });
    expect(mobile).toContain("access_type=offline");
    expect(mobile).toContain("gmail.send");
    const digest = new Uint8Array(await crypto.subtle.digest("SHA-256", new TextEncoder().encode("verifier")));
    expect(await codeChallenge("verifier")).toBe(encodeBase64Url(digest));
    expect(parseTokenResponse({ access_token: "tok", expires_in: 3600, refresh_token: "ref", id_token: "a.b.c" })).toMatchObject({ accessToken: "tok", refreshToken: "ref" });
    expect(parseTokenResponse({ access_token: "", expires_in: 1 })).toBeNull();
    const payload = encodeBase64Url(new TextEncoder().encode(JSON.stringify({ preferred_username: "crew@phxtech.com" })));
    expect(emailFromIdToken(`h.${payload}.s`)).toBe("crew@phxtech.com");
  });

  it("allows a fallback only after a proven no-side-effect failure", () => {
    const results: TransportResult[] = [
      { status: "no-side-effect-failure", error: "denied" },
      { status: "created-and-opened", location: "drafts" },
      { status: "created-not-opened", id: "1", location: "drafts", error: "tab" },
      { status: "indeterminate", error: "timeout" },
    ];
    expect(results.map(mayFallback)).toEqual([true, false, false, false]);
  });
});

describe("StrapSheets send order and pdf mime", () => {
  const online = { online: true, signedIn: true, mailVendor: "microsoft" as const, googleMailAllowed: false };
  it("keeps account, service, then share", () => {
    expect(sendOrder("account", online)).toEqual(["account", "service", "share"]);
    expect(sendOrder("account", { ...online, mailVendor: "google" })).toEqual(["service", "share"]);
    expect(sendOrder("account", { ...online, mailVendor: "google", googleMailAllowed: true })).toEqual(["account", "service", "share"]);
    expect(sendOrder("account", { ...online, online: false })).toEqual(["share"]);
    expect(sendOrder("service", { ...online, signedIn: false, mailVendor: null })).toEqual(["share"]);
    expect(sendOrder("share", online)).toEqual(["share"]);
  });

  it("attaches each PDF in base64", () => {
    const text = pdfTextMime({ to: ["a@example.com", "b@example.com"], subject: "Strap sheet: Motor", text: "Attached.", files: [{ name: "Motor MTR-001.pdf", pdf: new TextEncoder().encode("%PDF-1.4") }] }, "b1");
    expect(text).toContain("To: a@example.com, b@example.com");
    expect(text).toContain('filename="Motor MTR-001.pdf"');
    expect(text).toContain(btoa("%PDF-1.4"));
    expect(text.trimEnd().endsWith("--b1--")).toBe(true);
  });
});

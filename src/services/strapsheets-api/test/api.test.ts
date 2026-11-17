import { env } from "cloudflare:workers";
import { describe, expect, it } from "vitest";
import worker from "../src/index";
import { issueSession } from "../src/session";

async function call(path: string, init?: RequestInit): Promise<Response> {
  return worker.fetch(new Request(`http://localhost${path}`, init), env);
}

async function signedIn(path: string, init: RequestInit = {}): Promise<Response> {
  const session = await issueSession({ userId: "u1", email: "dd@example.com" }, env.TOKEN_SIGNING_KEY);
  return call(path, { ...init, headers: { authorization: `Bearer ${session.token}`, "content-type": "application/json" } });
}

const motorEntry = { which: "Motor", cells: { "2": "MTR-001", "3": "Steel", "8": "9.5" } };

describe("health", () => {
  it("answers", async () => {
    const response = await call("/health");
    expect(await response.json()).toEqual({ ok: true });
  });
});

describe("sessions", () => {
  it("refuses a route with no token", async () => {
    const response = await call("/v1/jobs");
    expect(response.status).toBe(401);
  });

  it("rejects a refresh token that was not signed here", async () => {
    const response = await call("/auth/refresh", { method: "POST", body: JSON.stringify({ refresh: "a.b" }) });
    expect(response.status).toBe(401);
  });
});

describe("sign-in", () => {
  it("sends Google sign-in to accounts.google.com with the PKCE challenge", async () => {
    const response = await call("/auth/google/start", { redirect: "manual" });
    expect(response.status).toBe(302);
    const location = new URL(response.headers.get("location") ?? "");
    expect(location.host).toBe("accounts.google.com");
    expect(location.searchParams.get("code_challenge_method")).toBe("S256");
    expect(location.searchParams.get("scope")).toBe("openid email profile");
  });

  it("accepts the phone deep link and the extension redirect", async () => {
    for (const redirect of ["com.phx.strapsheets://auth", "https://abcdef.chromiumapp.org/strapsheets"]) {
      const response = await call(`/auth/microsoft/start?redirect=${encodeURIComponent(redirect)}`, { redirect: "manual" });
      expect(response.status).toBe(302);
    }
  });

  it("refuses a redirect that would hand the session to another site", async () => {
    const response = await call(`/auth/google/start?redirect=${encodeURIComponent("https://evil.example/steal")}`, { redirect: "manual" });
    expect(response.status).toBe(400);
  });

  it("rejects a callback whose state was never issued", async () => {
    const response = await call("/auth/google/callback?code=abc&state=nope");
    expect(response.status).toBe(400);
  });

});

describe("first-run email check", () => {
  const sent: { to: string; from: string; subject: string; text: string }[] = [];
  const mailer = { send: async (message: { to: string; from: string; subject: string; text: string }) => { sent.push(message); return { messageId: "m1" }; } };
  const ask = async (path: string, body: unknown, ip: string): Promise<Response> =>
    await worker.fetch(new Request(`http://localhost${path}`, { method: "POST", headers: { "content-type": "application/json", "cf-connecting-ip": ip }, body: JSON.stringify(body) }), { ...env, EMAIL: mailer });
  const lastCode = (): string => /\b(\d{6})\b/.exec(sent[sent.length - 1]?.text ?? "")?.[1] ?? "";

  it("tells the person to try again, not the mail service's own words, when the mail cannot be sent", async () => {
    const broken = { send: async () => { throw new Error("could not find account config of sending domain"); } };
    const response = await worker.fetch(new Request("http://localhost/v1/onboard/email", { method: "POST", headers: { "content-type": "application/json", "cf-connecting-ip": "10.0.0.9" }, body: JSON.stringify({ email: "down@example.com" }) }), { ...env, EMAIL: broken });
    expect(response.status).toBe(502);
    const reply = (await response.json()) as { error: string };
    expect(reply.error).toBe("We could not send the code just now. Try again in a few minutes.");
    expect(JSON.stringify(reply)).not.toContain("sending domain");
  });

  it("mails a code to the address and accepts it once, with no account", async () => {
    const started = await ask("/v1/onboard/email", { email: " Crew@Example.com " }, "10.0.0.1");
    expect(started.status).toBe(200);
    expect(sent.at(-1)?.to).toBe("crew@example.com");
    expect(lastCode()).toMatch(/^\d{6}$/);
    const wrong = await ask("/v1/onboard/verify", { email: "crew@example.com", code: lastCode() === "000000" ? "111111" : "000000" }, "10.0.0.1");
    expect(wrong.status).toBe(400);
    const right = await ask("/v1/onboard/verify", { email: "CREW@example.com", code: ` ${lastCode()} ` }, "10.0.0.1");
    expect(await right.json()).toEqual({ verified: true, email: "crew@example.com" });
    const again = await ask("/v1/onboard/verify", { email: "crew@example.com", code: lastCode() }, "10.0.0.1");
    expect(((await again.json()) as { reason: string }).reason).toBe("expired");
  });

  it("refuses something that is not an address and a code that is not six digits", async () => {
    expect((await ask("/v1/onboard/email", { email: "not an address" }, "10.0.0.2")).status).toBe(400);
    expect((await ask("/v1/onboard/email", { email: "a@b.co, c@d.co" }, "10.0.0.2")).status).toBe(400);
    expect((await ask("/v1/onboard/verify", { email: "crew@example.com", code: "12ab56" }, "10.0.0.2")).status).toBe(400);
  });

  it("burns the code after five wrong tries, even for the right digits", async () => {
    await ask("/v1/onboard/email", { email: "burn@example.com" }, "10.0.0.3");
    const real = lastCode();
    const wrong = real === "123456" ? "654321" : "123456";
    for (let i = 0; i < 5; i += 1) await ask("/v1/onboard/verify", { email: "burn@example.com", code: wrong }, "10.0.0.3");
    const late = await ask("/v1/onboard/verify", { email: "burn@example.com", code: real }, "10.0.0.3");
    expect(((await late.json()) as { reason: string }).reason).toBe("locked");
  });

  it("stops after five codes an hour for one address", async () => {
    const statuses: number[] = [];
    for (let i = 0; i < 6; i += 1) statuses.push((await ask("/v1/onboard/email", { email: "flood@example.com" }, `10.1.0.${i}`)).status);
    expect(statuses).toEqual([200, 200, 200, 200, 200, 429]);
  });

  it("stops one caller who tries many addresses", async () => {
    const statuses: number[] = [];
    for (let i = 0; i < 21; i += 1) statuses.push((await ask("/v1/onboard/email", { email: `person${i}@example.com` }, "10.9.9.9")).status);
    expect(statuses.slice(0, 20).every((status) => status === 200)).toBe(true);
    expect(statuses[20]).toBe(429);
  });

  it("lets the web sheet's page read the answer", async () => {
    const origin = "https://n-abc123-0lu-script.googleusercontent.com";
    const response = await worker.fetch(new Request("http://localhost/v1/onboard/verify", { method: "OPTIONS", headers: { origin, "access-control-request-method": "POST" } }), { ...env, EMAIL: mailer });
    expect(response.headers.get("access-control-allow-origin")).toBe(origin);
  });
});

describe("inventory", () => {
  it("reports no items for a job that has never been published", async () => {
    const response = await signedIn("/v1/jobs/35783/inventory");
    expect(await response.json()).toEqual({ jobNumber: "35783", items: [], pulledAt: null });
  });
});

describe("options and ranges with no sheet", () => {
  it("offers the default thread list and stores ranges in D1", async () => {
    const empty = (await (await signedIn("/v1/options")).json()) as { sheet: boolean; threads: string[]; ranges: unknown[] };
    expect(empty.sheet).toBe(false);
    expect(empty.threads).toContain("4 1/2 IF [NC 50]");
    expect(empty.ranges).toEqual([]);
    const ranges = [{ family: "Motor", stepId: "d6", min: 5, max: 10 }];
    const saved = await signedIn("/v1/ranges", { method: "PUT", body: JSON.stringify({ ranges }) });
    expect(saved.status).toBe(200);
    const after = (await (await signedIn("/v1/options")).json()) as { ranges: unknown[] };
    expect(after.ranges).toEqual(ranges);
  });

  it("refuses a range that is not a number", async () => {
    const response = await signedIn("/v1/ranges", { method: "PUT", body: JSON.stringify({ ranges: [{ family: "Motor", stepId: "d6", min: "x", max: null }] }) });
    expect(response.status).toBe(400);
  });
});

describe("entries, PDF and email", () => {
  it("stores an entry with no sheet configured", async () => {
    const response = await signedIn("/v1/entries?jobNumber=35783", { method: "POST", body: JSON.stringify(motorEntry) });
    expect(response.status).toBe(201);
    const listed = (await (await signedIn("/v1/entries/35783")).json()) as { entries: { which: string }[] };
    expect(listed.entries.map((entry) => entry.which)).toContain("Motor");
  });

  it("lists the job's strapped serials, normalised, for every phone on the job", async () => {
    await signedIn("/v1/entries?jobNumber=90001", { method: "POST", body: JSON.stringify({ ...motorEntry, cells: { ...motorEntry.cells, "2": "mtr-777" } }) });
    const body = (await (await signedIn("/v1/jobs/90001/measured")).json()) as { serials: string[] };
    expect(body.serials).toEqual(["MTR777"]);
    const other = (await (await signedIn("/v1/jobs/90002/measured")).json()) as { serials: string[] };
    expect(other.serials).toEqual([]);
  });

  it("renders the Sheets card for each entry: one page per tool family, the drawing, a card each, the family title", async () => {
    const second = { ...motorEntry, cells: { ...motorEntry.cells, "2": "MTR-002" } };
    const jars = { which: "Jars", cells: { "115": "JAR-001" } };
    const response = await signedIn("/v1/pdf", { method: "POST", body: JSON.stringify({ entries: [motorEntry, second, jars] }) });
    expect(response.headers.get("content-type")).toBe("application/pdf");
    const text = new TextDecoder("latin1").decode(await response.arrayBuffer());
    expect(text.startsWith("%PDF-1.4")).toBe(true);
    expect(text).toContain("/Count 2");
    expect(text.split("/Subtype /Image")).toHaveLength(1);
    expect(text).toContain(" l ");
    for (const drawn of ["(MTR-001) Tj", "(MTR-002) Tj", "(JAR-001) Tj", "(MOTORS) Tj", "(JARS) Tj", "(H LEN) Tj"]) expect(text).toContain(drawn);
    expect(text).not.toContain("Strap Sheet -");
  });

  it("refuses a recipient who is not on the list", async () => {
    const response = await signedIn("/v1/email", { method: "POST", body: JSON.stringify({ to: ["someone@example.com"], entries: [motorEntry] }) });
    expect(response.status).toBe(403);
  });
});

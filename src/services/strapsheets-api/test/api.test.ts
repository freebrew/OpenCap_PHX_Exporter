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
    expect(text.split("/Subtype /Image")).toHaveLength(3);
    for (const drawn of ["(MTR-001) Tj", "(MTR-002) Tj", "(JAR-001) Tj", "(MOTORS) Tj", "(JARS) Tj", "(H LEN) Tj"]) expect(text).toContain(drawn);
    expect(text).not.toContain("Strap Sheet -");
  });

  it("refuses a recipient who is not on the list", async () => {
    const response = await signedIn("/v1/email", { method: "POST", body: JSON.stringify({ to: ["someone@example.com"], entries: [motorEntry] }) });
    expect(response.status).toBe(403);
  });
});

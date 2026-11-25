#!/usr/bin/env node
/**
 * Upload, check status, or publish a Chrome Web Store item using API v2.
 * Zero external dependencies (Node 20+ native fetch + crypto RS256 JWT).
 *
 * Auth (in priority order):
 *   1. CWS_SERVICE_ACCOUNT_JSON  — Google Cloud Service Account key JSON (recommended; never expires)
 *   2. CWS_CLIENT_ID + CWS_CLIENT_SECRET + CWS_REFRESH_TOKEN — OAuth 2.0 fallback
 *
 * Usage:
 *   node src/extension/scripts/cws-publish.mjs --publisher <id> --item <id> --action status
 *   node src/extension/scripts/cws-publish.mjs --publisher <id> --item <id> --zip dist/pkg.zip --action upload
 *   node src/extension/scripts/cws-publish.mjs --publisher <id> --item <id> --zip dist/pkg.zip --action publish
 */
import crypto from "node:crypto";
import fs from "node:fs";

const SCOPE = "https://www.googleapis.com/auth/chromewebstore";
const TOKEN_URL = "https://oauth2.googleapis.com/token";
const API_BASE = "https://chromewebstore.googleapis.com";

function parseArgs(argv) {
  const out = {};
  for (let i = 0; i < argv.length; i++) {
    const a = argv[i];
    if (a.startsWith("--")) {
      const key = a.slice(2);
      const next = argv[i + 1];
      if (!next || next.startsWith("--")) out[key] = true;
      else {
        out[key] = next;
        i++;
      }
    }
  }
  return out;
}

function base64url(input) {
  return Buffer.from(input)
    .toString("base64")
    .replace(/\+/g, "-")
    .replace(/\//g, "_")
    .replace(/=+$/g, "");
}

async function getAccessToken() {
  const saRaw = process.env.CWS_SERVICE_ACCOUNT_JSON?.trim();
  if (saRaw) {
    let sa;
    try {
      sa = JSON.parse(saRaw);
    } catch (err) {
      throw new Error(`CWS_SERVICE_ACCOUNT_JSON is not valid JSON: ${err.message}`);
    }
    if (!sa.client_email || !sa.private_key) {
      throw new Error("CWS_SERVICE_ACCOUNT_JSON must contain client_email and private_key.");
    }
    const now = Math.floor(Date.now() / 1000);
    const header = base64url(JSON.stringify({ alg: "RS256", typ: "JWT" }));
    const payload = base64url(JSON.stringify({
      iss: sa.client_email,
      scope: SCOPE,
      aud: sa.token_uri || TOKEN_URL,
      iat: now,
      exp: now + 3600,
    }));
    const unsigned = `${header}.${payload}`;
    const signer = crypto.createSign("RSA-SHA256");
    signer.update(unsigned);
    signer.end();
    const signature = base64url(signer.sign(sa.private_key));
    const assertion = `${unsigned}.${signature}`;

    const res = await fetch(sa.token_uri || TOKEN_URL, {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: new URLSearchParams({
        grant_type: "urn:ietf:params:oauth:grant-type:jwt-bearer",
        assertion,
      }),
    });
    const body = await res.json();
    if (!res.ok || !body.access_token) {
      throw new Error(`Service account token exchange failed (${res.status}): ${JSON.stringify(body)}`);
    }
    console.log(`Authenticated as service account: ${sa.client_email}`);
    return body.access_token;
  }

  const clientId = process.env.CWS_CLIENT_ID?.trim();
  const clientSecret = process.env.CWS_CLIENT_SECRET?.trim();
  const refreshToken = process.env.CWS_REFRESH_TOKEN?.trim();
  if (clientId && clientSecret && refreshToken) {
    const res = await fetch(TOKEN_URL, {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: new URLSearchParams({
        client_id: clientId,
        client_secret: clientSecret,
        refresh_token: refreshToken,
        grant_type: "refresh_token",
      }),
    });
    const body = await res.json();
    if (!res.ok || !body.access_token) {
      throw new Error(`OAuth token refresh failed (${res.status}): ${JSON.stringify(body)}`);
    }
    console.log("Authenticated via OAuth 2.0 refresh token.");
    return body.access_token;
  }

  throw new Error(
    "Missing Chrome Web Store credentials. Set CWS_SERVICE_ACCOUNT_JSON (recommended) or CWS_CLIENT_ID + CWS_CLIENT_SECRET + CWS_REFRESH_TOKEN.",
  );
}

async function apiJson(url, options) {
  const res = await fetch(url, options);
  const text = await res.text();
  let json;
  try {
    json = text ? JSON.parse(text) : {};
  } catch {
    json = { raw: text };
  }
  if (!res.ok) {
    throw new Error(`HTTP ${res.status} ${url}\n${JSON.stringify(json, null, 2)}`);
  }
  return json;
}

async function fetchStatus(token, publisherId, itemId) {
  const url = `${API_BASE}/v2/publishers/${publisherId}/items/${itemId}:fetchStatus`;
  return apiJson(url, {
    method: "GET",
    headers: { Authorization: `Bearer ${token}` },
  });
}

async function uploadZip(token, publisherId, itemId, zipPath) {
  if (!fs.existsSync(zipPath)) {
    throw new Error(`Package ZIP not found: ${zipPath}`);
  }
  const bytes = fs.readFileSync(zipPath);
  const url = `${API_BASE}/upload/v2/publishers/${publisherId}/items/${itemId}:upload`;
  console.log(`Uploading ${zipPath} (${(bytes.length / 1024).toFixed(1)} KB) -> item ${itemId}...`);
  const result = await apiJson(url, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/zip",
    },
    body: bytes,
  });
  console.log("Upload response:", JSON.stringify(result, null, 2));

  let state = result.uploadState;
  for (let attempt = 1; state === "IN_PROGRESS" && attempt <= 20; attempt++) {
    await new Promise((r) => setTimeout(r, 3000));
    const status = await fetchStatus(token, publisherId, itemId);
    state = status.lastAsyncUploadState || status.uploadState;
    console.log(`Poll #${attempt}: uploadState=${state}`);
    if (state === "FAILED") {
      throw new Error(`Async upload failed:\n${JSON.stringify(status, null, 2)}`);
    }
  }
  if (state && state !== "SUCCEEDED" && state !== "NOT_FOUND") {
    throw new Error(`Unexpected uploadState after upload: ${state}`);
  }
  return result;
}

async function publishItem(token, publisherId, itemId) {
  const url = `${API_BASE}/v2/publishers/${publisherId}/items/${itemId}:publish`;
  console.log(`Submitting item ${itemId} for review...`);
  const result = await apiJson(url, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
    },
    body: "{}",
  });
  console.log("Publish response:", JSON.stringify(result, null, 2));
  return result;
}

async function main() {
  const args = parseArgs(process.argv.slice(2));
  const publisherId = args.publisher || process.env.CWS_PUBLISHER_ID;
  const itemId = args.item || process.env.CWS_ITEM_ID;
  const action = args.action || "upload";
  const zipPath = args.zip;

  if (!publisherId || !itemId) {
    throw new Error("Required: --publisher <PUBLISHER_ID> --item <ITEM_ID>");
  }
  if (!["status", "upload", "publish"].includes(action)) {
    throw new Error(`Invalid --action "${action}". Use: status | upload | publish`);
  }

  const token = await getAccessToken();

  if (action === "status") {
    const status = await fetchStatus(token, publisherId, itemId);
    console.log(JSON.stringify(status, null, 2));
    return;
  }

  if (!zipPath) {
    throw new Error(`--zip <path> is required for --action ${action}`);
  }
  await uploadZip(token, publisherId, itemId, zipPath);

  if (action === "publish") {
    await publishItem(token, publisherId, itemId);
  } else {
    console.log("Draft uploaded (not submitted for review).");
  }

  const finalStatus = await fetchStatus(token, publisherId, itemId);
  console.log("Current store status:", JSON.stringify(finalStatus, null, 2));
}

main().catch((err) => {
  console.error(err.message || err);
  process.exit(1);
});

// Shared helpers for the Outlook (Microsoft Graph) connection.
// Delegated access only: each board user connects their own Microsoft account, and the
// board acts as that person. Refresh tokens are stored encrypted in public.ms_connections.
import { createClient } from "jsr:@supabase/supabase-js@2.45.4";

export const TENANT = "46d828d7-d38f-4762-af94-3e5b0759ef41";
export const CLIENT_ID = "50207bc9-ce28-4f06-8650-ce48a2cdc3c4";
export const REDIRECT = "https://aylvakfpaehrzgwgyoub.supabase.co/functions/v1/ms-callback";
export const SITE = "https://office-board-fawn.vercel.app";
export const ORIGINS = new Set([
  "https://office-board-fawn.vercel.app",
  "https://office-board-handy-equip.vercel.app",
  "https://office-board-git-main-handy-equip.vercel.app",
]);
export const SCOPES = "openid email offline_access User.Read Mail.ReadWrite Mail.Send Contacts.Read Calendars.ReadWrite";
export const GRAPH = "https://graph.microsoft.com/v1.0";

export const admin = createClient(
  Deno.env.get("SUPABASE_URL")!,
  Deno.env.get("SUPABASE_SERVICE_ROLE_KEY")!,
  { auth: { persistSession: false, autoRefreshToken: false } },
);

export class MsError extends Error {
  constructor(public code: string, message: string, public status = 400) { super(message); }
}

export function clientSecret(): string {
  const s = Deno.env.get("MS_CLIENT_SECRET") ?? "";
  if (!s) throw new MsError("not_configured", "Outlook isn't set up on the server yet (missing MS_CLIENT_SECRET).", 503);
  return s;
}

/* ---------- encryption of stored refresh tokens (AES-GCM, key derived from the client secret) ---------- */
async function aesKey(): Promise<CryptoKey> {
  const raw = await crypto.subtle.digest("SHA-256", new TextEncoder().encode(clientSecret() + "|office-board-token-key"));
  return crypto.subtle.importKey("raw", raw, "AES-GCM", false, ["encrypt", "decrypt"]);
}
const toB64 = (u: Uint8Array) => btoa(String.fromCharCode(...u));
const fromB64 = (s: string) => Uint8Array.from(atob(s), (c) => c.charCodeAt(0));

export async function encrypt(text: string): Promise<string> {
  const iv = crypto.getRandomValues(new Uint8Array(12));
  const ct = new Uint8Array(await crypto.subtle.encrypt({ name: "AES-GCM", iv }, await aesKey(), new TextEncoder().encode(text)));
  const out = new Uint8Array(iv.length + ct.length); out.set(iv); out.set(ct, iv.length);
  return toB64(out);
}
export async function decrypt(b64: string): Promise<string> {
  const all = fromB64(b64);
  const pt = await crypto.subtle.decrypt({ name: "AES-GCM", iv: all.slice(0, 12) }, await aesKey(), all.slice(12));
  return new TextDecoder().decode(pt);
}

/* ---------- Microsoft token endpoint ---------- */
export async function tokenRequest(params: Record<string, string>) {
  const body = new URLSearchParams({ client_id: CLIENT_ID, client_secret: clientSecret(), scope: SCOPES, ...params });
  const res = await fetch(`https://login.microsoftonline.com/${TENANT}/oauth2/v2.0/token`, {
    method: "POST", headers: { "Content-Type": "application/x-www-form-urlencoded" }, body,
  });
  const json = await res.json().catch(() => ({}));
  return { ok: res.ok, json };
}

const cache = new Map<string, { token: string; exp: number }>();

/** A fresh Graph access token for this board user, refreshing (and re-saving) as needed. */
export async function accessTokenFor(userId: string): Promise<string> {
  const hit = cache.get(userId);
  if (hit && hit.exp > Date.now() + 60_000) return hit.token;

  const { data: row } = await admin.from("ms_connections").select("token_enc").eq("user_id", userId).maybeSingle();
  if (!row) throw new MsError("server_not_connected", "Outlook isn't connected for you yet.", 409);

  let refresh: string;
  try { refresh = await decrypt(row.token_enc); }
  catch { await admin.from("ms_connections").delete().eq("user_id", userId); throw new MsError("needs_reauth", "Reconnect Outlook.", 409); }

  const { ok, json } = await tokenRequest({ grant_type: "refresh_token", refresh_token: refresh });
  if (!ok || !json.access_token) {
    if (json.error === "invalid_grant" || json.error === "interaction_required") {
      await admin.from("ms_connections").delete().eq("user_id", userId);
      throw new MsError("needs_reauth", "Reconnect Outlook.", 409);
    }
    throw new MsError("upstream_error", "Microsoft didn't answer. Try again.", 502);
  }
  if (json.refresh_token && json.refresh_token !== refresh) {
    await admin.from("ms_connections").update({ token_enc: await encrypt(json.refresh_token), updated_at: new Date().toISOString() }).eq("user_id", userId);
  }
  cache.set(userId, { token: json.access_token, exp: Date.now() + (Number(json.expires_in) || 3600) * 1000 });
  return json.access_token;
}

export async function graph(userId: string, path: string, init: RequestInit = {}) {
  const token = await accessTokenFor(userId);
  const res = await fetch(GRAPH + path, {
    ...init,
    headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json", ...(init.headers ?? {}) },
  });
  if (res.status === 401) { cache.delete(userId); throw new MsError("needs_reauth", "Reconnect Outlook.", 409); }
  if (res.status === 204 || res.status === 202) return {};
  const json = await res.json().catch(() => ({}));
  if (!res.ok) throw new MsError("tool_error", (json.error && json.error.message) || `Outlook error ${res.status}`, 502);
  return json;
}

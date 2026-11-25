// Outlook actions for signed-in Office Board users.
// POST { action: "status" | "start" | "disconnect" | "create_draft" | "send_draft", ... }
// with the user's Supabase session in the Authorization header.
import { createClient } from "jsr:@supabase/supabase-js@2.45.4";
import { admin, CLIENT_ID, graph, MsError, ORIGINS, REDIRECT, SCOPES, TENANT } from "../_shared/ms.ts";

function cors(origin: string | null): Record<string, string> {
  const h: Record<string, string> = {
    "Access-Control-Allow-Methods": "POST, OPTIONS",
    "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type",
    "Vary": "Origin",
  };
  if (origin && ORIGINS.has(origin)) h["Access-Control-Allow-Origin"] = origin;
  return h;
}
function reply(origin: string | null, status: number, body: unknown) {
  return new Response(JSON.stringify(body), { status, headers: { ...cors(origin), "Content-Type": "application/json" } });
}
const emails = (v: unknown) => (Array.isArray(v) ? v : String(v ?? "").split(/[;,\s]+/))
  .map((x) => String(x).trim()).filter((x) => /^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(x)).slice(0, 20);

Deno.serve(async (req) => {
  const origin = req.headers.get("origin");
  if (req.method === "OPTIONS") return new Response("ok", { headers: cors(origin) });
  if (req.method !== "POST" || !origin || !ORIGINS.has(origin)) return reply(origin, 403, { error: { code: "forbidden", message: "Not allowed" } });

  try {
    // Who is calling, and are they allowed on the board?
    const jwt = (req.headers.get("authorization") ?? "").replace(/^Bearer\s+/i, "");
    const { data: u } = await admin.auth.getUser(jwt);
    const user = u?.user;
    if (!user) throw new MsError("unauthenticated", "Sign in again.", 401);
    const asUser = createClient(Deno.env.get("SUPABASE_URL")!, Deno.env.get("SUPABASE_ANON_KEY")!, {
      global: { headers: { Authorization: `Bearer ${jwt}` } }, auth: { persistSession: false },
    });
    const { data: member } = await asUser.rpc("is_board_member");
    if (member !== true) throw new MsError("forbidden", "Not allowed.", 403);

    const b = await req.json().catch(() => ({}));
    switch (b.action) {
      case "status": {
        const { data } = await admin.from("ms_connections").select("ms_email").eq("user_id", user.id).maybeSingle();
        return reply(origin, 200, { connected: !!data, ms_email: data?.ms_email ?? null });
      }
      case "start": {
        const state = Array.from(crypto.getRandomValues(new Uint8Array(24)), (x) => x.toString(16).padStart(2, "0")).join("");
        await admin.from("ms_oauth_states").delete().lt("created_at", new Date(Date.now() - 3600_000).toISOString());
        await admin.from("ms_oauth_states").insert({ state, user_id: user.id });
        const q = new URLSearchParams({
          client_id: CLIENT_ID, response_type: "code", redirect_uri: REDIRECT, response_mode: "query",
          scope: SCOPES, state, login_hint: user.email ?? "", prompt: "select_account",
        });
        return reply(origin, 200, { url: `https://login.microsoftonline.com/${TENANT}/oauth2/v2.0/authorize?${q}` });
      }
      case "disconnect": {
        await admin.from("ms_connections").delete().eq("user_id", user.id);
        return reply(origin, 200, { connected: false });
      }
      case "create_draft": {
        const to = emails(b.to);
        if (!to.length) throw new MsError("invalid_argument", "Add the customer's email address first.");
        const msg = await graph(user.id, "/me/messages", {
          method: "POST",
          body: JSON.stringify({
            subject: String(b.subject ?? "").slice(0, 255),
            body: { contentType: b.bodyType === "text" ? "Text" : "HTML", content: String(b.body ?? "") },
            toRecipients: to.map((address) => ({ emailAddress: { address } })),
          }),
        });
        return reply(origin, 200, { id: msg.id, webLink: msg.webLink });
      }
      case "send_draft": {
        const id = String(b.messageId ?? "");
        if (!id) throw new MsError("invalid_argument", "No draft to send.");
        await graph(user.id, `/me/messages/${encodeURIComponent(id)}/send`, { method: "POST" });
        return reply(origin, 200, { sent: true });
      }
      default:
        throw new MsError("invalid_argument", "Unknown action.");
    }
  } catch (e) {
    const err = e instanceof MsError ? e : new MsError("upstream_error", "Something went wrong. Try again.", 500);
    if (!(e instanceof MsError)) console.error(e);
    return reply(origin, err.status, { error: { code: err.code, message: err.message } });
  }
});

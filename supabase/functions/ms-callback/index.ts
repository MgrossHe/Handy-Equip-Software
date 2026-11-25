// Microsoft sends the user back here after they approve access ("Connect Outlook").
// Swaps the one-time code for tokens, checks the Microsoft account matches the board
// user's email, stores the refresh token encrypted, and returns to the board.
import { admin, encrypt, GRAPH, MsError, REDIRECT, SITE, tokenRequest } from "../_shared/ms.ts";

const back = (params: Record<string, string>) =>
  new Response(null, { status: 302, headers: { Location: `${SITE}/?${new URLSearchParams(params)}` } });

Deno.serve(async (req) => {
  const url = new URL(req.url);
  const state = url.searchParams.get("state") ?? "";
  const code = url.searchParams.get("code") ?? "";
  if (url.searchParams.get("error")) return back({ outlook_error: "Microsoft sign-in was cancelled or refused." });
  if (!state || !code) return back({ outlook_error: "Missing sign-in details. Try Connect Outlook again." });

  try {
    // The state must be one we issued in the last 15 minutes; use it once.
    const { data: st } = await admin.from("ms_oauth_states").delete().eq("state", state).select("user_id, created_at").maybeSingle();
    if (!st || Date.parse(st.created_at) < Date.now() - 15 * 60_000) {
      return back({ outlook_error: "That sign-in expired. Try Connect Outlook again." });
    }

    const { ok, json } = await tokenRequest({ grant_type: "authorization_code", code, redirect_uri: REDIRECT });
    if (!ok || !json.refresh_token || !json.access_token) {
      console.error("token exchange failed", json.error, json.error_description);
      return back({ outlook_error: "Microsoft didn't accept the sign-in. Try again." });
    }

    const meRes = await fetch(`${GRAPH}/me?$select=mail,userPrincipalName,proxyAddresses,otherMails`, {
      headers: { Authorization: `Bearer ${json.access_token}` },
    });
    const me = await meRes.json().catch(() => ({}));
    if (!meRes.ok) console.error("graph /me failed", meRes.status, me?.error?.code, me?.error?.message);
    // Every address this Microsoft account answers to: primary mail, sign-in name, and aliases.
    const addresses = new Set<string>(
      [me.mail, me.userPrincipalName, ...(me.otherMails ?? []),
        ...(me.proxyAddresses ?? []).map((p: string) => String(p).replace(/^smtp:/i, ""))]
        .filter(Boolean).map((x: string) => String(x).toLowerCase()),
    );
    const msEmail = String(me.mail || me.userPrincipalName || "").toLowerCase();

    const { data: u } = await admin.auth.admin.getUserById(st.user_id);
    const boardEmail = String(u?.user?.email ?? "").toLowerCase();
    if (!boardEmail || !addresses.has(boardEmail)) {
      console.error("mailbox mismatch", { boardEmail, msEmail, addresses: [...addresses] });
      return back({ outlook_error: `Sign in to Microsoft as ${boardEmail} (you used ${msEmail || "another account"}).` });
    }

    const { error: saveErr } = await admin.from("ms_connections").upsert({
      user_id: st.user_id, ms_email: boardEmail, token_enc: await encrypt(json.refresh_token),
      scopes: String(json.scope ?? ""), updated_at: new Date().toISOString(),
    });
    if (saveErr) {
      console.error("saving connection failed", saveErr.message);
      return back({ outlook_error: "Couldn't save the Outlook connection. Try again." });
    }
    console.log("outlook connected", boardEmail);
    return back({ outlook: "connected" });
  } catch (e) {
    if (e instanceof MsError && e.code === "not_configured") return back({ outlook_error: e.message });
    console.error(e);
    return back({ outlook_error: "Couldn't connect Outlook. Try again." });
  }
});

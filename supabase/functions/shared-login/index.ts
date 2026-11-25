// Office Board shared-password sign-in.
// Anyone with an @handyequip.com email + the shared password gets a session.
// Only a salted PBKDF2 hash of the password is stored (public.server_settings,
// readable only by this function). Repeated wrong guesses are locked out, and
// every attempt is recorded in public.login_attempts.
import { createClient } from "jsr:@supabase/supabase-js@2.45.4";

const DOMAIN = "handyequip.com";
const ORIGINS = new Set([
  "https://office-board-fawn.vercel.app",
  "https://office-board-handy-equip.vercel.app",
  "https://office-board-git-main-handy-equip.vercel.app",
]);
const WINDOW_MIN = 15;
const MAX_FAILS_PER_IP = 10;
const MAX_FAILS_PER_EMAIL = 5;

const admin = createClient(
  Deno.env.get("SUPABASE_URL")!,
  Deno.env.get("SUPABASE_SERVICE_ROLE_KEY")!,
  { auth: { persistSession: false, autoRefreshToken: false } },
);

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
  return new Response(JSON.stringify(body), {
    status,
    headers: { ...cors(origin), "Content-Type": "application/json" },
  });
}

function b64(s: string): Uint8Array {
  return Uint8Array.from(atob(s), (c) => c.charCodeAt(0));
}

// stored format: pbkdf2-sha256$<iterations>$<salt b64>$<hash b64>
async function passwordMatches(given: string, stored: string): Promise<boolean> {
  const [alg, iter, salt, hash] = stored.split("$");
  if (alg !== "pbkdf2-sha256" || !iter || !salt || !hash) return false;
  const key = await crypto.subtle.importKey("raw", new TextEncoder().encode(given), "PBKDF2", false, ["deriveBits"]);
  const bits = new Uint8Array(await crypto.subtle.deriveBits(
    { name: "PBKDF2", hash: "SHA-256", salt: b64(salt), iterations: Number(iter) }, key, 256));
  const want = b64(hash);
  if (bits.length !== want.length) return false;
  let diff = 0;
  for (let i = 0; i < bits.length; i++) diff |= bits[i] ^ want[i];
  return diff === 0;
}

async function recentFails(column: "ip" | "email", value: string): Promise<number> {
  const since = new Date(Date.now() - WINDOW_MIN * 60_000).toISOString();
  const { count } = await admin.from("login_attempts").select("id", { count: "exact", head: true })
    .eq(column, value).eq("ok", false).gte("at", since);
  return count ?? 0;
}

Deno.serve(async (req) => {
  const origin = req.headers.get("origin");
  if (req.method === "OPTIONS") return new Response("ok", { headers: cors(origin) });
  if (req.method !== "POST") return reply(origin, 405, { error: "Method not allowed" });
  if (!origin || !ORIGINS.has(origin)) return reply(origin, 403, { error: "Not allowed" });

  const { data: setting } = await admin.from("server_settings").select("value").eq("name", "shared_password_hash").maybeSingle();
  if (!setting?.value) return reply(origin, 503, { error: "Sign-in isn't set up yet." });

  let email = "", password = "";
  try {
    const b = await req.json();
    email = String(b.email ?? "").trim().toLowerCase();
    password = String(b.password ?? "");
  } catch { return reply(origin, 400, { error: "Bad request" }); }

  if (!/^[^@\s]+@[^@\s]+$/.test(email) || email.split("@")[1] !== DOMAIN || email.length > 254) {
    return reply(origin, 400, { error: "Use your @" + DOMAIN + " email." });
  }
  const ip = (req.headers.get("x-forwarded-for") ?? "").split(",")[0].trim() || "unknown";

  if (await recentFails("ip", ip) >= MAX_FAILS_PER_IP || await recentFails("email", email) >= MAX_FAILS_PER_EMAIL) {
    return reply(origin, 429, { error: "Too many wrong tries. Wait 15 minutes, then try again." });
  }

  if (!(await passwordMatches(password, setting.value))) {
    await admin.from("login_attempts").insert({ email, ip, ok: false });
    await new Promise((r) => setTimeout(r, 600));
    return reply(origin, 401, { error: "Wrong email or password." });
  }

  // Make sure the account exists (ignore "already registered"), then issue a one-time sign-in token.
  const created = await admin.auth.admin.createUser({ email, email_confirm: true });
  if (created.error && !/already|registered|exists/i.test(created.error.message)) {
    return reply(origin, 500, { error: "Couldn't sign in. Try again." });
  }
  const link = await admin.auth.admin.generateLink({ type: "magiclink", email });
  const tokenHash = link.data?.properties?.hashed_token;
  if (link.error || !tokenHash) return reply(origin, 500, { error: "Couldn't sign in. Try again." });

  await admin.from("login_attempts").insert({ email, ip, ok: true });
  return reply(origin, 200, { token_hash: tokenHash });
});

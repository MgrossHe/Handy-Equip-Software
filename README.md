# Office Board

Handy Equip's internal office board: AR, customers (sales/service requests), purchasing and email follow-up.

- **Hosting:** Vercel (static site, no build step).
- **Data and sign-in:** Supabase project `office-board`. Staff sign in with their `@handyequip.com`
  email and the shared office password. The password is checked by the Supabase function
  `shared-login` (source in `supabase/functions/shared-login/`); only a salted hash is stored, in
  `server_settings`. Wrong guesses are locked out for 15 minutes and every attempt is logged in
  `login_attempts`. Only emails in `allowed_emails` or a domain in `allowed_domains` can see or
  change anything; every table uses row-level security.
- **No customer data lives in this repository.** The Squid snapshot, email summary, promises and
  notes are stored in Supabase and loaded only after sign-in.

## Files

| File | What it is |
|---|---|
| `index.html` | Page layout, styles and the sign-in screen |
| `app.js` | Sign-in, permission check, and the adapter that stores notes/requests in Supabase |
| `board.js` | The board itself (tabs, tables, Excel downloads) |
| `vercel.json` | Security headers; keeps the site out of search engines |
| `supabase/functions/shared-login` | Shared-password sign-in |
| `supabase/functions/ms-api`, `ms-callback`, `_shared/ms.ts` | Outlook connection (Microsoft Graph, delegated): Connect Outlook, save draft, send |

## Access

Everyone with an `@handyequip.com` email can sign in with the shared password. To change the
password, ask Claude to replace the hash in `server_settings` (row `shared_password_hash`).

See who signed in (Supabase SQL editor):

```sql
select at, email, ok from public.login_attempts order by at desc limit 50;
```

## Outlook

Each person clicks **Connect Outlook** once and approves with their own Microsoft login
(Entra app "Office Board", delegated permissions: User.Read, offline_access, Mail.ReadWrite,
Mail.Send, Contacts.Read, Calendars.ReadWrite). The Microsoft account must match the board
email. Refresh tokens are stored encrypted in `ms_connections` (server-only). The client
secret lives only in the Supabase Edge Function secret `MS_CLIENT_SECRET`; it expires after
24 months and must then be renewed in Azure and replaced in Supabase.

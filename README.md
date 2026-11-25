# Office Board

Handy Equip's internal office board: AR, customers (sales/service requests), purchasing and email follow-up.

- **Hosting:** Vercel (static site, no build step).
- **Data and sign-in:** Supabase project `office-board`. Sign-in is by emailed link. Only emails in
  `allowed_emails` (or a domain in `allowed_domains`) can see or change anything; every table uses
  row-level security.
- **No customer data lives in this repository.** The Squid snapshot, email summary, promises and
  notes are stored in Supabase and loaded only after sign-in.

## Files

| File | What it is |
|---|---|
| `index.html` | Page layout, styles and the sign-in screen |
| `app.js` | Sign-in, permission check, and the adapter that stores notes/requests in Supabase |
| `board.js` | The board itself (tabs, tables, Excel downloads) |
| `vercel.json` | Security headers; keeps the site out of search engines |

## Giving someone access

Run in the Supabase SQL editor:

```sql
insert into public.allowed_emails (email) values ('name@handyequip.com');
-- or everyone at the company:
insert into public.allowed_domains (domain) values ('handyequip.com');
```

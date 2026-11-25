# Automatic email refresh

A scheduled Claude task runs on workdays (about every 2 hours, 8am–6pm New York time).
It reads Mendel's Outlook through the Microsoft 365 connector and updates Supabase:

- `board_files` row `email`: rack offers, people waiting for a reply, recent drafts and
  payments, with `asOf` (date) and `checkedAt` (time of the run).
- `docs` (collection `callnotes`): customer replies about invoices, as notes with
  `source: "email"`, id `em-` + first 16 hex chars of md5(internet message id).
- `docs` (collection `requests`): clear new sales / service / purchasing requests, id
  `em-rq-` + first 14 hex chars of md5(internet message id).

New notes and requests are only added, never edited or deleted. The task never sends,
drafts, moves, deletes or marks email, and never changes Squid.

The board shows a "New email data is available" notice when the refresh lands while it is open.

# Everest Plunge Stock Sheet Agent

Automates writes to Everest Plunge's stock & fulfillment Google Sheet (the
per-batch client allocation tracker — batch tabs like "BATCH 12", each with
a product anchor row and stacked client rows above a "↳ Remaining" row).

## What this agent actually does right now

This is the **mechanism**, not the full automation. It provides:

- Read the sheet's tab list and any tab's raw contents.
- Insert a new client row into a named batch tab, in the correct physical
  position (directly above the "↳ Remaining" row) — the same spot a person
  inserts one manually today.

**What it deliberately does NOT do yet**: decide *when* to call that
insert, or *which* batch/product a closed deal should go into. That
requires cross-referencing Pipely, Shopify, Xero, and Mainfreight to know
what's actually in stock and where — and that matching logic isn't defined
yet (see project memory). Wiring an automatic trigger is a separate, later
step once that's worked out. For now, `/admin/add-client` is a manual/
scripted call, not something firing on its own.

## Automation Log — the working "log a sold deal / mark order sent" tool

Added 2026-08-31, this is what actually backs the ops-console app. Unlike
`/admin/add-client` above, this does NOT touch the real batch tabs — it
writes into a tab this agent creates and owns itself ("🤖 Automation Log",
auto-created on first use), specifically to avoid guessing at the
unconfirmed batch-tab column layout. Safe to use today.

- `POST /admin/log-sold-deal` — logs a new sold deal, runs a stock
  availability check against Stock Overview (informational, not a hard
  gate), generates an order ID.
- `POST /admin/mark-order-sent` — fills in courier/tracking/sent-date for
  an existing logged order, found by order ID.
- `GET /admin/automation-log` — lists every logged entry.

Once the real batch tab layout is confirmed, decide then whether this tab
becomes the permanent record or a sync step copies entries into the real
batch tabs — deliberately not decided yet.

## Not yet confirmed — verify against the REAL sheet before trusting this

This was written from the sheet's documented structure (per the "How To
Use" tab and prior notes), not from inspecting the live sheet directly —
I don't have access to it yet. Before relying on this:

- **The anchor row's assigned/remaining counts**: if those cells are live
  SUM formulas over a row range that spans the insertion point, they'll
  update automatically the moment a row is inserted inside that range —
  in which case this agent's job is done as soon as the row lands. If
  they're plain typed numbers (matching the sheet's own "someone manually
  updates the assigned count" instructions), inserting the row does NOT
  fix those counts — that'd need a second step this agent doesn't do yet.
  Open the real sheet and check before assuming either way.
- **Column order for a client row**: memory describes the columns as
  contact, delivery address, courier, tracking #, delivery date, install
  Y/N + installer + install date, payment status, invoice #, amount — but
  the exact order/count of columns in the real sheet hasn't been confirmed
  against this code. Get the real header row from `/admin/read-tab` first
  and build `rowValues` to match it exactly.
- The "↳ Remaining" row is found by a case-insensitive text search for
  "remaining" in column A, not an exact match on the arrow glyph — more
  robust to retyping, but confirm it doesn't also match some other label
  in a real tab (e.g. a note mentioning "remaining" elsewhere).

## Admin endpoints (require `x-api-key`)

- `GET /admin/tabs` — list every tab name in the sheet.
- `GET /admin/read-tab?tabName=BATCH%2012` — raw rows for one tab.
- `POST /admin/add-client` — body `{ "tabName": "BATCH 12", "rowValues": [...] }`,
  inserts a new client row above that tab's "↳ Remaining" row.
- `POST /admin/log-sold-deal` — body `{ source, customerName, email, sku, quantity, deliveryAddress, dealValue, depositStatus, notes }`, only `customerName` required.
- `POST /admin/mark-order-sent` — body `{ orderId, courier, trackingNumber, sentDate }`.
- `GET /admin/automation-log` — lists all logged entries.

## Setup checklist

1. Confirm the Operations sheet exists as an actual Google Sheet (not just
   the `.xlsx` file) and share edit access with whichever Google account
   you'll authorize this agent with.
2. Create a Google Cloud OAuth client (Sheets API enabled), fill in
   `GOOGLE_CLIENT_ID`/`SECRET`/`REDIRECT_URI` and `SHEET_ID` in Railway.
3. Deploy, visit `/oauth/start` once, log in as whichever Google account
   has edit access.
4. Call `GET /admin/tabs` and `GET /admin/read-tab` first to confirm real
   tab names and column layout before trusting `/admin/add-client` against
   a real batch.

## Deployment

Same pattern as the other agents: own GitHub repo, Railway service in the
Everest Plunge Railway project, env vars pasted into Railway's Variables
tab. No persistent volume needed (Google's refresh token doesn't rotate).

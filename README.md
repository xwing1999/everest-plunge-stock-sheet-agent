# Everest Plunge Stock Sheet Agent

Automates writes to Everest Plunge's stock & fulfillment Google Sheet (the
per-batch client allocation tracker — batch tabs like "BATCH 12", each with
a product anchor row and stacked client rows above a "↳ Remaining" row).

**2026-08-31 — audit fixes**: `recordNewOrderAgainstBatch` is now locked
per-SKU (was a non-atomic read-modify-write that could silently lose an
increment under concurrent calls), `/admin/set-final-payment-status`
refuses `status:"Paid"` (that value is only reachable via
`/admin/mark-final-payment-received`, closing a gap where any caller with
the API key could satisfy the release gate with no real payment behind
it), and `columnIndexToLetter` now throws on a missing/renamed column
instead of silently producing a malformed write range.

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
- `GET /admin/automation-log` — lists every logged entry, with a computed
  `Status` (Ready to organise / Awaiting stock (on water) / Awaiting stock
  (not yet ordered) / Sent) derived from Allocation + Order Sent Date.

Once the real batch tab layout is confirmed, decide then whether this tab
becomes the permanent record or a sync step copies entries into the real
batch tabs — deliberately not decided yet.

## Stock allocation model (added 2026-08-31)

Xavier's 4-bucket model for every sold deal: **On Shore** (ships now),
**On Water** (claimed against a named incoming batch), **Next Custom
Order** (nothing covers it yet — goes in the next China order). A person
picks the bucket when logging the deal (`allocation` field on
`/admin/log-sold-deal`), informed by the live balance already shown in the
ops console — deliberately NOT an automatic waterfall algorithm, which
risks getting concurrent claims against limited stock wrong in a way a
human glancing at the real number won't.

- `POST /admin/mark-batch-arrived` — body `{ batchReference }`. Flips
  every deal allocated to that batch from "On Water" to "On Shore" in one
  move — Xavier: "everyone already allocated to that stock needs to be
  organised." Does **not** touch Stock Overview's physical In Stock count
  or the current-batch column rollover ("Batch 10" → "Batch 11" in that
  tab's own headers) — that needs a real arrival manifest and a confirmed
  rollover process, not a guess. Stays manual in Stock Overview for now.
- `GET /admin/products-to-order` — bucket 4, computed rather than
  hand-maintained: every un-ordered "Next Custom Order" deal, grouped by
  SKU with summed quantities. The live shopping list for the next order.
- `POST /admin/mark-order-placed` — body `{ orderId }`. Marks one deal's
  stock as ordered, dropping it out of the products-to-order list.

## Batch ETA / Countdown (added 2026-09-01)

Xavier's countdown-to-shore model, replacing an earlier "ship date lives
somewhere unconfirmed" blocker: one arrival ETA per **batch/shipment**
(not per order) — matches how the real spreadsheet already tracks it
(each batch tab has one shared "ETA:" field in its info strip). Ops sets
it once per batch; every order allocated to that batch inherits the same
countdown via its existing "Batch Reference" field. `pipely-xero-agent`
polls this to fire the final 50% invoice automatically once a batch is
within its lead-time window — see that agent's README.

- `POST /admin/set-batch-eta` — body `{ batchReference, date }` (ISO
  `YYYY-MM-DD`). Not validated beyond non-empty — this agent doesn't know
  what "valid" means for a date any better than the caller does.
- `GET /admin/batch-etas` — every batch's current ETA + when it was last
  set.
- `GET /admin/automation-log` now includes a computed `Ship ETA` field per
  entry (not a real sheet column — resolved from the batch ETA store by
  "Batch Reference"), so consumers don't need to do their own join.

Deliberately NOT read from/written into the real batch tabs' own "ETA:"
cell — this agent has no parsing logic for those tabs' real layout yet
(see "Confirmed against the real spreadsheet" below), and guessing at
cell positions risks corrupting a manually-maintained sheet. Kept as this
agent's own small persisted store instead.

## Confirmed against the real spreadsheet (2026-09-01)

Inspected `Everest Plunge_Operations_v2.xlsx` directly (16 tabs). Real
findings, correcting/confirming earlier notes:

- **Real SKU catalog** (Stock Overview tab): SKU-001/002 Obsidian Sauna
  2/4 Person ($5,490/$7,990), SKU-003/004 Onyx Sauna 2/4 Person
  ($5,490/$7,990), SKU-005/006 Redlight Sauna 2/4 Person
  ($5,990/$7,990), SKU-007 Sienna Sauna 4 Person Black ($10,990,
  "Premium"). Two unlisted-SKU lines also appear in Master Order List:
  Traditional Sauna 5-6/7-8 Person.
- **Vulcan Sauna (White Aluminium) is real and active** — appears in real
  BATCH 12 invoice rows (Noel McGirr, $7,675 deposit, paid). Earlier
  memory guessed it might be discontinued since it didn't appear in an
  older CSV snapshot — that guess was wrong, don't repeat it.
- **Real batch tab structure** (BATCH 12 inspected directly): simpler
  than the "How To Use" tab's anchor-row/remaining-row description
  suggests in practice — real rows are closer to a flat invoice ledger
  (Invoice No. | Invoice Date | Customer Name | Product | Batch label |
  Invoice Amount | Invoice Type | Status), with extra detail (interior
  notes, freight cost) stacked as extra rows under just the Product
  column. The anchor/remaining structure may still exist elsewhere in the
  tab (not fully mapped) — don't assume the simpler ledger view is the
  complete picture.
- **Each batch tab has its own ETA field** in row 3 ("Shipment: BATCH-12
  | ETA: TBC | Route: TBC | Status: Pending") — this is what the Batch
  ETA feature above is modeled on, kept as a separate store rather than
  read/written directly (see above for why).
- **`STOCK_OVERVIEW_TAB` env var default ("📦 Stock Overview") does not
  exactly match either real tab name** in this workbook ("📦 Stock
  Overview_13.08.2026" and "📦 Stock Overview_XW") — verify the actual
  live Google Sheet's tab name before deploying; this may need updating.

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
- `POST /admin/log-sold-deal` — body `{ source, customerName, email, sku, quantity, deliveryAddress, dealValue, depositStatus, allocation, batchReference, expectedDate, notes }`, only `customerName` required. `allocation` should be one of `On Shore` / `On Water` / `Next Custom Order`; `batchReference` only matters when allocation is `On Water`.
- `POST /admin/mark-order-sent` — body `{ orderId, courier, trackingNumber, sentDate }`.
- `GET /admin/automation-log` — lists all logged entries with computed `Status`.
- `POST /admin/mark-batch-arrived` — body `{ batchReference }`.
- `GET /admin/products-to-order` — bucket-4 shopping list, grouped by SKU.
- `POST /admin/mark-order-placed` — body `{ orderId }`.

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

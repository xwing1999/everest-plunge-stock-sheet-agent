import 'dotenv/config';
import express from 'express';
import { google } from 'googleapis';
import fs from 'fs';
import path from 'path';

const app = express();
app.use(express.json());

// ---------------------------------------------------------------------------
// KEYED LOCK — added 2026-08-31 after an audit found recordNewOrderAgainstBatch
// does a non-atomic read-modify-write (read New Orders, add quantity, write
// back an absolute number) — two concurrent calls for the SAME SKU could
// both read the same stale value and one increment would silently overwrite
// the other, losing a real stock update. Serializes calls sharing a key
// (here, per-SKU) within this process — different SKUs still run in
// parallel since they touch different cells.
// ---------------------------------------------------------------------------
const locks = new Map();
function withLock(key, fn) {
  const prevTail = locks.get(key) || Promise.resolve();
  const run = prevTail.then(fn, fn);
  locks.set(key, run.then(() => {}, () => {}));
  return run;
}

// ---------------------------------------------------------------------------
// AUTH — shared-secret pattern, same as every other agent in this project.
// /oauth/* stays exempt, same reasoning as the other Google/Xero agents.
// ---------------------------------------------------------------------------
app.use((req, res, next) => {
  if (req.path === '/health' || req.path.startsWith('/oauth/')) return next();
  const provided = req.header('x-api-key');
  if (!process.env.API_KEY || provided !== process.env.API_KEY) {
    return res.status(401).json({ error: 'Unauthorized' });
  }
  next();
});

// ---------------------------------------------------------------------------
// GOOGLE OAUTH — same pattern as franchisor-revenue-agent. A Google refresh
// token obtained with access_type=offline does NOT rotate on every use, so
// no disk-persistence dance is needed — set it once in Railway's
// GOOGLE_REFRESH_TOKEN variable via /oauth/start and it keeps working.
// ---------------------------------------------------------------------------
const oauth2Client = new google.auth.OAuth2(
  process.env.GOOGLE_CLIENT_ID,
  process.env.GOOGLE_CLIENT_SECRET,
  process.env.GOOGLE_REDIRECT_URI
);
if (process.env.GOOGLE_REFRESH_TOKEN) {
  oauth2Client.setCredentials({ refresh_token: process.env.GOOGLE_REFRESH_TOKEN });
}
const sheets = google.sheets({ version: 'v4', auth: oauth2Client });

app.get('/oauth/start', (_req, res) => {
  const url = oauth2Client.generateAuthUrl({
    access_type: 'offline',
    prompt: 'consent',
    scope: ['https://www.googleapis.com/auth/spreadsheets']
  });
  res.redirect(url);
});

app.get('/oauth/callback', async (req, res) => {
  const { code, error } = req.query;
  if (error) return res.status(400).send(`Google returned an error: ${error}`);
  if (!code) return res.status(400).send('Missing code parameter.');

  try {
    const { tokens } = await oauth2Client.getToken(code);
    oauth2Client.setCredentials(tokens);
    res.send(`
      <h2>Google Sheets connected</h2>
      <p>Copy this into the <code>GOOGLE_REFRESH_TOKEN</code> Railway variable
      (this agent has no persistent storage of its own, so it must come from
      an env var on every restart):</p>
      <pre>${tokens.refresh_token ?? '(no refresh_token returned — you likely already granted consent once before; revoke access at myaccount.google.com/permissions for this app and try /oauth/start again)'}</pre>
      <p>Make sure this Google account actually has edit access to the
      Everest Plunge Operations sheet before continuing.</p>
      <p>You can close this tab.</p>
    `);
  } catch (err) {
    res.status(500).send(`Token exchange failed: ${err.message}`);
  }
});

// ---------------------------------------------------------------------------
// SHEET PRIMITIVES
// ---------------------------------------------------------------------------
async function getSheetMeta() {
  const meta = await sheets.spreadsheets.get({ spreadsheetId: process.env.SHEET_ID });
  return meta.data.sheets ?? [];
}

async function getSheetIdByTitle(tabName) {
  const sheetsMeta = await getSheetMeta();
  const found = sheetsMeta.find((s) => s.properties.title === tabName);
  if (!found) throw new Error(`No tab named "${tabName}" found in the sheet.`);
  return found.properties.sheetId;
}

// Finds the "↳ Remaining" row in a batch tab by scanning column A for the
// word "remaining" (case-insensitive substring match, not an exact-character
// match on the arrow glyph — more robust if the label is ever retyped
// slightly differently). Returns a 0-based row index (matches the Sheets
// API's row-index convention used by insertDimension below).
async function findRemainingRowIndex(tabName) {
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: process.env.SHEET_ID,
    range: `'${tabName}'!A:A`
  });
  const rows = res.data.values ?? [];
  const idx = rows.findIndex((row) => (row[0] ?? '').toString().toLowerCase().includes('remaining'));
  if (idx === -1) throw new Error(`Could not find a "↳ Remaining" row in tab "${tabName}" — is this actually a batch tab?`);
  return idx;
}

// ---------------------------------------------------------------------------
// INSERT A CLIENT ROW — the core write primitive. Inserts a new blank row
// directly above the "↳ Remaining" row and fills it with the given values,
// same physical position a person would insert one manually. Does NOT
// recompute the anchor row's assigned/remaining counts itself — if those
// cells are live SUM formulas over a range that spans this position, they
// pick up the new row automatically once inserted; if they're not (i.e. the
// sheet currently expects a person to update them by hand, per how the
// "How To Use" tab describes the current manual process), this call alone
// won't fix that. CONFIRM which is actually true for the real sheet before
// relying on this unattended — this was written from the sheet's documented
// structure, not by inspecting the live formulas.
// ---------------------------------------------------------------------------
async function insertClientRow(tabName, rowValues) {
  const sheetId = await getSheetIdByTitle(tabName);
  const remainingIdx = await findRemainingRowIndex(tabName);

  await sheets.spreadsheets.batchUpdate({
    spreadsheetId: process.env.SHEET_ID,
    requestBody: {
      requests: [{
        insertDimension: {
          range: { sheetId, dimension: 'ROWS', startIndex: remainingIdx, endIndex: remainingIdx + 1 },
          inheritFromBefore: true
        }
      }]
    }
  });

  const rowNumber = remainingIdx + 1; // 0-based index -> 1-based row number, and the new blank row now sits exactly here
  await sheets.spreadsheets.values.update({
    spreadsheetId: process.env.SHEET_ID,
    range: `'${tabName}'!A${rowNumber}`,
    valueInputOption: 'USER_ENTERED',
    requestBody: { values: [rowValues] }
  });

  return { tabName, rowNumber, rowValues };
}

// ---------------------------------------------------------------------------
// STOCK OVERVIEW — this part IS confirmed against a real export (the
// "📦 Stock Overview" tab, CSV pulled 2026-08-13), unlike the batch-tab
// logic above which is still written from the sheet's documented structure
// only. Confirmed real layout:
//   SKU | Product Name | Model / Size | In Stock | Batch N | Available |
//   New Orders (Batch N) | Balance | Batch N+1 | Batch N+2 | New Order |
//   Retail Price (NZD) | Warehouse | Notes
// The batch number is baked into the column HEADER TEXT itself ("Batch 10",
// "New Orders (Batch 10)") and shifts every time a new batch cycle starts —
// so columns are found by pattern match on each read, never by a fixed
// name or index.
//
// CORRECTED 2026-09-01 — Xavier: "the overview just reflects the different
// batches etc along the bottom of the sheet, all the sub headings just
// need to combine." Confirmed against real live data: the sheet has grown
// MULTIPLE "Batch N" reservation columns over time (Batch 10, 11, 12 all
// present simultaneously) plus more than one "New Order(s)" column, but
// the tab's own "Available"/"Balance" formula cells were only ever written
// to account for the FIRST batch column — they don't get updated as new
// batch columns are added. Real example proving this (SKU-002): In Stock
// 15, Batch 10 = 5, Batch 11 = 4 — actually available is 15-5-4=6, but the
// sheet's own "Available" cell still shows 10 (just 15-5, ignoring Batch
// 11 entirely). Trusting the sheet's own Available/Balance cells directly
// (the previous behavior of this function) meant every SKU with stock
// reserved against more than one open batch showed as having MORE stock
// to sell than actually exists.
//
// Fixed by finding EVERY column matching the batch/new-order patterns (not
// just the first) and summing them — Available and Balance are now
// computed here, not read from the sheet's stale single-batch formula
// cells. The sheet's own Available/Balance are still read back as
// `sheetAvailable`/`sheetBalance` for comparison/debugging, but nothing
// in this codebase should trust those two over the computed values.
// ---------------------------------------------------------------------------
const STOCK_OVERVIEW_TAB = process.env.STOCK_OVERVIEW_TAB || '📦 Stock Overview';

async function getStockOverview() {
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: process.env.SHEET_ID,
    range: STOCK_OVERVIEW_TAB
  });
  const rows = res.data.values ?? [];

  const headerRowIdx = rows.findIndex((row) => (row[0] ?? '').toString().trim().toUpperCase() === 'SKU');
  if (headerRowIdx === -1) throw new Error(`Could not find the header row (a cell reading "SKU") in "${STOCK_OVERVIEW_TAB}".`);
  const headers = rows[headerRowIdx].map((h) => (h ?? '').toString().trim());

  const singleCol = {
    sku: headers.findIndex((h) => h.toUpperCase() === 'SKU'),
    productName: headers.findIndex((h) => h.toUpperCase() === 'PRODUCT NAME'),
    modelSize: headers.findIndex((h) => h.toUpperCase().startsWith('MODEL')),
    inStock: headers.findIndex((h) => h.toUpperCase() === 'IN STOCK'),
    // Kept only for the sheetAvailable/sheetBalance comparison fields below
    // — not used for the real computed available/balance.
    available: headers.findIndex((h) => h.toUpperCase() === 'AVAILABLE'),
    balance: headers.findIndex((h) => h.toUpperCase() === 'BALANCE')
  };
  const missing = Object.entries(singleCol).filter(([, idx]) => idx === -1).map(([name]) => name);
  if (missing.length) throw new Error(`Stock Overview header row is missing expected column(s): ${missing.join(', ')}. Sheet structure may have changed — check "${STOCK_OVERVIEW_TAB}" by eye before trusting this.`);

  // Every batch-reservation and new-order column, not just the first —
  // this is the actual fix. "New Order" (no batch number, no parentheses)
  // is real data-entry inconsistency in the live sheet, not a different
  // concept — matched loosely on purpose.
  const batchCols = headers.map((h, i) => ({ h, i })).filter(({ h }) => /^batch\s*\d+$/i.test(h)).map(({ i }) => i);
  const newOrderCols = headers.map((h, i) => ({ h, i })).filter(({ h }) => /^new orders?\b/i.test(h)).map(({ i }) => i);
  if (!batchCols.length) throw new Error(`Stock Overview header row has no "Batch N" column(s) — sheet structure may have changed.`);

  const batchLabel = batchCols.map((i) => headers[i]).join(' + ');
  const sumCols = (row, cols) => cols.reduce((sum, i) => sum + Number(row[i] || 0), 0);

  const products = [];
  for (let r = headerRowIdx + 1; r < rows.length; r++) {
    const row = rows[r];
    const sku = (row[singleCol.sku] ?? '').toString().trim();
    if (!sku || sku.toUpperCase() === 'TOTALS') break;

    const inStock = Number(row[singleCol.inStock] || 0);
    const reserved = sumCols(row, batchCols);
    const newOrders = sumCols(row, newOrderCols);
    const available = inStock - reserved;
    const balance = available - newOrders;

    products.push({
      sku,
      productName: row[singleCol.productName] ?? '',
      modelSize: row[singleCol.modelSize] ?? '',
      inStock,
      reserved,
      available,
      newOrders,
      balance,
      sheetAvailable: Number(row[singleCol.available] || 0), // for comparison only, see comment above
      sheetBalance: Number(row[singleCol.balance] || 0),     // for comparison only, see comment above
      rowNumber: r + 1 // 1-based sheet row, for writing back to this exact row later
    });
  }

  // recordNewOrderAgainstBatch writes into the FIRST new-order column found
  // — the sheet has no way to know which specific batch a brand-new order
  // should count against, so this matches the previous (pre-fix) behavior
  // for that one write path rather than guessing which of several columns
  // is "correct" to increment.
  const col = { ...singleCol, reserved: batchCols[0], newOrders: newOrderCols[0] };

  return { batchLabel, products, columns: col };
}

async function checkStockAvailability(sku, quantity) {
  const { products, batchLabel } = await getStockOverview();
  const product = products.find((p) => p.sku === sku);
  if (!product) return { found: false, sku, reason: `No SKU "${sku}" found in Stock Overview` };
  return {
    found: true,
    sku,
    batchLabel,
    balance: product.balance,
    requested: quantity,
    fulfillable: product.balance >= quantity,
    shortfall: Math.max(0, quantity - product.balance)
  };
}

// Guards -1 explicitly (audit 2026-08-31): every caller passes the result
// of headers.indexOf(...), and a missing/renamed column previously
// produced columnIndexToLetter(-1) === '' silently — turning a targeted
// single-cell write into a malformed, row-only A1 range that could
// overwrite the WRONG columns (e.g. Order ID) instead of failing loudly.
function columnIndexToLetter(index) {
  if (index < 0) throw new Error(`columnIndexToLetter: no such column (index ${index}) — a header may have been renamed or is missing from the sheet.`);
  let letter = '';
  let n = index;
  while (n >= 0) {
    letter = String.fromCharCode('A'.charCodeAt(0) + (n % 26)) + letter;
    n = Math.floor(n / 26) - 1;
  }
  return letter;
}

// Increments the "New Orders (Batch N)" cell for one SKU by `quantity`.
// This is the one write into Stock Overview verified safe against the real
// formula relationship above — it does NOT touch In Stock, the batch-
// reserved column, Available, or Balance directly, since those are (or
// behave like) live formulas that should recalculate on their own once
// New Orders changes. Watch the first real write to confirm Balance
// actually updates as expected before relying on this unattended.
//
// Locked per-SKU (audit 2026-08-31) — this is a read-modify-write (read
// the current count, add quantity, write back an absolute number, not a
// Sheets formula increment). Two concurrent calls for the same SKU could
// otherwise both read the same stale value and the second write would
// silently clobber the first, losing a real stock update with no error.
// Different SKUs still run concurrently since they touch different cells.
async function recordNewOrderAgainstBatch(sku, quantity) {
  return withLock(`stock:${sku}`, () => recordNewOrderAgainstBatchLocked(sku, quantity));
}

async function recordNewOrderAgainstBatchLocked(sku, quantity) {
  const { products, columns } = await getStockOverview();
  const product = products.find((p) => p.sku === sku);
  if (!product) throw new Error(`Cannot record order — no SKU "${sku}" found in Stock Overview.`);

  const newOrdersColLetter = columnIndexToLetter(columns.newOrders);
  const newValue = product.newOrders + quantity;
  await sheets.spreadsheets.values.update({
    spreadsheetId: process.env.SHEET_ID,
    range: `'${STOCK_OVERVIEW_TAB}'!${newOrdersColLetter}${product.rowNumber}`,
    valueInputOption: 'USER_ENTERED',
    requestBody: { values: [[newValue]] }
  });

  return { sku, previousNewOrders: product.newOrders, addedQuantity: quantity, newNewOrders: newValue };
}

// ---------------------------------------------------------------------------
// AUTOMATION LOG (added 2026-08-31) — a tab this agent fully owns and
// creates itself, deliberately SEPARATE from the real batch tabs. Xavier
// needs a working "log a sold deal / mark an order sent" tool now, but the
// real batch-tab column layout is still unconfirmed — writing into an
// unverified structure risks corrupting the actual ops sheet. This tab
// sidesteps that entirely: known headers, plain append-only log, safe to
// build without guessing anyone else's structure. Once the real batch tab
// layout is confirmed, this can either become the permanent record, or a
// sync step can copy entries from here into the correct batch tab — that
// decision is deliberately deferred, not made here.
// ---------------------------------------------------------------------------
const AUTOMATION_LOG_TAB = process.env.AUTOMATION_LOG_TAB || '🤖 Automation Log';
// Allocation / Batch Reference / Expected Date / Order Placed added
// 2026-08-31 — Xavier's 4-bucket stock model: every sold deal is claimed
// against On Shore stock, an On Water incoming batch, or falls to Next
// Custom Order (not yet placed with the manufacturer). Deliberately a
// field a HUMAN sets when logging the deal (informed by the live Stock
// Overview balance already shown in the ops console), not an automatic
// waterfall — an automatic allocation engine risks getting concurrent
// claims against the same limited stock wrong, which a person glancing at
// the real number does not.
// External Ref / Final Payment Status / Ship Target Date added 2026-08-31
// alongside the final-payment release gate — External Ref lets another
// agent (e.g. pipely-xero-agent, which only knows a Pipely opportunity ID,
// not this tab's generated Order ID) look an entry up reliably instead of
// parsing it out of free-text Notes.
const AUTOMATION_LOG_HEADERS = [
  'Order ID', 'External Ref', 'Timestamp', 'Source', 'Customer Name', 'Email', 'SKU', 'Product',
  'Quantity', 'Delivery Address', 'Deal Value', 'Stock Check', 'Deposit Status',
  'Final Payment Status', 'Ship Target Date',
  'Allocation', 'Batch Reference', 'Expected Date', 'Order Placed',
  'Courier', 'Tracking #', 'Order Sent Date', 'Notes'
];

async function ensureAutomationLogTab() {
  const sheetsMeta = await getSheetMeta();
  const exists = sheetsMeta.some((s) => s.properties.title === AUTOMATION_LOG_TAB);
  if (exists) return;

  await sheets.spreadsheets.batchUpdate({
    spreadsheetId: process.env.SHEET_ID,
    requestBody: { requests: [{ addSheet: { properties: { title: AUTOMATION_LOG_TAB } } }] }
  });
  await sheets.spreadsheets.values.update({
    spreadsheetId: process.env.SHEET_ID,
    range: `'${AUTOMATION_LOG_TAB}'!A1`,
    valueInputOption: 'USER_ENTERED',
    requestBody: { values: [AUTOMATION_LOG_HEADERS] }
  });
}

async function getAutomationLogRows() {
  // Fixed 2026-09-01: on a brand-new sheet copy where the tab has never
  // been created yet (only logSoldDeal's write path called
  // ensureAutomationLogTab before), a plain values.get on a non-existent
  // tab name throws "Unable to parse range" instead of just returning
  // empty — broke every read endpoint (automation-log, products-to-order)
  // until the first deal was ever logged.
  await ensureAutomationLogTab();
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: process.env.SHEET_ID,
    range: AUTOMATION_LOG_TAB
  });
  const rows = res.data.values ?? [];
  const headerRowIdx = rows.findIndex((row) => (row[0] ?? '').toString().trim() === 'Order ID');
  if (headerRowIdx === -1) return { headers: AUTOMATION_LOG_HEADERS, entries: [] };
  const headers = rows[headerRowIdx];
  const entries = rows.slice(headerRowIdx + 1)
    .map((row, i) => ({ rowNumber: headerRowIdx + 2 + i, ...Object.fromEntries(headers.map((h, idx) => [h, row[idx] ?? ''])) }))
    .filter((e) => e['Order ID']);
  return { headers, entries };
}

async function logSoldDeal({ source, externalRef, customerName, email, sku, quantity, deliveryAddress, dealValue, depositStatus, finalPaymentStatus, shipTargetDate, allocation, batchReference, expectedDate, notes }) {
  await ensureAutomationLogTab();

  const orderId = `EP-${Date.now()}-${Math.random().toString(36).slice(2, 7)}`;
  let stockCheckNote = 'Not checked';
  let productName = '';
  if (sku && quantity) {
    try {
      const check = await checkStockAvailability(sku, Number(quantity));
      if (check.found) {
        stockCheckNote = check.fulfillable
          ? `OK — ${check.balance} available (${check.batchLabel})`
          : `SHORT by ${check.shortfall} — only ${check.balance} available (${check.batchLabel})`;
        const overview = await getStockOverview();
        productName = overview.products.find((p) => p.sku === sku)?.productName ?? '';
      } else {
        stockCheckNote = check.reason;
      }
    } catch (err) {
      stockCheckNote = `Stock check failed: ${err.message}`;
    }
  }

  const row = [
    orderId,
    externalRef || '',
    new Date().toISOString(),
    source || 'Manual',
    customerName || '',
    email || '',
    sku || '',
    productName,
    quantity ?? '',
    deliveryAddress || '',
    dealValue ?? '',
    stockCheckNote,
    depositStatus || '',
    finalPaymentStatus || '',
    shipTargetDate || '',
    allocation || '',
    batchReference || '',
    expectedDate || '',
    '', // Order Placed — only meaningful for Next Custom Order rows, see markOrderPlaced
    '', '', '', // Courier, Tracking #, Order Sent Date — filled in later via markOrderSent
    notes || ''
  ];

  await sheets.spreadsheets.values.append({
    spreadsheetId: process.env.SHEET_ID,
    range: `'${AUTOMATION_LOG_TAB}'!A:A`,
    valueInputOption: 'USER_ENTERED',
    insertDataOption: 'INSERT_ROWS',
    requestBody: { values: [row] }
  });

  return { orderId, stockCheck: stockCheckNote };
}

// RELEASE GATE — Xavier: the final 50% "must be paid before sending the
// product," no exceptions. Enforced here, not just in the UI, so it can't
// be bypassed by calling this endpoint directly. Shopify orders skip the
// gate entirely — they're paid in full at checkout (Deposit Status is set
// to "Paid in full" by shopify-xero-agent), there's no final payment step
// for them to wait on.
function assertReadyToShip(entry) {
  if (entry['Deposit Status'] === 'Paid in full') return;
  if (entry['Final Payment Status'] === 'Paid') return;
  throw new Error(
    `Order ${entry['Order ID']} cannot be marked sent — final payment is not confirmed paid ` +
    `(current status: "${entry['Final Payment Status'] || 'not invoiced'}"). ` +
    `Confirm payment in Xero, then call /admin/mark-final-payment-received first.`
  );
}

async function markOrderSent({ orderId, courier, trackingNumber, sentDate }) {
  const { headers, entries } = await getAutomationLogRows();
  const entry = entries.find((e) => e['Order ID'] === orderId);
  if (!entry) throw new Error(`No Automation Log entry found for order ID "${orderId}"`);
  assertReadyToShip(entry);

  // Courier, Tracking #, Order Sent Date are three consecutive columns in
  // AUTOMATION_LOG_HEADERS, so a single contiguous range covers all three.
  const courierCol = columnIndexToLetter(headers.indexOf('Courier'));
  const sentDateCol = columnIndexToLetter(headers.indexOf('Order Sent Date'));

  await sheets.spreadsheets.values.update({
    spreadsheetId: process.env.SHEET_ID,
    range: `'${AUTOMATION_LOG_TAB}'!${courierCol}${entry.rowNumber}:${sentDateCol}${entry.rowNumber}`,
    valueInputOption: 'USER_ENTERED',
    requestBody: { values: [[courier || '', trackingNumber || '', sentDate || new Date().toISOString().slice(0, 10)]] }
  });

  return { orderId, courier, trackingNumber, sentDate };
}

// Finds an entry (and its row) by either key — orderId (what the ops
// console knows) or externalRef (what pipely-xero-agent knows: the Pipely
// opportunity ID, not this tab's generated Order ID). Takes the already-
// fetched {headers, entries} rather than re-fetching, since callers
// usually need both the lookup and the headers/column positions anyway.
function findAutomationLogEntry({ entries }, { orderId, externalRef }) {
  const entry = orderId
    ? entries.find((e) => e['Order ID'] === orderId)
    : entries.find((e) => e['External Ref'] === externalRef);
  if (!entry) throw new Error(`No Automation Log entry found for ${orderId ? `order ID "${orderId}"` : `external ref "${externalRef}"`}`);
  return entry;
}

async function setFinalPaymentStatus({ orderId, externalRef, status }) {
  const log = await getAutomationLogRows();
  const entry = findAutomationLogEntry(log, { orderId, externalRef });
  const col = columnIndexToLetter(log.headers.indexOf('Final Payment Status'));
  await sheets.spreadsheets.values.update({
    spreadsheetId: process.env.SHEET_ID,
    range: `'${AUTOMATION_LOG_TAB}'!${col}${entry.rowNumber}`,
    valueInputOption: 'USER_ENTERED',
    requestBody: { values: [[status]] }
  });
  return { orderId: entry['Order ID'], status };
}

// ---------------------------------------------------------------------------
// BATCH ETA / COUNTDOWN (added 2026-09-01) — Xavier's countdown-to-shore
// model, corrected same day from an initial per-ORDER design to per-BATCH
// (shipment): "I want a countdown on orders arriving to shores from their
// container boats" — matches the real spreadsheet's own structure, where
// each batch tab has a single shared "ETA:" field in its info strip
// (Shipment: BATCH-12 | ETA: ... | Route: ... | Status: ...), not a
// per-client date. Every order allocated to that batch (via the
// Automation Log's existing "Batch Reference" field) shares one ETA.
//
// Deliberately NOT read from/written into the real batch tabs themselves
// (e.g. "BATCH 12") — this agent has no parsing logic for those yet (their
// real layout — anchor rows, client rows, the "↳ Remaining" row — isn't
// something this code understands), and guessing at cell positions there
// risks corrupting a manually-maintained sheet. Kept as this agent's own
// small persisted store instead, keyed by Batch Reference (the same
// string already used to link orders to shipments), same "capture what we
// have" pattern as the Automation Log itself. Real Mainfreight tracking
// integration isn't confirmed yet, so this stays human-entered, not
// pulled automatically — see project notes.
// ---------------------------------------------------------------------------
const BATCH_ETAS_FILE = process.env.BATCH_ETAS_FILE || '/data/batch-etas.json';

function loadBatchEtas() {
  try { return JSON.parse(fs.readFileSync(BATCH_ETAS_FILE, 'utf8')); } catch { return {}; }
}
function saveBatchEtas(map) {
  try {
    fs.mkdirSync(path.dirname(BATCH_ETAS_FILE), { recursive: true });
    fs.writeFileSync(BATCH_ETAS_FILE, JSON.stringify(map, null, 2));
  } catch (err) {
    console.warn('Could not persist batch ETAs to disk:', err.message);
  }
}

function setBatchEta(batchReference, date) {
  const map = loadBatchEtas();
  map[batchReference] = { eta: date, updatedAt: new Date().toISOString() };
  saveBatchEtas(map);
  return { batchReference, eta: date };
}

// Not stored — computed from Allocation + Order Sent Date so it can't
// drift out of sync with them.
function deriveOrderStatus(entry) {
  if (entry['Order Sent Date']) return 'Sent';
  if (entry['Allocation'] === 'On Shore') return 'Ready to organise';
  if (entry['Allocation'] === 'On Water') return 'Awaiting stock (on water)';
  if (entry['Allocation'] === 'Next Custom Order') return 'Awaiting stock (not yet ordered)';
  return 'Unallocated';
}

// ---------------------------------------------------------------------------
// BATCH ARRIVAL — when a shipment lands, every deal that was allocated
// against it needs to flip from "waiting" to "ready to organise" in one
// move, per Xavier: "everyone already allocated to that stock needs to be
// organised." Only flips the Allocation tracking in the Automation Log —
// deliberately does NOT touch Stock Overview's physical In Stock count or
// the current-batch column rollover (e.g. "Batch 10" -> "Batch 11" in that
// tab's own headers). That needs a real arrival manifest (exact per-SKU
// quantities, not just what's already allocated to specific deals) and a
// confirmed process for the column rollover — guessing either risks the
// live stock formulas. Stays a manual step in Stock Overview for now.
// ---------------------------------------------------------------------------
async function markBatchArrived(batchReference) {
  const { headers, entries } = await getAutomationLogRows();
  const allocationCol = columnIndexToLetter(headers.indexOf('Allocation'));
  const matching = entries.filter((e) => e['Allocation'] === 'On Water' && e['Batch Reference'] === batchReference);

  for (const entry of matching) {
    await sheets.spreadsheets.values.update({
      spreadsheetId: process.env.SHEET_ID,
      range: `'${AUTOMATION_LOG_TAB}'!${allocationCol}${entry.rowNumber}`,
      valueInputOption: 'USER_ENTERED',
      requestBody: { values: [['On Shore']] }
    });
  }

  return {
    batchReference,
    flippedCount: matching.length,
    flipped: matching.map((e) => ({ orderId: e['Order ID'], customerName: e['Customer Name'], sku: e['SKU'], quantity: e['Quantity'] }))
  };
}

// ---------------------------------------------------------------------------
// PRODUCTS TO ORDER — bucket 4 of Xavier's model, computed rather than
// hand-maintained: every "Next Custom Order" deal not yet marked ordered,
// grouped by SKU. This is the live shopping list for the next China order.
//
// EXTENDED 2026-09-01 — after the Stock Overview fix above surfaced a real
// oversold SKU (more reserved across batch columns than physically in
// stock), Xavier: "if it's oversold that means there should be another 3
// in the pending to be ordered tab." Added `stockOverviewShortfall` per
// SKU, from any product where getStockOverview()'s computed balance is
// negative. Kept SEPARATE from `totalQuantity` (the Automation Log-based
// count) rather than summed into one number — I can't confirm whether a
// Stock Overview shortfall and a logged "Next Custom Order" deal for the
// same SKU represent the same underlying units or different ones (the
// batch-column data predates this ops-console/Automation-Log system and
// may not overlap with it at all). Silently adding them risks
// overstating how much actually needs ordering. Show both, let a human
// reconcile before placing a real order.
// ---------------------------------------------------------------------------
async function getProductsToOrder() {
  const { entries } = await getAutomationLogRows();
  const pending = entries.filter((e) => e['Allocation'] === 'Next Custom Order' && !e['Order Placed']);

  const bySku = {};
  for (const e of pending) {
    const sku = e['SKU'] || '(no SKU)';
    if (!bySku[sku]) bySku[sku] = { sku, product: e['Product'] || '', totalQuantity: 0, orderIds: [] };
    bySku[sku].totalQuantity += Number(e['Quantity'] || 0);
    bySku[sku].orderIds.push(e['Order ID']);
  }

  const { products } = await getStockOverview();
  for (const p of products) {
    if (p.balance < 0) {
      if (!bySku[p.sku]) bySku[p.sku] = { sku: p.sku, product: p.productName, totalQuantity: 0, orderIds: [] };
      bySku[p.sku].stockOverviewShortfall = -p.balance;
    }
  }

  return Object.values(bySku);
}

async function markOrderPlaced(orderId) {
  const { headers, entries } = await getAutomationLogRows();
  const entry = entries.find((e) => e['Order ID'] === orderId);
  if (!entry) throw new Error(`No Automation Log entry found for order ID "${orderId}"`);

  const col = columnIndexToLetter(headers.indexOf('Order Placed'));
  await sheets.spreadsheets.values.update({
    spreadsheetId: process.env.SHEET_ID,
    range: `'${AUTOMATION_LOG_TAB}'!${col}${entry.rowNumber}`,
    valueInputOption: 'USER_ENTERED',
    requestBody: { values: [['TRUE']] }
  });
  return { orderId };
}

// ---------------------------------------------------------------------------
// ADMIN ENDPOINTS
//
// Deliberately admin endpoints, not the only interface — the logic that
// decides WHEN a deal should turn into a sheet row (which batch, which
// product, is there even stock left) depends on cross-referencing Pipely/
// Shopify/Xero/Mainfreight that isn't defined yet (see project memory).
// This agent exposes the mechanism; wiring an automatic trigger to it is a
// separate, later step.
// ---------------------------------------------------------------------------
app.get('/admin/tabs', async (_req, res) => {
  try {
    const sheetsMeta = await getSheetMeta();
    res.json({ tabs: sheetsMeta.map((s) => s.properties.title) });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.get('/admin/read-tab', async (req, res) => {
  const { tabName } = req.query;
  if (!tabName) return res.status(400).json({ error: 'tabName query param is required' });
  try {
    const result = await sheets.spreadsheets.values.get({ spreadsheetId: process.env.SHEET_ID, range: tabName });
    res.json({ rows: result.data.values ?? [] });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.post('/admin/add-client', async (req, res) => {
  const { tabName, rowValues } = req.body;
  if (!tabName || !Array.isArray(rowValues)) {
    return res.status(400).json({ error: 'tabName and rowValues[] are required' });
  }
  try {
    const result = await insertClientRow(tabName, rowValues);
    res.json({ ok: true, ...result });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.get('/admin/stock-overview', async (_req, res) => {
  try {
    res.json(await getStockOverview());
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.post('/admin/check-availability', async (req, res) => {
  const { sku, quantity } = req.body;
  if (!sku || !quantity) return res.status(400).json({ error: 'sku and quantity are required' });
  try {
    res.json(await checkStockAvailability(sku, Number(quantity)));
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// Called by other agents (e.g. the Shopify-Xero agent) after a sale is
// confirmed, to bump the current batch's "New Orders" count for a SKU.
// Does not decide fulfillment strategy or touch batch tabs — just the one
// verified-safe Stock Overview write.
app.post('/admin/record-order', async (req, res) => {
  const { sku, quantity } = req.body;
  if (!sku || !quantity) return res.status(400).json({ error: 'sku and quantity are required' });
  try {
    res.json({ ok: true, ...(await recordNewOrderAgainstBatch(sku, Number(quantity))) });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.get('/admin/automation-log', async (_req, res) => {
  try {
    const { entries } = await getAutomationLogRows();
    const batchEtas = loadBatchEtas();
    res.json({
      entries: entries.map((e) => ({
        ...e,
        Status: deriveOrderStatus(e),
        // Computed, not a real sheet column — resolved from the batch ETA
        // store by "Batch Reference", the same key that already links an
        // order to its shipment. Consumers (ops console, pipely-xero-
        // agent's final-invoice sweep) read this one field rather than
        // each re-implementing the lookup.
        'Ship ETA': batchEtas[e['Batch Reference']]?.eta ?? null
      }))
    });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.post('/admin/log-sold-deal', async (req, res) => {
  const {
    source, externalRef, customerName, email, sku, quantity, deliveryAddress, dealValue,
    depositStatus, finalPaymentStatus, shipTargetDate, allocation, batchReference, expectedDate, notes
  } = req.body;
  if (!customerName) return res.status(400).json({ error: 'customerName is required' });
  try {
    res.json({
      ok: true,
      ...(await logSoldDeal({
        source, externalRef, customerName, email, sku, quantity, deliveryAddress, dealValue,
        depositStatus, finalPaymentStatus, shipTargetDate, allocation, batchReference, expectedDate, notes
      }))
    });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// Called once the final 50% invoice has been created (by pipely-xero-agent,
// via externalRef) or by ops directly (via orderId).
//
// Deliberately refuses status:"Paid" here (audit 2026-08-31) — this
// endpoint exists for pipely-xero-agent to record "Invoiced" right after
// it actually creates a real Xero invoice, not as a general-purpose status
// setter. Without this restriction, any caller holding this service's
// shared API key could POST status:"Paid" directly with no real invoice
// or payment behind it, satisfying assertReadyToShip's release gate for
// an order that was never actually paid. "Paid" is only reachable through
// /admin/mark-final-payment-received below, which is the one deliberate,
// named action a human takes after confirming payment in Xero themselves.
app.post('/admin/set-final-payment-status', async (req, res) => {
  const { orderId, externalRef, status } = req.body;
  if ((!orderId && !externalRef) || !status) {
    return res.status(400).json({ error: 'orderId or externalRef, and status, are required' });
  }
  if (status === 'Paid') {
    return res.status(400).json({ error: 'Cannot set status "Paid" through this endpoint — call /admin/mark-final-payment-received instead, after confirming payment in Xero.' });
  }
  try {
    res.json({ ok: true, ...(await setFinalPaymentStatus({ orderId, externalRef, status })) });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.post('/admin/mark-final-payment-received', async (req, res) => {
  const { orderId, externalRef } = req.body;
  if (!orderId && !externalRef) return res.status(400).json({ error: 'orderId or externalRef is required' });
  try {
    res.json({ ok: true, ...(await setFinalPaymentStatus({ orderId, externalRef, status: 'Paid' })) });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.post('/admin/set-batch-eta', (req, res) => {
  const { batchReference, date } = req.body;
  if (!batchReference || !date) {
    return res.status(400).json({ error: 'batchReference and date are required' });
  }
  try {
    res.json({ ok: true, ...setBatchEta(batchReference, date) });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.get('/admin/batch-etas', (_req, res) => {
  res.json({ batchEtas: loadBatchEtas() });
});

app.post('/admin/mark-order-sent', async (req, res) => {
  const { orderId, courier, trackingNumber, sentDate } = req.body;
  if (!orderId) return res.status(400).json({ error: 'orderId is required' });
  try {
    res.json({ ok: true, ...(await markOrderSent({ orderId, courier, trackingNumber, sentDate })) });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// Flips every deal allocated against an incoming batch to "ready to
// organise" once it's physically landed. Does NOT touch Stock Overview's
// In Stock count — see the comment above markBatchArrived for why.
app.post('/admin/mark-batch-arrived', async (req, res) => {
  const { batchReference } = req.body;
  if (!batchReference) return res.status(400).json({ error: 'batchReference is required' });
  try {
    res.json({ ok: true, ...(await markBatchArrived(batchReference)) });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.get('/admin/products-to-order', async (_req, res) => {
  try {
    res.json({ products: await getProductsToOrder() });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.post('/admin/mark-order-placed', async (req, res) => {
  const { orderId } = req.body;
  if (!orderId) return res.status(400).json({ error: 'orderId is required' });
  try {
    res.json({ ok: true, ...(await markOrderPlaced(orderId)) });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.get('/health', (_req, res) => res.json({ ok: true }));

const port = process.env.PORT || 3009;
app.listen(port, () => console.log(`Everest Plunge Stock Sheet Agent listening on :${port}`));

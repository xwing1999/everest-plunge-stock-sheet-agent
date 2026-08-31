import 'dotenv/config';
import express from 'express';
import { google } from 'googleapis';

const app = express();
app.use(express.json());

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
//   New Orders (Batch N) | Balance | Retail Price (NZD) | Warehouse | Notes
// The batch number is baked into the column HEADER TEXT itself ("Batch 10",
// "New Orders (Batch 10)") and shifts every time a new batch cycle starts —
// so columns are found by pattern match on each read, never by a fixed
// name or index. Verified from real numbers that Available = In Stock -
// Batch N, and Balance = Available - New Orders (Batch N); e.g. SKU-002:
// 15 - 5 = 10 available, 10 - 2 = 8 balance. Treated as a live formula
// relationship, matching the tab's own header note ("'Available'
// auto-calculates").
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

  const col = {
    sku: headers.findIndex((h) => h.toUpperCase() === 'SKU'),
    productName: headers.findIndex((h) => h.toUpperCase() === 'PRODUCT NAME'),
    modelSize: headers.findIndex((h) => h.toUpperCase().startsWith('MODEL')),
    inStock: headers.findIndex((h) => h.toUpperCase() === 'IN STOCK'),
    reserved: headers.findIndex((h) => /^batch\s*\d+$/i.test(h)),
    available: headers.findIndex((h) => h.toUpperCase() === 'AVAILABLE'),
    newOrders: headers.findIndex((h) => /^new orders\s*\(batch\s*\d+\)$/i.test(h)),
    balance: headers.findIndex((h) => h.toUpperCase() === 'BALANCE')
  };
  const missing = Object.entries(col).filter(([, idx]) => idx === -1).map(([name]) => name);
  if (missing.length) throw new Error(`Stock Overview header row is missing expected column(s): ${missing.join(', ')}. Sheet structure may have changed — check "${STOCK_OVERVIEW_TAB}" by eye before trusting this.`);

  const batchLabel = headers[col.reserved];
  const products = [];
  for (let r = headerRowIdx + 1; r < rows.length; r++) {
    const row = rows[r];
    const sku = (row[col.sku] ?? '').toString().trim();
    if (!sku || sku.toUpperCase() === 'TOTALS') break;
    products.push({
      sku,
      productName: row[col.productName] ?? '',
      modelSize: row[col.modelSize] ?? '',
      inStock: Number(row[col.inStock] || 0),
      reserved: Number(row[col.reserved] || 0),
      available: Number(row[col.available] || 0),
      newOrders: Number(row[col.newOrders] || 0),
      balance: Number(row[col.balance] || 0),
      rowNumber: r + 1 // 1-based sheet row, for writing back to this exact row later
    });
  }

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

function columnIndexToLetter(index) {
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
async function recordNewOrderAgainstBatch(sku, quantity) {
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
const AUTOMATION_LOG_HEADERS = [
  'Order ID', 'Timestamp', 'Source', 'Customer Name', 'Email', 'SKU', 'Product',
  'Quantity', 'Delivery Address', 'Deal Value', 'Stock Check', 'Deposit Status',
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

async function logSoldDeal({ source, customerName, email, sku, quantity, deliveryAddress, dealValue, depositStatus, allocation, batchReference, expectedDate, notes }) {
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

async function markOrderSent({ orderId, courier, trackingNumber, sentDate }) {
  const { headers, entries } = await getAutomationLogRows();
  const entry = entries.find((e) => e['Order ID'] === orderId);
  if (!entry) throw new Error(`No Automation Log entry found for order ID "${orderId}"`);

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
    res.json({ entries: entries.map((e) => ({ ...e, Status: deriveOrderStatus(e) })) });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.post('/admin/log-sold-deal', async (req, res) => {
  const { source, customerName, email, sku, quantity, deliveryAddress, dealValue, depositStatus, allocation, batchReference, expectedDate, notes } = req.body;
  if (!customerName) return res.status(400).json({ error: 'customerName is required' });
  try {
    res.json({ ok: true, ...(await logSoldDeal({ source, customerName, email, sku, quantity, deliveryAddress, dealValue, depositStatus, allocation, batchReference, expectedDate, notes })) });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
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

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

app.get('/health', (_req, res) => res.json({ ok: true }));

const port = process.env.PORT || 3009;
app.listen(port, () => console.log(`Everest Plunge Stock Sheet Agent listening on :${port}`));

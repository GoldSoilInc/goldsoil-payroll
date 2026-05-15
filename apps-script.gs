/* ============================================================
   GoldSoil Payroll — Google Apps Script backend
   ----------------------------------------------------------------
   Deploy this as a Web App:
     1. Apps Script editor → New project (or attach to the Sheet)
     2. Paste this file in
     3. Fill SHEET_ID + TOKEN below
     4. Deploy → New deployment → Type: Web app
        - Execute as: Me (your Workspace account)
        - Who has access: Anyone
     5. Copy the /exec URL — paste it into script.js as APPS_SCRIPT_URL
   Security model:
     - "Access: Anyone" is required because cross-origin fetch from
       GitHub Pages can't follow Google login redirects cleanly.
     - The TOKEN below is a shared secret — the front-end appends
       ?token=... to every request and we reject mismatches.
     - The deployment URL itself is also a secret (don't paste in chat).
     - This is "URL-secret" privacy — fine for an internal tool. If
       you want stronger auth later, switch to HTMLService hosting
       (the app lives at script.google.com instead of github.io).
   ============================================================ */

// CONFIG — fill these in before deploying
const SHEET_ID = 'PASTE_YOUR_GOOGLE_SHEET_ID_HERE';
const TOKEN    = 'PASTE_A_LONG_RANDOM_STRING_HERE';

// Tabs we expect to find. Names must match the Sheet exactly.
const TABS = [
  'Monthly Comm Outreach / Sales',
  'Approved Leads',
  'Sales',
  'Contracted Deals',
];

function doGet(e) {
  // Token check — fail closed
  const t = (e && e.parameter && e.parameter.token) || '';
  if (t !== TOKEN) {
    return jsonOut({ error: 'Unauthorized' });
  }

  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const result = {};
    for (const tabName of TABS) {
      const sheet = ss.getSheetByName(tabName);
      if (!sheet) {
        result[tabName] = { error: `Tab "${tabName}" not found in sheet` };
        continue;
      }
      result[tabName] = sheetToObjects(sheet);
    }
    // Include server clock so the front-end can default to "last full month"
    // even if the user's machine is in a weird timezone.
    result._meta = {
      generatedAt: new Date().toISOString(),
      sheetId: SHEET_ID,
      tabsRead: TABS,
    };
    return jsonOut(result);
  } catch (err) {
    return jsonOut({ error: String(err && err.message || err) });
  }
}

function sheetToObjects(sheet) {
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values[0].map(h => String(h).trim());
  const rows = [];
  for (let i = 1; i < values.length; i++) {
    // Skip fully-empty rows
    if (values[i].every(v => v === '' || v === null)) continue;
    const obj = {};
    for (let j = 0; j < headers.length; j++) {
      if (!headers[j]) continue;  // ignore unnamed columns
      const v = values[i][j];
      // Serialize Date objects as YYYY-MM-DD so the front-end parses cleanly
      if (v instanceof Date) {
        obj[headers[j]] = Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      } else {
        obj[headers[j]] = v;
      }
    }
    rows.push(obj);
  }
  return rows;
}

function jsonOut(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ------------------------------------------------------------
   OPTIONAL: a manual test you can run from the editor to verify
   the token + sheet plumbing without deploying.
   ------------------------------------------------------------ */
function _test_doGet() {
  const out = doGet({ parameter: { token: TOKEN } });
  Logger.log(out.getContent().substring(0, 1000));
}

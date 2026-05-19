/**
 * Master Database — Apps Script web app
 * Receives POST submissions from the data-entry website
 * and appends a row to the MasterData tab.
 *
 * SETUP (one-time):
 * 1. Open your Google Sheet → Extensions → Apps Script
 * 2. Delete any existing code, paste this entire file
 * 3. Click Deploy → New deployment
 *    - Type: Web app
 *    - Execute as: Me
 *    - Who has access: Anyone   (required — no Google login for submitters)
 * 4. Authorize when prompted, copy the deployment URL
 * 5. Paste that URL into the ENDPOINT constant in index.html
 *
 * If you ever edit this script, you MUST create a NEW deployment
 * (or use Manage deployments → edit → New version) for changes to take effect.
 */

const SHEET_NAME = 'MasterData';

// Maps incoming JSON keys to MasterData columns A–O in order.
const COLUMN_ORDER = [
  'orgName',       // A — Organization Name
  'region',        // B — Region / Province
  'primaryFocus',  // C — Primary Focus
  'orgSize',       // D — Organization Size
  'studentGroup',  // E — Assigned Student Group
  'contactEmail',  // F — Primary Contact Email
  'website',       // G — Website / Social Link
  'fundingArch',   // H — Funding Architecture (comma-separated)
  'instFriction',  // I — Institutional Friction (comma-separated)
  'opTrauma',      // J — Operational Trauma (comma-separated)
  'diasporaReal',  // K — Diaspora Realities (comma-separated)
  'outreach',      // L — Outreach Status
  'consent',       // M — Data Consent Status
  'notes',         // N — Qualitative Notes
  'dateLogged',    // O — Date Logged
];

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
    if (!sheet) throw new Error(`Sheet "${SHEET_NAME}" not found.`);

    // Optional anti-duplication: reject if the same website URL already exists
    if (data.website) {
      const urlCol = sheet.getRange('G2:G').getValues().flat().filter(Boolean);
      if (urlCol.includes(data.website.trim())) {
        return jsonResponse({ ok: false, error: 'Duplicate URL — already in ledger.' });
      }
    }

    if (data.action === 'update') {
      const rowIndex = parseInt(data.rowIndex);
      if (!rowIndex || rowIndex <= 1) throw new Error('Invalid rowIndex for update.');
      const row = COLUMN_ORDER.map(k => data[k] || '');
      sheet.getRange(rowIndex, 1, 1, row.length).setValues([row]);
      return jsonResponse({ ok: true });
    }

    const row = COLUMN_ORDER.map(k => data[k] || '');
    sheet.appendRow(row);

    return jsonResponse({ ok: true, row: sheet.getLastRow() });
  } catch (err) {
    return jsonResponse({ ok: false, error: err.message });
  }
}

function doGet(e) {
  try {
    const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
    if (!sheet) throw new Error(`Sheet "${SHEET_NAME}" not found.`);

    const action = e.parameter.action;

    // DELETE a row by 1-based sheet row index
    if (action === 'delete') {
      const rowIndex = parseInt(e.parameter.rowIndex);
      if (!rowIndex || rowIndex <= 1) throw new Error('Invalid rowIndex.');
      sheet.deleteRow(rowIndex);
      return jsonResponse({ ok: true });
    }

    // GET all data rows
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return jsonResponse({ ok: true, rows: [] });

    const rows = data.slice(1).map((row, i) => {
      const obj = {};
      COLUMN_ORDER.forEach((key, j) => { obj[key] = row[j]; });
      obj._rowIndex = i + 2; // 1-based, offset by header row
      return obj;
    });

    return jsonResponse({ ok: true, rows });

  } catch (err) {
    return jsonResponse({ ok: false, error: err.message });
  }
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

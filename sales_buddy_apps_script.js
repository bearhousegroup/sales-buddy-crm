// Sales Buddy CRM — Google Apps Script Backend
// Paste this entire file into Extensions > Apps Script in your Google Sheet
// Then: Deploy > New deployment > Web app > Execute as Me > Anyone > Deploy

const SHEET_ID = '1WJOdOcU3hcn6JyYkty5P1UlSQVQtYEEhywnUGWjONB4';
const ACCOUNTS_TAB = 'Accounts';
const VISITS_TAB = 'Visits';

// ---- CORS helper ----
function cors(output) {
  return output
    .setMimeType(ContentService.MimeType.JSON)
    .addHeader('Access-Control-Allow-Origin', '*')
    .addHeader('Access-Control-Allow-Methods', 'GET, POST')
    .addHeader('Access-Control-Allow-Headers', 'Content-Type');
}

// ---- GET — load all CRM data ----
function doGet(e) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    ensureSheets(ss);

    const action = e.parameter.action || 'load';

    if (action === 'load') {
      const accounts = getAccounts(ss);
      const visits = getVisits(ss);
      return cors(ContentService.createTextOutput(JSON.stringify({
        ok: true,
        accounts: accounts,
        visits: visits
      })));
    }

    return cors(ContentService.createTextOutput(JSON.stringify({ ok: false, error: 'Unknown action' })));
  } catch (err) {
    return cors(ContentService.createTextOutput(JSON.stringify({ ok: false, error: err.toString() })));
  }
}

// ---- POST — save account update or visit log ----
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.openById(SHEET_ID);
    ensureSheets(ss);

    if (body.action === 'saveAccount') {
      upsertAccount(ss, body.data);
      return cors(ContentService.createTextOutput(JSON.stringify({ ok: true })));
    }

    if (body.action === 'logVisit') {
      appendVisit(ss, body.data);
      // Also update lastVisit on account
      updateAccountLastVisit(ss, body.data.licenseId, body.data.date, body.data.rep);
      return cors(ContentService.createTextOutput(JSON.stringify({ ok: true })));
    }

    return cors(ContentService.createTextOutput(JSON.stringify({ ok: false, error: 'Unknown action' })));
  } catch (err) {
    return cors(ContentService.createTextOutput(JSON.stringify({ ok: false, error: err.toString() })));
  }
}

// ---- Sheet setup ----
function ensureSheets(ss) {
  // Accounts tab
  let accounts = ss.getSheetByName(ACCOUNTS_TAB);
  if (!accounts) {
    accounts = ss.insertSheet(ACCOUNTS_TAB);
    accounts.appendRow([
      'license_id', 'name', 'address', 'type', 'program',
      'stage', 'contact', 'phone', 'email', 'rep',
      'tier', 'zone', 'notes', 'website', 'terms',
      'total_orders', 'avg_order', 'last_order', 'first_order',
      'last_visit', 'last_visit_rep',
      'from_leaflink', 'updated_at'
    ]);
    accounts.setFrozenRows(1);
    // Format header
    accounts.getRange(1, 1, 1, 23)
      .setBackground('#1a7a4a')
      .setFontColor('white')
      .setFontWeight('bold');
    accounts.setColumnWidth(1, 140);
    accounts.setColumnWidth(2, 220);
    accounts.setColumnWidth(3, 280);
  }

  // Visits tab
  let visits = ss.getSheetByName(VISITS_TAB);
  if (!visits) {
    visits = ss.insertSheet(VISITS_TAB);
    visits.appendRow([
      'id', 'license_id', 'account_name', 'rep',
      'date', 'type', 'notes', 'logged_at'
    ]);
    visits.setFrozenRows(1);
    visits.getRange(1, 1, 1, 8)
      .setBackground('#1a7a4a')
      .setFontColor('white')
      .setFontWeight('bold');
    visits.setColumnWidth(3, 220);
    visits.setColumnWidth(7, 300);
  }
}

// ---- Read accounts ----
function getAccounts(ss) {
  const sheet = ss.getSheetByName(ACCOUNTS_TAB);
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return {};

  const headers = data[0];
  const result = {};
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const obj = {};
    headers.forEach((h, j) => obj[h] = row[j]);
    if (obj.license_id) result[obj.license_id] = obj;
  }
  return result;
}

// ---- Read visits ----
function getVisits(ss) {
  const sheet = ss.getSheetByName(VISITS_TAB);
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return {};

  const headers = data[0];
  const result = {};
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const obj = {};
    headers.forEach((h, j) => obj[h] = row[j]);
    const lid = obj.license_id;
    if (lid) {
      if (!result[lid]) result[lid] = [];
      result[lid].push({ date: obj.date, type: obj.type, notes: obj.notes, rep: obj.rep });
    }
  }
  return result;
}

// ---- Upsert account (find by license_id, update or append) ----
function upsertAccount(ss, data) {
  const sheet = ss.getSheetByName(ACCOUNTS_TAB);
  const rows = sheet.getDataRange().getValues();
  const headers = rows[0];
  const licIdx = headers.indexOf('license_id');

  // Find existing row
  let targetRow = -1;
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][licIdx] === data.license_id) {
      targetRow = i + 1; // 1-indexed for Sheets
      break;
    }
  }

  const now = new Date().toISOString();
  const rowData = [
    data.license_id || '',
    data.name || '',
    data.address || '',
    data.type || '',
    data.program || '',
    data.stage || 'prospect',
    data.contact || '',
    data.phone || '',
    data.email || '',
    data.rep || '',
    data.tier || '',
    data.zone || '',
    data.notes || '',
    data.website || '',
    data.terms || '',
    data.total_orders || 0,
    data.avg_order || 0,
    data.last_order || '',
    data.first_order || '',
    data.last_visit || '',
    data.last_visit_rep || '',
    data.from_leaflink ? 'yes' : '',
    now
  ];

  if (targetRow > 0) {
    sheet.getRange(targetRow, 1, 1, rowData.length).setValues([rowData]);
  } else {
    sheet.appendRow(rowData);
  }
}

// ---- Append visit log ----
function appendVisit(ss, data) {
  const sheet = ss.getSheetByName(VISITS_TAB);
  const id = Utilities.getUuid();
  sheet.appendRow([
    id,
    data.licenseId || '',
    data.accountName || '',
    data.rep || '',
    data.date || '',
    data.type || '',
    data.notes || '',
    new Date().toISOString()
  ]);
}

// ---- Update last visit on account ----
function updateAccountLastVisit(ss, licenseId, date, rep) {
  const sheet = ss.getSheetByName(ACCOUNTS_TAB);
  const rows = sheet.getDataRange().getValues();
  const headers = rows[0];
  const licIdx = headers.indexOf('license_id');
  const lvIdx = headers.indexOf('last_visit');
  const lvrIdx = headers.indexOf('last_visit_rep');

  for (let i = 1; i < rows.length; i++) {
    if (rows[i][licIdx] === licenseId) {
      sheet.getRange(i + 1, lvIdx + 1).setValue(date);
      sheet.getRange(i + 1, lvrIdx + 1).setValue(rep || '');
      return;
    }
  }
}

/**
 * Guest book for Mazen & Manar's wedding page.
 * Paste this whole file into Google Sheets: Extensions > Apps Script.
 * Every wish is saved as a row in the "Wishes" tab.
 * To hide a wish from the page, type  no  in its "Show" column.
 */
const SHEET_NAME = 'Wishes';

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(['Time', 'Name', 'Wish', 'Show (yes / no)']);
    sh.setFrozenRows(1);
  }
  return sh;
}

// The page reads the wishes (newest first)
function doGet() {
  const values = sheet_().getDataRange().getValues().slice(1);
  const rows = values
    .filter(r => r[1] && r[2] && String(r[3]).trim().toLowerCase() !== 'no')
    .map(r => ({ time: r[0], name: String(r[1]), message: String(r[2]) }))
    .reverse()
    .slice(0, 300);
  return json_({ ok: true, rows: rows });
}

// The page sends a new wish
function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents || '{}');
    if (d.website) return json_({ ok: true });            // spam bot trap
    const name = String(d.name || '').trim().slice(0, 60);
    const message = String(d.message || '').trim().slice(0, 600);
    if (!name || !message) return json_({ ok: false, error: 'empty' });

    const lock = LockService.getScriptLock();
    lock.waitLock(5000);
    sheet_().appendRow([new Date(), safe_(name), safe_(message), 'yes']);
    lock.releaseLock();
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

// Stops text that starts with = + - @ from running as a spreadsheet formula
function safe_(s) { return /^[=+\-@]/.test(s) ? "'" + s : s; }

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

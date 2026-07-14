/**
 * Leads CRM — Google Apps Script backend
 * -----------------------------------------------------
 * Paste this into the Apps Script editor bound to the Google Sheet that
 * StaticForms/your form is already writing leads into. Deploy it as a Web App
 * and paste the resulting URL into the CRM's Setup panel.
 *
 * Expected sheet columns (row 1 = header, edit COLS below if yours differ):
 *   A Timestamp | B Name | C Email | D Phone | E Company | F Message | G Status | H Notes
 *
 * If your sheet doesn't have Status / Notes columns yet, this script adds
 * them automatically the first time it runs.
 */

const SHEET_NAME = 'Sheet1'; // change to your tab name
const COLS = {
  timestamp: 1, name: 2, email: 3, phone: 4, company: 5,
  message: 6, status: 7, notes: 8
};
const STATUSES = ['New', 'Contacted', 'Qualified', 'Proposal Sent', 'Won', 'Lost'];

function _sheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
  const header = sh.getRange(1, 1, 1, 8).getValues()[0];
  if (!header[6]) sh.getRange(1, 7).setValue('Status');
  if (!header[7]) sh.getRange(1, 8).setValue('Notes');
  return sh;
}

function _json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  const sh = _sheet();
  const last = sh.getLastRow();
  if (last < 2) return _json({ leads: [] });
  const rows = sh.getRange(2, 1, last - 1, 8).getValues();
  const leads = rows.map((r, i) => ({
    id: i + 2, // sheet row number, used as the stable id
    timestamp: r[0] ? new Date(r[0]).toISOString() : '',
    name: r[1] || '', email: r[2] || '', phone: r[3] || '',
    company: r[4] || '', message: r[5] || '',
    status: r[6] || 'New', notes: r[7] || ''
  })).filter(l => l.name || l.email);
  return _json({ leads: leads });
}

function doPost(e) {
  const body = JSON.parse(e.postData.contents || '{}');
  const sh = _sheet();
  if (body.action === 'update' && body.id) {
    const row = Number(body.id);
    if (body.status !== undefined) sh.getRange(row, COLS.status).setValue(body.status);
    if (body.notes !== undefined) sh.getRange(row, COLS.notes).setValue(body.notes);
    return _json({ ok: true });
  }
  return _json({ ok: false, error: 'Unknown action' });
}

/**
 * Deploy → New deployment → type "Web app".
 *   Execute as: Me
 *   Who has access: Anyone (this is what lets your static site call it
 *   without server-side code or login; keep the resulting URL private)
 * Copy the Web App URL into the CRM's Setup panel.
 */

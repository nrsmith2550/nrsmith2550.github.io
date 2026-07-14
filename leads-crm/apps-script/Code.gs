/**
 * Leads CRM — Google Apps Script backend
 * -----------------------------------------------------
 * Paste this into the Apps Script editor bound to the Google Sheet that
 * StaticForms/your form is already writing leads into. Deploy it as a Web App
 * and paste the resulting URL into the CRM's Setup panel.
 *
 * Expected sheet columns (row 1 = header, edit COLS below if yours differ):
 *   A submitted_at | B submission_id | C subject | D name | E phone | F email
 *   G address | H service | I sqft | J message | K Status | L notes
 *
 * submission_id (column B) is a stable unique id from StaticForms — it's used
 * to find the right row on update instead of a row number.
 *
 * If your sheet doesn't have Status / notes columns yet, this script adds
 * them automatically the first time it runs.
 */

const SHEET_NAME = 'Sheet1'; // change to your tab name
const COLS = {
  submittedAt: 1, submissionId: 2, subject: 3, name: 4, phone: 5, email: 6,
  address: 7, service: 8, sqft: 9, message: 10, status: 11, notes: 12
};
const STATUSES = ['New', 'Contacted', 'Qualified', 'Proposal Sent', 'Won', 'Lost'];

function _sheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
  const header = sh.getRange(1, 1, 1, 12).getValues()[0];
  if (!header[10]) sh.getRange(1, 11).setValue('Status');
  if (!header[11]) sh.getRange(1, 12).setValue('notes');
  return sh;
}

function _json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function _findRowBySubmissionId(sh, id) {
  const last = sh.getLastRow();
  if (last < 2) return -1;
  const ids = sh.getRange(2, COLS.submissionId, last - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) return i + 2;
  }
  return -1;
}

function doGet(e) {
  const sh = _sheet();
  const last = sh.getLastRow();
  if (last < 2) return _json({ leads: [] });
  const rows = sh.getRange(2, 1, last - 1, 12).getValues();
  const leads = rows.map(r => ({
    id: r[1] || '', // submission_id, used as the stable id
    timestamp: r[0] ? new Date(r[0]).toISOString() : '',
    subject: r[2] || '', name: r[3] || '', phone: r[4] || '', email: r[5] || '',
    address: r[6] || '', service: r[7] || '', sqft: r[8] || '', message: r[9] || '',
    status: r[10] || 'New', notes: r[11] || ''
  })).filter(l => l.name || l.email);
  return _json({ leads: leads });
}

function doPost(e) {
  const body = JSON.parse(e.postData.contents || '{}');
  const sh = _sheet();
  if (body.action === 'update' && body.id) {
    const row = _findRowBySubmissionId(sh, body.id);
    if (row === -1) return _json({ ok: false, error: 'Lead not found' });
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

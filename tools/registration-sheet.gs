// Saves registrations from /register/ into the Google Sheet this script is attached to.
// Setup: Sheet > Extensions > Apps Script > paste this > Deploy > New deployment > Web app
//        Execute as: Me. Who has access: Anyone. Copy the /exec URL into register/index.html.
const FIELDS = ['date', 'age', 'phone', 'email', 'course', 'mode'];

function doPost(e) {
  const p = (e && e.parameter) || {};
  const reply = ok => ContentService.createTextOutput(JSON.stringify({ ok: ok }))
    .setMimeType(ContentService.MimeType.JSON);
  if (p.website) return reply(true); // hidden field only bots fill in
  if (FIELDS.some(f => !p[f] || p[f].length > 200)) return reply(false);

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
    if (sheet.getLastRow() === 0) sheet.appendRow(['Received', 'Date', 'Age', 'Phone', 'Email', 'Course', 'Mode']);
    // the leading apostrophe stores text as text, so nothing typed can run as a formula
    sheet.appendRow([new Date()].concat(FIELDS.map(f => "'" + p[f])));
  } finally {
    lock.releaseLock();
  }
  return reply(true);
}

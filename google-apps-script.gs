/**
 * Christmas in the Cellar — RSVP → Google Sheet
 * Target sheet: https://docs.google.com/spreadsheets/d/1QP2cEtPGqzmNcGfA3hqro37F339oztPkVYPRlQAi2BY/edit
 *
 * SETUP (about 3 minutes) — do this ON the sheet linked above:
 *   1. Open that Google Sheet.
 *   2. Extensions ▸ Apps Script. Delete any starter code, paste EVERYTHING
 *      in this file, and click Save.
 *   3. Deploy ▸ New deployment ▸ (gear icon) Web app.
 *        - Description:      RSVP
 *        - Execute as:       Me
 *        - Who has access:   Anyone
 *      Click Deploy, authorize when prompted, then COPY the "Web app URL"
 *      (looks like https://script.google.com/macros/s/AKfy..../exec).
 *   4. Send me that /exec URL and I'll wire the form to it.
 *
 * The first submission creates an "RSVPs" tab with a bold, frozen header row,
 * then every RSVP is appended as a new row in the column order you asked for.
 */

var HEADERS = [
  'Name',
  'Plus One?',
  'Name of Plus One',
  'Jacket Size of Plus One',
  'Jacket Size of RSVP',
  'Need Hotel?',
  'Attending Happy Hour?',
  'Submitted'
];

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000); // prevent two submissions writing at once
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('RSVPs') || ss.insertSheet('RSVPs');

    if (sheet.getLastRow() === 0) {
      sheet.appendRow(HEADERS);
      sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
      sheet.setFrozenRows(1);
    }

    var d = JSON.parse(e.postData.contents);
    sheet.appendRow([
      d.fullName || '',
      d.plusOne || '',
      d.plusOneName || '',
      d.plusOneJacketSize || '',
      d.jacketSize || '',
      d.hotelRoom || '',
      d.attendingHH || '',
      d.timestamp ? new Date(d.timestamp) : new Date()
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

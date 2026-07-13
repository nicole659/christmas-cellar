/**
 * Christmas in the Cellar — RSVP → Google Sheet
 *
 * SETUP (about 3 minutes):
 *   1. Create a new Google Sheet (sheet.new). Name it whatever you like.
 *   2. In that Sheet: Extensions ▸ Apps Script. Delete any starter code,
 *      then paste EVERYTHING in this file and click Save.
 *   3. Click Deploy ▸ New deployment ▸ (gear icon) Web app.
 *        - Description: RSVP
 *        - Execute as: Me
 *        - Who has access: Anyone
 *      Click Deploy, authorize when prompted, and COPY the "Web app URL"
 *      (it looks like https://script.google.com/macros/s/AKfy..../exec).
 *   4. Send me that URL and I'll paste it into the site's FORM_ENDPOINT.
 *
 * The first submission creates a "RSVPs" tab with a header row, then every
 * RSVP is appended as a new row.
 */

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000); // avoid two submissions writing at once
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('RSVPs') || ss.insertSheet('RSVPs');

    if (sheet.getLastRow() === 0) {
      sheet.appendRow([
        'Submitted', 'Full Name', 'Plus One?', 'Plus One Name',
        'Plus One Jacket', 'Hotel Room?', 'Tuesday Happy Hour?', 'Jacket Size'
      ]);
    }

    var d = JSON.parse(e.postData.contents);
    sheet.appendRow([
      d.timestamp || new Date().toISOString(),
      d.fullName || '',
      d.plusOne || '',
      d.plusOneName || '',
      d.plusOneJacketSize || '',
      d.hotelRoom || '',
      d.attendingHH || '',
      d.jacketSize || ''
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

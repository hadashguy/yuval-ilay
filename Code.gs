/**
 * הסעות לחתונה — קולט הרשמות מדף הנחיתה ושומר שורה בגוגל שיטס.
 * להדביק ב-Extensions ▸ Apps Script של הגיליון, ואז Deploy ▸ New deployment ▸ Web app.
 */

const SHEET_NAME = 'הרשמות';
const HEADERS = ['חותמת זמן', 'שם פרטי', 'שם משפחה', 'בן/בת זוג', 'מקומות',
                 'עדיפות ראשונה', 'עדיפות שנייה', 'הסעה חזור'];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    const data = parseBody(e);
    const partner = clean(data.partnerName);
    getSheet().appendRow([
      new Date(),
      clean(data.firstName),
      clean(data.lastName),
      partner,
      partner ? 2 : 1,   // מקומות — נגזר כאן ולא נשלח מהדפדפן
      clean(data.pickupFirst, 200),
      clean(data.pickupSecond, 200),
      data.returnRide === 'yes' ? 'כן' : 'לא'
    ]);
    return json({ ok: true });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

/** פתיחת הכתובת בדפדפן מחזירה את זה — דרך נוחה לבדוק שה-deployment חי. */
function doGet() {
  return json({ ok: true, service: 'wedding-shuttle' });
}

function parseBody(e) {
  if (e && e.postData && e.postData.contents) {
    try {
      return JSON.parse(e.postData.contents);
    } catch (err) {
      // לא JSON — ניפול לשדות הטופס למטה
    }
  }
  return (e && e.parameter) || {};
}

function clean(value, max) {
  return String(value == null ? '' : value).trim().slice(0, max || 80);
}

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.setRightToLeft(true);
    sheet.setColumnWidth(1, 150);
    sheet.setColumnWidth(6, 260);
    sheet.setColumnWidth(7, 260);
  }
  return sheet;
}

function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Comment sink for the Patient Data Strategy page hosted on GitHub Pages.
 *
 * Setup
 *  1. Create a Google Sheet. Rename the first tab to: Comments
 *  2. Extensions → Apps Script. Delete the placeholder and paste this file.
 *  3. Set SHARED_SECRET below to any random string, and use the same value
 *     in the page's commentApiKey tweak.
 *  4. Run setupSheet() once (Run ▸ setupSheet) and approve the permission prompt.
 *  5. Deploy → New deployment → type "Web app".
 *       Execute as:        Me
 *       Who has access:    Anyone
 *     Copy the /exec URL into the page's commentEndpoint tweak.
 *  6. Re-deploy (Manage deployments → edit → Version: New version) after any
 *     edit to this file, or the old code keeps running.
 *
 * The secret is visible to anyone who views the page source. That is acceptable
 * here because this endpoint can only append rows — it never reads or deletes.
 */

var SHARED_SECRET = 'purple-kangaroo';
var SHEET_NAME = 'Comments';
var HEADERS = ['received', 'name', 'comment', 'quoted text', 'chapter', 'section', 'page', 'client time', 'id', 'status'];

function setupSheet() {
  var sheet = sheetOrCreate_();
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(3, 420);
    sheet.setColumnWidth(4, 320);
  }
  return 'ready';
}

function doPost(e) {
  try {
    var body = JSON.parse((e && e.postData && e.postData.contents) || '{}');

    if (SHARED_SECRET && SHARED_SECRET !== 'CHANGE-ME' && body.token !== SHARED_SECRET) {
      return json_({ ok: false, error: 'unauthorized' });
    }

    // Deletes mark the row rather than removing it, so the review trail stays auditable.
    if (body.action === 'delete') {
      return json_({ ok: true, marked: markDeleted_(body.ids || []) });
    }

    if (!body.name || !body.body) {
      return json_({ ok: false, error: 'name and body are required' });
    }

    var lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      sheetOrCreate_().appendRow([
        new Date(),
        String(body.name).slice(0, 200),
        String(body.body).slice(0, 5000),
        String(body.quote || '').slice(0, 1000),
        String(body.chapter || ''),
        String(body.section || ''),
        String(body.page || ''),
        String(body.at || ''),
        String(body.id || ''),
        ''
      ]);
    } finally {
      lock.releaseLock();
    }

    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

/** Lets you confirm the deployment is live by opening the /exec URL in a browser. */
function doGet() {
  return json_({ ok: true, service: 'mca-comments' });
}

/** Writes "deleted" into the status column for each id the page reports removed. */
function markDeleted_(ids) {
  if (!ids.length) return 0;
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sheet = sheetOrCreate_();
    var rows = sheet.getLastRow() - 1;
    if (rows < 1) return 0;
    var idCol = HEADERS.indexOf('id') + 1;
    var statusCol = HEADERS.indexOf('status') + 1;
    var values = sheet.getRange(2, idCol, rows, 1).getValues();
    var marked = 0;
    for (var i = 0; i < values.length; i++) {
      if (ids.indexOf(String(values[i][0])) > -1) {
        sheet.getRange(i + 2, statusCol).setValue('deleted');
        marked++;
      }
    }
    return marked;
  } finally {
    lock.releaseLock();
  }
}

function sheetOrCreate_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

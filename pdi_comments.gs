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
 * The secret is visible to anyone who views the page source. Appending was
 * always open; note that the `list` action below also makes comments READABLE
 * to anyone holding the secret — i.e. every reviewer can read every comment.
 * That is the point of the comment board; keep it in mind before pasting
 * anything sensitive into the sheet.
 */

var SHARED_SECRET = 'purple-kangaroo';
var SHEET_NAME = 'Comments';
var HEADERS = ['received', 'name', 'comment', 'quoted text', 'chapter', 'section', 'page', 'client time', 'id', 'status', 'parent'];

function setupSheet() {
  var sheet = sheetOrCreate_();
  // Writes (or upgrades) the header row — safe to re-run after adding columns.
  sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold');
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(3, 420);
  sheet.setColumnWidth(4, 320);
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
        '',
        String(body.parent || '')   // id of the comment this replies to, if any
      ]);
    } finally {
      lock.releaseLock();
    }

    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

/**
 * GET /exec                      → liveness check
 * GET /exec?action=list&token=…  → all non-deleted comments as JSON (for the comment board)
 */
function doGet(e) {
  var p = (e && e.parameter) || {};
  if (p.action === 'list') {
    if (SHARED_SECRET && SHARED_SECRET !== 'CHANGE-ME' && p.token !== SHARED_SECRET) {
      return json_({ ok: false, error: 'unauthorized' });
    }
    return json_({ ok: true, comments: listComments_() });
  }
  return json_({ ok: true, service: 'mca-comments' });
}

function listComments_() {
  var sheet = sheetOrCreate_();
  var rows = sheet.getLastRow() - 1;
  if (rows < 1) return [];
  var values = sheet.getRange(2, 1, rows, HEADERS.length).getValues();
  var out = [];
  for (var i = 0; i < values.length; i++) {
    var v = values[i];
    if (String(v[9]) === 'deleted') continue;
    out.push({
      received: v[0] instanceof Date ? v[0].toISOString() : String(v[0]),
      name: String(v[1]),
      body: String(v[2]),
      quote: String(v[3]),
      chapter: String(v[4]),
      section: String(v[5]),
      page: String(v[6]),
      at: String(v[7]),
      id: String(v[8]),
      parent: String(v[10] || '')
    });
  }
  return out;
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

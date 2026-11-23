/**
 * VM Dissemination Reporting — free API on top of the Google Sheet (Google Apps Script Web App).
 *
 * READ   GET  <WEB_APP_URL>?action=data
 *        -> { sheets:[{name, source, values:[[...]]}], updated }
 *        Every tab is returned as raw values. The dashboard finds the header row in each tab,
 *        auto-corrects names (Rain/Rainfall, CTG/Chittagong, …) and removes the same broadcast
 *        logged in two tabs. So you can keep adding tabs (one per project, Beneficiaries, …).
 *
 * WRITE  (needs the API key; run setup() once to create it). Rows go to WRITE_SHEET.
 *   Same column order as the sheet (date, district, upazila, content, total, success, failed, project, media):
 *        ?action=add&key=KEY&q=23-11-2026,Sylhet,Kanaighat,Lightning,1005,454,551,UNICEF,Kotha
 *   Short form (date, upazila, content, total, media[, project]):
 *        ?action=add&key=KEY&q=23-11-2026,Ulipur,Flood,2579,GP,Protistha
 *   Named fields:
 *        ?action=add&key=KEY&date=2026-11-23&district=Sylhet&upazila=Kanaighat&content=Lightning&count=1005&success=454&failed=551&project=UNICEF&media=Kotha
 *   Many rows (POST, send Content-Type: text/plain to avoid a CORS preflight):
 *        { "key":"KEY", "sheet":"Sheet2", "rows":[ {date, district, upazila, content, count, success, failed, project, media} | "csv line" ] }
 *   Unique beneficiaries (the Sheet2 "Unique Beneficiary" block):
 *        ?action=beneficiary&key=KEY&project=Protistha&upazila=Ulipur&count=7303
 */

var WRITE_SHEET = 'Sheet2';              // the all-projects log new rows are appended to
var BENEFICIARY_SHEET = 'Beneficiaries'; // Project | District | Upazila | Unique Beneficiaries
var SHEET_COLUMNS = ['Date', 'Districts', 'Upazila', 'Message content', 'Total Dissimination ', 'Success', 'Failed', 'Project', 'Disseminated Media'];

// Small lookup so short-form rows get a district. The dashboard has the full national gazetteer.
var UPAZILA_DISTRICT = {
  fulchari: 'Gaibandha', fulchhari: 'Gaibandha', sundarganj: 'Gaibandha', dewanganj: 'Jamalpur', ulipur: 'Kurigram', chilmari: 'Kurigram',
  rajarhat: 'Kurigram', kalapara: 'Patuakhali', derai: 'Sunamganj', tahirpur: 'Sunamganj', dharmapasha: 'Sunamganj', gowainghat: 'Sylhet',
  kanaighat: 'Sylhet', banshkhali: 'Chittagong', teknaf: "Cox's Bazar", lama: 'Bandarban', naikhongchhari: 'Bandarban', khaliajuri: 'Netrakona', madan: 'Netrakona'
};

/** Run once from the editor: creates the Beneficiaries tab and an API key. */
function setup() {
  var ss = SpreadsheetApp.getActive();
  if (!ss.getSheetByName(BENEFICIARY_SHEET)) {
    ss.insertSheet(BENEFICIARY_SHEET).appendRow(['Project', 'District', 'Upazila', 'Unique Beneficiaries']);
  }
  var props = PropertiesService.getScriptProperties();
  if (!props.getProperty('API_KEY')) props.setProperty('API_KEY', Utilities.getUuid().slice(0, 8));
  Logger.log('API key: ' + props.getProperty('API_KEY'));
}

function doGet(e) {
  var p = (e && e.parameter) || {};
  try {
    var action = p.action || 'data';
    if (action === 'data') return json_(readAll_());
    if (action === 'add') {
      checkKey_(p.key);
      return json_({ ok: true, added: appendRows_([p.q ? parseLine_(p.q) : p], p.sheet) });
    }
    if (action === 'beneficiary') {
      checkKey_(p.key);
      return json_({ ok: true, beneficiary: upsertBeneficiary_(p) });
    }
    return json_({ ok: false, error: 'Unknown action ' + action });
  } catch (err) {
    return json_({ ok: false, error: String(err.message || err) });
  }
}

function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents || '{}');
    checkKey_(body.key);
    var out = { ok: true, added: 0, beneficiaries: 0 };
    if (body.rows) out.added = appendRows_(body.rows.map(function (r) { return typeof r === 'string' ? parseLine_(r) : r; }), body.sheet);
    (body.beneficiaries || []).forEach(function (b) { upsertBeneficiary_(b); out.beneficiaries++; });
    return json_(out);
  } catch (err) {
    return json_({ ok: false, error: String(err.message || err) });
  }
}

// ---------- read: every tab, raw ----------
function readAll_() {
  var ss = SpreadsheetApp.getActive();
  return {
    sheets: ss.getSheets().map(function (sh) {
      var range = sh.getDataRange(), values = range.getValues();
      values = values.map(function (r) {
        return r.map(function (c) { return c instanceof Date ? Utilities.formatDate(c, ss.getSpreadsheetTimeZone(), "yyyy-MM-dd'T'00:00:00") : c; });
      });
      while (values.length && values[values.length - 1].every(function (c) { return c === '' || c === null; })) values.pop();
      return { name: sh.getName(), source: ss.getName(), values: values };
    }),
    updated: new Date().toISOString()
  };
}

// ---------- write ----------
function appendRows_(rows, sheetName) {
  var ss = SpreadsheetApp.getActive(), sh = ss.getSheetByName(sheetName || WRITE_SHEET);
  if (!sh) { sh = ss.insertSheet(sheetName || WRITE_SHEET); sh.appendRow(SHEET_COLUMNS); }
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var data = rows.map(function (r) {
      var c = clean_(r);
      return [c.date, c.district, c.upazila, c.content, c.count, c.success, c.failed, c.project, c.media];
    });
    if (!data.length) return 0;
    var start = sh.getLastRow() + 1;
    sh.getRange(start, 1, data.length, SHEET_COLUMNS.length).setValues(data);
    sh.getRange(start, 1, data.length, 1).setNumberFormat('d mmmm yyyy'); // shows "23 November 2026" like the existing rows
    return data.length;
  } finally { lock.releaseLock(); }
}

function upsertBeneficiary_(p) {
  var ss = SpreadsheetApp.getActive();
  if (!ss.getSheetByName(BENEFICIARY_SHEET)) setup();
  var sh = ss.getSheetByName(BENEFICIARY_SHEET), project = p.project || 'Unassigned', n = Number(p.count);
  if (!p.upazila || isNaN(n)) throw new Error('Need upazila and count');
  var v = sh.getDataRange().getValues();
  for (var i = 1; i < v.length; i++) {
    if (norm_(v[i][0]) === norm_(project) && norm_(v[i][2]) === norm_(p.upazila)) {
      sh.getRange(i + 1, 4).setValue(n);
      return { project: project, upazila: p.upazila, count: n, updated: true };
    }
  }
  sh.appendRow([project, p.district || UPAZILA_DISTRICT[norm_(p.upazila).replace(/[^a-z]/g, '')] || '', p.upazila, n]);
  return { project: project, upazila: p.upazila, count: n, created: true };
}

// 9 fields = sheet order; 5-6 fields = short form
function parseLine_(q) {
  var a = String(q).split(/[,;|]/).map(function (s) { return s.trim(); });
  if (a.length >= 9) return { date: a[0], district: a[1], upazila: a[2], content: a[3], count: a[4], success: a[5], failed: a[6], project: a[7], media: a[8] };
  return { date: a[0], upazila: a[1], content: a[2], count: a[3], media: a[4], project: a[5] };
}

function clean_(r) {
  var count = Number(r.count != null && r.count !== '' ? r.count : r.total);
  if (isNaN(count)) throw new Error('Total (count) must be a number');
  var up = String(r.upazila || '').trim();
  var numOrBlank = function (v) { return v === '' || v == null || isNaN(Number(v)) ? '' : Number(v); };
  return {
    date: parseDate_(r.date),
    district: r.district || UPAZILA_DISTRICT[norm_(up).replace(/[^a-z]/g, '')] || '',
    upazila: up,
    content: r.content || '',
    count: count,
    success: numOrBlank(r.success),
    failed: numOrBlank(r.failed),
    project: r.project || '',
    media: r.media || ''
  };
}

function parseDate_(s) {
  if (!s) return new Date();
  if (s instanceof Date) return s;
  var m = String(s).match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m) return new Date(+m[1], +m[2] - 1, +m[3]);
  m = String(s).match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{4})/);
  if (m) return new Date(+m[3], +m[2] - 1, +m[1]);
  var d = new Date(s);  // "10 March 2026"
  if (!isNaN(d)) return d;
  throw new Error('Bad date: ' + s + ' (use DD-MM-YYYY, YYYY-MM-DD or "10 March 2026")');
}

function checkKey_(k) {
  var key = PropertiesService.getScriptProperties().getProperty('API_KEY');
  if (!key || k !== key) throw new Error('Invalid or missing API key');
}
function norm_(s) { return String(s == null ? '' : s).trim().toLowerCase(); }
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

/**
 * ============================================================
 *  NEXUS // ENG — REGISTRATION BACKEND  (Google Apps Script)
 * ============================================================
 *  HOW TO DEPLOY (full walkthrough in SETUP.md):
 *   1. Open your Google Sheet  →  Extensions  →  Apps Script
 *   2. Replace the default code with this file  →  Save
 *   3. Deploy  →  New deployment  →  Type: Web app
 *        Execute as: Me
 *        Who has access: Anyone
 *   4. Authorize, then copy the Web App URL (.../exec)
 *   5. Paste that URL into SHEET_URL in index.html AND owner.html
 *
 *  SECURITY NOTE: API_KEY must match SHEET_KEY in both pages.
 *  Change both if you want a private key.
 * ============================================================
 */

var SHEET_NAME = 'Registrations';
var API_KEY = 'NEXUS-S04-KEY-8F3A2B'; // ← change in pages too
var HEADERS = ['id', 'ts', 'lead', 'email', 'size', 'track', 'institution', 'edited', 'source'];

/** Find (or create) the registrations sheet. */
function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(HEADERS);
    return sh;
  }
  var first = sh.getRange(1, 1, 1, HEADERS.length).getValues()[0];
  if (String(first[0]) !== 'id') {
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  }
  return sh;
}

function iso_(d) {
  if (d instanceof Date && !isNaN(d.getTime())) return d.toISOString();
  return String(d || '');
}

/** Read every registration as JSON-friendly objects (newest first). */
function readAll_() {
  var sh = getSheet_();
  var last = sh.getLastRow();
  if (last < 2) return [];
  var values = sh.getRange(2, 1, last - 1, HEADERS.length).getValues();
  var rows = [];
  for (var i = 0; i < values.length; i++) {
    var r = values[i];
    if (!r[0]) continue;
    rows.push({
      id: String(r[0]),
      ts: iso_(r[1]),
      lead: String(r[2] || ''),
      email: String(r[3] || ''),
      size: String(r[4] || '2'),
      track: String(r[5] || ''),
      institution: String(r[6] || ''),
      edited: r[7] ? iso_(r[7]) : '',
      source: String(r[8] || 'web')
    });
  }
  rows.sort(function (a, b) {
    return (b.ts || '').localeCompare(a.ts || '');
  });
  return rows;
}

/** 1-based sheet row number for an id, or -1. */
function findRow_(sh, id) {
  var last = sh.getLastRow();
  if (last < 2) return -1;
  var ids = sh.getRange(2, 1, last - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) return i + 2;
  }
  return -1;
}

/** JSON or JSONP response (+ CORS headers when the API supports them). */
function json_(obj, cb) {
  var payload = JSON.stringify(obj);
  if (cb) {
    // JSONP: only allow safe callback names (prevents script injection)
    cb = String(cb).replace(/[^A-Za-z0-9_$]/g, '');
    return ContentService.createTextOutput('window["' + cb + '"](' + payload + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  var out = ContentService.createTextOutput(payload)
    .setMimeType(ContentService.MimeType.JSON);
  try {
    out.setHeaders({
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type'
    });
  } catch (e) {
    // setHeaders unavailable in this account — clients fall back to JSONP
  }
  return out;
}

function params_(e) {
  var p = {};
  if (e && e.parameter) {
    for (var k in e.parameter) p[k] = e.parameter[k];
  }
  if (e && e.postData && e.postData.contents) {
    try {
      var body = JSON.parse(e.postData.contents);
      for (var k2 in body) p[k2] = body[k2];
    } catch (err) {}
  }
  return p;
}

function handle_(p) {
  var cb = p.cb || '';
  if (p.key !== API_KEY) return json_({ ok: false, error: 'BAD_KEY' }, cb);

  var op = p.op || 'list';
  var sh, row, payload, r;

  if (op === 'list') {
    return json_({ ok: true, rows: readAll_() }, cb);
  }

  if (op === 'add') {
    payload = JSON.parse(p.payload || '{}');
    sh = getSheet_();
    var id = payload.id || 'e' + Date.now() + Math.random().toString(36).slice(2, 7);
    sh.appendRow([
      id,
      payload.ts || new Date().toISOString(),
      payload.lead || '',
      payload.email || '',
      payload.size || '2',
      payload.track || '',
      payload.institution || '',
      '',
      payload.source || 'web'
    ]);
    return json_({ ok: true, id: id }, cb);
  }

  if (op === 'update') {
    payload = JSON.parse(p.payload || '{}'); // { id, changes: {...} }
    sh = getSheet_();
    row = findRow_(sh, payload.id);
    if (row < 0) return json_({ ok: false, error: 'NOT_FOUND' }, cb);
    var current = null;
    var all = readAll_();
    for (var i = 0; i < all.length; i++) {
      if (String(all[i].id) === String(payload.id)) current = all[i];
    }
    if (!current) return json_({ ok: false, error: 'NOT_FOUND' }, cb);
    var c = payload.changes || {};
    var merged = {
      id: current.id,
      ts: current.ts,
      lead: 'lead' in c ? c.lead : current.lead,
      email: 'email' in c ? c.email : current.email,
      size: 'size' in c ? c.size : current.size,
      track: 'track' in c ? c.track : current.track,
      institution: 'institution' in c ? c.institution : current.institution,
      edited: new Date().toISOString(),
      source: current.source
    };
    sh.getRange(row, 1, 1, HEADERS.length).setValues([
      [merged.id, merged.ts, merged.lead, merged.email, merged.size,
       merged.track, merged.institution, merged.edited, merged.source]
    ]);
    return json_({ ok: true, row: merged }, cb);
  }

  if (op === 'delete') {
    sh = getSheet_();
    row = findRow_(sh, p.id);
    if (row < 0) return json_({ ok: false, error: 'NOT_FOUND' }, cb);
    sh.deleteRow(row);
    return json_({ ok: true }, cb);
  }

  if (op === 'purge') {
    sh = getSheet_();
    var last = sh.getLastRow();
    if (last > 1) sh.deleteRows(2, last - 1);
    return json_({ ok: true }, cb);
  }

  return json_({ ok: false, error: 'BAD_OP' }, cb);
}

/** GET  ?op=list|add|update|delete|purge&key=...[&payload=...][&cb=...] */
function doGet(e) {
  try {
    return handle_(params_(e));
  } catch (err) {
    return json_({ ok: false, error: String(err) }, e && e.parameter ? e.parameter.cb : '');
  }
}

/** POST with JSON body (same operations as GET). */
function doPost(e) {
  try {
    return handle_(params_(e));
  } catch (err) {
    return json_({ ok: false, error: String(err) }, e && e.parameter ? e.parameter.cb : '');
  }
}

/** CORS preflight. */
function doOptions() {
  try {
    return ContentService.createTextOutput('')
      .setMimeType(ContentService.MimeType.JSON)
      .setHeaders({
        'Access-Control-Allow-Origin': '*',
        'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
        'Access-Control-Allow-Headers': 'Content-Type'
      });
  } catch (err) {
    return ContentService.createTextOutput('').setMimeType(ContentService.MimeType.JSON);
  }
}

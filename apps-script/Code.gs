/**
 * EYECON OPTICS — Google Apps Script backend.
 * Every app section is stored in its own sheet of the spreadsheet this script is bound to.
 * Columns: A id | B updatedAt (ms) | C summary (for reading only) | D… record JSON (split into chunks).
 * Users live in the "_users" sheet; deletions are logged in "_deleted" so other devices sync them.
 */
var CHUNK = 45000;
var TOKEN_TTL = 12 * 3600 * 1000;
var RESERVED = { _users: 1, _deleted: 1 };

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('EYECON OPTICS')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
}

/** JSON API used by the GitHub Pages site: POST {fn, args} as text/plain. */
function doPost(e) {
  var out;
  try {
    var req = JSON.parse(e.postData.contents);
    var name = 'api_' + String(req.fn || '');
    var f = globalThis[name];
    if (!/^api_[A-Za-z]+$/.test(name) || typeof f !== 'function') throw new Error('bad_fn');
    out = { ok: f.apply(null, Array.isArray(req.args) ? req.args : []) };
  } catch (err) { out = { err: String(err && err.message || err) }; }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

/* ---------- helpers ---------- */
function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }
function props_() { return PropertiesService.getScriptProperties(); }
function sheet_(name, create) {
  var sh = ss_().getSheetByName(name);
  if (!sh && create) {
    sh = ss_().insertSheet(name);
    if (name === '_users') sh.appendRow(['username', 'name', 'role', 'active', 'salt', 'hash', 'created']);
    else if (name === '_deleted') sh.appendRow(['collection', 'id', 'time']);
    else sh.appendRow(['id', 'updatedAt', 'summary', 'data']);
    sh.setFrozenRows(1);
  }
  return sh;
}
function validCol_(c) {
  if (typeof c !== 'string' || !/^[A-Za-z][A-Za-z0-9_]{0,40}$/.test(c) || RESERVED[c]) throw new Error('bad_collection');
  return c;
}
function validId_(id) {
  if (typeof id !== 'string' || !/^[A-Za-z0-9_\-.]{1,80}$/.test(id)) throw new Error('bad_id');
  return id;
}
function hash_(salt, pass) {
  var h = salt + '|' + pass;
  for (var i = 0; i < 300; i++) {
    var b = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, h, Utilities.Charset.UTF_8);
    h = Utilities.base64Encode(b);
  }
  return h;
}
function uid_() { return Utilities.getUuid().replace(/-/g, ''); }
function summary_(c, d) {
  var keys = ['code', 'no', 'name', 'brand', 'customer', 'supplier', 'title', 'date', 'phone', 'note', 'total', 'amount', 'qty', 'status'];
  var out = [];
  for (var i = 0; i < keys.length && out.length < 6; i++) {
    var v = d[keys[i]];
    if (v !== undefined && v !== null && v !== '' && typeof v !== 'object') out.push(String(v).slice(0, 60));
  }
  return out.join(' | ');
}
function rowFor_(c, id, d, t) {
  var json = JSON.stringify(d), sm = summary_(c, d);
  var row = [id, t, /^[=+\-@]/.test(sm) ? ' ' + sm : sm];
  for (var i = 0; i < json.length; i += CHUNK) row.push('~' + json.slice(i, i + CHUNK));
  return row;
}
function readRow_(r) {
  var json = '';
  for (var i = 3; i < r.length; i++) { if (r[i] === '' || r[i] === null) break; var x = String(r[i]); json += x.charAt(0) === '~' ? x.slice(1) : x; }
  try { return JSON.parse(json || '{}'); } catch (e) { return {}; }
}
function findRow_(sh, id) {
  if (sh.getLastRow() < 2) return 0;
  var f = sh.getRange(2, 1, sh.getLastRow() - 1, 1).createTextFinder(id).matchEntireCell(true).findNext();
  return f ? f.getRow() : 0;
}
function writeRow_(sh, rowNo, row) {
  var width = Math.max(sh.getLastColumn(), row.length);
  var full = row.slice();
  while (full.length < width) full.push('');
  if (sh.getMaxColumns() < width) sh.insertColumnsAfter(sh.getMaxColumns(), width - sh.getMaxColumns());
  var rng = sh.getRange(rowNo, 1, 1, width);
  rng.setNumberFormat('@');
  rng.setValues([full.map(function (x) { return typeof x === 'number' ? String(x) : x; })]);
}
function touch_(c, t) { props_().setProperty('m_' + c, String(t)); props_().setProperty('ver', String(t)); }
function now_() { return Date.now(); }

/* ---------- auth ---------- */
function session_(token) {
  if (!token || typeof token !== 'string') throw new Error('auth');
  var raw = props_().getProperty('tok_' + token);
  if (!raw) throw new Error('auth');
  var s = JSON.parse(raw);
  if (s.exp < now_()) { props_().deleteProperty('tok_' + token); throw new Error('auth'); }
  return s;
}
function writer_(token) { var s = session_(token); if (s.role === 'viewer') throw new Error('perm'); return s; }
function admin_(token) { var s = session_(token); if (s.role !== 'admin') throw new Error('perm'); return s; }
function users_() {
  var sh = sheet_('_users', true), n = sh.getLastRow();
  if (n < 2) return [];
  return sh.getRange(2, 1, n - 1, 7).getValues().map(function (r, i) {
    return { row: i + 2, username: String(r[0]), name: String(r[1]), role: String(r[2]) || 'staff', active: r[3] === true || String(r[3]).toUpperCase() === 'TRUE', salt: String(r[4]), hash: String(r[5]) };
  });
}
function cleanTokens_() {
  var all = props_().getProperties(), t = now_();
  for (var k in all) if (k.indexOf('tok_') === 0) { try { if (JSON.parse(all[k]).exp < t) props_().deleteProperty(k); } catch (e) { props_().deleteProperty(k); } }
}
function issue_(u) {
  cleanTokens_();
  var tok = uid_() + uid_();
  props_().setProperty('tok_' + tok, JSON.stringify({ u: u.username, role: u.role, name: u.name, exp: now_() + TOKEN_TTL }));
  return { token: tok, user: { username: u.username, name: u.name, role: u.role } };
}

function api_status() { return JSON.stringify({ hasUsers: users_().length > 0 }); }

function api_setup(username, password, name) {
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    if (users_().length) throw new Error('exists');
    username = String(username || '').trim().toLowerCase();
    if (!/^[a-z0-9_.\-]{3,30}$/.test(username) || String(password || '').length < 6) throw new Error('weak');
    var salt = uid_();
    sheet_('_users', true).appendRow([username, String(name || username), 'admin', true, salt, hash_(salt, password), new Date()]);
    return JSON.stringify(issue_({ username: username, name: String(name || username), role: 'admin' }));
  } finally { lock.releaseLock(); }
}

function api_login(username, password) {
  username = String(username || '').trim().toLowerCase();
  var u = users_().filter(function (x) { return x.username === username; })[0];
  Utilities.sleep(300);
  if (!u || !u.active || hash_(u.salt, String(password || '')) !== u.hash) throw new Error('login');
  return JSON.stringify(issue_(u));
}

function api_logout(token) { props_().deleteProperty('tok_' + token); return '1'; }

function api_me(token) { var s = session_(token); return JSON.stringify({ username: s.u, name: s.name, role: s.role }); }

function api_users(token) {
  admin_(token);
  return JSON.stringify(users_().map(function (u) { return { username: u.username, name: u.name, role: u.role, active: u.active }; }));
}

function api_saveUser(token, o) {
  var me = admin_(token);
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var username = String(o.username || '').trim().toLowerCase();
    if (!/^[a-z0-9_.\-]{3,30}$/.test(username)) throw new Error('weak');
    var role = ['admin', 'staff', 'viewer'].indexOf(o.role) >= 0 ? o.role : 'staff';
    var list = users_(), u = list.filter(function (x) { return x.username === username; })[0];
    var sh = sheet_('_users', true);
    if (u && username === me.u && (role !== 'admin' || o.active === false)) throw new Error('self');
    if (!u) {
      if (String(o.password || '').length < 6) throw new Error('weak');
      var salt = uid_();
      sh.appendRow([username, String(o.name || username), role, o.active !== false, salt, hash_(salt, o.password), new Date()]);
    } else {
      sh.getRange(u.row, 2, 1, 3).setValues([[String(o.name || u.name), role, o.active !== false]]);
      if (o.password) {
        if (String(o.password).length < 6) throw new Error('weak');
        var s2 = uid_(); sh.getRange(u.row, 5, 1, 2).setValues([[s2, hash_(s2, o.password)]]);
      }
      if (o.password || o.active === false || role !== u.role) {
        var all = props_().getProperties();
        for (var k in all) if (k.indexOf('tok_') === 0) { try { if (JSON.parse(all[k]).u === username) props_().deleteProperty(k); } catch (e) {} }
      }
    }
    return '1';
  } finally { lock.releaseLock(); }
}

function api_changePassword(token, oldPass, newPass) {
  var s = session_(token);
  var u = users_().filter(function (x) { return x.username === s.u; })[0];
  if (!u || hash_(u.salt, String(oldPass || '')) !== u.hash) throw new Error('login');
  if (String(newPass || '').length < 6) throw new Error('weak');
  var salt = uid_();
  sheet_('_users', true).getRange(u.row, 5, 1, 2).setValues([[salt, hash_(salt, newPass)]]);
  return '1';
}

/* ---------- data ---------- */
function readCol_(c, since) {
  var sh = sheet_(c, false), out = [];
  if (!sh || sh.getLastRow() < 2) return out;
  var vals = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
  for (var i = 0; i < vals.length; i++) {
    var r = vals[i]; if (!r[0]) continue;
    var t = Number(r[1]) || 0;
    if (since && t <= since) continue;
    out.push([String(r[0]), readRow_(r), t]);
  }
  return out;
}

function api_load(token) {
  session_(token);
  var res = { t: now_(), cols: {} };
  ss_().getSheets().forEach(function (sh) {
    var n = sh.getName(); if (RESERVED[n] || !/^[A-Za-z][A-Za-z0-9_]{0,40}$/.test(n)) return;
    res.cols[n] = readCol_(n, 0).map(function (x) { return [x[0], x[1]]; });
  });
  return JSON.stringify(res);
}

function api_poll(token, since) {
  session_(token);
  since = Number(since) || 0;
  var t = now_(), all = props_().getProperties(), res = { t: t, cols: {}, del: [] };
  if (Number(all.ver || 0) <= since) return JSON.stringify(res);
  for (var k in all) {
    if (k.indexOf('m_') !== 0 || Number(all[k]) <= since) continue;
    var c = k.slice(2);
    if (c === '_deleted') {
      var dsh = sheet_('_deleted', false);
      if (dsh && dsh.getLastRow() > 1) dsh.getRange(2, 1, dsh.getLastRow() - 1, 3).getValues().forEach(function (r) { if (Number(r[2]) > since) res.del.push([String(r[0]), String(r[1])]); });
    } else res.cols[c] = readCol_(c, since).map(function (x) { return [x[0], x[1]]; });
  }
  return JSON.stringify(res);
}

function api_set(token, c, id, json, merge) {
  writer_(token); validCol_(c); validId_(id);
  var d = JSON.parse(json);
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var sh = sheet_(c, true), rowNo = findRow_(sh, id), t = now_();
    if (merge && rowNo) {
      var old = readRow_(sh.getRange(rowNo, 1, 1, sh.getLastColumn()).getValues()[0]);
      for (var k in d) old[k] = d[k];
      d = old;
    }
    if (rowNo) sh.getRange(rowNo, 4, 1, Math.max(1, sh.getLastColumn() - 3)).clearContent();
    writeRow_(sh, rowNo || sh.getLastRow() + 1, rowFor_(c, id, d, t));
    touch_(c, t);
    return JSON.stringify({ t: t, data: d });
  } finally { lock.releaseLock(); }
}

function api_del(token, c, id) {
  writer_(token); validCol_(c); validId_(id);
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var sh = sheet_(c, false), rowNo = sh ? findRow_(sh, id) : 0, t = now_();
    if (rowNo) sh.deleteRow(rowNo);
    sheet_('_deleted', true).appendRow([c, id, t]);
    touch_('_deleted', t);
    return JSON.stringify({ t: t });
  } finally { lock.releaseLock(); }
}

function api_incr(token, path, min) {
  writer_(token);
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var key = 'ctr_' + String(path).replace(/[^A-Za-z0-9_\/]/g, '');
    var n = Math.max(Number(props_().getProperty(key) || 0) + 1, Number(min) || 0);
    props_().setProperty(key, String(n));
    return String(n);
  } finally { lock.releaseLock(); }
}

/** Import an app backup (JSON from «تحميل نسخة احتياطية»). Admin only; existing ids are overwritten. */
function api_import(token, json) {
  admin_(token);
  var data = JSON.parse(json), n = 0;
  var lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    var t = now_();
    var put = function (c, id, d) {
      validCol_(c); validId_(id);
      var sh = sheet_(c, true), rowNo = findRow_(sh, id);
      if (rowNo) sh.getRange(rowNo, 4, 1, Math.max(1, sh.getLastColumn() - 3)).clearContent();
      writeRow_(sh, rowNo || sh.getLastRow() + 1, rowFor_(c, id, d, t)); n++;
    };
    if (data.settings) put('settings', 'main', data.settings);
    if (data.logoDoc) put('settings', 'logo', data.logoDoc); else if (data.logo) put('settings', 'logo', { data: data.logo });
    if (data.brandLogos) put('settings', 'brands', { logos: data.brandLogos });
    var touched = { settings: 1 };
    for (var c in data) {
      if (!Array.isArray(data[c]) || RESERVED[c]) continue;
      data[c].forEach(function (r) { if (r && r.id) { var d = {}; for (var k in r) if (k !== 'id') d[k] = r[k]; put(c, String(r.id), d); } });
      touched[c] = 1;
    }
    for (var k2 in touched) touch_(k2, t);
    return String(n);
  } finally { lock.releaseLock(); }
}

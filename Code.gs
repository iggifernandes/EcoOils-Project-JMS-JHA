// VERSION 1.0.0 — EcoOils JMS & JHA Builder backend (Google Apps Script)
// Deploy: Extensions > Apps Script (personal Gmail) > paste > run setup() once > Deploy > New deployment > Web app
//   Execute as: Me   |   Who has access: Anyone   (every data call still needs a login token)

var CFG = {
  ROOT: 'EcoOils JMS-JHA',
  SS_NAME: 'EcoOils JMS-JHA Database',
  TOKEN_HOURS: 12,
  RESET_MINUTES: 15,
  MAX_FAILS: 5,
  LOCK_MINUTES: 15
};
var HEAD = {
  Documents: ['id', 'type', 'ref', 'rev', 'status', 'date', 'title', 'ptw', 'updated', 'updatedBy', 'fileId', 'deleted'],
  Revisions: ['id', 'ref', 'rev', 'date', 'by', 'reason', 'fileId', 'issuedAt'],
  Lessons: ['id', 'date', 'title', 'source', 'categories', 'lsr', 'keywords', 'lesson', 'action', 'active', 'addedBy'],
  Users: ['email', 'name', 'role', 'notify', 'salt', 'pinHash'],
  Log: ['ts', 'email', 'action', 'detail']
};

/* ============ one-time setup (run manually once) ============ */
function setup() {
  var p = PropertiesService.getScriptProperties();
  var root = folder_();
  if (!p.getProperty('SS_ID')) {
    var ss = SpreadsheetApp.create(CFG.SS_NAME);
    p.setProperty('SS_ID', ss.getId());
    DriveApp.getFileById(ss.getId()).moveTo(root);
  }
  var ss2 = SpreadsheetApp.openById(p.getProperty('SS_ID'));
  Object.keys(HEAD).forEach(function (name) {
    var sh = ss2.getSheetByName(name) || ss2.insertSheet(name);
    if (sh.getLastRow() === 0) sh.appendRow(HEAD[name]);
  });
  var def = ss2.getSheetByName('Sheet1');
  if (def && ss2.getSheets().length > 1) ss2.deleteSheet(def);
  var owner = Session.getEffectiveUser().getEmail();
  if (!findRow_('Users', 'email', owner)) {
    ss2.getSheetByName('Users').appendRow([owner, 'Owner', 'admin', 'FALSE', '', '']);
  }
  Logger.log('Setup done. Owner: ' + owner + ' | Drive folder: ' + root.getUrl() + ' | Sheet: ' + ss2.getUrl());
}

/* ============ web app entry points ============ */
function doGet(e) { return json_({ ok: true, version: version_(), service: 'jms-jha' }); }

function doPost(e) {
  var req;
  try { req = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: 'BAD_REQUEST' }); }
  try {
    var a = req.action;
    if (a === 'ping') return json_({ ok: true, version: version_() });
    if (a === 'requestReset') return json_(requestReset_(req));
    if (a === 'setPin') return json_(setPin_(req));
    if (a === 'login') return json_(login_(req));
    var user = auth_(req.token);
    if (!user) return json_({ ok: false, error: 'AUTH' });
    var write = ['save', 'delete', 'lessonSave', 'lessonDelete', 'addUser', 'removeUser'].indexOf(a) >= 0;
    if (write) {
      if (user.role === 'viewer') return json_({ ok: false, error: 'FORBIDDEN' });
      var lock = LockService.getScriptLock();
      lock.waitLock(20000);
      try { return json_(route_(a, req, user)); } finally { lock.releaseLock(); }
    }
    return json_(route_(a, req, user));
  } catch (err) {
    log_('', 'ERROR', String(err));
    return json_({ ok: false, error: 'SERVER', message: String(err) });
  }
}

function route_(a, req, user) {
  switch (a) {
    case 'list': return list_();
    case 'get': return getDoc_(req.id);
    case 'save': return save_(req.doc, user);
    case 'delete': return need_(user, 'admin') || del_(req.id, user);
    case 'lessons': return { ok: true, lessons: lessons_() };
    case 'lessonSave': return lessonSave_(req.lesson, user);
    case 'lessonDelete': return lessonDelete_(req.id, user);
    case 'verify': return verify_(req.ref, req.rev);
    case 'users': return need_(user, 'admin') || { ok: true, users: rows_('Users').map(function (u) { return { email: u.email, name: u.name, role: u.role, notify: u.notify, hasPin: !!u.pinHash }; }) };
    case 'addUser': return need_(user, 'admin') || addUser_(req.user);
    case 'removeUser': return need_(user, 'admin') || removeUser_(req.email);
    case 'me': return { ok: true, email: user.email, name: user.name, role: user.role };
  }
  return { ok: false, error: 'UNKNOWN_ACTION' };
}

/* ============ auth: PIN can only be set through the email reset flow ============ */
function requestReset_(req) {
  var email = norm_(req.email);
  var u = findRow_('Users', 'email', email);
  // same answer whether or not the email is registered (no account probing)
  if (u) {
    var code = String(Math.floor(100000 + Math.random() * 900000));
    var salt = Utilities.getUuid();
    var rec = { h: hash_(salt, code), salt: salt, exp: Date.now() + CFG.RESET_MINUTES * 60000, tries: 0 };
    PropertiesService.getScriptProperties().setProperty('R_' + email, JSON.stringify(rec));
    MailApp.sendEmail({
      to: email,
      subject: 'EcoOils JMS & JHA Builder – PIN reset code',
      body: 'Your one-time code is ' + code + '\nIt expires in ' + CFG.RESET_MINUTES + ' minutes.\nIf you did not ask for this, ignore this email.'
    });
    log_(email, 'RESET_REQUEST', '');
  }
  return { ok: true, message: 'If this email is registered, a code has been sent.' };
}

function setPin_(req) {
  var email = norm_(req.email), pin = String(req.pin || ''), code = String(req.code || '');
  if (!/^\d{6}$/.test(pin)) return { ok: false, error: 'PIN_FORMAT', message: 'PIN must be 6 digits.' };
  var p = PropertiesService.getScriptProperties();
  var raw = p.getProperty('R_' + email);
  if (!raw) return { ok: false, error: 'BAD_CODE' };
  var rec = JSON.parse(raw);
  if (Date.now() > rec.exp || rec.tries >= 5) { p.deleteProperty('R_' + email); return { ok: false, error: 'BAD_CODE' }; }
  if (hash_(rec.salt, code) !== rec.h) { rec.tries++; p.setProperty('R_' + email, JSON.stringify(rec)); return { ok: false, error: 'BAD_CODE' }; }
  var salt = Utilities.getUuid();
  setCell_('Users', 'email', email, 'salt', salt);
  setCell_('Users', 'email', email, 'pinHash', hash_(salt, pin));
  p.deleteProperty('R_' + email);
  CacheService.getScriptCache().remove('F_' + email);
  log_(email, 'PIN_SET', '');
  return { ok: true };
}

function login_(req) {
  var email = norm_(req.email), pin = String(req.pin || '');
  var cache = CacheService.getScriptCache();
  var fails = Number(cache.get('F_' + email) || 0);
  if (fails >= CFG.MAX_FAILS) return { ok: false, error: 'LOCKED', message: 'Too many attempts. Try again in ' + CFG.LOCK_MINUTES + ' minutes.' };
  var u = findRow_('Users', 'email', email);
  if (!u || !u.pinHash || hash_(u.salt, pin) !== u.pinHash) {
    cache.put('F_' + email, String(fails + 1), CFG.LOCK_MINUTES * 60);
    return { ok: false, error: 'LOGIN' };
  }
  cache.remove('F_' + email);
  var token = Utilities.getUuid() + Utilities.getUuid();
  PropertiesService.getScriptProperties().setProperty('T_' + token, JSON.stringify({ email: email, exp: Date.now() + CFG.TOKEN_HOURS * 3600000 }));
  log_(email, 'LOGIN', '');
  return { ok: true, token: token, name: u.name, role: u.role, email: email };
}

function auth_(token) {
  if (!token) return null;
  var p = PropertiesService.getScriptProperties();
  var raw = p.getProperty('T_' + token);
  if (!raw) return null;
  var t = JSON.parse(raw);
  if (Date.now() > t.exp) { p.deleteProperty('T_' + token); return null; }
  var u = findRow_('Users', 'email', t.email);
  return u ? { email: u.email, name: u.name, role: u.role } : null;
}
function need_(user, role) { return user.role === role ? null : { ok: false, error: 'FORBIDDEN' }; }
function hash_(salt, v) {
  var d = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + ':' + v);
  return d.map(function (b) { return ('0' + ((b < 0 ? b + 256 : b)).toString(16)).slice(-2); }).join('');
}
function norm_(e) { return String(e || '').trim().toLowerCase(); }

/* ============ documents (JSON in Drive, index in Sheet) ============ */
function list_() {
  var docs = rows_('Documents').filter(function (r) { return String(r.deleted) !== 'TRUE'; }).map(function (r) {
    return { id: r.id, type: r.type, ref: r.ref, rev: Number(r.rev), status: r.status, date: r.date, title: r.title, ptw: r.ptw, updated: Number(r.updated) };
  });
  return { ok: true, docs: docs };
}

function getDoc_(id) {
  var r = findRow_('Documents', 'id', id);
  if (!r || String(r.deleted) === 'TRUE') return { ok: false, error: 'NOT_FOUND' };
  var txt = DriveApp.getFileById(r.fileId).getBlob().getDataAsString();
  return { ok: true, doc: JSON.parse(txt) };
}

function save_(doc, user) {
  if (!doc || !doc.id) return { ok: false, error: 'BAD_DOC' };
  var existing = findRow_('Documents', 'id', doc.id);
  if (existing && Number(existing.updated) > Number(doc.updated)) {
    return { ok: true, conflict: true, serverUpdated: Number(existing.updated) }; // last write wins: client keeps newer server copy
  }
  var json = JSON.stringify(doc);
  var fileId;
  if (existing && existing.fileId) {
    DriveApp.getFileById(existing.fileId).setContent(json);
    fileId = existing.fileId;
  } else {
    fileId = sub_('Documents').createFile(doc.ref + '_' + doc.id + '.json', json, 'application/json').getId();
  }
  var title = doc.type === 'jha' ? doc.data.task : doc.data.title;
  var row = [doc.id, doc.type, doc.ref, doc.rev, doc.status, doc.date, title || '', doc.data.ptw || '', doc.updated, user.email, fileId, 'FALSE'];
  upsert_('Documents', 'id', doc.id, row);
  // archive every issued revision once, and tell the people who must reprint
  var newly = [];
  (doc.history || []).forEach(function (h) {
    var key = doc.ref + '|' + h.rev;
    var ex = findRow_('Revisions', 'id', key);
    var snap = JSON.stringify({ ref: doc.ref, type: doc.type, rev: h.rev, date: h.date, by: h.by, reason: h.reason, lessons: h.lessons || [], data: h.data });
    if (ex) return;
    var f = sub_('Revisions').createFile(doc.ref + '_Rev' + pad2_(h.rev) + '.json', snap, 'application/json');
    upsert_('Revisions', 'id', key, [key, doc.ref, h.rev, h.date, h.by, h.reason, f.getId(), new Date().toISOString()]);
    newly.push(h);
  });
  newly.forEach(function (h) { notify_(doc, h); });
  return { ok: true, updated: doc.updated };
}

function del_(id, user) {
  setCell_('Documents', 'id', id, 'deleted', 'TRUE');
  log_(user.email, 'DELETE', id);
  return { ok: true };
}

function verify_(ref, rev) {
  var r = findRow_('Documents', 'ref', ref);
  if (!r) return { ok: true, status: 'UNKNOWN' };
  var cur = Number(r.rev), issued = r.status === 'issued';
  var latestIssued = issued ? cur : cur - 1;
  if (Number(rev) === latestIssued) return { ok: true, status: 'CURRENT', latest: latestIssued };
  return { ok: true, status: Number(rev) < latestIssued ? 'SUPERSEDED' : 'NOT_ISSUED', latest: latestIssued };
}

function notify_(doc, h) {
  var users = rows_('Users').filter(function (u) { return String(u.notify).toUpperCase() === 'TRUE'; });
  if (!users.length) return;
  var title = doc.type === 'jha' ? doc.data.task : doc.data.title;
  var body = doc.ref + ' (' + (doc.type === 'jha' ? 'JHA' : 'JMS') + ') – ' + (title || '') + '\n' +
    'Revision ' + pad2_(h.rev) + ' issued on ' + h.date + '.\nReason: ' + (h.reason || '') + '\n\n' +
    'PTW: ' + (doc.data.ptw || '-') + (h.rev > 0 ? '\nAll earlier printed copies are superseded. Print the new revision from the app and replace the old copy.' : '');
  users.forEach(function (u) {
    try { MailApp.sendEmail({ to: u.email, subject: '[' + doc.ref + '] Revision ' + pad2_(h.rev) + (h.rev > 0 ? ' issued – reprint required' : ' issued'), body: body }); } catch (e) { log_(u.email, 'MAIL_FAIL', String(e)); }
  });
}

/* ============ lessons learned ============ */
function lessons_() {
  return rows_('Lessons').filter(function (r) { return String(r.active).toUpperCase() !== 'FALSE'; }).map(function (r) {
    return { id: r.id, date: r.date, title: r.title, source: r.source, categories: split_(r.categories), lsr: split_(r.lsr).map(Number).filter(Boolean), keywords: split_(r.keywords), lesson: r.lesson, action: r.action };
  });
}
function lessonSave_(l, user) {
  if (!l || !l.title) return { ok: false, error: 'BAD_LESSON' };
  var id = l.id || ('L' + Date.now().toString(36));
  upsert_('Lessons', 'id', id, [id, l.date || '', l.title, l.source || '', (l.categories || []).join(','), (l.lsr || []).join(','), (l.keywords || []).join(','), l.lesson || '', l.action || '', 'TRUE', user.email]);
  log_(user.email, 'LESSON_SAVE', id);
  return { ok: true, id: id };
}
function lessonDelete_(id, user) {
  setCell_('Lessons', 'id', id, 'active', 'FALSE');
  log_(user.email, 'LESSON_OFF', id);
  return { ok: true };
}

/* ============ users ============ */
function addUser_(u) {
  var email = norm_(u && u.email);
  if (!email) return { ok: false, error: 'BAD_USER' };
  var role = ['admin', 'editor', 'viewer'].indexOf(u.role) >= 0 ? u.role : 'editor';
  upsert_('Users', 'email', email, [email, u.name || '', role, u.notify ? 'TRUE' : 'FALSE',
    (findRow_('Users', 'email', email) || {}).salt || '', (findRow_('Users', 'email', email) || {}).pinHash || '']);
  return { ok: true, message: 'User saved. They set their PIN with "Forgot PIN" on first login.' };
}
function removeUser_(email) {
  var sh = sheet_('Users'), rows = sh.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) if (norm_(rows[i][0]) === norm_(email)) { sh.deleteRow(i + 1); return { ok: true }; }
  return { ok: false, error: 'NOT_FOUND' };
}

/* ============ sheet / drive helpers ============ */
function version_() { return '1.0.0'; }
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function pad2_(n) { return ('0' + n).slice(-2); }
function split_(s) { return String(s || '').split(',').map(function (x) { return x.trim(); }).filter(String); }
function ss_() {
  var id = PropertiesService.getScriptProperties().getProperty('SS_ID');
  if (!id) throw new Error('Run setup() first');
  return SpreadsheetApp.openById(id);
}
function sheet_(n) { return ss_().getSheetByName(n); }
function rows_(name) {
  var v = sheet_(name).getDataRange().getValues();
  var h = v[0];
  return v.slice(1).map(function (r) { var o = {}; h.forEach(function (k, i) { o[k] = r[i]; }); return o; });
}
function findRow_(name, key, val) {
  var rs = rows_(name);
  for (var i = 0; i < rs.length; i++) {
    var a = rs[i][key];
    if (key === 'email' ? norm_(a) === norm_(val) : String(a) === String(val)) return rs[i];
  }
  return null;
}
function rowIndex_(name, key, val) {
  var sh = sheet_(name), v = sh.getDataRange().getValues(), c = v[0].indexOf(key);
  for (var i = 1; i < v.length; i++) if (key === 'email' ? norm_(v[i][c]) === norm_(val) : String(v[i][c]) === String(val)) return i + 1;
  return -1;
}
function upsert_(name, key, val, row) {
  var i = rowIndex_(name, key, val), sh = sheet_(name);
  if (i < 0) sh.appendRow(row); else sh.getRange(i, 1, 1, row.length).setValues([row]);
}
function setCell_(name, key, val, col, v) {
  var i = rowIndex_(name, key, val);
  if (i < 0) return;
  var sh = sheet_(name), c = sh.getDataRange().getValues()[0].indexOf(col) + 1;
  sh.getRange(i, c).setValue(v);
}
function folder_() {
  var p = PropertiesService.getScriptProperties(), id = p.getProperty('ROOT_ID');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) { } }
  var f = DriveApp.createFolder(CFG.ROOT);
  p.setProperty('ROOT_ID', f.getId());
  return f;
}
function sub_(name) {
  var root = folder_(), it = root.getFoldersByName(name);
  return it.hasNext() ? it.next() : root.createFolder(name);
}
function log_(email, action, detail) {
  try { sheet_('Log').appendRow([new Date().toISOString(), email, action, detail]); } catch (e) { }
}

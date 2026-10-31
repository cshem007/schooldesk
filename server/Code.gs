/**
 * School Desk sync backend (Google Apps Script web app).
 *
 * Deploy:  Deploy > New deployment > type "Web app"
 *          Execute as: Me      Who has access: Anyone
 * Copy the web app URL (ends in /exec) into config.js of the app.
 *
 * Data lives in a Drive folder called "SchoolDeskData" in the account that deploys this script.
 * Accounts are a name and a 4 to 6 digit PIN, kept (hashed) in Script Properties.
 */

const ROOT_NAME = 'SchoolDeskData';
const MAX_FAILS = 5;                 // wrong PINs in a row before a lock
const LOCK_MS = 10 * 60 * 1000;      // lock length
const TOKEN_MS = 120 * 24 * 3600 * 1000;
const MAX_DOC = 9 * 1024 * 1024;     // characters in one notebook file
const MAX_IMG = 4 * 1024 * 1024;     // characters in one photo (data URL)

function doGet() {
  return json_({ ok: true, app: 'school-desk' });
}

function doPost(e) {
  let out;
  try {
    const req = JSON.parse(e.postData.contents);
    out = handle_(req);
  } catch (err) {
    out = { error: 'BAD_REQUEST', detail: String((err && err.message) || err) };
  }
  return json_(out);
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

/* ---------------- routing ---------------- */

function handle_(req) {
  const a = req.action;
  const readOnly = (a === 'get' || a === 'getImg' || a === 'list');
  const lock = LockService.getScriptLock();
  if (!readOnly) lock.waitLock(25000);
  try {
    switch (a) {
      case 'register': return register_(req);
      case 'login':    return login_(req);
      case 'logout':   return logout_(req);
      case 'list':     return list_(auth_(req));
      case 'get':      return get_(auth_(req), req);
      case 'put':      return put_(auth_(req), req);
      case 'putImg':   return putImg_(auth_(req), req);
      case 'getImg':   return getImg_(auth_(req), req);
      case 'del':      return del_(auth_(req), req);
      default:         return { error: 'BAD_ACTION' };
    }
  } catch (err) {
    if (err && err.error) return err;           // thrown { error: '...' }
    throw err;
  } finally {
    if (!readOnly) lock.releaseLock();
  }
}

/* ---------------- accounts ---------------- */

function props_() { return PropertiesService.getScriptProperties(); }

function sha_(s) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s, Utilities.Charset.UTF_8)
    .map(function (b) { return ('0' + ((b < 0 ? b + 256 : b).toString(16))).slice(-2); }).join('');
}

function cleanName_(n) {
  n = String(n || '').trim().replace(/\s+/g, ' ');
  if (n.length < 2 || n.length > 30) throw { error: 'NAME' };
  return n;
}
function cleanPin_(p) {
  p = String(p || '');
  if (!/^\d{4,6}$/.test(p)) throw { error: 'PIN' };
  return p;
}
function userKey_(name) { return 'u:' + name.toLowerCase(); }
function fileKey_(name) { return sha_('sd:' + name.toLowerCase()).slice(0, 16); }

function newToken_(name) {
  const token = (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, '');
  props_().setProperty('t:' + token, JSON.stringify({ u: name, exp: Date.now() + TOKEN_MS }));
  pruneTokens_();
  return token;
}

function pruneTokens_() {
  if (Math.random() > 0.05) return;                      // now and then
  const p = props_(), all = p.getProperties(), now = Date.now();
  Object.keys(all).forEach(function (k) {
    if (k.indexOf('t:') !== 0) return;
    try { if (JSON.parse(all[k]).exp < now) p.deleteProperty(k); } catch (e) { p.deleteProperty(k); }
  });
}

function register_(req) {
  const name = cleanName_(req.name), pin = cleanPin_(req.pin), p = props_();
  if (p.getProperty(userKey_(name))) return { error: 'EXISTS' };
  const salt = Utilities.getUuid();
  p.setProperty(userKey_(name), JSON.stringify({ name: name, salt: salt, hash: sha_(salt + ':' + pin), fails: 0, until: 0 }));
  return { token: newToken_(name), name: name };
}

function login_(req) {
  const name = cleanName_(req.name), pin = cleanPin_(req.pin), p = props_();
  const raw = p.getProperty(userKey_(name));
  if (!raw) return { error: 'AUTH_FAIL' };
  const u = JSON.parse(raw);
  if (u.until && Date.now() < u.until) return { error: 'LOCKED' };
  if (sha_(u.salt + ':' + pin) !== u.hash) {
    u.fails = (u.fails || 0) + 1;
    if (u.fails >= MAX_FAILS) { u.until = Date.now() + LOCK_MS; u.fails = 0; }
    p.setProperty(userKey_(name), JSON.stringify(u));
    return { error: u.until && Date.now() < u.until ? 'LOCKED' : 'AUTH_FAIL' };
  }
  u.fails = 0; u.until = 0;
  p.setProperty(userKey_(name), JSON.stringify(u));
  return { token: newToken_(u.name), name: u.name };
}

function logout_(req) {
  if (req.token) props_().deleteProperty('t:' + String(req.token));
  return { ok: true };
}

function auth_(req) {
  const t = String(req.token || '');
  if (!/^[a-f0-9]{64}$/.test(t)) throw { error: 'AUTH' };
  const raw = props_().getProperty('t:' + t);
  if (!raw) throw { error: 'AUTH' };
  const o = JSON.parse(raw);
  if (o.exp < Date.now()) throw { error: 'AUTH' };
  return { name: o.u, fk: fileKey_(o.u) };
}

/* ---------------- Drive storage ---------------- */

function root_() {
  const p = props_(), id = p.getProperty('rootId');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) { /* recreate */ } }
  const it = DriveApp.getFoldersByName(ROOT_NAME);
  const f = it.hasNext() ? it.next() : DriveApp.createFolder(ROOT_NAME);
  p.setProperty('rootId', f.getId());
  return f;
}
function readFile_(name) {
  const it = root_().getFilesByName(name);
  return it.hasNext() ? it.next().getBlob().getDataAsString() : null;
}
function writeFile_(name, text) {
  const r = root_(), it = r.getFilesByName(name);
  if (it.hasNext()) it.next().setContent(text); else r.createFile(name, text, 'text/plain');
}
function trashFile_(name) {
  const it = root_().getFilesByName(name);
  while (it.hasNext()) it.next().setTrashed(true);
}
function fileExists_(name) { return root_().getFilesByName(name).hasNext(); }

function id_(v) {
  v = String(v || '');
  if (!/^[a-z0-9]{4,40}$/.test(v)) throw { error: 'BAD_ID' };
  return v;
}
function readIndex_(u) {
  const t = readFile_(u.fk + '__index.json');
  return t ? JSON.parse(t) : { nbs: {} };
}
function writeIndex_(u, idx) { writeFile_(u.fk + '__index.json', JSON.stringify(idx)); }

/* ---------------- notebooks ---------------- */

function list_(u) {
  return { nbs: readIndex_(u).nbs };
}

function get_(u, req) {
  const id = id_(req.id), t = readFile_(u.fk + '__nb_' + id + '.json');
  if (!t) return { error: 'NOT_FOUND' };
  const doc = JSON.parse(t);
  return { doc: doc, rev: doc.rev };
}

function put_(u, req) {
  const doc = req.doc || {};
  const id = id_(doc.id), base = Number(req.baseRev) || 0;
  const text0 = JSON.stringify(doc);
  if (text0.length > MAX_DOC) return { error: 'TOO_BIG' };
  const idx = readIndex_(u), cur = idx.nbs[id];
  if (cur && cur.del) return { error: 'DELETED' };
  if (cur && cur.rev !== base) return { error: 'CONFLICT', rev: cur.rev };
  if (!cur && base !== 0) return { error: 'CONFLICT', rev: 0 };
  const rev = (cur ? cur.rev : 0) + 1;
  doc.rev = rev;
  writeFile_(u.fk + '__nb_' + id + '.json', JSON.stringify(doc));
  idx.nbs[id] = { name: String(doc.name || '').slice(0, 60), color: String(doc.color || ''), rev: rev, updated: Date.now() };
  writeIndex_(u, idx);
  return { rev: rev };
}

function del_(u, req) {
  const id = id_(req.id), idx = readIndex_(u), cur = idx.nbs[id];
  if (!cur) return { error: 'NOT_FOUND' };
  const t = readFile_(u.fk + '__nb_' + id + '.json');
  if (t) {
    try {
      (JSON.parse(t).images || []).forEach(function (im) { trashFile_(u.fk + '__img_' + id_(im.id) + '.txt'); });
    } catch (e) { /* ignore */ }
    trashFile_(u.fk + '__nb_' + id + '.json');
  }
  idx.nbs[id] = { del: true, rev: cur.rev + 1, updated: Date.now() };
  writeIndex_(u, idx);
  return { ok: true };
}

/* ---------------- photos ---------------- */

function putImg_(u, req) {
  const id = id_(req.id), src = String(req.src || '');
  if (src.indexOf('data:image/') !== 0) return { error: 'BAD_IMG' };
  if (src.length > MAX_IMG) return { error: 'TOO_BIG' };
  const name = u.fk + '__img_' + id + '.txt';
  if (!fileExists_(name)) writeFile_(name, src);       // photos never change once saved
  return { ok: true };
}

function getImg_(u, req) {
  const t = readFile_(u.fk + '__img_' + id_(req.id) + '.txt');
  return t ? { src: t } : { error: 'NOT_FOUND' };
}

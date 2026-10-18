// Demo mode: the same backend code as Code.gs, running in the browser with data kept on this device.
window.TS_GAS = function (SpreadsheetApp, ContentService, Utilities, PropertiesService, LockService, CacheService, Session) {
/**
 * MozBlue Timesheets: Google Apps Script backend.
 *
 * The Google Sheet that holds this script stores everything:
 *   Staff    one row per person (PINs are stored only as salted hashes)
 *   Entries  one row per timesheet entry
 *   Flags    questions and flags from the Project Director, with replies
 *   Config   lists (work types, activity areas) as JSON
 *
 * Setup (once): set ADMIN_PIN below, run setup(), then Deploy > New deployment > Web app
 * (Execute as: Me, Who has access: Anyone). Paste the web app URL into config.js of the app.
 */

// Set Faruk's first PIN (4 to 8 digits) before running setup(). He can change it in the app.
const ADMIN_PIN = '';
const ADMIN_ID = 'ftavares';

const COLS = {
  Staff:   ['id', 'name', 'position', 'station', 'lineManager', 'role', 'active', 'pinHash', 'updatedAt'],
  Entries: ['id', 'staffId', 'date', 'workType', 'area', 'activity', 'location', 'summary', 'createdAt', 'updatedAt', 'deleted'],
  Flags:   ['id', 'entryId', 'staffId', 'entryDate', 'message', 'createdBy', 'createdAt', 'status', 'reply', 'repliedAt', 'resolvedAt', 'staffSeen'],
  Config:  ['key', 'value']
};
const DATE_FIELDS = {date: 1, entryDate: 1};
const TS_FIELDS = {createdAt: 1, updatedAt: 1, repliedAt: 1, resolvedAt: 1};

// ---------- web entry points
function doGet(e) { return out_(handle_({action: 'bootstrap'})); }
function doPost(e) {
  var req;
  try { req = JSON.parse(e.postData.contents); } catch (x) { return out_({ok: false, error: 'bad_request'}); }
  return out_(handle_(req));
}
function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

function handle_(req) {
  try {
    var a = req && req.action;
    if (a === 'bootstrap') return {ok: true, staff: publicStaff_(), lists: getConfig_('lists'), today: today_()};
    if (a === 'login') return login_(req);
    var me = auth_(req);
    switch (a) {
      case 'myData':      return myData_(me, req);
      case 'saveEntries': return locked_(function () { return saveEntries_(me, req.entries || []); });
      case 'replyFlag':   return locked_(function () { return replyFlag_(me, req); });
      case 'markSeen':    return locked_(function () { return markSeen_(me, req.ids || []); });
      case 'changePin':   return locked_(function () { return changePin_(me, req); });
      case 'teamData':    admin_(me); return teamData_(req);
      case 'flagEntry':   admin_(me); return locked_(function () { return flagEntry_(me, req); });
      case 'resolveFlag': admin_(me); return locked_(function () { return resolveFlag_(req); });
      case 'saveStaff':   admin_(me); return locked_(function () { return saveStaff_(req.staff || {}); });
      case 'resetPin':    admin_(me); return locked_(function () { return resetPin_(req.targetId); });
      case 'saveLists':   admin_(me); return locked_(function () { return saveLists_(req.lists); });
      default: return {ok: false, error: 'unknown_action'};
    }
  } catch (err) {
    return {ok: false, error: String(err && err.message || err)};
  }
}

// ---------- auth
function hash_(staffId, pin) {
  var salt = prop_('SALT');
  var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + '|' + staffId + '|' + pin, Utilities.Charset.UTF_8);
  return hex_(bytes);
}
function token_(staffId, pinHash) { return hex_(Utilities.computeHmacSha256Signature(staffId + '|' + pinHash, prop_('SECRET'))); }
function hex_(bytes) { return bytes.map(function (b) { return ('0' + (b & 255).toString(16)).slice(-2); }).join(''); }
function validPin_(p) { return /^\d{4,8}$/.test(String(p || '')); }

function login_(req) {
  var s = findStaff_(req.staffId);
  if (!s || !truthy_(s.active)) return {ok: false, error: 'unknown_user'};
  var cache = CacheService.getScriptCache(), key = 'fail_' + s.id, fails = +(cache.get(key) || 0);
  if (fails >= 8) return {ok: false, error: 'locked'};
  if (!s.pinHash) {
    if (!validPin_(req.newPin)) return {ok: false, error: 'pin_not_set'};
    return locked_(function () {
      var fresh = findStaff_(s.id);
      if (fresh.pinHash) return {ok: false, error: 'pin_already_set'};
      fresh.pinHash = hash_(fresh.id, req.newPin); fresh.updatedAt = now_(); writeRow_('Staff', fresh);
      return session_(fresh, true);
    });
  }
  if (hash_(s.id, req.pin) !== s.pinHash) { cache.put(key, String(fails + 1), 900); return {ok: false, error: 'wrong_pin'}; }
  cache.remove(key);
  return session_(s, false);
}
function session_(s, first) { return {ok: true, firstTime: first, token: token_(s.id, s.pinHash), me: publicOne_(s)}; }
function auth_(req) {
  var s = findStaff_(req.staffId);
  if (!s || !truthy_(s.active) || !s.pinHash || req.token !== token_(s.id, s.pinHash)) throw new Error('not_signed_in');
  return s;
}
function admin_(me) { if (me.role !== 'admin') throw new Error('admin_only'); }
function changePin_(me, req) {
  if (hash_(me.id, req.oldPin) !== me.pinHash) return {ok: false, error: 'wrong_pin'};
  if (!validPin_(req.newPin)) return {ok: false, error: 'invalid_pin'};
  me.pinHash = hash_(me.id, req.newPin); me.updatedAt = now_(); writeRow_('Staff', me);
  return session_(me, false);
}

// ---------- staff-side actions
function myData_(me, req) {
  var from = req.from || '0000', to = req.to || '9999';
  var entries = readAll_('Entries').filter(function (e) { return e.staffId === me.id && !truthy_(e.deleted) && e.date >= from && e.date <= to; });
  var flags = readAll_('Flags').filter(function (f) { return f.staffId === me.id; });
  return {ok: true, me: publicOne_(me), entries: entries.map(clean_), flags: flags.map(clean_), lists: getConfig_('lists'), today: today_()};
}
function saveEntries_(me, list) {
  var all = readAll_('Entries'), byId = {}, saved = [], skipped = [];
  all.forEach(function (e) { byId[e.id] = e; });
  list.forEach(function (x) {
    if (!x || !x.id || !/^\d{4}-\d{2}-\d{2}$/.test(String(x.date || ''))) { skipped.push(x && x.id); return; }
    var cur = byId[x.id];
    if (cur && cur.staffId !== me.id) { skipped.push(x.id); return; }
    var row = cur || {id: String(x.id).slice(0, 40), staffId: me.id, createdAt: now_()};
    ['date', 'workType', 'area', 'activity', 'location', 'summary'].forEach(function (k) { row[k] = text_(x[k], k === 'summary' ? 4000 : 200); });
    row.deleted = x.deleted ? true : false;
    row.updatedAt = now_();
    writeRow_('Entries', row);
    if (!cur) byId[row.id] = row;
    saved.push({id: row.id, updatedAt: row.updatedAt});
  });
  return {ok: true, saved: saved, skipped: skipped};
}
function replyFlag_(me, req) {
  var f = readAll_('Flags').filter(function (x) { return x.id === req.flagId; })[0];
  if (!f || f.staffId !== me.id) return {ok: false, error: 'not_found'};
  f.reply = text_(req.reply, 2000); f.repliedAt = now_(); f.status = 'answered'; f.staffSeen = true;
  writeRow_('Flags', f);
  return {ok: true, flag: clean_(f)};
}
function markSeen_(me, ids) {
  readAll_('Flags').forEach(function (f) { if (f.staffId === me.id && ids.indexOf(f.id) >= 0 && !truthy_(f.staffSeen)) { f.staffSeen = true; writeRow_('Flags', f); } });
  return {ok: true};
}

// ---------- Project Director (admin) actions
function teamData_(req) {
  var from = req.from || '0000', to = req.to || '9999';
  var entries = readAll_('Entries').filter(function (e) { return !truthy_(e.deleted) && e.date >= from && e.date <= to; });
  return {ok: true, staff: publicStaff_(true), entries: entries.map(clean_), flags: readAll_('Flags').map(clean_), lists: getConfig_('lists'), today: today_()};
}
function flagEntry_(me, req) {
  // Either an existing entry (entryId) or a day without an entry (targetId + date).
  var e;
  if (req.entryId) { e = readAll_('Entries').filter(function (x) { return x.id === req.entryId; })[0]; if (!e) return {ok: false, error: 'not_found'}; }
  else { if (!findStaff_(req.targetId) || !/^\d{4}-\d{2}-\d{2}$/.test(String(req.date || ''))) return {ok: false, error: 'not_found'}; e = {id: '', staffId: req.targetId, date: req.date}; }
  if (!String(req.message || '').trim()) return {ok: false, error: 'message_required'};
  var f = {id: uid_(), entryId: e.id, staffId: e.staffId, entryDate: e.date, message: text_(req.message, 2000), createdBy: me.name, createdAt: now_(), status: 'open', reply: '', repliedAt: '', resolvedAt: '', staffSeen: false};
  writeRow_('Flags', f);
  return {ok: true, flag: clean_(f)};
}
function resolveFlag_(req) {
  var f = readAll_('Flags').filter(function (x) { return x.id === req.flagId; })[0];
  if (!f) return {ok: false, error: 'not_found'};
  f.status = req.reopen ? (f.reply ? 'answered' : 'open') : 'resolved'; f.resolvedAt = req.reopen ? '' : now_();
  writeRow_('Flags', f);
  return {ok: true, flag: clean_(f)};
}
function saveStaff_(x) {
  if (!x.name) return {ok: false, error: 'name_required'};
  var all = readAll_('Staff'), cur = x.id ? all.filter(function (s) { return s.id === x.id; })[0] : null;
  if (!cur) {
    var base = (String(x.name).split(/\s+/)[0].charAt(0) + String(x.name).split(/\s+/).slice(-1)[0]).toLowerCase().normalize('NFD').replace(/[^a-z]/g, '') || 'staff', id = base, n = 2;
    while (all.some(function (s) { return s.id === id; })) id = base + n++;
    cur = {id: id, pinHash: ''};
  }
  ['name', 'position', 'station', 'lineManager'].forEach(function (k) { if (x[k] != null) cur[k] = text_(x[k], 200); });
  if (x.role === 'admin' || x.role === 'staff') cur.role = x.role;
  if (!cur.role) cur.role = 'staff';
  cur.active = x.active === false ? false : true;
  if (cur.id === ADMIN_ID) { cur.role = 'admin'; cur.active = true; }
  cur.updatedAt = now_();
  writeRow_('Staff', cur);
  return {ok: true, staff: publicOne_(cur, true)};
}
function resetPin_(id) {
  var s = findStaff_(id); if (!s) return {ok: false, error: 'not_found'};
  s.pinHash = ''; s.updatedAt = now_(); writeRow_('Staff', s);
  return {ok: true};
}
function saveLists_(lists) {
  if (!lists || !Array.isArray(lists.workTypes) || !Array.isArray(lists.areas)) return {ok: false, error: 'bad_lists'};
  setConfig_('lists', lists);
  return {ok: true, lists: lists};
}

// ---------- sheet helpers
function sheet_(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet(), s = ss.getSheetByName(name);
  if (!s) {
    s = ss.insertSheet(name);
    s.getRange(1, 1, 1, COLS[name].length).setValues([COLS[name]]);
    s.getRange(1, 1, s.getMaxRows(), COLS[name].length).setNumberFormat('@');
    s.setFrozenRows(1);
  }
  return s;
}
function readAll_(name) {
  var s = sheet_(name), last = s.getLastRow(), cols = COLS[name];
  if (last < 2) return [];
  var vals = s.getRange(2, 1, last - 1, cols.length).getValues(), tz = Session.getScriptTimeZone();
  return vals.map(function (r, i) {
    var o = {_row: i + 2};
    cols.forEach(function (k, j) {
      var v = r[j];
      if (v instanceof Date) v = DATE_FIELDS[k] ? Utilities.formatDate(v, tz, 'yyyy-MM-dd') : v.toISOString();
      o[k] = v;
    });
    return o;
  }).filter(function (o) { return o.id !== '' && o.id != null || o.key; });
}
function writeRow_(name, o) {
  var s = sheet_(name), cols = COLS[name];
  var row = cols.map(function (k) { var v = o[k]; if (v == null) return ''; if (typeof v === 'string' && /^[=+\-@]/.test(v)) return "'" + v; return v; });
  if (o._row) s.getRange(o._row, 1, 1, cols.length).setValues([row]);
  else { var r = s.getLastRow() + 1; s.getRange(r, 1, 1, cols.length).setNumberFormat('@').setValues([row]); o._row = r; }
}
function findStaff_(id) { return readAll_('Staff').filter(function (s) { return s.id === id; })[0] || null; }
function publicOne_(s, full) {
  var o = {id: s.id, name: s.name, position: s.position || '', station: s.station || '', lineManager: s.lineManager || '', role: s.role || 'staff', active: truthy_(s.active), hasPin: !!s.pinHash};
  return o;
}
function publicStaff_(all) { return readAll_('Staff').filter(function (s) { return all || truthy_(s.active); }).map(function (s) { return publicOne_(s, all); }); }
function clean_(o) { var c = {}; for (var k in o) if (k !== '_row') c[k] = o[k]; c.deleted = truthy_(c.deleted); if ('staffSeen' in c) c.staffSeen = truthy_(c.staffSeen); return c; }
function getConfig_(key) {
  var row = readAll_('Config').filter(function (r) { return r.key === key; })[0];
  try { return row ? JSON.parse(row.value) : null; } catch (e) { return null; }
}
function setConfig_(key, val) {
  var row = readAll_('Config').filter(function (r) { return r.key === key; })[0] || {key: key};
  row.value = JSON.stringify(val); writeRow_('Config', row);
}
function locked_(fn) {
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}
function prop_(k) { return PropertiesService.getScriptProperties().getProperty(k) || ''; }
function truthy_(v) { return v === true || v === 'TRUE' || v === 'true' || v === 1; }
function text_(v, max) { return String(v == null ? '' : v).slice(0, max || 500); }
function now_() { return new Date().toISOString(); }
function today_() { return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'); }
function uid_() { return Utilities.getUuid().replace(/-/g, '').slice(0, 12); }

// ---------- one-time setup
function setup() { setupWith_(ADMIN_PIN); }
function setupWith_(adminPin) {
  var props = PropertiesService.getScriptProperties();
  if (!props.getProperty('SALT')) props.setProperty('SALT', Utilities.getUuid());
  if (!props.getProperty('SECRET')) props.setProperty('SECRET', Utilities.getUuid() + Utilities.getUuid());
  ['Staff', 'Entries', 'Flags', 'Config'].forEach(sheet_);
  if (!readAll_('Staff').length) SEED.staff.forEach(function (s) { writeRow_('Staff', {id: s.id, name: s.name, position: s.position, station: s.station, lineManager: s.lineManager, role: s.role, active: true, pinHash: '', updatedAt: now_()}); });
  if (!readAll_('Entries').length) {
    var stamp = now_(), rows = SEED.entries.map(function (e) { return COLS.Entries.map(function (k) { return k === 'createdAt' || k === 'updatedAt' ? stamp : k === 'deleted' ? false : (e[k] == null ? '' : e[k]); }); });
    if (rows.length) sheet_('Entries').getRange(2, 1, rows.length, COLS.Entries.length).setNumberFormat('@').setValues(rows);
  }
  if (!getConfig_('lists')) setConfig_('lists', SEED.lists);
  var admin = findStaff_(ADMIN_ID);
  if (admin && !admin.pinHash) {
    if (!validPin_(adminPin)) throw new Error('Set ADMIN_PIN (4 to 8 digits) at the top of Code.gs, then run setup() again.');
    admin.pinHash = hash_(ADMIN_ID, adminPin); admin.updatedAt = now_(); writeRow_('Staff', admin);
  }
  return 'Setup complete';
}

// ---------- seed data (staff, lists and entries imported from the October 2026 timesheets)
var SEED = {"staff":[{"id":"emabuto","name":"Edson Mabuto","position":"Restoration Manager","station":"Quelimane","lineManager":"Senior Project Director","role":"staff","active":true},{"id":"ftavares","name":"Faruk Tavares","position":"Senior Project Director","station":"","lineManager":"","role":"admin","active":true},{"id":"jmafuca","name":"Jorge Mafuca","position":"","station":"","lineManager":"","role":"staff","active":true},{"id":"klevi","name":"Kleber Levi","position":"GIS analist","station":"BF Office, Quelimane and Field","lineManager":"Nelio Sitoe","role":"staff","active":true},{"id":"lalfainho","name":"Lucia Alfainho","position":"","station":"","lineManager":"","role":"staff","active":true},{"id":"lsalgado","name":"Lucia Salgado","position":"Coordenadora de Mobilização Comuniaria","station":"","lineManager":"","role":"staff","active":true},{"id":"nsitoe","name":"Nelio Sitoe","position":"","station":"","lineManager":"","role":"staff","active":true},{"id":"nfuraca","name":"Noca Furaca","position":"Community Engagement, Training and Research Manager","station":"Office and Field","lineManager":"Senior Director of MozBlue Project 1","role":"staff","active":true},{"id":"omuganiua","name":"Odete Muganiua","position":"Community Programs Coordinator","station":"Quelimane","lineManager":"Faruk Tavares","role":"staff","active":true},{"id":"ocumbe","name":"Orcelio Cumbe","position":"Milling Program Coordinator","station":"Quelimane","lineManager":"","role":"staff","active":true},{"id":"pmiguel","name":"Paixao Miguel","position":"Biodiversity Maneger","station":"","lineManager":"","role":"staff","active":true},{"id":"sali","name":"Shees Ali","position":"Carbon Certification Manager","station":"BF Office, Quelimane","lineManager":"Faruk Tavares","role":"staff","active":true}],"entries":[{"staffId":"emabuto","date":"2026-09-16","workType":"Office work","area":"Project Management & Coordination","activity":"","location":"Office","summary":"Mangrove S3 Strategy","id":"seed000"},{"staffId":"emabuto","date":"2026-09-17","workType":"Office work","area":"Project Management & Coordination","activity":"","location":"Office","summary":"Mangrove S3 Strategy","id":"seed001"},{"staffId":"emabuto","date":"2026-09-18","workType":"Office work","area":"Project Management & Coordination","activity":"","location":"Office","summary":"AM Nursery","id":"seed002"},{"staffId":"emabuto","date":"2026-09-21","workType":"Office work","area":"Project Management & Coordination","activity":"","location":"Office","summary":"Mangrove S3 Strategy","id":"seed003"},{"staffId":"emabuto","date":"2026-09-22","workType":"Office work","area":"Project Management & Coordination","activity":"","location":"Office","summary":"Mangrove S3 Strategy","id":"seed004"},{"staffId":"emabuto","date":"2026-09-23","workType":"Office work","area":"Project Management & Coordination","activity":"","location":"Office","summary":"Mangrove S3 Strategy","id":"seed005"},{"staffId":"emabuto","date":"2026-09-24","workType":"Office work","area":"Project Management & Coordination","activity":"","location":"Office","summary":"Mangrove S3 Strategy","id":"seed006"},{"staffId":"emabuto","date":"2026-09-25","workType":"Office work","area":"Project Management & Coordination","activity":"","location":"Office","summary":"Mangrove S3 Strategy","id":"seed007"},{"staffId":"emabuto","date":"2026-09-26","workType":"Office work","area":"","activity":"","location":"Office","summary":"Mangrove S3 Strategy","id":"seed008"},{"staffId":"emabuto","date":"2026-09-27","workType":"Office work","area":"","activity":"","location":"Office","summary":"Mangrove S3 Strategy","id":"seed009"},{"staffId":"emabuto","date":"2026-09-28","workType":"Office work","area":"Project Management & Coordination","activity":"","location":"Office","summary":"MozBlue Season 3 Restoration Strategy and Implementation Plan","id":"seed010"},{"staffId":"emabuto","date":"2026-09-29","workType":"Office work","area":"Project Management & Coordination","activity":"","location":"Office","summary":"MozBlue Season 3 Restoration Strategy and Implementation Plan","id":"seed011"},{"staffId":"emabuto","date":"2026-09-30","workType":"Office work","area":"Project Management & Coordination","activity":"","location":"Office","summary":"MozBlue Season 3 Restoration Strategy and Implementation Plan","id":"seed012"},{"staffId":"emabuto","date":"2026-10-01","workType":"Office work","area":"Project Management & Coordination","activity":"","location":"Office","summary":"MozBlue Season 3 Restoration Strategy and Implementation Plan","id":"seed013"},{"staffId":"emabuto","date":"2026-10-02","workType":"Office work","area":"Project Management & Coordination","activity":"","location":"Office","summary":"MozBlue Season 3 Restoration Strategy and Implementation Plan","id":"seed014"},{"staffId":"emabuto","date":"2026-10-03","workType":"Office work","area":"","activity":"","location":"Office","summary":"MozBlue Season 3 Restoration Strategy and Implementation Plan","id":"seed015"},{"staffId":"emabuto","date":"2026-10-04","workType":"Office work","area":"","activity":"","location":"Office","summary":"MozBlue Season 3 Restoration Strategy and Implementation Plan","id":"seed016"},{"staffId":"emabuto","date":"2026-10-05","workType":"Public holiday","area":"Project Management & Coordination","activity":"","location":"Office","summary":"MozBlue Season 3 Restoration Strategy and Implementation Plan","id":"seed017"},{"staffId":"emabuto","date":"2026-10-06","workType":"Office work","area":"Project Management & Coordination","activity":"","location":"Office","summary":"MozBlue Season 3 Restoration Strategy and Implementation Plan","id":"seed018"},{"staffId":"emabuto","date":"2026-10-07","workType":"Office work","area":"Project Management & Coordination","activity":"","location":"Office","summary":"MozBlue Season 3 Restoration Strategy and Implementation Plan","id":"seed019"},{"staffId":"emabuto","date":"2026-10-08","workType":"Fieldwork","area":"ANR – Assisted Natural Regeneration","activity":"","location":"Mucussa","summary":"Carbon Field Study for Silvestrum","id":"seed020"},{"staffId":"emabuto","date":"2026-10-09","workType":"Office work","area":"Project Management & Coordination","activity":"","location":"Office","summary":"ANR Field Validation Plan","id":"seed021"},{"staffId":"emabuto","date":"2026-10-12","workType":"Office work","area":"Project Management & Coordination","activity":"","location":"Office","summary":"Terms of Reference and Supporting Information – ANR Technical Assessment","id":"seed022"},{"staffId":"emabuto","date":"2026-10-13","workType":"Fieldwork","area":"ANR – Assisted Natural Regeneration","activity":"","location":"Mucussa","summary":"Carbon Field Study for Silvestrum","id":"seed023"},{"staffId":"emabuto","date":"2026-10-14","workType":"Office work","area":"Project Management & Coordination","activity":"","location":"Office","summary":"Revised ANR Field Validation SOP","id":"seed024"},{"staffId":"emabuto","date":"2026-10-15","workType":"Office work","area":"Project Management & Coordination","activity":"","location":"Office","summary":"Revised ANR Field Validation SOP","id":"seed025"},{"staffId":"klevi","date":"2026-10-05","workType":"Office work","area":"GIS, Mapping & Data Management","activity":"","location":"BF office","summary":"Prepared the sampling points required for 6-month monitoring assessments (March and April)","id":"seed026"},{"staffId":"klevi","date":"2026-10-06","workType":"Office work","area":"GIS, Mapping & Data Management","activity":"","location":"BF office","summary":"Updated the field monitoring form in Survey123 for data collection and Produced orientation and field navigation maps to guide monitors during field activities.","id":"seed027"},{"staffId":"klevi","date":"2026-10-07","workType":"Office work","area":"GIS, Mapping & Data Management","activity":"","location":"BF office","summary":"Developed the Survey123 form for Assisted Natural Regeneration (ANR) and extracted monitoring points with the highest natural regeneration rates of Avicennia marina (AM) and Rhizophora mucronata (RM) to support the identification and selection of suitable sites for ANR interventions and Produced orientation and field navigation maps to guide monitors during field activities.","id":"seed028"},{"staffId":"klevi","date":"2026-10-08","workType":"Fieldwork","area":"ANR – Assisted Natural Regeneration","activity":"","location":"Mucussa","summary":"Participation in ANR training, field drone survey support for CNN visit, and completion of the interview session at the BF office.","id":"seed029"},{"staffId":"klevi","date":"2026-10-09","workType":"Office work","area":"Monitoring, Evaluation & Restoration Monitoring","activity":"","location":"BF office","summary":"Distribution of sampling points to the field team for march planted areas monitoring.","id":"seed030"},{"staffId":"klevi","date":"2026-10-12","workType":"Office work","area":"GIS, Mapping & Data Management","activity":"","location":"BF office","summary":"Processing of field data collected on Saturday and update the Survey123 form for ANR data collection","id":"seed031"},{"staffId":"klevi","date":"2026-10-13","workType":"Office work","area":"GIS, Mapping & Data Management","activity":"","location":"BF office","summary":"Processing of field data collected on Monday, creation of a visualization map","id":"seed032"},{"staffId":"klevi","date":"2026-10-14","workType":"Office work","area":"","activity":"","location":"BF office","summary":"","id":"seed033"},{"staffId":"lsalgado","date":"2026-10-05","workType":"Fieldwork","area":"Visits, Missions & Events","activity":"","location":"Idugo","summary":"","id":"seed034"},{"staffId":"lsalgado","date":"2026-10-06","workType":"Fieldwork","area":"Visits, Missions & Events","activity":"","location":"Mucussa","summary":"","id":"seed035"},{"staffId":"lsalgado","date":"2026-10-07","workType":"Fieldwork","area":"Other","activity":"","location":"Office","summary":"","id":"seed036"},{"staffId":"lsalgado","date":"2026-10-12","workType":"Office work","area":"Other","activity":"","location":"Office","summary":"Meeting e plano semanal","id":"seed037"},{"staffId":"lsalgado","date":"2026-10-13","workType":"Office work","area":"Other","activity":"","location":"Office","summary":"preparação de cartas e termos de consentimentos para Marrongane e Zalala","id":"seed038"},{"staffId":"lsalgado","date":"2026-10-14","workType":"Office work","area":"Other","activity":"","location":"Office","summary":"Revisção de SOP e partilha","id":"seed039"},{"staffId":"lsalgado","date":"2026-10-15","workType":"Fieldwork","area":"Other","activity":"","location":"Idugo","summary":"Actividades de Clube escolar dia de Lavagem das maos e importancia do Mangal como filtro","id":"seed040"},{"staffId":"lsalgado","date":"2026-10-15","workType":"Fieldwork","area":"Other","activity":"","location":"Zalala","summary":"Sensibilização e assinatura de cartas de consentimentos Zalala (Tate e Safuri)","id":"seed041"},{"staffId":"nfuraca","date":"2026-10-06","workType":"Fieldwork","area":"Visits, Missions & Events","activity":"","location":"Mucussa","summary":"Visit Beekiping in Mucussa with CNN. Another activity was rekated Naminane. Community meeting to sign the minute of meeting previous happened in community.","id":"seed042"},{"staffId":"nfuraca","date":"2026-10-07","workType":"Fieldwork","area":"","activity":"","location":"","summary":"","id":"seed043"},{"staffId":"nfuraca","date":"2026-10-08","workType":"Fieldwork","area":"Training & Capacity Building","activity":"","location":"Field - Mucussa","summary":"2 Plot was used for training to sampling ANR and Carbon sampling","id":"seed044"},{"staffId":"nfuraca","date":"2026-10-09","workType":"Office work","area":"ANR – Assisted Natural Regeneration","activity":"","location":"At office","summary":"Planing to restart field work of ANR and Carbon assessment","id":"seed045"},{"staffId":"nfuraca","date":"2026-10-12","workType":"Office work","area":"ANR – Assisted Natural Regeneration","activity":"","location":"BF Office","summary":"Preparing field work of ANR and Carbon sampling. The team is ready to work during next 4 days.","id":"seed046"},{"staffId":"nfuraca","date":"2026-10-13","workType":"Fieldwork","area":"ANR – Assisted Natural Regeneration","activity":"","location":"Field - Mucussa","summary":"","id":"seed047"},{"staffId":"ocumbe","date":"2026-10-13","workType":"Office work","area":"Social Program Implementation","activity":"","location":"Office","summary":"COAST meeting with BF, elaborate rice milling article for Mozambique club","id":"seed050"},{"staffId":"ocumbe","date":"2026-10-14","workType":"Office work","area":"Social Program Implementation","activity":"","location":"Office","summary":"Elaboracao de relatorio de treinamento de operador de Moagem em Mussere","id":"seed051"},{"staffId":"ocumbe","date":"2026-10-15","workType":"Office work","area":"Social Program Implementation","activity":"","location":"Office","summary":"Elaboracao de relatorio de treinamento de operador de Moagem em Mussere","id":"seed048"},{"staffId":"ocumbe","date":"2026-10-15","workType":"Office work","area":"Social Program Implementation","activity":"","location":"Office","summary":"Elaboracao de relatorio de treinamento de operador de Moagem em Mussere","id":"seed052"},{"staffId":"ocumbe","date":"2026-10-16","workType":"Fieldwork","area":"Social Program Implementation","activity":"","location":"Office and Mucai","summary":"Elaboracao de relatorio de treinamento de operador de Moagem em Mussere","id":"seed049"},{"staffId":"omuganiua","date":"2026-09-16","workType":"Office work","area":"Project Management & Coordination","activity":"","location":"Office","summary":"","id":"seed053"},{"staffId":"omuganiua","date":"2026-09-17","workType":"Office work","area":"Project Management & Coordination","activity":"","location":"Office","summary":"","id":"seed054"},{"staffId":"omuganiua","date":"2026-09-18","workType":"Office work","area":"Project Management & Coordination","activity":"","location":"Office","summary":"","id":"seed055"},{"staffId":"omuganiua","date":"2026-09-21","workType":"Annual leave","area":"","activity":"","location":"","summary":"","id":"seed056"},{"staffId":"omuganiua","date":"2026-09-22","workType":"Annual leave","area":"","activity":"","location":"","summary":"","id":"seed057"},{"staffId":"omuganiua","date":"2026-09-23","workType":"Annual leave","area":"","activity":"","location":"","summary":"","id":"seed058"},{"staffId":"omuganiua","date":"2026-09-24","workType":"Annual leave","area":"","activity":"","location":"","summary":"","id":"seed059"},{"staffId":"omuganiua","date":"2026-09-25","workType":"Annual leave","area":"","activity":"","location":"","summary":"","id":"seed060"},{"staffId":"omuganiua","date":"2026-09-28","workType":"Annual leave","area":"","activity":"","location":"","summary":"","id":"seed061"},{"staffId":"omuganiua","date":"2026-09-29","workType":"Annual leave","area":"","activity":"","location":"","summary":"","id":"seed062"},{"staffId":"omuganiua","date":"2026-09-30","workType":"Annual leave","area":"","activity":"","location":"","summary":"","id":"seed063"},{"staffId":"omuganiua","date":"2026-10-01","workType":"Annual leave","area":"","activity":"","location":"","summary":"","id":"seed064"},{"staffId":"omuganiua","date":"2026-10-02","workType":"Annual leave","area":"","activity":"","location":"","summary":"","id":"seed065"},{"staffId":"omuganiua","date":"2026-10-05","workType":"Annual leave","area":"","activity":"","location":"","summary":"","id":"seed066"},{"staffId":"omuganiua","date":"2026-10-06","workType":"Annual leave","area":"","activity":"","location":"","summary":"","id":"seed067"},{"staffId":"omuganiua","date":"2026-10-07","workType":"Annual leave","area":"","activity":"","location":"","summary":"","id":"seed068"},{"staffId":"omuganiua","date":"2026-10-08","workType":"Annual leave","area":"","activity":"","location":"","summary":"","id":"seed069"},{"staffId":"omuganiua","date":"2026-10-09","workType":"Office work","area":"Project Management & Coordination","activity":"","location":"Office","summary":"","id":"seed070"},{"staffId":"omuganiua","date":"2026-10-12","workType":"Office work","area":"Project Management & Coordination","activity":"","location":"Office","summary":"Weekly staff meeting, COAST Quarterly Report and PSD","id":"seed071"},{"staffId":"omuganiua","date":"2026-10-13","workType":"Office work","area":"Project Management & Coordination","activity":"","location":"Office","summary":"COAST check-in meeting Quarterly Report and PSD","id":"seed072"},{"staffId":"omuganiua","date":"2026-10-14","workType":"Office work","area":"Project Management & Coordination","activity":"","location":"Office","summary":"COAST Quarterly Report and PSD Submission","id":"seed073"},{"staffId":"omuganiua","date":"2026-10-15","workType":"Fieldwork","area":"Community Engagement","activity":"","location":"Mudenga","summary":"Community Consultation","id":"seed074"},{"staffId":"sali","date":"2026-10-02","workType":"Office work","area":"Project Management & Coordination","activity":"","location":"BF Office","summary":"Finalizing the Season 3 restoration model and strategy document","id":"seed075"},{"staffId":"sali","date":"2026-10-03","workType":"Office work","area":"Project Management & Coordination","activity":"","location":"BF Office","summary":"Finalizing the Season 3 restoration model and strategy document","id":"seed076"},{"staffId":"sali","date":"2026-10-04","workType":"Office work","area":"Project Management & Coordination","activity":"","location":"BF Office","summary":"Finalizing the Season 3 restoration model and strategy document","id":"seed077"},{"staffId":"sali","date":"2026-10-05","workType":"Office work","area":"ANR – Assisted Natural Regeneration","activity":"","location":"BF Office","summary":"Restoration Strategy corrections and review with Removall","id":"seed078"},{"staffId":"sali","date":"2026-10-06","workType":"Office work","area":"Project Management & Coordination","activity":"","location":"BF Office","summary":"ANR chronosequence field methodology design","id":"seed079"},{"staffId":"sali","date":"2026-10-07","workType":"Office work","area":"Project Management & Coordination","activity":"","location":"BF Office","summary":"ANR chronosequence field methodology design","id":"seed080"},{"staffId":"sali","date":"2026-10-08","workType":"Fieldwork","area":"ANR – Assisted Natural Regeneration","activity":"","location":"Mucussa","summary":"ANR chronosequence field survey","id":"seed081"},{"staffId":"sali","date":"2026-10-09","workType":"Office work","area":"Project Management & Coordination","activity":"","location":"BF Office","summary":"ANR chronosequence field methodology design","id":"seed082"},{"staffId":"sali","date":"2026-10-10","workType":"Office work","area":"Project Management & Coordination","activity":"","location":"BF Office","summary":"ANR chronosequence field methodology design","id":"seed083"}],"lists":{"workTypes":[{"name":"Office work","working":true},{"name":"Fieldwork","working":true},{"name":"Trip","working":true},{"name":"Days off","working":false},{"name":"Annual leave","working":false},{"name":"Excused absence","working":false},{"name":"Public holiday","working":false},{"name":"Weekly rest","working":false}],"areas":[{"name":"Community Engagement","activities":["Community engagement","Community mobilization","Community consultation","Community awareness and sensitization","Community meeting","Community leadership meeting","Community validation meeting","Community grievance management","Community conflict resolution","Community feedback session","Community agreement preparation and signing","Community organization/committee establishment","Community committee capacity building","Community field visit","Community activity monitoring","Community stakeholder mapping"]},{"name":"Stakeholder & Institutional Engagement","activities":["Stakeholder consultation","Stakeholder meeting","Government stakeholder meeting","Government coordination meeting","Government official visit","District Government visit","Provincial Government visit","Partner meeting","Partner visit","Investor/Donor visit","Board visit","Technical partner visit","Service provider meeting","Technical meeting with service provider","Coordination meeting with implementation partner","Workshop","Blue Forest workshop","Technical workshop","Planning workshop","Validation workshop"]},{"name":"ANR – Assisted Natural Regeneration","activities":["ANR desktop assessment","ANR site identification","ANR field assessment","ANR site validation","ANR baseline assessment","ANR mapping","ANR polygon delineation","ANR community consultation","ANR intervention planning","ANR implementation","ANR hydrological assessment","ANR hydrological intervention","ANR access/control intervention","ANR protection activities","ANR natural recruitment assessment","ANR field supervision","ANR monitoring","ANR survival/recruitment monitoring","ANR effectiveness assessment","ANR data collection","ANR reporting"]},{"name":"Planting Operation","activities":["Planting planning","Planting site assessment","Planting site selection","Planting site validation","Planting area mapping","Micro-polygon mapping","Planting readiness assessment","Propagule availability assessment","Propagule collection","Propagule procurement","Propagule transportation","Propagule distribution","Propagule storage and handling","Planting team mobilization","Planting team training","Planting operation","Planting supervision","Planting quality control","Planting density assessment","Planting progress assessment","Daily planting verification","Planted area verification","Planting completion assessment","Planting data collection","Planting reporting"]},{"name":"Replanting / Refilling Operation","activities":["Replanting area assessment","Replanting eligibility assessment","Low-density polygon identification","Replanting area prioritization","Replanting field validation","Replanting planning","Refilling planning","Replanting propagule collection","Replanting propagule transportation","Replanting operation","Refilling operation","Replanting supervision","Replanting quality control","Replanting density verification","Replanting monitoring","Replanting completion assessment","Replanting data collection","Replanting reporting"]},{"name":"Avicennia marina Pilot Project","activities":["Avicennia pilot planning","Avicennia site assessment","Avicennia site selection","Avicennia site validation","Avicennia baseline assessment","Avicennia hydrological assessment","Avicennia seed availability assessment","Avicennia seed collection","Avicennia seed handling and storage","Avicennia nursery establishment","Community-based Avicennia seedling production","Avicennia seedling production","Avicennia nursery monitoring","Avicennia seedling quality assessment","Avicennia seedling transportation","Avicennia direct seeding","Avicennia seedling planting","Avicennia pilot implementation","Avicennia planting supervision","Avicennia survival monitoring","Avicennia performance assessment","Avicennia pilot evaluation","Avicennia lessons-learned assessment","Avicennia pilot reporting"]},{"name":"Mangrove Nursery","activities":["Nursery potential assessment","Nursery site identification","Nursery site validation","Nursery planning","Nursery construction","Nursery material procurement","Nursery equipment installation","Nursery operator training","Seed/propagule collection","Seedling production","Seedling maintenance","Nursery supervision","Nursery monitoring","Seedling quality control","Seedling inventory","Nursery production assessment","Nursery completion report"]},{"name":"Social Program Implementation","activities":["Social program planning","Social program community consultation","Social program implementation","Social program supervision","Social program field visit","Social program monitoring","Social program beneficiary registration","Social program beneficiary validation","Social program training","Social program distribution","Social program impact assessment","Social program reporting"]},{"name":"Livelihoods Implementation","activities":["Livelihoods baseline assessment","Livelihoods community consultation","Livelihoods opportunity assessment","Livelihoods value-chain assessment","Livelihoods beneficiary identification","Livelihoods beneficiary validation","Livelihoods implementation","Livelihoods training","Livelihoods business development support","Livelihoods technical assistance","Livelihoods equipment distribution","Livelihoods supervision","Livelihoods field visit","Livelihoods monitoring","Livelihoods performance assessment","Livelihoods income monitoring","Livelihoods impact assessment","Livelihoods reporting"]},{"name":"Specific Livelihood / IGA Activities","activities":["Beekeeping implementation","Beekeeping monitoring","Honey value-chain development","Salt value-chain assessment","Salt business implementation","Salt Committee establishment","Salt Committee training","Salt procurement","Salt business monitoring","Salt business/accountability training","Milling program implementation","Solar mill construction/installation","Mill operator training","Milling business monitoring","Bamboo nursery implementation","Bamboo nursery monitoring","Bamboo value-chain development","Savings group implementation","Savings group monitoring","Crab aquaculture assessment","Crab aquaculture implementation"]},{"name":"Monitoring, Evaluation & Restoration Monitoring","activities":["Monitoring of planted areas","Restoration monitoring","Survival assessment","Plant density assessment","Permanent monitoring plot establishment","Monitoring plot assessment","Baseline monitoring","Post-planting monitoring","Replanting monitoring","ANR monitoring","Nursery monitoring","Livelihoods monitoring","Social program monitoring","Biodiversity monitoring","Environmental monitoring","Social monitoring","Field data collection","GPS data collection","Geotagged photo collection","Monitoring data validation","Monitoring data analysis","Monitoring report preparation","Monitoring field supervision"]},{"name":"GIS, Mapping & Data Management","activities":["GIS mapping","Desktop spatial analysis","Satellite imagery analysis","Drone/imagery assessment","Restoration area mapping","Planting polygon mapping","Micro-polygon mapping","Field polygon validation","GPS track collection","GIS data validation","GIS QA/QC","Spatial database update","ArcGIS data management","Map preparation","Area calculation and verification"]},{"name":"Biodiversity","activities":["Biodiversity baseline assessment","Biodiversity field survey","Biodiversity monitoring","Species identification","Habitat assessment","Biodiversity data collection","Biodiversity stakeholder consultation","Biodiversity management planning","Biodiversity field visit","Biodiversity reporting"]},{"name":"Procurement & Logistics","activities":["Procurement planning","Request for quotation","Quotation assessment","Supplier identification","Supplier evaluation","Procurement","Purchase of materials","Purchase of equipment","Purchase of restoration inputs","Purchase of nursery materials","Purchase of livelihood equipment","Material delivery","Material distribution","Logistics planning","Transport coordination","Boat logistics","Vehicle logistics","Equipment mobilization","Asset management","Inventory management","Fuel management"]},{"name":"Service Provider Management","activities":["Service provider mobilization","Service provider coordination","Service provider meeting","Technical meeting with service provider","Service provider field supervision","Service provider performance monitoring","Service provider work verification","Service provider deliverable review","Service provider invoice verification","Service provider compliance assessment","Contractor performance assessment"]},{"name":"Project Management & Coordination","activities":["Project planning","Operational planning","Weekly planning","Monthly planning","Workplan preparation","Budget planning","Budget review","Operational coordination","Internal coordination meeting","Technical coordination meeting","Management meeting","Team meeting","Field team coordination","Project progress review","Project performance review","Risk assessment","Risk management","Issue resolution","Project reporting","Management reporting","Board reporting","Donor/investor reporting"]},{"name":"Training & Capacity Building","activities":["Community training","Staff training","Restoration training","Planting training","Nursery training","Monitoring training","Livelihoods training","Business management training","Bookkeeping/accountability training","Health and safety training","Equipment operation training","Service provider training","Refresher training","Training evaluation"]},{"name":"Visits, Missions & Events","activities":["Field visit","Technical field visit","Community visit","Government official visit","Government visit","Partner visit","Service provider visit","Investor visit","Donor visit","Board visit","Monitoring mission","Technical mission","Due-diligence visit","External stakeholder visit","Project demonstration visit","Community event","Project event"]},{"name":"Environmental & Social / Safeguards","activities":["Environmental assessment","Social assessment","Environmental and Social monitoring","Safeguards assessment","Community grievance follow-up","Health and safety inspection","Environmental compliance monitoring","Social compliance monitoring","Incident assessment","Stakeholder engagement documentation"]},{"name":"Administrative & Financial Activities","activities":["Cash advance preparation","Cash advance management","Expense justification","Invoice review","Payment verification","Financial documentation","Budget monitoring","Cost tracking","Financial reporting","Contract administration","Documentation and filing"]},{"name":"Finance","activities":[]},{"name":"Project Management","activities":[]},{"name":"Government","activities":[]},{"name":"Auditing","activities":[]},{"name":"Donor","activities":[]},{"name":"Other","activities":["Travel","Other (describe in summary)"]}],"periodStartDay":15}};

return {doPost: doPost, setupWith_: setupWith_};
};

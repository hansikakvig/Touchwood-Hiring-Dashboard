/**
 * Touchwood Hiring Dashboard — Google Apps Script backend
 * --------------------------------------------------------
 * 1. Change FIRST_ADMIN_PASSWORD below (and LOGO_FILE_ID if you have the logo in Drive).
 * 2. Run setup() once from the editor and approve the permissions.
 * 3. Deploy > New deployment > Web app (Execute as: Me, Who has access: Anyone).
 */

const CONFIG = {
  APP_TITLE: 'Touchwood Hiring',
  FIRST_ADMIN_USERNAME: 'admin',
  FIRST_ADMIN_PASSWORD: 'BeePositive',   // change this before running setup()
  FIRST_ADMIN_NAME: 'HR Admin',
  LOGO_FILE_ID: '',                         // Drive file ID or link of the Touchwood World logo (PNG/JPG/SVG)
  SESSION_MINUTES: 30,                      // sign out after this many minutes without activity
  SYNC_EVERY_MINUTES: 10,                   // 1, 5, 10, 15 or 30
  MAX_LOGIN_ATTEMPTS: 5,                    // then locked for 15 minutes
  // Openings created on first setup, linked to the application sheets already in Drive.
  SEED_JOBS: [
    { business: 'lifespaces', title: 'Business Head', location: 'Nashik', sheetId: '1-IN1FQ5ss8ZZY6nAlBARyGoet7JA-KAqWn1RKvXEMec',
      applyLink: 'https://script.google.com/macros/s/AKfycbzEKbYKvTHXaAbjQDleTwyzU2riRf1pXCgWJocVS_f2InGTQa8oC3O5Zt3bhVXtRVAZhA/exec' },
    { business: 'bliss', title: 'General Manager', location: 'Nashik', sheetId: '1xAUpN75iYk9Z_uoeKjnEhuAbspGh3E013SDN0bfpAWw' },
    { business: 'bliss', title: 'Business Development Manager (Hospitality Sales)', department: 'Sales', location: 'Nashik', sheetId: '1C1vtzSMRJTJDqdkZGCgWKhtLVb8ACNTGGkUvIjdPC8Y' },
    { business: 'bliss', title: 'Wedding Sales Manager', department: 'Sales', location: 'Nashik', sheetId: '1eXea8hU2dIbUVx1e_2uZZc3YExHA7tQcFfMe-PNtgow' }
  ]
};

const SHEETS = {
  admins: ['id', 'username', 'name', 'role', 'salt', 'hash', 'active', 'createdAt', 'lastLogin'],
  jobs: ['id', 'business', 'title', 'department', 'location', 'vacancies', 'experience', 'education', 'employmentType', 'salary', 'status',
    'description', 'skills', 'applyLink', 'sheetId', 'createdAt', 'updatedAt'],
  candidates: ['id', 'business', 'jobId', 'jobTitle', 'name', 'mobile', 'email', 'experience', 'location', 'appliedDate', 'source', 'status',
    'resumeName', 'resumeLink', 'notes', 'details', 'sheetId', 'sourceKey', 'appliedAt', 'deleted', 'createdAt', 'updatedAt'],
  interviews: ['id', 'candidateId', 'business', 'jobId', 'date', 'time', 'round', 'mode', 'interviewer', 'venue', 'status', 'notes',
    'feedback', 'rating', 'history', 'createdAt', 'updatedAt'],
  settings: ['key', 'value']
};
const JSON_FIELDS = { candidates: ['details'], interviews: ['history'] };
const EDITABLE = {
  jobs: SHEETS.jobs.filter(function (k) { return k !== 'id'; }),
  candidates: SHEETS.candidates.filter(function (k) { return ['id', 'sourceKey', 'deleted'].indexOf(k) < 0; }),
  interviews: SHEETS.interviews.filter(function (k) { return k !== 'id'; })
};

/* =========================== web app =========================== */

function doGet() {
  const t = HtmlService.createTemplateFromFile('Index');
  t.logo = getLogoDataUri_();
  t.appTitle = CONFIG.APP_TITLE;
  return t.evaluate()
    .setTitle(CONFIG.APP_TITLE)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
}

/* =========================== setup =========================== */

function setup() {
  const props = PropertiesService.getScriptProperties();
  let ss;
  const existing = props.getProperty('DB_ID');
  if (existing) {
    ss = SpreadsheetApp.openById(existing);
  } else {
    ss = SpreadsheetApp.create('Touchwood Hiring Database');
    props.setProperty('DB_ID', ss.getId());
  }
  Object.keys(SHEETS).forEach(function (name) {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    const cols = SHEETS[name];
    sh.getRange(1, 1, 1, cols.length).setValues([cols]).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.getRange(1, 1, sh.getMaxRows(), cols.length).setNumberFormat('@');
  });
  const def = ss.getSheetByName('Sheet1');
  if (def && ss.getSheets().length > 1) ss.deleteSheet(def);

  if (!props.getProperty('RESUME_FOLDER_ID')) {
    props.setProperty('RESUME_FOLDER_ID', DriveApp.createFolder('Touchwood Hiring Resumes').getId());
  }

  if (!readAll_('admins').length) {
    const salt = Utilities.getUuid();
    upsert_('admins', {
      id: newId_(), username: CONFIG.FIRST_ADMIN_USERNAME.toLowerCase(), name: CONFIG.FIRST_ADMIN_NAME, role: 'Admin',
      salt: salt, hash: hashPw_(CONFIG.FIRST_ADMIN_PASSWORD, salt), active: 'TRUE', createdAt: nowIso_(), lastLogin: ''
    });
  }

  if (!readAll_('jobs').length) {
    CONFIG.SEED_JOBS.forEach(function (j) {
      upsert_('jobs', Object.assign({
        id: newId_(), department: '', vacancies: '1', experience: '', education: '', employmentType: 'Full-time', salary: '',
        status: 'Active', description: '', skills: '', applyLink: '', createdAt: nowIso_(), updatedAt: nowIso_()
      }, j));
    });
  }

  if (CONFIG.LOGO_FILE_ID && !getSetting_('logoFileId', '')) setSetting_('logoFileId', fileIdFrom_(CONFIG.LOGO_FILE_ID));

  ScriptApp.getProjectTriggers().forEach(function (tr) {
    if (tr.getHandlerFunction() === 'syncApplications') ScriptApp.deleteTrigger(tr);
  });
  ScriptApp.newTrigger('syncApplications').timeBased().everyMinutes(CONFIG.SYNC_EVERY_MINUTES).create();

  syncApplications();
  Logger.log('Setup done. Database: ' + ss.getUrl());
  Logger.log('Sign in with username "' + CONFIG.FIRST_ADMIN_USERNAME + '" and the password in CONFIG, then change it.');
}

/* =========================== auth =========================== */

function login(username, password) {
  try {
    username = String(username || '').trim().toLowerCase();
    if (!username || !password) return { error: 'BAD_LOGIN', message: 'Enter your username and password.' };
    const cache = CacheService.getScriptCache();
    const failKey = 'fail_' + username;
    const fails = Number(cache.get(failKey) || 0);
    if (fails >= CONFIG.MAX_LOGIN_ATTEMPTS) {
      return { error: 'LOCKED', message: 'Too many wrong attempts. Try again in 15 minutes, or ask an admin to reset your password.' };
    }
    const u = readAll_('admins').filter(function (a) { return String(a.username).toLowerCase() === username; })[0];
    if (!u || u.active !== 'TRUE' || hashPw_(String(password), u.salt) !== u.hash) {
      cache.put(failKey, String(fails + 1), 900);
      Utilities.sleep(500);
      return { error: 'BAD_LOGIN', message: 'Wrong username or password.' };
    }
    cache.remove(failKey);
    const token = (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, '');
    const sess = { id: u.id, username: u.username, name: u.name, role: u.role };
    cache.put('sess_' + token, JSON.stringify(sess), CONFIG.SESSION_MINUTES * 60);
    rememberToken_(u.id, token);
    patchRow_('admins', u.id, { lastLogin: nowIso_() });
    return { ok: true, token: token, data: snapshot_(sess) };
  } catch (e) {
    return { error: 'SERVER', message: String(e && e.message || e) };
  }
}

function logout(token) {
  if (token) CacheService.getScriptCache().remove('sess_' + token);
  return { ok: true };
}

function changePassword(token, current, next) {
  return call_(token, false, function (s) {
    const u = findRow_('admins', s.id);
    if (!u || hashPw_(String(current || ''), u.salt) !== u.hash) return { error: 'BAD_PASSWORD', message: 'Your current password is wrong.' };
    const err = pwProblem_(next);
    if (err) return { error: 'WEAK', message: err };
    const salt = Utilities.getUuid();
    patchRow_('admins', s.id, { salt: salt, hash: hashPw_(String(next), salt) });
    return { ok: true };
  });
}

/* =========================== data API =========================== */

function getData(token) {
  return call_(token, false, function (s) { return snapshot_(s); });
}

function saveRecord(token, col, id, data) {
  return call_(token, false, function (s) {
    checkCol_(col);
    const now = nowIso_();
    const prev = id ? findRow_(col, id) : null;
    const obj = prev ? Object.assign({}, prev) : { id: newId_(), createdAt: now, deleted: '' };
    EDITABLE[col].forEach(function (k) { if (data && data[k] !== undefined) obj[k] = data[k]; });
    if (!prev && col === 'candidates') obj.sourceKey = '';
    obj.updatedAt = now;
    if (!obj.createdAt) obj.createdAt = now;
    withLock_(function () { upsert_(col, obj); });
    return { ok: true, id: obj.id, data: snapshot_(s) };
  });
}

function patchRecord(token, col, id, data) {
  return call_(token, false, function (s) {
    checkCol_(col);
    const clean = {};
    EDITABLE[col].forEach(function (k) { if (data && data[k] !== undefined) clean[k] = data[k]; });
    clean.updatedAt = nowIso_();
    withLock_(function () { patchRow_(col, id, clean); });
    return { ok: true, data: snapshot_(s) };
  });
}

function deleteRecord(token, col, id) {
  return call_(token, false, function (s) {
    checkCol_(col);
    withLock_(function () {
      if (col === 'candidates') {
        readAll_('interviews').filter(function (i) { return i.candidateId === id; })
          .forEach(function (i) { deleteRow_('interviews', i.id); });
        const c = findRow_('candidates', id);
        // Imported candidates are hidden, not erased, so the next sync does not bring them back.
        if (c && c.sourceKey) patchRow_('candidates', id, { deleted: 'TRUE', updatedAt: nowIso_() });
        else deleteRow_('candidates', id);
      } else {
        deleteRow_(col, id);
      }
    });
    return { ok: true, data: snapshot_(s) };
  });
}

function bulkAddCandidates(token, rows) {
  return call_(token, false, function (s) {
    const now = nowIso_();
    const objs = (rows || []).slice(0, 2000).map(function (r) {
      const o = { id: newId_(), createdAt: now, updatedAt: now, deleted: '', sourceKey: '' };
      EDITABLE.candidates.forEach(function (k) { if (r[k] !== undefined) o[k] = r[k]; });
      return o;
    });
    withLock_(function () { appendRows_('candidates', objs); });
    return { ok: true, added: objs.length, data: snapshot_(s) };
  });
}

function saveSettings(token, cfg) {
  return call_(token, false, function (s) {
    const cur = getSetting_('config', {});
    const next = {
      businesses: (cfg && cfg.businesses) || cur.businesses || [],
      interviewers: (cfg && cfg.interviewers) || cur.interviewers || [],
      departments: (cfg && cfg.departments) || cur.departments || []
    };
    withLock_(function () { setSetting_('config', next); });
    return { ok: true, data: snapshot_(s) };
  });
}

function saveLogo(token, link) {
  return call_(token, true, function () {
    const id = fileIdFrom_(link);
    if (link && !id) return { error: 'BAD_LINK', message: 'That doesn’t look like a Google Drive file link.' };
    if (id) {
      const blob = DriveApp.getFileById(id).getBlob();
      if (!/^image\//.test(blob.getContentType())) return { error: 'BAD_TYPE', message: 'The logo must be an image file (PNG, JPG or SVG).' };
      if (blob.getBytes().length > 1500000) return { error: 'TOO_BIG', message: 'The logo file is over 1.5 MB. Upload a smaller PNG.' };
    }
    setSetting_('logoFileId', id || '');
    return { ok: true, logo: getLogoDataUri_() };
  });
}

function uploadResume(token, file) {
  return call_(token, false, function () {
    if (!file || !file.data) return { error: 'NO_FILE', message: 'No file received.' };
    const bytes = Utilities.base64Decode(file.data);
    if (bytes.length > 10 * 1024 * 1024) return { error: 'TOO_BIG', message: 'The resume is over 10 MB.' };
    const folder = DriveApp.getFolderById(PropertiesService.getScriptProperties().getProperty('RESUME_FOLDER_ID'));
    const f = folder.createFile(Utilities.newBlob(bytes, file.type || 'application/octet-stream', String(file.name || 'resume')));
    try { f.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) { /* domain policy may block this */ }
    return { ok: true, url: f.getUrl(), name: f.getName() };
  });
}

function findSheets(token) {
  return call_(token, false, function () {
    const q = "mimeType = 'application/vnd.google-apps.spreadsheet' and trashed = false and " +
      "(title contains 'Application' or title contains 'Applicant' or title contains 'Candidate' or title contains 'Hiring')";
    const it = DriveApp.searchFiles(q);
    const out = [];
    while (it.hasNext() && out.length < 30) {
      const f = it.next();
      out.push({ id: f.getId(), title: f.getName(), modifiedTime: Utilities.formatDate(f.getLastUpdated(), 'UTC', "yyyy-MM-dd'T'HH:mm:ss'Z'") });
    }
    return { ok: true, files: out };
  });
}

function syncNow(token) {
  return call_(token, false, function (s) {
    const r = syncApplications();
    return { ok: true, added: r ? r.added : 0, busy: !r, data: snapshot_(s) };
  });
}

/* =========================== user management (Admin only) =========================== */

function saveUser(token, u) {
  return call_(token, true, function (s) {
    u = u || {};
    const username = String(u.username || '').trim().toLowerCase();
    if (!/^[a-z0-9._-]{3,40}$/.test(username)) return { error: 'BAD_USERNAME', message: 'Username needs 3–40 letters, numbers, dots, dashes or underscores.' };
    const all = readAll_('admins');
    if (all.some(function (a) { return a.username === username && a.id !== u.id; })) return { error: 'TAKEN', message: 'That username is already used.' };
    const role = u.role === 'Admin' ? 'Admin' : 'HR';
    const prev = u.id ? findRow_('admins', u.id) : null;
    if (prev && prev.role === 'Admin' && (role !== 'Admin' || u.active === 'FALSE') && activeAdmins_(all, prev.id) === 0) {
      return { error: 'LAST_ADMIN', message: 'Keep at least one active admin.' };
    }
    const obj = prev ? Object.assign({}, prev) : { id: newId_(), createdAt: nowIso_(), lastLogin: '', active: 'TRUE' };
    obj.username = username;
    obj.name = String(u.name || username).trim();
    obj.role = role;
    if (u.active === 'TRUE' || u.active === 'FALSE') obj.active = u.active;
    if (u.password || !prev) {
      const err = pwProblem_(u.password);
      if (err) return { error: 'WEAK', message: err };
      obj.salt = Utilities.getUuid();
      obj.hash = hashPw_(String(u.password), obj.salt);
    }
    withLock_(function () { upsert_('admins', obj); });
    if (prev && (obj.active === 'FALSE' || u.password || obj.role !== prev.role)) dropSessions_(obj.id);
    return { ok: true, data: snapshot_(s) };
  });
}

function deleteUser(token, id) {
  return call_(token, true, function (s) {
    if (id === s.id) return { error: 'SELF', message: 'You can’t delete your own account.' };
    const all = readAll_('admins');
    const u = findRow_('admins', id);
    if (u && u.role === 'Admin' && u.active === 'TRUE' && activeAdmins_(all, id) === 0) return { error: 'LAST_ADMIN', message: 'Keep at least one active admin.' };
    withLock_(function () { deleteRow_('admins', id); });
    dropSessions_(id);
    return { ok: true, data: snapshot_(s) };
  });
}

/* =========================== application sheet sync =========================== */

const HMAP = [
  ['email', ['email', 'e-mail', 'mail id']],
  ['mobile', ['mobile', 'phone', 'contact no', 'contact number', 'whatsapp', 'number']],
  ['resumeName', ['cv file name', 'resume file name', 'file name']],
  ['resume', ['resume', 'cv']],
  ['company', ['current company', 'company', 'employer', 'organisation', 'organization']],
  ['designation', ['designation']],
  ['ectc', ['expected ctc', 'expected salary']],
  ['ctc', ['current ctc', 'ctc', 'current salary']],
  ['notice', ['notice']],
  ['linkedin', ['linkedin']],
  ['position', ['position', 'job', 'role', 'applied for', 'title']],
  ['appliedDate', ['applied on', 'application date', 'submitted', 'date', 'timestamp']],
  ['domainExp', ['real estate exp', 'industry exp', 'domain exp', 'hospitality exp', 'relevant exp']],
  ['experience', ['experience', 'exp']],
  ['location', ['location', 'city', 'address']],
  ['source', ['utm_source', 'source']],
  ['name', ['candidate name', 'full name', 'name', 'candidate']]
];
const DETAIL_KEYS = ['designation', 'company', 'domainExp', 'ctc', 'ectc', 'notice', 'linkedin'];

function syncApplications() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(25000)) return null;
  try {
    const jobs = readAll_('jobs').filter(function (j) { return j.sheetId; });
    const cands = readAll_('candidates');
    const keys = {};
    cands.forEach(function (c) {
      if (c.sourceKey) keys[c.sourceKey] = 1;
      const k = digits_(c.mobile) || String(c.email || '').toLowerCase();
      if (k) keys[c.jobId + '|' + k] = 1;
    });
    const results = {};
    const newRows = [];
    const now = nowIso_();
    jobs.forEach(function (job) {
      try {
        const ss = SpreadsheetApp.openById(job.sheetId);
        const values = ss.getSheets()[0].getDataRange().getDisplayValues();
        const head = values[0] || [];
        const m = mapHeaders_(head);
        if (m.name == null) throw new Error('No name column found in this sheet');
        const g = function (r, k) { return m[k] != null ? String(r[m[k]] || '').trim() : ''; };
        const body = values.slice(1);
        const order = dateOrder_(body.map(function (r) { return g(r, 'appliedDate'); }));
        let added = 0;
        body.forEach(function (r) {
          const name = g(r, 'name');
          if (!name) return;
          const mob = digits_(g(r, 'mobile'));
          const em = g(r, 'email').toLowerCase();
          const key = mob || em;
          if (!key) return;
          const sk = job.sheetId + '|' + key;
          if (keys[sk] || keys[job.id + '|' + key]) return;
          keys[sk] = 1; keys[job.id + '|' + key] = 1;
          let exp = g(r, 'experience');
          if (/^\d+(\.\d+)?$/.test(exp)) exp += ' yrs';
          const res = g(r, 'resume');
          newRows.push({
            id: newId_(), business: job.business, jobId: job.id, jobTitle: job.title, name: name, mobile: mob || g(r, 'mobile'),
            email: g(r, 'email'), experience: exp, location: g(r, 'location'), appliedDate: normDate_(g(r, 'appliedDate'), order),
            source: srcName_(g(r, 'source')), status: 'New', resumeName: g(r, 'resumeName'), resumeLink: /^https?:\/\//i.test(res) ? res : '',
            notes: '', details: DETAIL_KEYS.filter(function (k) { return g(r, k); }).map(function (k) { return { k: String(head[m[k]]).trim(), v: g(r, k) }; }),
            sheetId: job.sheetId, sourceKey: sk, appliedAt: g(r, 'appliedDate'), deleted: '', createdAt: now, updatedAt: now
          });
          added++;
        });
        results[job.id] = { ok: true, rows: body.length, added: added, title: ss.getName() };
      } catch (e) {
        results[job.id] = { ok: false, message: String(e && e.message || e) };
      }
    });
    if (newRows.length) appendRows_('candidates', newRows);
    const info = { at: now, added: newRows.length, results: results };
    setSetting_('lastSync', info);
    return info;
  } finally {
    lock.releaseLock();
  }
}

function mapHeaders_(h) {
  const m = {};
  h.forEach(function (raw, ix) {
    const x = String(raw).toLowerCase().trim();
    for (let i = 0; i < HMAP.length; i++) {
      const k = HMAP[i][0];
      if (m[k] == null && HMAP[i][1].some(function (kw) { return x.indexOf(kw) >= 0; })) { m[k] = ix; break; }
    }
  });
  return m;
}
function dateOrder_(vals) {
  let dmy = 0, mdy = 0;
  vals.forEach(function (v) {
    const m = String(v || '').match(/^(\d{1,2})[\/.-](\d{1,2})[\/.-]\d{2,4}/);
    if (m) { if (+m[1] > 12) dmy++; if (+m[2] > 12) mdy++; }
  });
  return dmy > mdy ? 'dmy' : 'mdy';
}
function normDate_(s, order) {
  s = String(s || '').trim();
  const pad = function (n) { return ('0' + n).slice(-2); };
  let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m) return m[1] + '-' + pad(m[2]) + '-' + pad(m[3]);
  m = s.match(/^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{2,4})/);
  if (m) {
    let y = +m[3]; if (y < 100) y += 2000;
    const a = +m[1], b = +m[2];
    return y + '-' + pad(order === 'mdy' ? a : b) + '-' + pad(order === 'mdy' ? b : a);
  }
  const d = s ? new Date(s) : new Date();
  return Utilities.formatDate(isNaN(d) ? new Date() : d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}
function srcName_(v) {
  const k = String(v || '').trim().toLowerCase();
  if (!k) return 'Application form';
  const map = { ig: 'Instagram', instagram: 'Instagram', fb: 'Facebook', facebook: 'Facebook', meta: 'Meta ads', google: 'Google',
    linkedin: 'LinkedIn', indeed: 'Indeed', naukri: 'Naukri', whatsapp: 'WhatsApp', wa: 'WhatsApp' };
  return map[k] || (k.charAt(0).toUpperCase() + k.slice(1));
}
function digits_(s) { return String(s || '').replace(/\D/g, '').slice(-10); }

/* =========================== helpers =========================== */

function call_(token, adminOnly, fn) {
  try {
    const s = auth_(token);
    if (!s) return { error: 'AUTH', message: 'Your session ended. Sign in again.' };
    if (adminOnly && s.role !== 'Admin') return { error: 'FORBIDDEN', message: 'Only admins can do this.' };
    return fn(s);
  } catch (e) {
    return { error: 'SERVER', message: String(e && e.message || e) };
  }
}
function auth_(token) {
  if (!token) return null;
  const cache = CacheService.getScriptCache();
  const raw = cache.get('sess_' + token);
  if (!raw) return null;
  cache.put('sess_' + token, raw, CONFIG.SESSION_MINUTES * 60); // sliding expiry
  return JSON.parse(raw);
}
function rememberToken_(userId, token) {
  const cache = CacheService.getScriptCache();
  const list = JSON.parse(cache.get('toks_' + userId) || '[]').slice(-9);
  list.push(token);
  cache.put('toks_' + userId, JSON.stringify(list), 21600);
}
function dropSessions_(userId) {
  const cache = CacheService.getScriptCache();
  JSON.parse(cache.get('toks_' + userId) || '[]').forEach(function (t) { cache.remove('sess_' + t); });
  cache.remove('toks_' + userId);
}
function hashPw_(pw, salt) {
  let d = salt + '|' + pw, bytes;
  for (let i = 0; i < 1000; i++) {
    bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, d, Utilities.Charset.UTF_8);
    d = Utilities.base64Encode(bytes) + salt;
  }
  return Utilities.base64Encode(bytes);
}
function pwProblem_(pw) {
  pw = String(pw || '');
  if (pw.length < 8) return 'Use at least 8 characters for the password.';
  if (!/[A-Za-z]/.test(pw) || !/\d/.test(pw)) return 'Use both letters and numbers in the password.';
  return '';
}
function activeAdmins_(all, exceptId) {
  return all.filter(function (a) { return a.role === 'Admin' && a.active === 'TRUE' && a.id !== exceptId; }).length;
}
function snapshot_(s) {
  const cfg = getSetting_('config', {});
  const out = {
    jobs: readAll_('jobs').map(function (j) { j.vacancies = Number(j.vacancies) || 1; return j; }),
    cands: readAll_('candidates').filter(function (c) { return c.deleted !== 'TRUE'; }),
    ints: readAll_('interviews'),
    cfg: { businesses: cfg.businesses || [], interviewers: cfg.interviewers || [], departments: cfg.departments || [] },
    sync: getSetting_('lastSync', null),
    me: { id: s.id, username: s.username, name: s.name, role: s.role }
  };
  if (s.role === 'Admin') {
    out.users = readAll_('admins').map(function (a) {
      return { id: a.id, username: a.username, name: a.name, role: a.role, active: a.active, lastLogin: a.lastLogin, createdAt: a.createdAt };
    });
  }
  return out;
}
function checkCol_(col) { if (!EDITABLE[col]) throw new Error('Unknown collection'); }
function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try { return fn(); } finally { lock.releaseLock(); }
}
function newId_() { return Utilities.getUuid().replace(/-/g, '').slice(0, 16); }
function nowIso_() { return new Date().toISOString(); }
function fileIdFrom_(x) {
  x = String(x || '').trim();
  const m = x.match(/\/d\/([A-Za-z0-9_-]{20,})/) || x.match(/[?&]id=([A-Za-z0-9_-]{20,})/);
  if (m) return m[1];
  return /^[A-Za-z0-9_-]{20,}$/.test(x) ? x : '';
}
function getLogoDataUri_() {
  try {
    const id = getSetting_('logoFileId', '') || fileIdFrom_(CONFIG.LOGO_FILE_ID);
    if (!id) return '';
    const b = DriveApp.getFileById(id).getBlob();
    return 'data:' + b.getContentType() + ';base64,' + Utilities.base64Encode(b.getBytes());
  } catch (e) {
    return '';
  }
}

/* ---- sheet storage ---- */
function db_() {
  const id = PropertiesService.getScriptProperties().getProperty('DB_ID');
  if (!id) throw new Error('Run setup() in the Apps Script editor first.');
  return SpreadsheetApp.openById(id);
}
function sheet_(name) { return db_().getSheetByName(name); }
function readAll_(name) {
  const sh = sheet_(name);
  const last = sh.getLastRow();
  if (last < 2) return [];
  const cols = SHEETS[name];
  const vals = sh.getRange(2, 1, last - 1, cols.length).getDisplayValues();
  return vals.filter(function (r) { return r[0] !== ''; }).map(function (r) {
    const o = {};
    cols.forEach(function (k, i) { o[k] = r[i]; });
    (JSON_FIELDS[name] || []).forEach(function (k) {
      try { o[k] = o[k] ? JSON.parse(o[k]) : []; } catch (e) { o[k] = []; }
    });
    return o;
  });
}
function toRow_(name, obj) {
  return SHEETS[name].map(function (k) {
    let v = obj[k];
    if ((JSON_FIELDS[name] || []).indexOf(k) >= 0) v = JSON.stringify(v || []);
    if (v === undefined || v === null) v = '';
    if (typeof v === 'boolean') v = v ? 'TRUE' : 'FALSE';
    return String(v);
  });
}
function rowIndex_(name, id) {
  const sh = sheet_(name);
  const last = sh.getLastRow();
  if (last < 2) return -1;
  const ids = sh.getRange(2, 1, last - 1, 1).getDisplayValues();
  for (let i = 0; i < ids.length; i++) if (ids[i][0] === String(id)) return i + 2;
  return -1;
}
function findRow_(name, id) {
  return readAll_(name).filter(function (r) { return r.id === String(id); })[0] || null;
}
function upsert_(name, obj) {
  const sh = sheet_(name);
  let row = rowIndex_(name, obj.id);
  if (row < 0) row = sh.getLastRow() + 1;
  sh.getRange(row, 1, 1, SHEETS[name].length).setNumberFormat('@').setValues([toRow_(name, obj)]);
}
function appendRows_(name, objs) {
  if (!objs.length) return;
  const sh = sheet_(name);
  sh.getRange(sh.getLastRow() + 1, 1, objs.length, SHEETS[name].length)
    .setNumberFormat('@').setValues(objs.map(function (o) { return toRow_(name, o); }));
}
function patchRow_(name, id, patch) {
  const cur = findRow_(name, id);
  if (!cur) throw new Error('Record not found. It may have been deleted.');
  upsert_(name, Object.assign(cur, patch));
}
function deleteRow_(name, id) {
  const row = rowIndex_(name, id);
  if (row > 0) sheet_(name).deleteRow(row);
}
function getSetting_(key, fallback) {
  const r = readAll_('settings').filter(function (x) { return x.key === key; })[0];
  if (!r) return fallback;
  try { return JSON.parse(r.value); } catch (e) { return fallback; }
}
function setSetting_(key, value) {
  const sh = sheet_('settings');
  const last = sh.getLastRow();
  const keys = last > 1 ? sh.getRange(2, 1, last - 1, 1).getDisplayValues().map(function (r) { return r[0]; }) : [];
  const ix = keys.indexOf(key);
  const row = ix >= 0 ? ix + 2 : last + 1;
  sh.getRange(row, 1, 1, 2).setNumberFormat('@').setValues([[key, JSON.stringify(value)]]);
}

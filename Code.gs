/****************************************************
 * Karyamas – GAS Backend (Web App)
 * Endpoint:
 *  POST action=login         { username, password }
 *  POST action=pushScore     { row: { ...front item... } }
 *  POST action=syncUsers     { users: [ {username,password,role,region,unit,status}, ... ] }
 *  GET  ?action=getUsers
 *  GET  ?action=getMasterData[&region=...&unit=...]
 ****************************************************/

const CFG = {
  // Jika dikosongkan, skrip akan membuat Spreadsheet baru dan menyimpan ID di Script Properties.
  SPREADSHEET_ID: '1Q1DXzVzivova2G3zegjsJ9I1ep4na3oMMYnUppK8YmA',
  PROP_KEY: 'MANDOR_APP_SSID',
  SHEETS: {
    USERS:  'Users',
    MASTER: 'MasterMandor',
    SCORES: 'Scores',
  },
  DEFAULT_ADMIN: {
    username: 'admin',
    password: 'user123',
    role:     'admin',
    name:     'Administrator',
    region:   '',     // kosong = semua
    unit:     '',     // kosong = semua
    status:   'active'
  }
};

/* ===== Entry points ===== */
function doPost(e) {
  try {
    ensureSetup();
    const req = JSON.parse(e.postData.contents || '{}');
    const action = String(req.action || '').trim();

    switch (action) {
      case 'login':
        return jsonOK(login(req.username, req.password));

      case 'pushScore':
        return jsonOK(pushScore(req.row));

      case 'syncUsers':
        return jsonOK(syncUsers(req.users));

      default:
        return jsonErr('Unknown POST action');
    }
  } catch (err) {
    return jsonErr(err.message || String(err));
  }
}

function doGet(e) {
  try {
    ensureSetup();
    const action = String(e.parameter.action || '').trim();

    switch (action) {
      case 'getUsers':
        return jsonOK(getUsers());

      case 'getMasterData':
        return jsonOK(getMasterData({
          region: e.parameter.region || '',
          unit:   e.parameter.unit   || ''
        }));

      default:
        return jsonErr('Unknown GET action');
    }
  } catch (err) {
    return jsonErr(err.message || String(err));
  }
}

/* ===== Setup helpers ===== */
function ensureSetup() {
  const ss = getSpreadsheet_(); // selalu valid, auto-create bila perlu

  // Pastikan sheet & header
  ensureSheet(ss, CFG.SHEETS.USERS,  ['username','password','role','name','region','unit','status']);
  ensureSheet(ss, CFG.SHEETS.MASTER, ['NIP','Nama','Region','Unit','Divisi','KodeJabatan','Grade']);
  ensureSheet(ss, CFG.SHEETS.SCORES, ['_id','NIP','Nama','Region','Unit','Divisi','KodeJabatan','Grade','NilaiIsian','NilaiBKM','Total','_synced','createdBy','createdAt','syncAt']);

  // Seed admin default bila kosong
  const shUsers = ss.getSheetByName(CFG.SHEETS.USERS);
  const rows = shUsers.getDataRange().getValues();
  if (rows.length <= 1) {
    shUsers.appendRow([
      CFG.DEFAULT_ADMIN.username,
      CFG.DEFAULT_ADMIN.password,
      CFG.DEFAULT_ADMIN.role,
      CFG.DEFAULT_ADMIN.name,
      CFG.DEFAULT_ADMIN.region,
      CFG.DEFAULT_ADMIN.unit,
      CFG.DEFAULT_ADMIN.status
    ]);
  }
}

// === TAMBAHKAN helper ini ===
function getSpreadsheet_() {
  const props = PropertiesService.getScriptProperties();

  // 1) Utamakan SSID di Script Properties
  let ssid = props.getProperty(CFG.PROP_KEY);

  // 2) Atau pakai CFG.SPREADSHEET_ID jika ada
  if (!ssid && CFG.SPREADSHEET_ID && String(CFG.SPREADSHEET_ID).trim()) {
    ssid = String(CFG.SPREADSHEET_ID).trim();
    props.setProperty(CFG.PROP_KEY, ssid);
  }

  // 3) Coba open; jika gagal → buat baru & simpan
  try {
    if (ssid && String(ssid).trim()) {
      return SpreadsheetApp.openById(ssid);
    }
  } catch (e) {
    // ssid lama invalid → buang & buat baru
  }

  const ssNew = SpreadsheetApp.create('Mandor Competency DB');
  const newId = ssNew.getId();
  props.setProperty(CFG.PROP_KEY, newId);
  return ssNew;
}


function ensureSheet(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);

  const w = headers.length;
  const cur = sh.getRange(1,1,1,w).getValues()[0];
  const same = cur.every((v,i) => String(v).trim() === headers[i]);
  if (!same) {
    sh.clear();
    sh.getRange(1,1,1,w).setValues([headers]);
  }
}

/* ===== Actions ===== */

// LOGIN
function login(username, password) {
  if (!username || !password) throw new Error('Username/password wajib.');

  const ss = getSpreadsheet_();
  const sh = ss.getSheetByName(CFG.SHEETS.USERS);
  const values = sh.getDataRange().getValues();
  const head = values[0], rows = values.slice(1);

  const idx = (h) => head.indexOf(h);
  const iUser = idx('username'), iPass = idx('password'), iRole = idx('role'),
        iName = idx('name'), iReg = idx('region'), iUnit = idx('unit'), iStat = idx('status');

  const found = rows.find(r =>
    String(r[iUser]).trim().toLowerCase() === String(username).toLowerCase() &&
    String(r[iPass]).trim() === String(password).trim()
  );
  if (!found) throw new Error('Username/password salah.');
  if (String(found[iStat]).toLowerCase() !== 'active') throw new Error('User tidak aktif.');

  const user = {
    username: String(found[iUser]),
    role:     String(found[iRole] || 'kerani'),
    name:     String(found[iName] || ''),
    region:   String(found[iReg]  || ''), // '' = semua
    unit:     String(found[iUnit] || ''), // '' = semua
    status:   String(found[iStat] || 'active')
  };

  return user; // {status:'success', data:user}
}

// PUSH SCORE (upsert by _id)
function pushScore(row) {
  if (!row) throw new Error('Row kosong.');
  const id = String(row.id || row._id || '').trim();
  if (!id) throw new Error('Missing _id/id.');

  // map frontend → backend columns
  const rec = {
    _id:         id,
    NIP:         String(row.nip || row.NIP || ''),
    Nama:        String(row.nama || row.Nama || ''),
    Region:      String(row.region || row.Region || ''),
    Unit:        String(row.unit || row.Unit || ''),
    Divisi:      String(row.divisi || row.Divisi || ''),
    KodeJabatan: String(row.jabatan || row.KodeJabatan || ''),
    Grade:       String(row.grade || row.Grade || ''),
    NilaiIsian:  toNumber(row.nilaiIsian || row.NilaiIsian || 0),
    NilaiBKM:    toNumber(row.nilaiBKM   || row.NilaiBKM   || 0),
    Total:       toNumber(row.totalNilai || row.Total      || 0),
    _synced:     true,
    createdBy:   String(row.inputBy || row.createdBy || ''),
    createdAt:   String(row.timestamp || row.createdAt || new Date().toISOString()),
    syncAt:      new Date()
  };

  const ss = getSpreadsheet_();
  const sh = ss.getSheetByName(CFG.SHEETS.SCORES);
  const all = sh.getDataRange().getValues();

  // cari baris dengan _id sama
  const rowIdx = all.findIndex(r => r[0] === rec._id); // kolom A = _id
  const arr = [
    rec._id, rec.NIP, rec.Nama, rec.Region, rec.Unit, rec.Divisi, rec.KodeJabatan, rec.Grade,
    rec.NilaiIsian, rec.NilaiBKM, rec.Total, rec._synced, rec.createdBy, rec.createdAt, rec.syncAt
  ];

  if (rowIdx >= 0) {
    sh.getRange(rowIdx+1, 1, 1, arr.length).setValues([arr]);
  } else {
    sh.appendRow(arr);
  }
  return { ok: true };
}

// SYNC USERS (upsert per username)
function syncUsers(users) {
  if (!Array.isArray(users)) throw new Error('users harus array.');

  const ss = getSpreadsheet_();
  const sh = ss.getSheetByName(CFG.SHEETS.USERS);
  const values = sh.getDataRange().getValues();
  const head = values[0];
  const idx = (h) => head.indexOf(h);

  // Buat map username → rowIndex
  const map = new Map();
  values.slice(1).forEach((r, i) => {
    map.set(String(r[idx('username')]).trim().toLowerCase(), i + 2); // +2 karena header + 1-based
  });

  users.forEach(u => {
    const username = String(u.username || '').trim();
    if (!username) return;
    const line = [
      username,
      String(u.password || ''),
      String(u.role || 'kerani'),
      String(u.name || ''),
      String(u.region || ''),
      String(u.unit || ''),
      String(u.status || 'active')
    ];

    const key = username.toLowerCase();
    if (map.has(key)) {
      sh.getRange(map.get(key), 1, 1, line.length).setValues([line]);
    } else {
      sh.appendRow(line);
    }
  });

  return { ok: true, count: users.length };
}

// GET USERS (semua)
function getUsers() {
  const ss = getSpreadsheet_();
  const sh = ss.getSheetByName(CFG.SHEETS.USERS);
  const values = sh.getDataRange().getValues();
  const head = values[0], rows = values.slice(1);
  const idx = (h) => head.indexOf(h);

  const out = rows.map(r => ({
    username: String(r[idx('username')]),
    role:     String(r[idx('role')]),
    name:     String(r[idx('name')]),
    region:   String(r[idx('region')]),
    unit:     String(r[idx('unit')]),
    status:   String(r[idx('status')])
  }));
  return out;
}

// GET MASTER (opsional filter region/unit)
function getMasterData(filter) {
  const region = String(filter?.region || '').trim();
  const unit   = String(filter?.unit   || '').trim();

  const ss = getSpreadsheet_();
  const sh = ss.getSheetByName(CFG.SHEETS.MASTER);
  const values = sh.getDataRange().getValues();
  const head = values[0], rows = values.slice(1);
  const idx = (h) => head.indexOf(h);

  let out = rows.map(r => ({
    NIP:         clean(r[idx('NIP')]),
    Nama:        clean(r[idx('Nama')]),
    Region:      clean(r[idx('Region')]),
    Unit:        clean(r[idx('Unit')]),
    Divisi:      clean(r[idx('Divisi')]),
    KodeJabatan: clean(r[idx('KodeJabatan')]),
    Grade:       clean(r[idx('Grade')])
  }));

  if (region) out = out.filter(x => String(x.Region) === region);
  if (unit)   out = out.filter(x => String(x.Unit)   === unit);

  return out;
}

/* ===== utils ===== */
function jsonOK(data) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'success', data }))
    .setMimeType(ContentService.MimeType.JSON);
}
function jsonErr(message) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'error', message: String(message) }))
    .setMimeType(ContentService.MimeType.JSON);
}
function toNumber(v) {
  const n = Number(v); return isFinite(n) ? n : 0;
}
function clean(v) {
  return (v === null || v === undefined) ? '' : v;
}

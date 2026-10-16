/**
 * MANAJEMEN & ARSIP SURAT + QR VALIDASI (Google Apps Script) - v4
 * Role: superadmin (akses penuh + kelola pengguna) dan dosen (hanya melihat arsip).
 * Arsip: Tahun > Unit > Tujuan surat. Nomor default: URUT.KODE.UNIT.TAHUN
 * Jalankan setup() sekali sebelum deploy.
 */
var HEAD = ['ID','Dibuat','Nomor','Tanggal','Perihal','Tujuan','Pejabat','Jabatan','Unit','UnitKey','Status','Hash','File','Tahun','KatKey','Publik','Sumber'];
var UHEAD = ['Username','Nama','Role','Salt','Hash','Status','Dibuat'];
var BULAN = ['Januari','Februari','Maret','April','Mei','Juni','Juli','Agustus','September','Oktober','November','Desember'];
var MAX_MB = 25; // batas ukuran file (MB). Jangan lebih dari ~30 karena batas Apps Script
var DEFAULT_PW = 'ganti-password-ini';

// ===== KONFIGURASI =====
var DEFAULT_FMT = '{urut}.{kode}.{unit}.{tahun}';
var UNITS = {
  fakultas:      { n: 'Fakultas Pertanian',           c: 'FP' },
  agroteknologi: { n: 'Program Studi Agroteknologi',  c: 'BP' },
  agribisnis:    { n: 'Program Studi Agribisnis',     c: 'EP' }
};
// Jenis surat bawaan. Dipakai sampai superadmin mengubahnya lewat tab "Jenis surat".
// k = id internal, n = nama, c = kode pada nomor, units = unit yang memakai jenis ini
var ALLU = ['fakultas', 'agroteknologi', 'agribisnis'];
var DEFAULT_CATS = [
  { k: 'dekan',  n: 'Surat Ke Dekan',     c: 'FP',     units: ['agroteknologi', 'agribisnis'] },
  { k: 'mhs',    n: 'Surat ke Mahasiswa', c: 'M',      units: ALLU },
  { k: 'ks',     n: 'Surat Kerjasama',    c: 'KS',     units: ALLU },
  { k: 'dosen',  n: 'Surat Ke Dosen',     c: 'D',      units: ALLU },
  { k: 'pkl',    n: 'Surat PKL',          c: 'PL',     units: ALLU },
  { k: 'magang', n: 'Surat Magang',       c: 'Magang', units: ALLU },
  { k: 'rektor', n: 'Surat Ke Rektor',    c: 'R',      units: ALLU },
  { k: 'luar',   n: 'Surat Ke Luar',      c: 'L',      units: ALLU }
];

/* ---------- SETUP (jalankan sekali dari editor) ---------- */
function setup() {
  var p = PropertiesService.getScriptProperties();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('Surat') || ss.insertSheet('Surat');
  sh.getRange('A:Q').setNumberFormat('@');
  sh.getRange(1, 1, 1, HEAD.length).setValues([HEAD]);
  var us = ss.getSheetByName('User') || ss.insertSheet('User');
  us.getRange('A:G').setNumberFormat('@');
  us.getRange(1, 1, 1, UHEAD.length).setValues([UHEAD]);
  if (us.getLastRow() < 2) addUserRow_('admin', 'Super Admin', 'superadmin', p.getProperty('ADMIN_PASSWORD') || DEFAULT_PW);
  if (!p.getProperty('FOLDER_ID')) p.setProperty('FOLDER_ID', DriveApp.createFolder('Surat Keluar').getId());
  if (!p.getProperty('YEARS')) p.setProperty('YEARS', JSON.stringify([new Date().getFullYear()]));
  Logger.log('Setup selesai. Login pertama: username "admin", password "' + (p.getProperty('ADMIN_PASSWORD') || DEFAULT_PW) + '". Segera ganti di menu Akun saya.');
}

// Darurat (lupa password superadmin): jalankan dari editor, password "admin" kembali ke default
function resetAdmin() {
  var u = findUser_('admin');
  if (!u) return addUserRow_('admin', 'Super Admin', 'superadmin', DEFAULT_PW);
  setPw_(u, DEFAULT_PW);
  usersSheet_().getRange(u.row, 3).setValue('superadmin');
  usersSheet_().getRange(u.row, 6).setValue('AKTIF');
  Logger.log('Password admin direset ke: ' + DEFAULT_PW);
}

/* ---------- WEB APP ---------- */
function doGet(e) {
  var v = e && e.parameter && e.parameter.v;
  if (v) {
    var t = HtmlService.createTemplateFromFile('Verify');
    var r = getLetter_(v);
    if (r) { var cm = catsMap_()[r.KatKey]; r.Jenis = cm ? cm.n : '-'; }
    t.r = r;
    t.logo = r ? logo_(r.UnitKey) : '';
    return t.evaluate().setTitle('Validasi Surat').addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
  var a = HtmlService.createTemplateFromFile('Index');
  a.meta = meta_();
  return a.evaluate().setTitle('Arsip Surat').addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function sheet_() { return SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Surat'); }
function usersSheet_() { return SpreadsheetApp.getActiveSpreadsheet().getSheetByName('User'); }
function prop_(k) { return PropertiesService.getScriptProperties().getProperty(k); }
function tanggalID_(s) { var d = s.split('-'); return +d[2] + ' ' + BULAN[+d[1] - 1] + ' ' + d[0]; }
function baseUrl_() { return prop_('WEBAPP_URL') || ScriptApp.getService().getUrl(); }
function hex_(bytes) { return bytes.map(function (b) { return ('0' + (b & 0xff).toString(16)).slice(-2); }).join(''); }
function sha_(s) { return hex_(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s, Utilities.Charset.UTF_8)); }

/* ---------- USER & SESI ---------- */
function users_() {
  return usersSheet_().getDataRange().getValues().slice(1).map(function (r, i) {
    return { row: i + 2, username: String(r[0]), nama: String(r[1]), role: String(r[2]), salt: String(r[3]), hash: String(r[4]), status: String(r[5]) };
  });
}
function findUser_(name) {
  name = String(name || '').trim().toLowerCase();
  var l = users_();
  for (var i = 0; i < l.length; i++) if (l[i].username === name) return l[i];
  return null;
}
function activeSupers_() { return users_().filter(function (u) { return u.role === 'superadmin' && u.status === 'AKTIF'; }); }
function addUserRow_(username, nama, role, pw) {
  var salt = Utilities.getUuid();
  usersSheet_().appendRow([username, nama, role, salt, sha_(salt + pw), 'AKTIF', Utilities.formatDate(new Date(), 'Asia/Jakarta', 'dd/MM/yyyy HH:mm')]);
}
function setPw_(u, pw) {
  var salt = Utilities.getUuid();
  usersSheet_().getRange(u.row, 4, 1, 2).setValues([[salt, sha_(salt + pw)]]);
}

function login(username, pw) {
  var name = String(username || '').trim().toLowerCase(), c = CacheService.getScriptCache(), fk = 'F_' + name;
  var n = +(c.get(fk) || 0);
  if (n >= 5) throw new Error('Terlalu banyak percobaan gagal. Coba lagi 10 menit lagi.');
  var u = findUser_(name);
  if (!u || u.status !== 'AKTIF' || sha_(u.salt + String(pw || '')) !== u.hash) {
    c.put(fk, String(n + 1), 600);
    throw new Error('Username atau password salah');
  }
  c.remove(fk);
  var token = Utilities.getUuid() + Utilities.getUuid();
  c.put('T_' + token, u.username, 21600); // sesi 6 jam
  return { token: token, nama: u.nama, role: u.role, username: u.username };
}
function logout(token) { CacheService.getScriptCache().remove('T_' + token); }

// Cek sesi + role dari Sheet setiap panggilan (user dinonaktifkan = langsung tidak bisa akses)
function auth_(token, superOnly) {
  var name = CacheService.getScriptCache().get('T_' + token);
  if (!name) throw new Error('Sesi berakhir, silakan login ulang');
  var u = findUser_(name);
  if (!u || u.status !== 'AKTIF') throw new Error('Akun tidak aktif');
  if (superOnly && u.role !== 'superadmin') throw new Error('Akses ditolak: khusus superadmin');
  return u;
}

function listUsers(token) {
  auth_(token, true);
  return users_().map(function (u) { return { username: u.username, nama: u.nama, role: u.role, status: u.status }; });
}
function addUser(token, d) {
  auth_(token, true);
  var name = String(d.username || '').trim().toLowerCase();
  if (!/^[a-z0-9._-]{3,30}$/.test(name)) throw new Error('Username 3-30 karakter: huruf kecil, angka, titik, strip, garis bawah');
  if (!String(d.nama || '').trim()) throw new Error('Nama wajib diisi');
  if (d.role !== 'superadmin' && d.role !== 'dosen') throw new Error('Role tidak valid');
  if (String(d.password || '').length < 6) throw new Error('Password minimal 6 karakter');
  if (findUser_(name)) throw new Error('Username sudah dipakai');
  addUserRow_(name, d.nama.trim(), d.role, d.password);
  return true;
}
function toggleUser(token, username) {
  var me = auth_(token, true), u = findUser_(username);
  if (!u) throw new Error('User tidak ditemukan');
  if (u.status === 'AKTIF') {
    if (u.username === me.username) throw new Error('Anda tidak bisa menonaktifkan akun sendiri');
    if (u.role === 'superadmin' && activeSupers_().length < 2) throw new Error('Harus ada minimal satu superadmin aktif');
  }
  var s = u.status === 'AKTIF' ? 'NONAKTIF' : 'AKTIF';
  usersSheet_().getRange(u.row, 6).setValue(s);
  return s;
}
function resetUserPw(token, username, newPw) {
  auth_(token, true);
  var u = findUser_(username);
  if (!u) throw new Error('User tidak ditemukan');
  if (String(newPw || '').length < 6) throw new Error('Password minimal 6 karakter');
  setPw_(u, newPw);
  return true;
}
function deleteUser(token, username) {
  var me = auth_(token, true), u = findUser_(username);
  if (!u) throw new Error('User tidak ditemukan');
  if (u.username === me.username) throw new Error('Anda tidak bisa menghapus akun sendiri');
  if (u.role === 'superadmin' && u.status === 'AKTIF' && activeSupers_().length < 2) throw new Error('Harus ada minimal satu superadmin aktif');
  usersSheet_().deleteRow(u.row);
  return true;
}
function changeMyPw(token, oldPw, newPw) {
  var u = auth_(token);
  if (sha_(u.salt + String(oldPw || '')) !== u.hash) throw new Error('Password lama salah');
  if (String(newPw || '').length < 6) throw new Error('Password baru minimal 6 karakter');
  setPw_(u, newPw);
  return true;
}

/* ---------- META, LOGO ---------- */
function years_() {
  var y = [];
  try { y = JSON.parse(prop_('YEARS') || '[]'); } catch (e) {}
  if (!y.length) y = [new Date().getFullYear()];
  return y;
}
/* ---------- JENIS SURAT & FORMAT PENOMORAN (dikelola superadmin) ---------- */
var ROMAN_ = ['I', 'II', 'III', 'IV', 'V', 'VI', 'VII', 'VIII', 'IX', 'X', 'XI', 'XII'];
function cats_() {
  var l = null;
  try { l = JSON.parse(prop_('CATS') || 'null'); } catch (e) {}
  if (!l || !l.length) l = JSON.parse(JSON.stringify(DEFAULT_CATS));
  l.forEach(function (x) { x.fmt = x.fmt || DEFAULT_FMT; });
  return l;
}
function catsMap_() { var m = {}; cats_().forEach(function (x) { m[x.k] = x; }); return m; }

// Isi format: {urut} {kode} {unit} {tahun} {bulan} {romawi}
function fmtNo_(cat, unit, tahun, urut, bulan) {
  var b = bulan >= 1 && bulan <= 12 ? bulan : (new Date().getMonth() + 1);
  return String(cat.fmt || DEFAULT_FMT).replace(/\{(urut|kode|unit|tahun|bulan|romawi)\}/g, function (_, k) {
    return k === 'urut' ? urut : k === 'kode' ? cat.c : k === 'unit' ? UNITS[unit].c : k === 'tahun' ? tahun : k === 'bulan' ? b : ROMAN_[b - 1];
  });
}
// Regex untuk membaca nomor urut dari nomor surat yang sudah ada (sesuai format jenis surat)
function urutRegex_(cat, unit, tahun) {
  var esc = function (s) { return String(s).replace(/[.*+?^${}()|[\]\\\/]/g, '\\$&'); };
  var re = String(cat.fmt || DEFAULT_FMT).split(/(\{(?:urut|kode|unit|tahun|bulan|romawi)\})/).map(function (p) {
    switch (p) {
      case '{urut}': return '(\\d+)';
      case '{kode}': return esc(cat.c);
      case '{unit}': return esc(UNITS[unit].c);
      case '{tahun}': return esc(tahun);
      case '{bulan}': return '\\d{1,2}';
      case '{romawi}': return '[IVX]+';
      default: return esc(p);
    }
  }).join('');
  return new RegExp('^' + re + '(?:-\\d+)?$');
}
function catUse_(kat, unit) {
  return sheet_().getDataRange().getValues().slice(1).filter(function (r) { return String(r[14]) === kat && (!unit || String(r[9]) === unit); }).length;
}
function saveCats_(l) {
  PropertiesService.getScriptProperties().setProperty('CATS', JSON.stringify(l));
  return { cats: l, meta: meta_() };
}

function listCats(token) { auth_(token, true); return cats_(); }

function saveCat(token, d) {
  auth_(token, true);
  var n = String(d.n || '').trim(), c = String(d.c || '').trim(), fmt = String(d.fmt || '').trim() || DEFAULT_FMT;
  if (n.length < 3 || n.length > 60) throw new Error('Nama jenis surat 3-60 karakter');
  if (!/^[A-Za-z0-9.\-]{1,15}$/.test(c)) throw new Error('Kode 1-15 karakter: huruf, angka, titik, atau strip');
  if (fmt.length > 60) throw new Error('Format penomoran maksimal 60 karakter');
  if ((fmt.match(/\{urut\}/g) || []).length !== 1) throw new Error('Format harus memuat {urut} tepat satu kali');
  var rest = fmt.replace(/\{(urut|kode|unit|tahun|bulan|romawi)\}/g, '');
  if (/[{}]/.test(rest)) throw new Error('Placeholder tidak dikenal. Gunakan: {urut} {kode} {unit} {tahun} {bulan} {romawi}');
  if (!/^[A-Za-z0-9.\-\/_ ]*$/.test(rest)) throw new Error('Selain placeholder, format hanya boleh huruf, angka, titik, strip, garis miring, garis bawah, dan spasi');
  var units = (d.units || []).filter(function (u) { return UNITS[u]; });
  if (!units.length) throw new Error('Pilih minimal satu unit');

  var l = cats_(), cur = null;
  if (d.k) { cur = l.filter(function (x) { return x.k === d.k; })[0]; if (!cur) throw new Error('Jenis surat tidak ditemukan'); }
  if (l.some(function (x) { return x !== cur && x.n.toLowerCase() === n.toLowerCase(); })) throw new Error('Nama jenis surat sudah ada');
  if (cur) {
    cur.units.forEach(function (u) {
      if (units.indexOf(u) < 0 && catUse_(cur.k, u) > 0) throw new Error('Unit ' + UNITS[u].n + ' tidak bisa dilepas: sudah ada surat jenis ini di unit tersebut');
    });
    cur.n = n; cur.c = c; cur.fmt = fmt; cur.units = units;
  } else {
    if (l.length >= 40) throw new Error('Maksimal 40 jenis surat');
    l.push({ k: 'j' + Utilities.getUuid().replace(/-/g, '').slice(0, 6), n: n, c: c, fmt: fmt, units: units, off: false });
  }
  return saveCats_(l);
}

// Nonaktif = tidak muncul di form pendaftaran, tetapi arsip lama tetap tampil
function toggleCat(token, k) {
  auth_(token, true);
  var l = cats_(), x = l.filter(function (y) { return y.k === k; })[0];
  if (!x) throw new Error('Jenis surat tidak ditemukan');
  x.off = !x.off;
  return saveCats_(l);
}

function deleteCat(token, k) {
  auth_(token, true);
  var l = cats_(), x = l.filter(function (y) { return y.k === k; })[0];
  if (!x) throw new Error('Jenis surat tidak ditemukan');
  var n = catUse_(k);
  if (n > 0) throw new Error('Tidak bisa dihapus: ada ' + n + ' surat berjenis "' + x.n + '". Gunakan Nonaktifkan agar tidak muncul di form, atau hapus suratnya terlebih dahulu.');
  return saveCats_(l.filter(function (y) { return y.k !== k; }));
}

/* ---------- DAFTAR PEJABAT PENGESAH (dikelola superadmin) ---------- */
function pejabat_() { try { return JSON.parse(prop_('PEJABAT') || '[]'); } catch (e) { return []; } }
function savePejabat_(l) { PropertiesService.getScriptProperties().setProperty('PEJABAT', JSON.stringify(l)); return l; }
function cleanPj_(d) {
  var n = String((d && d.nama) || '').trim(), j = String((d && d.jabatan) || '').trim();
  if (!n || !j) throw new Error('Nama dan jabatan wajib diisi');
  if (n.length > 100 || j.length > 100) throw new Error('Maksimal 100 karakter');
  return { nama: n, jabatan: j };
}
function listPejabat(token) { auth_(token, true); return pejabat_(); }
function addPejabat(token, d) {
  auth_(token, true);
  var c = cleanPj_(d), l = pejabat_();
  l.push({ id: Utilities.getUuid().slice(0, 8), nama: c.nama, jabatan: c.jabatan });
  return savePejabat_(l);
}
function updatePejabat(token, id, d) {
  auth_(token, true);
  var c = cleanPj_(d), l = pejabat_(), p = l.filter(function (x) { return x.id === id; })[0];
  if (!p) throw new Error('Pejabat tidak ditemukan');
  p.nama = c.nama; p.jabatan = c.jabatan;
  return savePejabat_(l);
}
function deletePejabat(token, id) {
  auth_(token, true);
  return savePejabat_(pejabat_().filter(function (x) { return x.id !== id; }));
}

function meta_() {
  var u = {}, cl = cats_();
  Object.keys(UNITS).forEach(function (k) {
    u[k] = { n: UNITS[k].n, c: UNITS[k].c, cats: cl.filter(function (x) { return x.units.indexOf(k) >= 0; }).map(function (x) { return { k: x.k, n: x.n, off: !!x.off }; }) };
  });
  return { years: years_(), units: u, maxMb: MAX_MB, base: baseUrl_(), publicArsip: cfg_().enabled };
}
function logoId_(v) { var m = String(v || '').trim().match(/[-\w]{25,}/); return m ? m[0] : ''; }
var MIME_ = { png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', webp: 'image/webp', svg: 'image/svg+xml' };
function logo_(key) {
  try {
    var id = logoId_(prop_('LOGO_' + key));
    if (!id) return '';
    var f = DriveApp.getFileById(id), b = f.getBlob();
    var type = b.getContentType() || '';
    if (type.indexOf('image/') !== 0) type = MIME_[f.getName().split('.').pop().toLowerCase()] || '';
    if (!type) return '';
    return 'data:' + type + ';base64,' + Utilities.base64Encode(b.getBytes());
  } catch (e) { console.error('Logo ' + key + ': ' + e.message); return ''; }
}
function cekLogo() {
  Object.keys(UNITS).forEach(function (k) {
    var raw = prop_('LOGO_' + k), id = logoId_(raw);
    if (!raw) return Logger.log('[' + k + '] GAGAL: properti LOGO_' + k + ' belum dibuat / kosong');
    if (!id) return Logger.log('[' + k + '] GAGAL: nilai bukan ID Drive yang valid: "' + raw + '"');
    try {
      var f = DriveApp.getFileById(id), b = f.getBlob();
      Logger.log('[' + k + '] ' + (logo_(k) ? 'OK' : 'GAGAL (bukan file gambar)') + ' | ' + f.getName() + ' | ' + b.getContentType() + ' | ' + Math.round(b.getBytes().length / 1024) + ' KB');
    } catch (e) { Logger.log('[' + k + '] GAGAL membuka file: ' + e.message); }
  });
}

/* ---------- DATA SURAT ---------- */
function getLetter_(id) {
  id = String(id || '').trim().toUpperCase();
  var rows = sheet_().getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === id) {
      var o = {}; HEAD.forEach(function (h, j) { o[h] = rows[i][j] == null ? '' : String(rows[i][j]); });
      return o;
    }
  }
  return null;
}
function nextNo_(unit, kat, tahun, bulan) {
  var cat = catsMap_()[kat], re = urutRegex_(cat, unit, tahun), rows = sheet_().getDataRange().getValues(), max = 0;
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][9]) === unit && String(rows[i][14]) === kat && String(rows[i][13]) === String(tahun)) {
      var m = String(rows[i][2]).match(re);
      var n = m ? parseInt(m[1], 10) : parseInt(String(rows[i][2]), 10);
      if (n > max) max = n;
    }
  }
  return fmtNo_(cat, unit, tahun, max + 1, bulan);
}
function check_(unit, kat, tahun) {
  if (!UNITS[unit]) throw new Error('Unit tidak valid');
  var cat = catsMap_()[kat];
  if (!cat || cat.units.indexOf(unit) < 0) throw new Error('Jenis surat tidak tersedia untuk unit ini');
  if (cat.off) throw new Error('Jenis surat ini sedang dinonaktifkan');
  if (years_().map(String).indexOf(String(tahun)) < 0) throw new Error('Tahun belum ditambahkan');
}
function sub_(parent, name) {
  var it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}
function fileIdOf_(url) { var m = String(url).match(/[-\w]{25,}/); return m ? m[0] : ''; }

/* ---------- FUNGSI SURAT ---------- */
// Semua role: lihat arsip
function listLetters(token) {
  auth_(token);
  return lettersList_();
}
function lettersList_() {
  return sheet_().getDataRange().getValues().slice(1).reverse().map(function (r) {
    return { id: String(r[0]), no: String(r[2]), tgl: String(r[3]), perihal: String(r[4]), tujuan: String(r[5]), pejabat: String(r[6]), jabatan: String(r[7]),
             unit: String(r[8]), unitKey: String(r[9]), status: String(r[10]), tahun: String(r[13]), kat: String(r[14]), pub: String(r[15]) === 'TIDAK' ? 'TIDAK' : 'YA', hasFile: String(r[12]).trim() !== '' };
  });
}
// Publik (tanpa login): hanya daftar arsip. Tidak ada tautan/unduhan file, sidik, maupun QR.
function publicList() {
  var c = cfg_();
  if (!c.enabled) throw new Error('Arsip tidak ditampilkan untuk umum. Silakan login.');
  return lettersList_().filter(function (r) {
    return r.pub !== 'TIDAK' &&
      (!c.years.length || c.years.indexOf(r.tahun) >= 0) &&
      (!c.units.length || c.units.indexOf(r.unitKey) >= 0) &&
      (!c.cats.length || c.cats.indexOf(r.kat) >= 0) &&
      !(c.hideRevoked && r.status !== 'VALID');
  }).map(function (r) {
    return { id: r.id, no: r.no, tgl: r.tgl, perihal: r.perihal, tujuan: c.showTujuan ? r.tujuan : '', pejabat: c.showPengesah ? r.pejabat : '',
             jabatan: c.showPengesah ? r.jabatan : '', unit: r.unit, unitKey: r.unitKey, status: r.status, tahun: r.tahun, kat: r.kat };
  });
}

/* ---------- PENYARINGAN DATA PUBLIK (superadmin) ---------- */
// Kosong pada years/units/cats = tampilkan semua
function cfg_() {
  var d = { enabled: true, years: [], units: [], cats: [], hideRevoked: false, showTujuan: true, showPengesah: true }, s = {};
  try { s = JSON.parse(prop_('PUBLIC_CFG') || '{}'); } catch (e) {}
  Object.keys(d).forEach(function (k) { if (s[k] !== undefined) d[k] = s[k]; });
  return d;
}
function getPublicCfg(token) { auth_(token, true); return cfg_(); }
function savePublicCfg(token, c) {
  auth_(token, true);
  var ys = years_().map(String);
  var out = {
    enabled: !!c.enabled, hideRevoked: !!c.hideRevoked, showTujuan: !!c.showTujuan, showPengesah: !!c.showPengesah,
    years: (c.years || []).map(String).filter(function (y) { return ys.indexOf(y) >= 0; }),
    units: (c.units || []).filter(function (u) { return UNITS[u]; }),
    cats: (c.cats || []).filter(function (k) { return catsMap_()[k]; })
  };
  PropertiesService.getScriptProperties().setProperty('PUBLIC_CFG', JSON.stringify(out));
  return out;
}
// Sembunyikan / tampilkan satu surat pada daftar publik (QR validasi tetap berfungsi)
function setLetterPublic(token, id, isPublic) {
  auth_(token, true);
  var sh = sheet_(), rows = sh.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === id) { sh.getRange(i + 1, 16).setValue(isPublic ? 'YA' : 'TIDAK'); return true; }
  }
  throw new Error('Surat tidak ditemukan');
}
// Semua role: unduh file arsip (lewat server, tidak perlu akses Drive)
function getFile(token, id) {
  auth_(token);
  var l = getLetter_(id);
  if (!l) throw new Error('Surat tidak ditemukan');
  var fid = fileIdOf_(l.File);
  if (!fid) throw new Error('File tidak tersedia');
  var f = DriveApp.getFileById(fid), b = f.getBlob();
  var ext = (f.getName().match(/\.[A-Za-z0-9]{1,8}$/) || [''])[0];
  return { dataUri: 'data:' + (b.getContentType() || 'application/octet-stream') + ';base64,' + Utilities.base64Encode(b.getBytes()), name: 'Surat-' + l.Nomor.replace(/[\/\\:*?"<>|\s]/g, '-') + ext };
}

// Khusus superadmin
function nextNumber(token, unit, kat, tahun, bulan) {
  auth_(token, true); check_(unit, kat, tahun);
  return nextNo_(unit, kat, tahun, parseInt(bulan, 10));
}
function addYear(token, y) {
  auth_(token, true);
  y = parseInt(y, 10);
  if (!(y >= 2000 && y <= 2100)) throw new Error('Tahun tidak valid (2000-2100)');
  var list = years_();
  if (list.indexOf(y) < 0) list.push(y);
  list.sort(function (a, b) { return b - a; });
  PropertiesService.getScriptProperties().setProperty('YEARS', JSON.stringify(list));
  return list;
}
// Simpan berkas ke Drive (Tahun/Unit/Tujuan) -> { url, hash }
function saveFile_(d, unit, kat, tahun, no, id) {
  var bytes = Utilities.base64Decode(d.fileData);
  if (bytes.length > MAX_MB * 1024 * 1024) throw new Error('Ukuran file maksimal ' + MAX_MB + ' MB');
  var ext = (String(d.fileName || '').match(/\.[A-Za-z0-9]{1,8}$/) || [''])[0].toLowerCase();
  var mime = String(d.mime || '') || 'application/octet-stream';
  var name = 'Surat ' + String(no).replace(/[\/\\:*?"<>|]/g, '-') + ' [' + id + ']' + ext;
  var folder = sub_(sub_(sub_(DriveApp.getFolderById(prop_('FOLDER_ID')), String(tahun)), UNITS[unit].n), catsMap_()[kat].n);
  var file = folder.createFile(Utilities.newBlob(bytes, mime, name));
  return { url: file.getUrl(), hash: hex_(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, bytes)) };
}

// Langkah 1: daftarkan data surat -> QR langsung bisa dibuat. File surat OPSIONAL (boleh diunggah belakangan).
function createLetter(token, d) {
  auth_(token, true);
  check_(d.unit, d.kat, d.tahun);
  if (!prop_('FOLDER_ID')) throw new Error('Setup belum dijalankan. Jalankan fungsi setup() di editor Apps Script.');
  ['no', 'tanggal', 'perihal', 'tujuan'].forEach(function (k) { if (!d[k] || !String(d[k]).trim()) throw new Error('Kolom ' + k + ' wajib diisi'); });
  var pj = pejabat_().filter(function (p) { return p.id === d.pejabatId; })[0];
  if (!pj) throw new Error('Pilih pejabat pengesah');

  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var sh = sheet_(), nos = sh.getRange(2, 3, Math.max(sh.getLastRow() - 1, 1), 1).getValues();
    if (nos.some(function (r) { return String(r[0]) === d.no.trim(); })) throw new Error('Nomor surat sudah terdaftar');
    var id = Utilities.getUuid().replace(/-/g, '').slice(0, 8).toUpperCase();
    var f = d.fileData ? saveFile_(d, d.unit, d.kat, d.tahun, d.no.trim(), id) : { url: '', hash: '' };
    var now = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'dd/MM/yyyy HH:mm');
    sh.appendRow([id, now, d.no.trim(), tanggalID_(d.tanggal), d.perihal.trim(), d.tujuan.trim(), pj.nama, pj.jabatan, UNITS[d.unit].n, d.unit, 'VALID', f.hash, f.url, String(d.tahun), d.kat]);
    return { id: id, verifyUrl: baseUrl_() + '?v=' + id };
  } finally { lock.releaseLock(); }
}

// Langkah 2: unggah (atau ganti) berkas surat final untuk surat yang sudah terdaftar
function attachFile(token, id, d) {
  auth_(token, true);
  if (!d || !d.fileData) throw new Error('Pilih file surat terlebih dahulu');
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var sh = sheet_(), rows = sh.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === id) {
        var old = fileIdOf_(rows[i][12]);
        var f = saveFile_(d, String(rows[i][9]), String(rows[i][14]), String(rows[i][13]), String(rows[i][2]), id);
        sh.getRange(i + 1, 12, 1, 2).setValues([[f.hash, f.url]]);
        if (old) { try { DriveApp.getFileById(old).setTrashed(true); } catch (e) {} }
        return true;
      }
    }
    throw new Error('Surat tidak ditemukan');
  } finally { lock.releaseLock(); }
}
/* ---------- IMPOR ARSIP DARI GOOGLE SHEET (hasil Google Form) ---------- */
function openSrc_(src) {
  var m = String(src || '').match(/[-\w]{25,}/);
  if (!m) throw new Error('Tautan / ID Google Sheet tidak valid');
  try { return SpreadsheetApp.openById(m[0]); }
  catch (e) { throw new Error('Spreadsheet tidak dapat dibuka. Pastikan akun pemilik script punya akses ke sheet tersebut.'); }
}
function cellStr_(c) { return c instanceof Date ? Utilities.formatDate(c, 'Asia/Jakarta', 'dd/MM/yyyy') : String(c); }
function guess_(h) {
  var l = h.map(function (x) { return String(x).toLowerCase(); });
  function find(fn) { for (var i = 0; i < l.length; i++) if (fn(l[i])) return i; return -1; }
  return {
    no: find(function (s) { return /no(mor)?\.?\s*surat/.test(s) || s === 'no' || s === 'nomor'; }),
    tgl: find(function (s) { return /tanggal/.test(s) && !/timestamp/.test(s); }),
    perihal: find(function (s) { return /perihal|isi surat|tentang/.test(s); }),
    tujuan: find(function (s) { return /tujuan/.test(s); }),
    file: find(function (s) { return /(upload|unggah|file)/.test(s) && !/lampiran/.test(s); })
  };
}
function monthOf_(v) {
  var m;
  if (v instanceof Date && !isNaN(v)) return +Utilities.formatDate(v, 'Asia/Jakarta', 'M');
  var s = String(v || '').trim();
  if ((m = s.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{4})/))) return +m[2];
  if ((m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/))) return +m[2];
  return 0;
}
function parseTgl_(v) {
  var m;
  if (v instanceof Date && !isNaN(v)) return tanggalID_(Utilities.formatDate(v, 'Asia/Jakarta', 'yyyy-MM-dd'));
  var s = String(v || '').trim();
  if ((m = s.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{4})/)) && +m[2] >= 1 && +m[2] <= 12) return tanggalID_(m[3] + '-' + m[2] + '-' + m[1]);
  if ((m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/)) && +m[2] >= 1 && +m[2] <= 12) return tanggalID_(m[1] + '-' + m[2] + '-' + m[3]);
  return s || '-';
}

function importSheets(token, src) {
  auth_(token, true);
  var ss = openSrc_(src);
  return { title: ss.getName(), sheets: ss.getSheets().map(function (s) { return s.getName(); }) };
}
function importPreview(token, src, sheetName) {
  auth_(token, true);
  var sh = openSrc_(src).getSheetByName(sheetName);
  if (!sh) throw new Error('Tab sheet tidak ditemukan');
  var v = sh.getDataRange().getValues();
  if (!v.length) throw new Error('Sheet kosong');
  var head = v[0].map(String), rows = v.slice(1);
  var filled = rows.filter(function (r) { return r.some(function (x) { return String(x).trim() !== ''; }); });
  return { headers: head, total: rows.length, filled: filled.length, sample: filled.slice(0, 3).map(function (r) { return r.map(cellStr_); }), guess: guess_(head) };
}

// Salin berkas dari tautan Drive ke struktur arsip (Tahun/Unit/Tujuan)
function importFile_(link, copy, unit, kat, tahun, nomor, id) {
  var m = String(link || '').match(/[-\w]{25,}/);
  if (!m) return { url: '', hash: '', note: '' };
  try {
    var f = DriveApp.getFileById(m[0]);
    if (!copy) return { url: f.getUrl(), hash: '', note: '' };
    var ext = (f.getName().match(/\.[A-Za-z0-9]{1,8}$/) || [''])[0].toLowerCase();
    var name = 'Surat ' + String(nomor).replace(/[\/\\:*?"<>|]/g, '-') + ' [' + id + ']' + ext;
    var folder = sub_(sub_(sub_(DriveApp.getFolderById(prop_('FOLDER_ID')), String(tahun)), UNITS[unit].n), catsMap_()[kat].n);
    var cp = f.makeCopy(name, folder), hash = '';
    if (cp.getSize() <= MAX_MB * 1024 * 1024) hash = hex_(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, cp.getBlob().getBytes()));
    return { url: cp.getUrl(), hash: hash, note: '' };
  } catch (e) { return { url: '', hash: '', note: 'berkas tidak dapat disalin (' + e.message + '). Unggah manual dari tab Arsip.' }; }
}

// Dipanggil berulang oleh browser per kelompok baris. Aman diulang: baris yang sudah diimpor dilewati.
function importBatch(token, o, start, count) {
  auth_(token, true);
  if (!UNITS[o.unit]) throw new Error('Unit tidak valid');
  var tahun = parseInt(o.tahun, 10);
  if (!(tahun >= 2000 && tahun <= 2100)) throw new Error('Tahun tidak valid (2000-2100)');
  var ys = years_();
  if (ys.indexOf(tahun) < 0) { ys.push(tahun); ys.sort(function (a, b) { return b - a; }); PropertiesService.getScriptProperties().setProperty('YEARS', JSON.stringify(ys)); }
  check_(o.unit, o.kat, tahun);
  if (!prop_('FOLDER_ID')) throw new Error('Setup belum dijalankan.');
  var pj = null;
  if (o.pejabatId) { pj = pejabat_().filter(function (p) { return p.id === o.pejabatId; })[0]; if (!pj) throw new Error('Pejabat tidak ditemukan'); }
  var ss = openSrc_(o.src), src = ss.getSheetByName(o.sheet);
  if (!src) throw new Error('Tab sheet tidak ditemukan');
  var vals = src.getDataRange().getValues(), total = vals.length - 1, end = Math.min(start + count, total), m = o.map || {};
  function cell(r, k) { var i = parseInt(m[k], 10); return i >= 0 && i < r.length ? r[i] : ''; }

  var lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    var sh = sheet_(), ex = sh.getDataRange().getValues(), nos = {}, srcs = {};
    for (var x = 1; x < ex.length; x++) { nos[String(ex[x][2])] = 1; if (ex[x][16]) srcs[String(ex[x][16])] = 1; }
    var out = [], notes = [], skipped = 0, now = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'dd/MM/yyyy HH:mm');
    for (var i = start; i < end; i++) {
      var r = vals[i + 1], rowNum = i + 2;
      var no = String(cell(r, 'no')).trim(), perihal = String(cell(r, 'perihal')).trim();
      if (!no && !perihal) continue;
      if (!no) { skipped++; notes.push('Baris ' + rowNum + ': tidak ada nomor surat, dilewati'); continue; }
      var key = ss.getId() + '|' + o.sheet + '|' + rowNum;
      if (srcs[key]) { skipped++; continue; }
      var nomor = fmtNo_(catsMap_()[o.kat], o.unit, tahun, no, monthOf_(cell(r, 'tgl')));
      if (nos[nomor]) {
        var n = 2; while (nos[nomor + '-' + n]) n++;
        notes.push('Baris ' + rowNum + ': nomor ' + nomor + ' sudah ada, disimpan sebagai ' + nomor + '-' + n);
        nomor = nomor + '-' + n;
      }
      nos[nomor] = 1;
      var id = Utilities.getUuid().replace(/-/g, '').slice(0, 8).toUpperCase();
      var f = importFile_(String(cell(r, 'file')), !!o.copyFiles, o.unit, o.kat, tahun, nomor, id);
      if (f.note) notes.push('Baris ' + rowNum + ' (' + nomor + '): ' + f.note);
      out.push([id, now, nomor, parseTgl_(cell(r, 'tgl')), perihal || '-', String(cell(r, 'tujuan')).trim() || '-', pj ? pj.nama : '-', pj ? pj.jabatan : 'Pejabat berwenang',
                UNITS[o.unit].n, o.unit, 'VALID', f.hash, f.url, String(tahun), o.kat, 'YA', key]);
    }
    if (out.length) sh.getRange(sh.getLastRow() + 1, 1, out.length, HEAD.length).setValues(out);
    return { next: end, total: total, imported: out.length, skipped: skipped, notes: notes, years: years_() };
  } finally { lock.releaseLock(); }
}

function getQr(token, id) {
  auth_(token, true);
  var l = getLetter_(id);
  if (!l) throw new Error('Surat tidak ditemukan');
  var url = 'https://api.qrserver.com/v1/create-qr-code/?size=600x600&margin=16&data=' + encodeURIComponent(baseUrl_() + '?v=' + l.ID);
  var blob = UrlFetchApp.fetch(url).getBlob();
  return { dataUri: 'data:image/png;base64,' + Utilities.base64Encode(blob.getBytes()), name: 'QR-' + l.Nomor.replace(/[\/\\:*?"<>|\s]/g, '-') + '.png' };
}
function deleteLetter(token, id) {
  auth_(token, true);
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var sh = sheet_(), rows = sh.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === id) {
        var fid = fileIdOf_(rows[i][12]);
        if (fid) { try { DriveApp.getFileById(fid).setTrashed(true); } catch (e) {} }
        sh.deleteRow(i + 1);
        return true;
      }
    }
    throw new Error('Surat tidak ditemukan');
  } finally { lock.releaseLock(); }
}
function toggleRevoke(token, id) {
  auth_(token, true);
  var sh = sheet_(), rows = sh.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === id) {
      var s = rows[i][10] === 'VALID' ? 'DICABUT' : 'VALID';
      sh.getRange(i + 1, 11).setValue(s);
      return s;
    }
  }
  throw new Error('Surat tidak ditemukan');
}

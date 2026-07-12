// FASE 2b: identitas request-scoped, di-set oleh apiCall() dari session token.
// Null di luar konteks dispatcher.
var __REQ_USER_ = null;

function getCurrentUser() {
  // Tidy-up: identitas TUNGGAL dari session token (di-set apiCall → __REQ_USER_).
  // Fallback sesi Google dihapus → V2 tertutup 100% (pada deploy anonim,
  // Session tak lagi bisa mengklaim identitas apa pun).
  return __REQ_USER_ || null;
}

function getGoogleEmail() {
  try { return Session.getActiveUser().getEmail() || ''; } catch(e) { return ''; }
}

// ── HAK AKSES (PERMISSION) ──────────────────────────────
// Matriks { role: { cap: bool } }, gabungan default + override sheet "Hak Akses".
function getPermMatrix_() {
  var cache = CacheService.getScriptCache();
  var cached = cache.get('perm_matrix');
  if (cached) { try { return JSON.parse(cached); } catch(e) {} }

  var matrix = getDefaultPermMatrix_(); // { role: { cap: bool } }
  try {
    var ss = SpreadsheetApp.openById(getSpreadsheetId());
    var sheet = ss.getSheetByName(CONFIG.SHEETS.HAK_AKSES);
    if (sheet && sheet.getLastRow() > 1) {
      var rows = sheet.getDataRange().getValues();
      var header = rows[0];
      // kolom role mulai index 2 (0=Capability, 1=Keterangan)
      var roleCols = {};
      for (var c = 2; c < header.length; c++) {
        var rname = String(header[c] || '').trim();
        if (rname) roleCols[rname] = c;
      }
      for (var i = 1; i < rows.length; i++) {
        var cap = String(rows[i][0] || '').trim();
        if (!cap) continue;
        Object.keys(roleCols).forEach(function(role) {
          if (!matrix[role]) matrix[role] = {};
          var v = rows[i][roleCols[role]];
          matrix[role][cap] = (v === true || String(v).toUpperCase() === 'TRUE' || v === 1 || String(v) === '1');
        });
      }
    }
  } catch(e) {}

  // ADMIN selalu penuh (anti-lockout)
  if (matrix[CONFIG.ROLES.ADMIN]) {
    getCapabilities_().forEach(function(c) { matrix[CONFIG.ROLES.ADMIN][c.code] = true; });
  }
  try { cache.put('perm_matrix', JSON.stringify(matrix), 300); } catch(e) {}
  return matrix;
}

function getPermsForRole_(role) {
  var matrix = getPermMatrix_();
  return matrix[role] || {};
}

function userCan_(role, cap) {
  if (role === CONFIG.ROLES.ADMIN) return true;
  var matrix = getPermMatrix_();
  return !!(matrix[role] && matrix[role][cap]);
}

// FASE 5.1: enforcement input per-arah (masuk/keluar) dengan fallback ke
// kapabilitas lama 'trx.input' bila matriks tersimpan belum punya kolom baru.
// tipe: 'masuk' → trx.input.masuk; 'keluar'/'mutasi' → trx.input.keluar.
function userCanInput_(role, tipe) {
  if (role === CONFIG.ROLES.ADMIN) return true;
  var cap = (tipe === 'masuk') ? 'trx.input.masuk' : 'trx.input.keluar';
  var matrix = getPermMatrix_();
  var row = matrix[role] || {};
  // Bila kapabilitas baru sudah terdefinisi di matriks, pakai itu.
  if (Object.prototype.hasOwnProperty.call(row, cap)) return !!row[cap];
  // Kompatibilitas: matriks lama hanya punya 'trx.input'.
  return !!row['trx.input'];
}

function requirePermInput_(tipe) {
  var user = getCurrentUser();
  if (!user) return { success: false, message: 'Belum login' };
  if (!userCanInput_(user.role, tipe)) return { success: false, message: 'Akses ditolak' };
  return { success: true, user: user };
}

// Seperti checkAuth tapi cek capability tertentu.
function requirePerm(cap) {
  var user = getCurrentUser();
  if (!user) return { success: false, message: 'Belum login' };
  if (!userCan_(user.role, cap)) return { success: false, message: 'Akses ditolak' };
  return { success: true, user: user };
}

function getUserList_() {
  var cache = CacheService.getScriptCache();
  var cached = cache.get('user_list');
  if (cached) {
    try { return JSON.parse(cached); } catch(e) {}
  }
  var ss = SpreadsheetApp.openById(getSpreadsheetId());
  var sheet = ss.getSheetByName(CONFIG.SHEETS.USER);
  if (!sheet) return [];
  var rows = sheet.getDataRange().getValues();
  var h = headerMap_(rows[0]);
  var list = [];
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][0]) list.push({
      email: String(hGet_(rows[i], h, 'email', 0) || ''),
      nama: String(hGet_(rows[i], h, 'nama', 1) || ''),
      role: String(hGet_(rows[i], h, 'role', 2) || ''),
      status: String(hGet_(rows[i], h, 'status', 3) || 'Aktif'),
      created: toDateStr_(hGet_(rows[i], h, 'created', 4))
    });
  }
  try { cache.put('user_list', JSON.stringify(list), 300); } catch(e) {}
  return list;
}

function loginWithEmail(email) {
  try {
    // FASE 1: jalur ini TIDAK lagi menerbitkan identitas tanpa bukti (V1 di
    // docs/KEAMANAN.md). Login hanya berhasil bila email yang diklaim COCOK
    // dengan akun Google terverifikasi (Session). Akan digantikan alur OTP +
    // session token di Fase 2.
    var googleEmail = '';
    try { googleEmail = Session.getActiveUser().getEmail(); } catch(e) {}
    if (!googleEmail) {
      return { success: false, message: 'Tidak dapat memverifikasi akun Google. Login berbasis OTP akan tersedia (Fase 2).' };
    }
    var emailNorm = String(email || '').toLowerCase().trim();
    if (emailNorm && emailNorm !== googleEmail.toLowerCase().trim()) {
      return { success: false, message: 'Email tidak cocok dengan akun Google Anda.' };
    }
    // Identitas dipakai dari akun Google terverifikasi, bukan input mentah.
    var emailVerified = googleEmail.toLowerCase().trim();
    var users = getUserList_();
    for (var i = 0; i < users.length; i++) {
      if (users[i].email.toLowerCase().trim() === emailVerified) {
        var user = users[i];
        if (String(user.status || 'Aktif').toLowerCase() === 'nonaktif') {
          return { success: false, message: 'Akun nonaktif. Hubungi Admin.' };
        }
        logActivity(user.email, 'LOGIN', 'Login berhasil (Google terverifikasi)');
        user.perms = getPermsForRole_(user.role);
        return { success: true, user: user };
      }
    }
    return { success: false, message: 'Email tidak terdaftar. Hubungi Admin.' };
  } catch(e) {
    return { success: false, message: 'Error: ' + e.message };
  }
}

function logoutUser() {
  try {
    var user = getCurrentUser();
    if (user) logActivity(user.email, 'LOGOUT', 'Logout');
    PropertiesService.getUserProperties().deleteAllProperties();
    return { success: true };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function checkAuth(requiredRoles) {
  var user = getCurrentUser();
  if (!user) return { success: false, message: 'Belum login' };
  if (requiredRoles && requiredRoles.length > 0) {
    if (requiredRoles.indexOf(user.role) === -1) {
      return { success: false, message: 'Akses ditolak' };
    }
  }
  return { success: true, user: user };
}

// ── API untuk halaman admin "Kelola Hak Akses" ──────────
function getPermMatrix() {
  var auth = checkAuth([CONFIG.ROLES.ADMIN]);
  if (!auth.success) return auth;
  return {
    success: true,
    capabilities: getCapabilities_(),
    roles: getAllRoles_(),
    matrix: getPermMatrix_()
  };
}

// matrix: { role: { cap: bool } }
function savePermMatrix(matrix) {
  var auth = checkAuth([CONFIG.ROLES.ADMIN]);
  if (!auth.success) return auth;
  try {
    // T12: rekam DIFF matriks (sebelum → sesudah) sebagai aktivitas istimewa.
    var lama = getPermMatrix_();
    var ss = SpreadsheetApp.openById(getSpreadsheetId());
    var sheet = ss.getSheetByName(CONFIG.SHEETS.HAK_AKSES);
    if (!sheet) {
      sheet = ss.insertSheet(CONFIG.SHEETS.HAK_AKSES);
    }
    var roles = getAllRoles_();
    var caps = getCapabilities_();
    var header = ['Capability', 'Keterangan'].concat(roles);
    var out = [header];
    caps.forEach(function(c) {
      var row = [c.code, c.label];
      roles.forEach(function(role) {
        var val = (role === CONFIG.ROLES.ADMIN) ? true
          : !!(matrix && matrix[role] && matrix[role][c.code]);
        row.push(val);
      });
      out.push(row);
    });
    sheet.clearContents();
    sheet.getRange(1, 1, out.length, header.length).setValues(out);
    sheet.getRange(1, 1, 1, header.length)
      .setFontWeight('bold').setBackground('#1E40AF').setFontColor('#FFFFFF');
    sheet.setFrozenRows(1);

    try { CacheService.getScriptCache().remove('perm_matrix'); } catch(e) {}
    // Susun diff perubahan izin.
    var diffs = [];
    caps.forEach(function(c) {
      roles.forEach(function(role) {
        if (role === CONFIG.ROLES.ADMIN) return;
        var before = !!(lama[role] && lama[role][c.code]);
        var after = !!(matrix && matrix[role] && matrix[role][c.code]);
        if (before !== after) diffs.push(role + '.' + c.code + ': ' + (before ? 'ON' : 'OFF') + '→' + (after ? 'ON' : 'OFF'));
      });
    });
    logActivityWajib_(auth.user.email, 'PRIVILEGED_HAK_AKSES', diffs.length ? diffs.join('; ') : 'tanpa perubahan');
    return { success: true };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

// ── KELOLA USER (CRUD) ──────────────────────────────────
function getUserListAdmin() {
  var auth = requirePerm('user.manage');
  if (!auth.success) return auth;
  try { CacheService.getScriptCache().remove('user_list'); } catch(e) {}
  return { success: true, data: getUserList_(), roles: getAllRoles_() };
}

function countActiveAdmins_() {
  var users = getUserList_();
  var n = 0;
  for (var i = 0; i < users.length; i++) {
    if (users[i].role === CONFIG.ROLES.ADMIN && String(users[i].status || 'Aktif').toLowerCase() !== 'nonaktif') n++;
  }
  return n;
}

function findUserRow_(sheet, email) {
  var rows = sheet.getDataRange().getValues();
  var emailNorm = String(email).toLowerCase().trim();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).toLowerCase().trim() === emailNorm) return { rowIndex: i + 1, row: rows[i] };
  }
  return null;
}

function addUser(data) {
  var auth = requirePerm('user.manage');
  if (!auth.success) return auth;
  try {
    var email = String((data && data.email) || '').trim();
    if (!email) return { success: false, message: 'Email wajib diisi' };
    if (email.indexOf('@') < 0) return { success: false, message: 'Format email tidak valid' };
    var role = String((data && data.role) || '').trim();
    if (getAllRoles_().indexOf(role) < 0) return { success: false, message: 'Role tidak valid' };
    var users = getUserList_();
    for (var i = 0; i < users.length; i++) {
      if (users[i].email.toLowerCase().trim() === email.toLowerCase()) {
        return { success: false, message: 'Email sudah terdaftar' };
      }
    }
    var ss = SpreadsheetApp.openById(getSpreadsheetId());
    var sheet = ss.getSheetByName(CONFIG.SHEETS.USER);
    if (!sheet) return { success: false, message: 'Sheet Master User tidak ditemukan' };
    var status = (String((data && data.status) || 'Aktif').toLowerCase() === 'nonaktif') ? 'Nonaktif' : 'Aktif';
    sheet.appendRow([email, String((data && data.nama) || ''), role, status, new Date()]);
    try { CacheService.getScriptCache().remove('user_list'); } catch(e) {}
    logActivity(auth.user.email, 'USER', 'Tambah user: ' + email);
    return { success: true };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function updateUser(data) {
  var auth = requirePerm('user.manage');
  if (!auth.success) return auth;
  try {
    var email = String((data && data.email) || '').trim();
    if (!email) return { success: false, message: 'Email wajib diisi' };
    var role = String((data && data.role) || '').trim();
    if (getAllRoles_().indexOf(role) < 0) return { success: false, message: 'Role tidak valid' };
    var status = (String((data && data.status) || 'Aktif').toLowerCase() === 'nonaktif') ? 'Nonaktif' : 'Aktif';

    var ss = SpreadsheetApp.openById(getSpreadsheetId());
    var sheet = ss.getSheetByName(CONFIG.SHEETS.USER);
    if (!sheet) return { success: false, message: 'Sheet Master User tidak ditemukan' };
    var found = findUserRow_(sheet, email);
    if (!found) return { success: false, message: 'User tidak ditemukan' };

    var wasActiveAdmin = (String(found.row[2]) === CONFIG.ROLES.ADMIN && String(found.row[3] || 'Aktif').toLowerCase() !== 'nonaktif');
    var willBeActiveAdmin = (role === CONFIG.ROLES.ADMIN && status === 'Aktif');
    if (wasActiveAdmin && !willBeActiveAdmin && countActiveAdmins_() <= 1) {
      return { success: false, message: 'Tidak bisa: harus ada minimal 1 ADMIN aktif' };
    }

    var h = headerMap_(sheet.getDataRange().getValues()[0]);
    sheet.getRange(found.rowIndex, (h['nama'] !== undefined ? h['nama'] : 1) + 1).setValue(String((data && data.nama) || ''));
    sheet.getRange(found.rowIndex, (h['role'] !== undefined ? h['role'] : 2) + 1).setValue(role);
    sheet.getRange(found.rowIndex, (h['status'] !== undefined ? h['status'] : 3) + 1).setValue(status);
    try { CacheService.getScriptCache().remove('user_list'); } catch(e) {}
    logActivity(auth.user.email, 'USER', 'Update user: ' + email);
    return { success: true };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function setUserStatus(email, status) {
  var auth = requirePerm('user.manage');
  if (!auth.success) return auth;
  try {
    email = String(email || '').trim();
    var st = (String(status).toLowerCase() === 'nonaktif') ? 'Nonaktif' : 'Aktif';
    var ss = SpreadsheetApp.openById(getSpreadsheetId());
    var sheet = ss.getSheetByName(CONFIG.SHEETS.USER);
    if (!sheet) return { success: false, message: 'Sheet Master User tidak ditemukan' };
    var found = findUserRow_(sheet, email);
    if (!found) return { success: false, message: 'User tidak ditemukan' };
    if (st === 'Nonaktif' && String(found.row[2]) === CONFIG.ROLES.ADMIN && countActiveAdmins_() <= 1) {
      return { success: false, message: 'Tidak bisa: harus ada minimal 1 ADMIN aktif' };
    }
    var h = headerMap_(sheet.getDataRange().getValues()[0]);
    sheet.getRange(found.rowIndex, (h['status'] !== undefined ? h['status'] : 3) + 1).setValue(st);
    try { CacheService.getScriptCache().remove('user_list'); } catch(e) {}
    logActivity(auth.user.email, 'USER', 'Set status ' + email + ' = ' + st);
    return { success: true };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function deleteUser(email) {
  var auth = requirePerm('user.manage');
  if (!auth.success) return auth;
  try {
    email = String(email || '').trim();
    if (email.toLowerCase() === String(auth.user.email).toLowerCase()) {
      return { success: false, message: 'Tidak bisa menghapus akun sendiri' };
    }
    var ss = SpreadsheetApp.openById(getSpreadsheetId());
    var sheet = ss.getSheetByName(CONFIG.SHEETS.USER);
    if (!sheet) return { success: false, message: 'Sheet Master User tidak ditemukan' };
    var found = findUserRow_(sheet, email);
    if (!found) return { success: false, message: 'User tidak ditemukan' };
    var isActiveAdmin = (String(found.row[2]) === CONFIG.ROLES.ADMIN && String(found.row[3] || 'Aktif').toLowerCase() !== 'nonaktif');
    if (isActiveAdmin && countActiveAdmins_() <= 1) {
      return { success: false, message: 'Tidak bisa: harus ada minimal 1 ADMIN aktif' };
    }
    sheet.deleteRow(found.rowIndex);
    try { CacheService.getScriptCache().remove('user_list'); } catch(e) {}
    logActivity(auth.user.email, 'USER', 'Hapus user: ' + email);
    return { success: true };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

// ════════════════════════════════════════════════════════
// FASE 2a — HELPER KEAMANAN (kripto, OTP, perangkat, PIN, sesi)
// Semua fungsi di bawah bersifat helper internal (akhiran _).
// Belum di-wire ke checkAuth/requirePerm (itu Fase 2b) & belum
// dipanggil frontend (Fase 2c). Rahasia hanya disimpan sebagai hash.
// ════════════════════════════════════════════════════════

// TTL (milidetik)
var SEC_OTP_TTL_MS_ = 5 * 60 * 1000;           // OTP berlaku 5 menit
var SEC_SESSION_TTL_MS_ = 12 * 60 * 60 * 1000; // sesi berlaku 12 jam

// FASE 4: rate-limit & lockout (CacheService, per email)
var SEC_MAX_FAIL_ = 5;         // maks percobaan gagal sebelum terkunci
var SEC_LOCK_TTL_S_ = 15 * 60; // durasi kunci 15 menit (detik)

function _authFailKey_(email) { return 'authfail_' + String(email || '').toLowerCase().trim(); }
function _authLockKey_(email) { return 'authlock_' + String(email || '').toLowerCase().trim(); }

// Cek apakah email sedang terkunci akibat terlalu banyak percobaan gagal.
function isAuthLocked_(email) {
  try { return !!CacheService.getScriptCache().get(_authLockKey_(email)); } catch(e) { return false; }
}
// Catat 1 percobaan gagal; kunci 15 menit bila mencapai batas.
function recordAuthFail_(email) {
  try {
    var c = CacheService.getScriptCache();
    var n = parseInt(c.get(_authFailKey_(email)) || '0', 10) + 1;
    c.put(_authFailKey_(email), String(n), SEC_LOCK_TTL_S_);
    if (n >= SEC_MAX_FAIL_) c.put(_authLockKey_(email), '1', SEC_LOCK_TTL_S_);
    return n;
  } catch(e) { return 0; }
}
// Reset hitungan gagal & kunci setelah sukses.
function clearAuthFail_(email) {
  try { var c = CacheService.getScriptCache(); c.remove(_authFailKey_(email)); c.remove(_authLockKey_(email)); } catch(e) {}
}
// Throttle pengiriman OTP: cooldown 60 dtk + maks 5/jam per email (anti-spam).
function _otpThrottleOk_(email) {
  try {
    var norm = String(email || '').toLowerCase().trim();
    var c = CacheService.getScriptCache();
    if (c.get('otpcd_' + norm)) return false;
    var cnt = parseInt(c.get('otpcnt_' + norm) || '0', 10);
    if (cnt >= 5) return false;
    c.put('otpcd_' + norm, '1', 60);
    c.put('otpcnt_' + norm, String(cnt + 1), 3600);
    return true;
  } catch(e) { return true; }
}

// Pepper rahasia dari Script Properties (dibuat sekali bila belum ada).
function getPepper_() {
  var props = PropertiesService.getScriptProperties();
  var p = props.getProperty('SECURITY_PEPPER');
  if (!p) {
    p = Utilities.getUuid() + Utilities.getUuid();
    props.setProperty('SECURITY_PEPPER', p);
  }
  return p;
}

// Hash SHA-256 atas (pepper + salt + plain) → Base64. Tidak reversible.
function hashSecret_(plain, salt) {
  var raw = getPepper_() + '|' + String(salt || '') + '|' + String(plain);
  var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, raw, Utilities.Charset.UTF_8);
  return Utilities.base64Encode(bytes);
}

// Token acak kuat (±256-bit) untuk device/session token.
function randomToken_() {
  return (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, '');
}

// Cari user (aktif/nonaktif) berdasarkan email.
function getUserByEmail_(email) {
  var norm = String(email || '').toLowerCase().trim();
  var users = getUserList_();
  for (var i = 0; i < users.length; i++) {
    if (users[i].email.toLowerCase().trim() === norm) return users[i];
  }
  return null;
}

// ── OTP (bootstrap perangkat baru) ──────────────────────
// Kirim OTP 6 digit ke email terdaftar. Pesan netral (tidak
// membocorkan apakah email ada). Hash OTP disimpan di CacheService.
function requestOtp_(email) {
  var neutral = { success: true, message: 'Jika email terdaftar, kode OTP telah dikirim.' };
  try {
    var norm = String(email || '').toLowerCase().trim();
    if (!norm) return neutral;
    var user = getUserByEmail_(norm);
    if (!user || String(user.status || 'Aktif').toLowerCase() === 'nonaktif') return neutral;
    // FASE 4: jangan kirim bila akun terkunci atau throttle (pesan tetap netral).
    if (isAuthLocked_(norm) || !_otpThrottleOk_(norm)) return neutral;
    var otp = '' + Math.floor(100000 + Math.random() * 900000); // 6 digit
    var payload = JSON.stringify({ h: hashSecret_(otp, norm), e: (new Date().getTime() + SEC_OTP_TTL_MS_) });
    CacheService.getScriptCache().put('otp_' + norm, payload, Math.floor(SEC_OTP_TTL_MS_ / 1000));
    MailApp.sendEmail(user.email, 'Kode OTP Kas Kelompok',
      'Assalamu\'alaikum ' + (user.nama || '') + ',\n\n' +
      'Kode OTP Anda: ' + otp + '\n' +
      'Berlaku 5 menit. Jangan bagikan kode ini kepada siapa pun.\n\n— Kas Kelompok');
    logActivity(user.email, 'OTP_REQUEST', 'OTP dikirim');
    return neutral;
  } catch(e) {
    return { success: false, message: 'Gagal mengirim OTP. Coba lagi.' };
  }
}

// Verifikasi OTP → terbitkan token perangkat + session token.
function verifyOtp_(email, otp, deviceName) {
  try {
    var norm = String(email || '').toLowerCase().trim();
    if (isAuthLocked_(norm)) return { success: false, message: 'Terlalu banyak percobaan. Coba lagi dalam 15 menit.' };
    var cache = CacheService.getScriptCache();
    var raw = cache.get('otp_' + norm);
    if (!raw) return { success: false, message: 'OTP tidak ditemukan atau kadaluarsa. Minta OTP baru.' };
    var data = JSON.parse(raw);
    if (new Date().getTime() > data.e) { cache.remove('otp_' + norm); return { success: false, message: 'OTP kadaluarsa. Minta OTP baru.' }; }
    if (hashSecret_(String(otp || ''), norm) !== data.h) {
      recordAuthFail_(norm);
      return { success: false, message: 'Kode OTP salah.' };
    }
    cache.remove('otp_' + norm); // OTP sekali pakai
    clearAuthFail_(norm);
    var user = getUserByEmail_(norm);
    if (!user || String(user.status || 'Aktif').toLowerCase() === 'nonaktif') return { success: false, message: 'Akun tidak aktif.' };
    var dev = issueDeviceToken_(user.email, deviceName);
    var sess = issueSession_(user.email, dev.deviceId);
    logActivity(user.email, 'LOGIN_OTP', 'Perangkat baru: ' + (deviceName || '-'));
    user.perms = getPermsForRole_(user.role);
    return { success: true, sessionToken: sess.sessionToken, deviceToken: dev.deviceToken, expiresAt: sess.expiresAt, user: user };
  } catch(e) {
    return { success: false, message: 'Error: ' + e.message };
  }
}

// ── PERANGKAT TERPERCAYA ────────────────────────────────
// Terbitkan token perangkat; simpan HANYA hash-nya.
function issueDeviceToken_(email, deviceName) {
  var ss = getSS_();
  var sheet = ss.getSheetByName(CONFIG.SHEETS.PERANGKAT);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEETS.PERANGKAT);
    sheet.appendRow(['ID', 'Email', 'Device Hash', 'Nama Perangkat', 'Dibuat', 'Terakhir Dipakai', 'Status']);
  }
  var token = randomToken_();
  var deviceId = generateID('DEV');
  var norm = String(email).toLowerCase().trim();
  var now = new Date();
  sheet.appendRow([deviceId, norm, hashSecret_(token, norm), deviceName || 'Perangkat', now, now, 'Aktif']);
  return { deviceToken: token, deviceId: deviceId };
}

// Cocokkan token perangkat ke sheet Perangkat (status Aktif).
function verifyDevice_(email, deviceToken) {
  try {
    var norm = String(email || '').toLowerCase().trim();
    if (!norm || !deviceToken) return { ok: false };
    var hash = hashSecret_(String(deviceToken), norm);
    var ss = getSS_();
    var sheet = ss.getSheetByName(CONFIG.SHEETS.PERANGKAT);
    if (!sheet) return { ok: false };
    var rows = sheet.getDataRange().getValues();
    var h = headerMap_(rows[0]);
    for (var i = 1; i < rows.length; i++) {
      if (String(hGet_(rows[i], h, 'email', 1)).toLowerCase().trim() !== norm) continue;
      if (String(hGet_(rows[i], h, 'devicehash', 2)) !== hash) continue;
      if (String(hGet_(rows[i], h, 'status', 6)) !== 'Aktif') return { ok: false };
      var colTerakhir = (h['terakhirdipakai'] !== undefined ? h['terakhirdipakai'] : 5) + 1;
      sheet.getRange(i + 1, colTerakhir).setValue(new Date());
      return { ok: true, deviceId: String(hGet_(rows[i], h, 'id', 0)) };
    }
    return { ok: false };
  } catch(e) {
    return { ok: false };
  }
}

// ── PIN (login harian di perangkat terpercaya) ──────────
// Set/ubah PIN 6 digit; simpan hash+salt di Master User.
function setPin_(email, pin) {
  try {
    var norm = String(email || '').toLowerCase().trim();
    if (!/^\d{6}$/.test(String(pin || ''))) return { success: false, message: 'PIN harus 6 digit angka.' };
    var ss = getSS_();
    var sheet = ss.getSheetByName(CONFIG.SHEETS.USER);
    if (!sheet) return { success: false, message: 'Master User tidak ditemukan.' };
    var found = findUserRow_(sheet, norm);
    if (!found) return { success: false, message: 'User tidak ditemukan.' };
    var salt = randomToken_();
    var h = headerMap_(sheet.getDataRange().getValues()[0]);
    var colHash = (h['pinhash'] !== undefined ? h['pinhash'] : 5) + 1;
    var colSalt = (h['pinsalt'] !== undefined ? h['pinsalt'] : 6) + 1;
    sheet.getRange(found.rowIndex, colHash).setValue(hashSecret_(String(pin), salt));
    sheet.getRange(found.rowIndex, colSalt).setValue(salt);
    try { CacheService.getScriptCache().remove('user_list'); } catch(e) {}
    logActivity(norm, 'SET_PIN', 'PIN diperbarui');
    return { success: true };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

// Cek PIN. Kembalikan {ok} atau {ok:false, noPin:true} bila belum diset.
function verifyPin_(email, pin) {
  var norm = String(email || '').toLowerCase().trim();
  var ss = getSS_();
  var sheet = ss.getSheetByName(CONFIG.SHEETS.USER);
  if (!sheet) return { ok: false };
  var rows = sheet.getDataRange().getValues();
  var h = headerMap_(rows[0]);
  for (var i = 1; i < rows.length; i++) {
    if (String(hGet_(rows[i], h, 'email', 0)).toLowerCase().trim() !== norm) continue;
    var hash = String(hGet_(rows[i], h, 'pinhash', 5) || '');
    var salt = String(hGet_(rows[i], h, 'pinsalt', 6) || '');
    if (!hash) return { ok: false, noPin: true };
    return { ok: (hashSecret_(String(pin || ''), salt) === hash) };
  }
  return { ok: false };
}

// Login harian: perangkat terpercaya + PIN → session token.
function loginWithPin_(email, pin, deviceToken) {
  try {
    var norm = String(email || '').toLowerCase().trim();
    if (isAuthLocked_(norm)) return { success: false, message: 'Terlalu banyak percobaan. Coba lagi dalam 15 menit.' };
    var dev = verifyDevice_(email, deviceToken);
    if (!dev.ok) return { success: false, message: 'Perangkat tidak dikenali. Silakan verifikasi via OTP.' };
    var pinCheck = verifyPin_(email, pin);
    if (pinCheck.noPin) return { success: false, message: 'PIN belum diatur. Masuk via OTP lalu atur PIN.' };
    if (!pinCheck.ok) {
      recordAuthFail_(norm);
      return { success: false, message: 'PIN salah.' };
    }
    clearAuthFail_(norm);
    var user = getUserByEmail_(email);
    if (!user || String(user.status || 'Aktif').toLowerCase() === 'nonaktif') return { success: false, message: 'Akun tidak aktif.' };
    var sess = issueSession_(user.email, dev.deviceId);
    logActivity(user.email, 'LOGIN_PIN', 'Perangkat: ' + dev.deviceId);
    user.perms = getPermsForRole_(user.role);
    return { success: true, sessionToken: sess.sessionToken, expiresAt: sess.expiresAt, user: user };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

// ── SESI ────────────────────────────────────────────────
// Terbitkan session token; simpan HANYA hash + kadaluarsa (12 jam).
function issueSession_(email, deviceId) {
  var ss = getSS_();
  var sheet = ss.getSheetByName(CONFIG.SHEETS.SESI);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEETS.SESI);
    sheet.appendRow(['Token Hash', 'Email', 'Device ID', 'Dibuat', 'Kadaluarsa', 'Status']);
  }
  var token = randomToken_();
  var now = new Date();
  var exp = new Date(now.getTime() + SEC_SESSION_TTL_MS_);
  sheet.appendRow([hashSecret_(token, ''), String(email).toLowerCase().trim(), deviceId || '', now, exp, 'Aktif']);
  return { sessionToken: token, expiresAt: exp.getTime() };
}

// Verifikasi session token → user (email, role, perms) atau null.
// Sliding expiry: perpanjang bila sisa < 6 jam (mengurangi tulisan).
function verifySession_(token) {
  try {
    if (!token) return null;
    var hash = hashSecret_(String(token), '');
    var ss = getSS_();
    var sheet = ss.getSheetByName(CONFIG.SHEETS.SESI);
    if (!sheet) return null;
    var rows = sheet.getDataRange().getValues();
    var h = headerMap_(rows[0]);
    for (var i = 1; i < rows.length; i++) {
      if (String(hGet_(rows[i], h, 'tokenhash', 0)) !== hash) continue;
      if (String(hGet_(rows[i], h, 'status', 5)) !== 'Aktif') return null;
      var expVal = hGet_(rows[i], h, 'kadaluarsa', 4);
      var expMs = (expVal instanceof Date) ? expVal.getTime() : new Date(expVal).getTime();
      var nowMs = new Date().getTime();
      if (nowMs > expMs) {
        var colStat = (h['status'] !== undefined ? h['status'] : 5) + 1;
        sheet.getRange(i + 1, colStat).setValue('Kadaluarsa');
        return null;
      }
      var email = String(hGet_(rows[i], h, 'email', 1)).toLowerCase().trim();
      var user = getUserByEmail_(email);
      if (!user || String(user.status || 'Aktif').toLowerCase() === 'nonaktif') return null;
      if ((expMs - nowMs) < (6 * 60 * 60 * 1000)) {
        var colExp = (h['kadaluarsa'] !== undefined ? h['kadaluarsa'] : 4) + 1;
        sheet.getRange(i + 1, colExp).setValue(new Date(nowMs + SEC_SESSION_TTL_MS_));
      }
      user.perms = getPermsForRole_(user.role);
      return user;
    }
    return null;
  } catch(e) {
    return null;
  }
}

// Cabut sesi (logout / device revoke). Set status Dicabut pada baris token.
function invalidateSession_(token) {
  try {
    if (!token) return false;
    var hash = hashSecret_(String(token), '');
    var ss = getSS_();
    var sheet = ss.getSheetByName(CONFIG.SHEETS.SESI);
    if (!sheet) return false;
    var rows = sheet.getDataRange().getValues();
    var h = headerMap_(rows[0]);
    for (var i = 1; i < rows.length; i++) {
      if (String(hGet_(rows[i], h, 'tokenhash', 0)) !== hash) continue;
      var colStat = (h['status'] !== undefined ? h['status'] : 5) + 1;
      sheet.getRange(i + 1, colStat).setValue('Dicabut');
      return true;
    }
    return false;
  } catch(e) { return false; }
}

// ════════════════════════════════════════════════════════
// FASE 2b — RESOLUSI IDENTITAS TOKEN + DISPATCHER
// Sumber identitas TUNGGAL untuk panggilan token-based = session token.
// Pola dispatcher (docs/KEAMANAN.md §d): 1 endpoint apiCall menetapkan
// __REQ_USER_ dari token lalu memanggil fungsi tujuan. checkAuth() &
// requirePerm() TIDAK berubah tanda tangan — cukup membaca getCurrentUser()
// yang kini mengutamakan __REQ_USER_. Regresi minimal: ~65 endpoint tak
// disentuh; hanya frontend diarahkan lewat apiCall (Fase 2c).
// ════════════════════════════════════════════════════════

// Pengganti getCurrentUser lama untuk jalur token: identitas dari session token.
function resolveUser_(token) {
  return verifySession_(token);
}

// Denylist fungsi global yang TIDAK boleh dipanggil via dispatcher
// (setup/migrasi/util/rahasia). Helper privat (akhiran '_') otomatis ditolak.
var API_DENYLIST_ = {
  'apiCall': true, 'doGet': true, 'include': true,
  'loginWithEmail': true, 'getGoogleEmail': true, 'getCurrentUser': true,
  'setupSheets': true, 'repairSheetHeaders': true, 'migratePosSetoran': true,
  'migrasiKeamanan': true, 'migrasiPengendalian': true, 'bersihkanCache': true,
  'fmtRp': true, 'fmtTanggal': true, 'generateID': true, 'logActivity': true,
  'getSpreadsheetId': true, 'calculateSaldo': true, 'buildPDFHTML': true,
  'hitungJumlahBulan': true, 'mapPosNamaToIRColName': true, 'mapPosNamaToBukuIRCol': true,
  'testDashboard': true, 'testFindPeriode': true, 'testBukuIR': true
};

function isAllowedEndpoint_(fnName) {
  if (!fnName || typeof fnName !== 'string') return false;
  if (fnName.charAt(fnName.length - 1) === '_') return false; // helper privat
  if (API_DENYLIST_[fnName]) return false;
  return true;
}

// Dispatcher tunggal: semua panggilan frontend yang butuh auth lewat sini.
// token = session token; fnName = nama endpoint; argsArray = argumen asli.
// Endpoint tetap menegakkan requirePerm/checkAuth sendiri (defense-in-depth).
function apiCall(token, fnName, argsArray) {
  try {
    __REQ_USER_ = resolveUser_(token); // null bila token invalid/kadaluarsa
    if (!isAllowedEndpoint_(fnName)) return { success: false, message: 'Endpoint tidak dikenal.' };
    var fn = (typeof globalThis !== 'undefined') ? globalThis[fnName] : this[fnName];
    if (typeof fn !== 'function') return { success: false, message: 'Endpoint tidak ditemukan.' };
    var args = Array.isArray(argsArray) ? argsArray : [];
    return fn.apply(null, args);
  } catch(e) {
    return { success: false, message: e.message };
  } finally {
    __REQ_USER_ = null; // reset agar tidak bocor antar panggilan
  }
}

// ── Endpoint publik alur login (tanpa sesi — bootstrap) ──
function apiRequestOtp(email) {
  return requestOtp_(email);
}
function apiVerifyOtp(email, otp, deviceName) {
  return verifyOtp_(email, otp, deviceName);
}
function apiLoginPin(email, pin, deviceToken) {
  return loginWithPin_(email, pin, deviceToken);
}
// Set/ubah PIN — butuh session token valid.
function apiSetPin(token, pin) {
  var user = resolveUser_(token);
  if (!user) return { success: false, message: 'Sesi tidak valid. Silakan login ulang.' };
  return setPin_(user.email, pin);
}
// Logout — cabut sesi saat ini.
function apiLogout(token) {
  try {
    var user = resolveUser_(token);
    if (user) logActivity(user.email, 'LOGOUT', 'Logout (token)');
  } catch(e) {}
  invalidateSession_(token);
  return { success: true };
}
// Pulihkan sesi saat app dibuka: kembalikan user bila token masih valid.
function apiCurrentUser(token) {
  var user = resolveUser_(token);
  if (!user) return { success: false };
  return { success: true, user: user };
}

// Cek apakah email sudah punya PIN → tentukan jalur login (PIN vs OTP).
// Tidak membocorkan status: email tak terdaftar juga diarahkan ke OTP (netral).
function apiCheckLogin(email) {
  try {
    var norm = String(email || '').toLowerCase().trim();
    var user = getUserByEmail_(norm);
    if (!user || String(user.status || 'Aktif').toLowerCase() === 'nonaktif') return { ok: true, needOtp: true };
    var chk = verifyPin_(norm, '__cek__'); // hanya untuk cek ada/tidaknya PIN
    return { ok: true, needOtp: !!chk.noPin };
  } catch(e) {
    return { ok: true, needOtp: true };
  }
}

// Login harian TANPA token perangkat: cukup Email + PIN (server-verified).
// Faktor "perangkat" dilepas karena storage iframe GAS di HP sering terhapus;
// tetap ada proteksi lockout (5x gagal → kunci 15 menit).
function apiLoginEmailPin(email, pin) {
  try {
    var norm = String(email || '').toLowerCase().trim();
    if (isAuthLocked_(norm)) return { success: false, message: 'Terlalu banyak percobaan. Coba lagi dalam 15 menit.' };
    var chk = verifyPin_(norm, pin);
    if (chk.noPin) return { success: false, needOtp: true, message: 'PIN belum diatur. Masuk via OTP lalu atur PIN.' };
    if (!chk.ok) { recordAuthFail_(norm); return { success: false, message: 'PIN salah.' }; }
    clearAuthFail_(norm);
    var user = getUserByEmail_(norm);
    if (!user || String(user.status || 'Aktif').toLowerCase() === 'nonaktif') return { success: false, message: 'Akun tidak aktif.' };
    var sess = issueSession_(user.email, 'web');
    logActivity(user.email, 'LOGIN_PIN', 'Email + PIN');
    user.perms = getPermsForRole_(user.role);
    return { success: true, sessionToken: sess.sessionToken, user: user };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

// ════════════════════════════════════════════════════════
// FASE 4 — KELOLA PERANGKAT (Admin) + HOUSEKEEPING SESI
// Endpoint dipanggil lewat dispatcher apiCall (token-aware); tetap
// menegakkan requirePerm('user.manage') sendiri.
// ════════════════════════════════════════════════════════

// Cabut semua sesi milik sebuah perangkat (dipakai saat perangkat dicabut).
function invalidateSessionsByDevice_(deviceId) {
  try {
    if (!deviceId) return 0;
    var ss = getSS_();
    var sheet = ss.getSheetByName(CONFIG.SHEETS.SESI);
    if (!sheet) return 0;
    var rows = sheet.getDataRange().getValues();
    var h = headerMap_(rows[0]);
    var colStat = (h['status'] !== undefined ? h['status'] : 5) + 1;
    var n = 0;
    for (var i = 1; i < rows.length; i++) {
      if (String(hGet_(rows[i], h, 'deviceid', 2)) !== String(deviceId)) continue;
      if (String(hGet_(rows[i], h, 'status', 5)) !== 'Aktif') continue;
      sheet.getRange(i + 1, colStat).setValue('Dicabut');
      n++;
    }
    return n;
  } catch(e) { return 0; }
}

// Daftar perangkat terpercaya (semua user) untuk halaman Admin.
function getDeviceList() {
  var auth = requirePerm('user.manage');
  if (!auth.success) return auth;
  try {
    var ss = getSS_();
    var sheet = ss.getSheetByName(CONFIG.SHEETS.PERANGKAT);
    if (!sheet) return { success: true, data: [] };
    var rows = sheet.getDataRange().getValues();
    var h = headerMap_(rows[0]);
    var list = [];
    for (var i = 1; i < rows.length; i++) {
      if (!hGet_(rows[i], h, 'id', 0)) continue;
      list.push({
        id: String(hGet_(rows[i], h, 'id', 0)),
        email: String(hGet_(rows[i], h, 'email', 1) || ''),
        nama: String(hGet_(rows[i], h, 'namaperangkat', 3) || 'Perangkat'),
        dibuat: toDateStr_(hGet_(rows[i], h, 'dibuat', 4)),
        terakhir: toDateStr_(hGet_(rows[i], h, 'terakhirdipakai', 5)),
        status: String(hGet_(rows[i], h, 'status', 6) || '')
      });
    }
    // Terbaru dipakai di atas.
    list.sort(function(a, b) { return (a.terakhir < b.terakhir) ? 1 : (a.terakhir > b.terakhir ? -1 : 0); });
    return { success: true, data: list };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

// Cabut satu perangkat: Status=Dicabut + invalidasi sesi terkait.
function revokeDevice(deviceId) {
  var auth = requirePerm('user.manage');
  if (!auth.success) return auth;
  try {
    var ss = getSS_();
    var sheet = ss.getSheetByName(CONFIG.SHEETS.PERANGKAT);
    if (!sheet) return { success: false, message: 'Sheet Perangkat tidak ditemukan.' };
    var rows = sheet.getDataRange().getValues();
    var h = headerMap_(rows[0]);
    for (var i = 1; i < rows.length; i++) {
      if (String(hGet_(rows[i], h, 'id', 0)) !== String(deviceId)) continue;
      var colStat = (h['status'] !== undefined ? h['status'] : 6) + 1;
      sheet.getRange(i + 1, colStat).setValue('Dicabut');
      var n = invalidateSessionsByDevice_(deviceId);
      logActivity(auth.user.email, 'CABUT_PERANGKAT', deviceId + ' (' + n + ' sesi dicabut)');
      return { success: true, sesiDicabut: n };
    }
    return { success: false, message: 'Perangkat tidak ditemukan.' };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

// Housekeeping: tandai semua sesi kadaluarsa. Bisa dipanggil dari editor
// atau dijadwalkan trigger harian. verifySession_ juga menandai saat diakses.
function cleanupSessions_() {
  try {
    var ss = getSS_();
    var sheet = ss.getSheetByName(CONFIG.SHEETS.SESI);
    if (!sheet) return { success: true, expired: 0 };
    var rows = sheet.getDataRange().getValues();
    var h = headerMap_(rows[0]);
    var colStat = (h['status'] !== undefined ? h['status'] : 5) + 1;
    var nowMs = new Date().getTime();
    var n = 0;
    for (var i = 1; i < rows.length; i++) {
      if (String(hGet_(rows[i], h, 'status', 5)) !== 'Aktif') continue;
      var expVal = hGet_(rows[i], h, 'kadaluarsa', 4);
      var expMs = (expVal instanceof Date) ? expVal.getTime() : new Date(expVal).getTime();
      if (nowMs > expMs) { sheet.getRange(i + 1, colStat).setValue('Kadaluarsa'); n++; }
    }
    return { success: true, expired: n };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

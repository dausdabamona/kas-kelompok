function getCurrentUser() {
  try {
    var props = PropertiesService.getUserProperties();
    var email = props.getProperty('userEmail');
    var nama = props.getProperty('userName');
    var role = props.getProperty('userRole');
    if (!email) return null;
    return { email: email, nama: nama, role: role, perms: getPermsForRole_(role) };
  } catch(e) {
    return null;
  }
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
  var list = [];
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][0]) list.push({ email: rows[i][0], nama: rows[i][1], role: rows[i][2] });
  }
  try { cache.put('user_list', JSON.stringify(list), 300); } catch(e) {}
  return list;
}

function loginWithEmail(email) {
  try {
    if (!email) return { success: false, message: 'Email tidak boleh kosong' };
    var emailNorm = email.toLowerCase().trim();
    var users = getUserList_();
    for (var i = 0; i < users.length; i++) {
      if (users[i].email.toLowerCase().trim() === emailNorm) {
        var user = users[i];
        var props = PropertiesService.getUserProperties();
        props.setProperty('userEmail', user.email);
        props.setProperty('userName', user.nama);
        props.setProperty('userRole', user.role);
        logActivity(user.email, 'LOGIN', 'Login berhasil');
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
    logActivity(auth.user.email, 'HAK_AKSES', 'Update matriks hak akses');
    return { success: true };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

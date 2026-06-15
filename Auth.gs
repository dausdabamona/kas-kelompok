function getCurrentUser() {
  try {
    var props = PropertiesService.getUserProperties();
    var email = props.getProperty('userEmail');
    var nama = props.getProperty('userName');
    var role = props.getProperty('userRole');
    if (!email) return null;
    return { email: email, nama: nama, role: role };
  } catch(e) {
    return null;
  }
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

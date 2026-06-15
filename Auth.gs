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

function loginWithEmail(email) {
  try {
    if (!email) return { success: false, message: 'Email tidak boleh kosong' };
    var ss = SpreadsheetApp.openById(getSpreadsheetId());
    var sheet = ss.getSheetByName(CONFIG.SHEETS.USER);
    if (!sheet) return { success: false, message: 'Sheet User tidak ditemukan' };
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString().toLowerCase().trim() === email.toLowerCase().trim()) {
        var user = { email: data[i][0], nama: data[i][1], role: data[i][2] };
        var props = PropertiesService.getUserProperties();
        props.setProperty('userEmail', user.email);
        props.setProperty('userName', user.nama);
        props.setProperty('userRole', user.role);
        logActivity(user.email, 'LOGIN', 'Login berhasil');
        return { success: true, user: user };
      }
    }
    return { success: false, message: 'Email tidak terdaftar' };
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

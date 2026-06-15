const CONFIG = {
  SHEETS: {
    KELOMPOK: 'Master Kelompok',
    USER: 'Master User',
    PEMASUKAN: 'Master Pemasukan',
    PENGELUARAN: 'Master Pengeluaran',
    MUSYAWARAH: 'Master Musyawaroh',
    PERIOD: 'Master Period',
    ANGGOTA: 'Anggota',
    INPUT_PENERIMAAN: 'Input Penerimaan',
    BUKU_IR: 'Detail Buku IR',
    INPUT_PENGELUARAN: 'Input Pengeluaran',
    INPUT_SETORAN: 'Input Setoran Bank',
    INPUT_SALDO_BANK: 'Input Saldo Bank',
    POS_SETORAN: 'Pos Setoran',
    SETORAN_DESA: 'Setoran Desa',
    BANK_DAILY: 'Bank Daily',
    BANK_PENDING: 'Bank Pending',
    SALDO_TUTUP_BUKU: 'Saldo Tutup Buku',
    LOG: 'Activity Log',
  },
  ROLES: {
    ADMIN: 'ADMIN',
    BENDAHARA_1: 'BENDAHARA_1',
    BENDAHARA_2: 'BENDAHARA_2',
    PENULIS: 'PENULIS'
  },
  STATUS: { OPEN: 'OPEN', CLOSED: 'CLOSED' },
};

function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Kas Kelompok')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function fmtRp(angka) {
  if (!angka || isNaN(angka)) return 'Rp 0';
  return 'Rp ' + Number(angka).toLocaleString('id-ID');
}

function fmtTanggal(tgl) {
  if (!tgl) return '-';
  try {
    var d = new Date(tgl);
    var months = ['Jan','Feb','Mar','Apr','Mei','Jun','Jul','Agu','Sep','Okt','Nov','Des'];
    return d.getDate() + ' ' + months[d.getMonth()] + ' ' + d.getFullYear();
  } catch(e) {
    return String(tgl);
  }
}

function generateID(prefix) {
  return prefix + '_' + new Date().getTime() + '_' + Math.random().toString(36).substr(2, 5).toUpperCase();
}

function logActivity(user, action, detail) {
  try {
    var ss = SpreadsheetApp.openById(getSpreadsheetId());
    var sheet = ss.getSheetByName(CONFIG.SHEETS.LOG);
    if (!sheet) {
      sheet = ss.insertSheet(CONFIG.SHEETS.LOG);
      sheet.appendRow(['Timestamp', 'User', 'Action', 'Detail']);
    }
    sheet.appendRow([new Date(), user || 'System', action || '', detail || '']);
  } catch(e) {
    // Log silently fails — don't break main flow
  }
}

function getSpreadsheetId() {
  return '1nR-NkKy4h-IB2D9_WJ8kUD1MMd1KHD811Olf33nIy9k';
}

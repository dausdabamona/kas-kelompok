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
    PEMBELAAN: 'Pembelaan',
    PATUNGAN: 'Terobosan Kelompok',
    TAGIHAN_PATUNGAN: 'Penerobosan',
    HAK_AKSES: 'Hak Akses',
    LOG: 'Activity Log',
    KAS_PENEROBOS: 'Kas Penerobos',
    SERAH_TERIMA: 'Serah Terima',
    // FASE 2: identitas berbasis token
    PERANGKAT: 'Perangkat',
    SESI: 'Sesi',
    // FASE 4: lampiran bukti
    LAMPIRAN: 'Lampiran',
  },
  ROLES: {
    ADMIN: 'ADMIN',
    BENDAHARA_1: 'BENDAHARA_1',
    BENDAHARA_2: 'BENDAHARA_2',
    PENULIS: 'PENULIS',
    PENEROBOS: 'PENEROBOS'
  },
  STATUS: { OPEN: 'OPEN', CLOSED: 'CLOSED' },
};

// ══════════════════════════════════════════════════════
// HAK AKSES (CAPABILITY) — definisi & default per role
// ══════════════════════════════════════════════════════
// Daftar capability (sumber tunggal). grup: 'lihat' | 'edit'
function getCapabilities_() {
  return [
    // Lihat (menu/UI)
    { code: 'view.bukuIR', label: 'Lihat: Buku IR', grup: 'lihat' },
    { code: 'view.kelolaBuku', label: 'Lihat: Kelola Buku', grup: 'lihat' },
    { code: 'view.setoranDesa', label: 'Lihat: Setoran Desa', grup: 'lihat' },
    { code: 'view.jamaah', label: 'Lihat: Jamaah', grup: 'lihat' },
    { code: 'view.laporanSetoran', label: 'Lihat: Laporan Setoran', grup: 'lihat' },
    { code: 'view.pemeriksaanSaldo', label: 'Lihat: Pemeriksaan Saldo', grup: 'lihat' },
    { code: 'view.laporanPDF', label: 'Lihat: Laporan PDF', grup: 'lihat' },
    { code: 'view.terobosan', label: 'Lihat: Terobosan Kelompok', grup: 'lihat' },
    { code: 'view.setting', label: 'Lihat: Setting', grup: 'lihat' },
    // Edit/Aksi (server-enforced)
    { code: 'trx.input', label: 'Input Transaksi (lama — mencakup masuk & keluar)', grup: 'edit' },
    { code: 'trx.input.masuk', label: 'Input Pemasukan', grup: 'edit' },
    { code: 'trx.input.keluar', label: 'Input Pengeluaran', grup: 'edit' },
    { code: 'bukuIR.input', label: 'Input Rincian Buku IR', grup: 'edit' },
    { code: 'jamaah.add', label: 'Tambah Jamaah', grup: 'edit' },
    { code: 'jamaah.edit', label: 'Edit Jamaah', grup: 'edit' },
    { code: 'jamaah.delete', label: 'Hapus Jamaah', grup: 'edit' },
    { code: 'terobosan.create', label: 'Buat Terobosan', grup: 'edit' },
    { code: 'terobosan.bayar', label: 'Bayar Penerobosan', grup: 'edit' },
    { code: 'terobosan.batal', label: 'Batal Bayar Penerobosan', grup: 'edit' },
    { code: 'grade.edit', label: 'Atur Grade Jamaah', grup: 'edit' },
    { code: 'trx.approve', label: 'Setujui Pengeluaran (maker-checker)', grup: 'edit' },
    { code: 'saldo.input', label: 'Input/Pemeriksaan Saldo', grup: 'edit' },
    { code: 'bank.input', label: 'Input Transaksi Bank', grup: 'edit' },
    { code: 'bank.manage', label: 'Kelola Bank (rekon/pending/daily)', grup: 'edit' },
    { code: 'setoran.realisasi', label: 'Realisasi Setoran Desa', grup: 'edit' },
    { code: 'pembelaan.manage', label: 'Kelola Pembelaan', grup: 'edit' },
    { code: 'master.manage', label: 'Kelola Master (Pos/Musyawaroh)', grup: 'edit' },
    { code: 'user.manage', label: 'Kelola User', grup: 'edit' },
    { code: 'trx.edit.tunai', label: 'Edit/Hapus Transaksi Tunai', grup: 'edit' },
    { code: 'trx.edit.bank', label: 'Edit/Hapus Transaksi Bank', grup: 'edit' },
    { code: 'periode.manage', label: 'Kelola Periode (Tutup/Buka)', grup: 'edit' },
    { code: 'view.kasPenerobos', label: 'Lihat: Kas Penerobos', grup: 'lihat' },
    { code: 'view.serahTerima', label: 'Lihat: Serah Terima', grup: 'lihat' },
    { code: 'penerobos.input', label: 'Input Kas Penerobos', grup: 'edit' },
    { code: 'penerobos.kelola', label: 'Kelola/Batalkan Kas Penerobos (duplikat)', grup: 'edit' },
    { code: 'serahterima.konfirmasi', label: 'Konfirmasi Serah Terima', grup: 'edit' }
  ];
}

// Default izin per role (meniru perilaku hardcode lama). ADMIN selalu true.
function getDefaultPermMatrix_() {
  var R = CONFIG.ROLES;
  var allRoles = [R.ADMIN, R.BENDAHARA_1, R.BENDAHARA_2, R.PENULIS, R.PENEROBOS];
  var nonPenerobos = [R.ADMIN, R.BENDAHARA_1, R.BENDAHARA_2, R.PENULIS];
  // map capability → daftar role yang diizinkan secara default
  var def = {
    'view.bukuIR': nonPenerobos,
    'view.kelolaBuku': nonPenerobos,
    'view.setoranDesa': [R.ADMIN, R.BENDAHARA_1, R.BENDAHARA_2],
    'view.jamaah': [R.ADMIN, R.BENDAHARA_1, R.PENULIS, R.PENEROBOS],
    'view.laporanSetoran': nonPenerobos,
    'view.pemeriksaanSaldo': [R.ADMIN, R.BENDAHARA_1, R.BENDAHARA_2],
    'view.laporanPDF': nonPenerobos,
    'view.terobosan': allRoles,
    'view.setting': [R.ADMIN],
    'trx.input': allRoles, // dipertahankan demi kompatibilitas; enforcement kini via masuk/keluar
    'trx.input.masuk': [R.ADMIN, R.BENDAHARA_1, R.BENDAHARA_2, R.PENULIS, R.PENEROBOS],
    'trx.input.keluar': [R.ADMIN, R.BENDAHARA_1, R.BENDAHARA_2, R.PENULIS],
    'bukuIR.input': [R.ADMIN, R.BENDAHARA_1, R.BENDAHARA_2, R.PENULIS],
    'jamaah.add': [R.ADMIN, R.BENDAHARA_1, R.PENULIS, R.PENEROBOS],
    'jamaah.edit': [R.ADMIN, R.BENDAHARA_1, R.PENULIS, R.PENEROBOS],
    'jamaah.delete': [R.ADMIN],
    'terobosan.create': [R.ADMIN, R.BENDAHARA_1, R.PENULIS],
    'terobosan.bayar': [R.ADMIN, R.BENDAHARA_1, R.BENDAHARA_2, R.PENEROBOS],
    'terobosan.batal': [R.ADMIN, R.BENDAHARA_1, R.PENEROBOS],
    'grade.edit': [R.ADMIN, R.BENDAHARA_1],
    'trx.approve': [R.ADMIN, R.BENDAHARA_1],
    'saldo.input': [R.ADMIN, R.BENDAHARA_1, R.BENDAHARA_2],
    'bank.input': [R.ADMIN, R.BENDAHARA_1, R.BENDAHARA_2],
    'bank.manage': [R.ADMIN, R.BENDAHARA_1],
    'setoran.realisasi': [R.ADMIN, R.BENDAHARA_1, R.BENDAHARA_2],
    'pembelaan.manage': [R.ADMIN, R.BENDAHARA_1, R.BENDAHARA_2],
    'master.manage': [R.ADMIN],
    'user.manage': [R.ADMIN],
    'trx.edit.tunai': [R.ADMIN, R.BENDAHARA_1],
    'trx.edit.bank': [R.ADMIN, R.BENDAHARA_2],
    'periode.manage': [R.ADMIN],
    'view.kasPenerobos': [R.ADMIN, R.PENEROBOS],
    'view.serahTerima': [R.ADMIN, R.BENDAHARA_1, R.BENDAHARA_2],
    'penerobos.input': [R.ADMIN, R.PENEROBOS],
    'penerobos.kelola': [R.ADMIN],
    'serahterima.konfirmasi': [R.ADMIN, R.BENDAHARA_1, R.BENDAHARA_2]
  };
  // bentuk matriks { role: { cap: bool } }
  var matrix = {};
  allRoles.forEach(function(role) { matrix[role] = {}; });
  getCapabilities_().forEach(function(c) {
    var allowed = def[c.code] || [];
    allRoles.forEach(function(role) {
      matrix[role][c.code] = (role === R.ADMIN) || (allowed.indexOf(role) >= 0);
    });
  });
  return matrix;
}

function getAllRoles_() {
  var R = CONFIG.ROLES;
  return [R.ADMIN, R.BENDAHARA_1, R.BENDAHARA_2, R.PENULIS, R.PENEROBOS];
}

function doGet(e) {
  // Konsol admin desktop terpisah: ?view=admin (khusus ADMIN, gate di server).
  var view = (e && e.parameter && e.parameter.view) ? String(e.parameter.view) : '';
  if (view === 'admin') {
    return HtmlService.createTemplateFromFile('AdminConsole')
      .evaluate()
      .setTitle('Admin Console · Kas Kelompok')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Kas Kelompok')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// URL web app (untuk membuka Admin Console desktop dari app mobile).
function getAppUrl() {
  try { return { success: true, url: ScriptApp.getService().getUrl() }; }
  catch(e) { return { success: false, message: e.message }; }
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

// FASE 3: kunci skrip untuk operasi uang agar atomik (mencegah race saldo
// saat dua submit bersamaan). Auth dijalankan DI LUAR lock; hanya bagian
// tulis dibungkus. Gagal ambil lock 15 dtk → pesan "sibuk".
function withLock_(fn) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
  } catch(e) {
    return { success: false, message: 'Sistem sedang sibuk, coba lagi sebentar.' };
  }
  try {
    return fn();
  } finally {
    try { lock.releaseLock(); } catch(e) {}
  }
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
  // Tidy-up (V6): utamakan Script Properties; fallback ke konstanta lama &
  // simpan sekali agar tidak hardcoded lagi ke depan.
  var FALLBACK = '1nR-NkKy4h-IB2D9_WJ8kUD1MMd1KHD811Olf33nIy9k';
  try {
    var props = PropertiesService.getScriptProperties();
    var id = props.getProperty('SPREADSHEET_ID');
    if (!id) { id = FALLBACK; props.setProperty('SPREADSHEET_ID', id); }
    return id;
  } catch(e) {
    return FALLBACK;
  }
}

// ══════════════════════════════════════════════════════
// Sumber tunggal definisi header semua sheet
// Dipakai oleh setupSheets() & repairSheetHeaders()
// ══════════════════════════════════════════════════════
function getSheetSchema_() {
  return [
    {
      name: CONFIG.SHEETS.KELOMPOK,
      headers: ['ID', 'Nama Kelompok', 'Alamat', 'Ketua', 'Bendahara', 'Tahun Berdiri', 'Keterangan'],
      sample: ['KEL-001', 'Kelompok Jamaah', '', '', '', '', '']
    },
    {
      name: CONFIG.SHEETS.USER,
      // FASE 2: 'PIN Hash' & 'PIN Salt' menyimpan HANYA hash PIN (SHA-256 + salt + pepper),
      // bukan PIN asli. Ditambahkan idempoten via migrasiKeamanan() untuk sheet lama.
      headers: ['Email', 'Nama', 'Role', 'Status', 'Created', 'PIN Hash', 'PIN Salt'],
      sample: ['dausdaba@gmail.com', 'Firdaus', 'ADMIN', 'Aktif', new Date(), '', '']
    },
    {
      name: CONFIG.SHEETS.PEMASUKAN,
      headers: ['Kode', 'Nama Pemasukan', 'Kategori', '% Kelompok', '% Desa', '% Daerah', 'Input Tipe', 'Status'],
      note: 'Kategori: Buku IR / Jamaah / Umum | Status: Aktif / Nonaktif'
    },
    {
      name: CONFIG.SHEETS.PENGELUARAN,
      headers: ['Kode', 'Nama Pengeluaran', 'Kategori', 'Status'],
      note: 'Status: Aktif / Nonaktif'
    },
    {
      name: CONFIG.SHEETS.MUSYAWARAH,
      headers: ['ID', 'Nama', 'Nilai', 'Keterangan'],
      note: 'Nilai dalam Rupiah (angka, tanpa Rp)'
    },
    {
      name: CONFIG.SHEETS.PERIOD,
      headers: ['Periode', 'Nama', 'Tgl Mulai', 'Tgl Tutup', 'Status', 'Saldo Awal Tunai', 'Saldo Awal Bank', 'Catatan'],
      note: 'Status: OPEN / CLOSED | Hanya boleh 1 baris OPEN'
    },
    {
      name: CONFIG.SHEETS.ANGGOTA,
      headers: ['ID', 'Nama', 'No Telp', 'Alamat', 'Status', 'IR', '1/10 IR', 'Index', 'Infak Daerah', 'Grade'],
      note: 'Status: Aktif / Nonaktif | IR s.d. Infak Daerah dalam Rupiah | Grade: A/B/C/D/E untuk Terobosan Kelompok'
    },
    {
      name: CONFIG.SHEETS.INPUT_PENERIMAAN,
      headers: ['ID', 'Periode ID', 'Jenis ID', 'Anggota ID', 'Tanggal', 'Nominal', 'Sumber Kas', 'Catatan', 'Created By', 'Created At', 'Status', 'Dibatalkan By', 'Dibatalkan At', 'Alasan Batal', 'No Bukti'],
      note: 'Sumber Kas: Tunai / Bank | Status: Aktif / Dibatalkan (soft delete) | JANGAN edit manual'
    },
    {
      name: CONFIG.SHEETS.BUKU_IR,
      headers: ['ID', 'Transaksi ID', 'Periode ID', 'Anggota ID', 'Tanggal', 'IR', 'IR 1/10', 'Cicilan', 'Infak Daerah', 'Index', 'Created By', 'Created At'],
      note: 'JANGAN edit manual — diisi otomatis dari form Buku IR | Total = derived (dihitung saat baca)'
    },
    {
      name: CONFIG.SHEETS.INPUT_PENGELUARAN,
      headers: ['ID', 'Periode ID', 'Jenis ID', 'Tanggal', 'Nominal', 'Sumber Kas', 'Catatan', 'Created By', 'Created At', 'Status', 'Dibatalkan By', 'Dibatalkan At', 'Alasan Batal', 'No Bukti', 'Status Approval', 'Disetujui By', 'Disetujui At'],
      note: 'Sumber Kas: Tunai / Bank | Status: Aktif / Dibatalkan (soft delete) | JANGAN edit manual'
    },
    {
      name: CONFIG.SHEETS.INPUT_SETORAN,
      headers: ['ID', 'Periode ID', 'Tanggal', 'Nominal', 'Arah', 'Created By', 'Created At', 'Status', 'Dibatalkan By', 'Dibatalkan At', 'Alasan Batal'],
      note: 'Arah: setor (Tunai→Bank) / tarik (Bank→Tunai) | Status: Aktif / Dibatalkan | JANGAN edit manual'
    },
    {
      name: CONFIG.SHEETS.INPUT_SALDO_BANK,
      headers: ['ID', 'Periode ID', 'Tanggal', 'Saldo', 'Catatan', 'Created By', 'Created At'],
      note: 'JANGAN edit manual'
    },
    {
      name: CONFIG.SHEETS.POS_SETORAN,
      headers: ['ID', 'Nama', 'Sumber Tipe', 'Sumber Ref', 'Status', 'Target', 'Pengeluaran Ref'],
      note: 'Sumber Tipe: bukuir / pemasukan / manual | Sumber Ref: kolom Buku IR (ir/ir10/cicilan/infakdaerah/index) atau Kode Pemasukan | Pengeluaran Ref: Kode Master Pengeluaran (FK) untuk mencatat realisasi setoran sebagai pengeluaran kas | Status: Aktif / Nonaktif'
    },
    {
      name: CONFIG.SHEETS.SETORAN_DESA,
      headers: ['ID', 'PosID', 'PeriodeID', 'Realisasi', 'Catatan', 'Sumber Kas', 'Pengeluaran ID', 'Updated By', 'Updated At'],
      note: 'JANGAN edit manual — diisi dari halaman Setoran Desa | Pengeluaran ID = FK ke Input Pengeluaran (otomatis)'
    },
    {
      name: CONFIG.SHEETS.BANK_DAILY,
      headers: ['ID', 'Periode ID', 'Tanggal', 'Saldo Awal', 'Pemasukan', 'Pengeluaran', 'Saldo Akhir Teoritis', 'Saldo Akhir Actual', 'Selisih', 'Status', 'Catatan', 'Last Updated'],
      note: 'Status: Balance / Selisih | JANGAN edit manual'
    },
    {
      name: CONFIG.SHEETS.BANK_PENDING,
      headers: ['ID', 'Periode ID', 'Tanggal', 'Keterangan', 'Nominal', 'Sumber', 'Status', 'Tanggal Found', 'Catatan', 'Last Updated'],
      note: 'Status: Pending / Found / Adjusted / Waived'
    },
    {
      name: CONFIG.SHEETS.SALDO_TUTUP_BUKU,
      headers: ['ID', 'Periode ID', 'Tanggal Tutup', 'Saldo Tunai Akhir', 'Saldo Bank Akhir', 'Total Kas', 'Status', 'Catatan', 'Created By', 'Created At', 'Saldo Tunai Sistem', 'Saldo Bank Sistem', 'Selisih Tunai', 'Selisih Bank', 'Selisih Total', 'Alasan Selisih', 'Arsip File ID', 'Arsip URL', 'Arsip Hash'],
      note: 'JANGAN edit manual — diisi saat Tutup Buku | Saldo *Akhir = aktual (cash count), *Sistem = hitungan aplikasi, Selisih = aktual − sistem'
    },
    {
      name: CONFIG.SHEETS.PATUNGAN,
      headers: ['ID', 'Nama', 'Tanggal', 'PeriodeID', 'Status', 'GradeConfig', 'Catatan', 'CreatedBy', 'CreatedAt'],
      note: 'Terobosan Kelompok — JANGAN edit manual | GradeConfig = JSON nominal per grade'
    },
    {
      name: CONFIG.SHEETS.TAGIHAN_PATUNGAN,
      headers: ['ID', 'PatunganID', 'AnggotaID', 'AnggotaNama', 'Grade', 'Nominal', 'StatusBayar', 'TanggalBayar', 'Catatan', 'CreatedBy', 'CreatedAt'],
      note: 'Penerobosan per jamaah — JANGAN edit manual | StatusBayar: Belum / Lunas'
    },
    {
      name: CONFIG.SHEETS.KAS_PENEROBOS,
      headers: ['ID', 'Periode ID', 'Tanggal', 'Jenis ID', 'Anggota ID', 'Nominal', 'Sumber Kas', 'Catatan', 'Penerobos Email', 'Status', 'Serah Terima ID', 'Created At'],
      note: 'Status: Aktif / Diserahkan | Sumber Kas: Tunai / Bank | JANGAN edit manual'
    },
    {
      name: CONFIG.SHEETS.SERAH_TERIMA,
      headers: ['ID', 'Periode ID', 'Tanggal Serah', 'Penerobos Email', 'Total', 'Sumber Tujuan', 'Catatan', 'Status', 'Dikonfirmasi By', 'Dikonfirmasi At', 'Penerimaan ID', 'Created At'],
      note: 'Sumber Tujuan: Tunai / Bank | Status: Menunggu / Dikonfirmasi | JANGAN edit manual'
    },
    {
      name: CONFIG.SHEETS.HAK_AKSES,
      headers: ['Capability', 'Keterangan'].concat(getAllRoles_()),
      note: 'Matriks hak akses per role — diatur lewat menu Setting → Kelola Hak Akses. TRUE = boleh.'
    },
    {
      name: CONFIG.SHEETS.LOG,
      headers: ['Timestamp', 'User', 'Action', 'Detail'],
      note: 'Log otomatis — JANGAN edit manual'
    },
    {
      name: CONFIG.SHEETS.LAMPIRAN,
      headers: ['ID', 'Transaksi ID', 'Tipe', 'Nama File', 'Drive File ID', 'URL', 'Diunggah By', 'Diunggah At', 'Status'],
      note: 'FASE 4: bukti/lampiran transaksi (foto nota). Status: Aktif / Dihapus'
    },
    {
      // FASE 2: perangkat terpercaya. Simpan HANYA hash token perangkat.
      name: CONFIG.SHEETS.PERANGKAT,
      headers: ['ID', 'Email', 'Device Hash', 'Nama Perangkat', 'Dibuat', 'Terakhir Dipakai', 'Status'],
      note: 'Status: Aktif / Dicabut | Device Hash = hash token perangkat (bukan token asli) | JANGAN edit manual'
    },
    {
      // FASE 2: sesi aktif. Simpan HANYA hash session token.
      name: CONFIG.SHEETS.SESI,
      headers: ['Token Hash', 'Email', 'Device ID', 'Dibuat', 'Kadaluarsa', 'Status'],
      note: 'Status: Aktif / Kadaluarsa / Dicabut | Token Hash = hash session token | JANGAN edit manual'
    }
  ];
}

// ══════════════════════════════════════════════════════
// MIGRASI KEAMANAN (FASE 2) — idempoten, jalankan dari editor:
//   Run → migrasiKeamanan
// - Buat sheet "Perangkat" & "Sesi" bila belum ada.
// - Tambah kolom "PIN Hash" & "PIN Salt" di Master User bila belum ada.
// - Pastikan pepper keamanan tersimpan di Script Properties (dibuat sekali).
// Tidak menghapus/mengubah data yang sudah ada.
// ══════════════════════════════════════════════════════
function migrasiKeamanan() {
  if (!_setupAccessOk_()) return { success: false, message: 'Akses ditolak.' };
  var ss = SpreadsheetApp.openById(getSpreadsheetId());
  var log = [];

  // 1. Buat sheet Perangkat & Sesi (pakai definisi dari getSheetSchema_)
  var SCHEMA = getSheetSchema_();
  [CONFIG.SHEETS.PERANGKAT, CONFIG.SHEETS.SESI].forEach(function(nama) {
    var def = null;
    for (var i = 0; i < SCHEMA.length; i++) { if (SCHEMA[i].name === nama) { def = SCHEMA[i]; break; } }
    if (!def) return;
    var sheet = ss.getSheetByName(nama);
    if (!sheet) {
      sheet = ss.insertSheet(nama);
      sheet.appendRow(def.headers);
      sheet.getRange(1, 1, 1, def.headers.length)
        .setFontWeight('bold').setBackground('#1E40AF').setFontColor('#FFFFFF');
      sheet.setFrozenRows(1);
      if (def.note) sheet.getRange(1, 1).setNote(def.note);
      log.push('✅ DIBUAT: ' + nama);
    } else {
      log.push('⏭️ SUDAH ADA: ' + nama);
    }
  });

  // 2. Tambah kolom PIN di Master User (idempoten via ensureColumns_)
  var sheetUser = ss.getSheetByName(CONFIG.SHEETS.USER);
  if (sheetUser) {
    ensureColumns_(sheetUser, ['PIN Hash', 'PIN Salt']);
    log.push('🔧 KOLOM PIN dipastikan di ' + CONFIG.SHEETS.USER);
  } else {
    log.push('⚠️ Master User tidak ditemukan');
  }

  // 3. Pastikan pepper ada di Script Properties
  var pepper = getPepper_();
  log.push(pepper ? '🔒 Pepper keamanan siap (Script Properties)' : '⚠️ Gagal menyiapkan pepper');

  Logger.log(log.join('\n'));
  try {
    var ui = SpreadsheetApp.getUi();
    ui.alert('Migrasi Keamanan Selesai ✅', log.join('\n'), ui.ButtonSet.OK);
  } catch(e) {}
  return { success: true, log: log };
}

// ══════════════════════════════════════════════════════
// MIGRASI PENGENDALIAN (FASE 1) — idempoten, jalankan dari editor:
//   Run → migrasiPengendalian
// Menambah kolom soft-delete (Status/Dibatalkan By/At/Alasan Batal) pada
// sheet transaksi. Tidak menghapus/mengubah data lama. Baris lama tanpa
// Status dianggap 'Aktif' oleh pembaca (barisDibatalkan_).
// ══════════════════════════════════════════════════════
// Nomor urut periode (1-based) di Master Period — untuk kode No Bukti.
function _periodeSeqCode_(ss, periodeId) {
  var sheet = ss.getSheetByName(CONFIG.SHEETS.PERIOD);
  if (!sheet) return 1;
  var rows = sheet.getDataRange().getValues();
  var seq = 0;
  for (var i = 1; i < rows.length; i++) { if (rows[i][0]) { seq++; if (String(rows[i][0]) === String(periodeId)) return seq; } }
  return seq || 1;
}

// Isi No Bukti untuk data lama secara berurutan per periode (urut tanggal + ID).
function _backfillNoBukti_(ss) {
  var total = 0;
  [[CONFIG.SHEETS.INPUT_PENERIMAAN, 'BKM', 4], [CONFIG.SHEETS.INPUT_PENGELUARAN, 'BKK', 3]].forEach(function(cfg) {
    var sheet = ss.getSheetByName(cfg[0]);
    if (!sheet || sheet.getLastRow() < 2) return;
    var rows = sheet.getDataRange().getValues();
    var h = headerMap_(rows[0]);
    var cNB = h['nobukti']; if (cNB === undefined) return;
    var cPid = h['periodeid'] !== undefined ? h['periodeid'] : 1;
    var cTgl = h['tanggal'] !== undefined ? h['tanggal'] : cfg[2];
    var byPeriod = {}, maxSeq = {};
    for (var i = 1; i < rows.length; i++) {
      if (!rows[i][0]) continue;
      var pid = String(rows[i][cPid] || '');
      var nb = String(rows[i][cNB] || '');
      var m = nb.match(/-(\d+)$/);
      if (m) { var n = parseInt(m[1], 10); if (!(pid in maxSeq) || n > maxSeq[pid]) maxSeq[pid] = n; }
      else { if (!byPeriod[pid]) byPeriod[pid] = []; byPeriod[pid].push({ i: i, tgl: String(rows[i][cTgl] || ''), id: String(rows[i][0]) }); }
    }
    Object.keys(byPeriod).forEach(function(pid) {
      var arr = byPeriod[pid];
      arr.sort(function(a, b) { if (a.tgl !== b.tgl) return a.tgl < b.tgl ? -1 : 1; return a.id < b.id ? -1 : 1; });
      var seq = maxSeq[pid] || 0;
      var pseq = _periodeSeqCode_(ss, pid);
      arr.forEach(function(o) { seq++; sheet.getRange(o.i + 1, cNB + 1).setValue(cfg[1] + '-' + pseq + '-' + ('0000' + seq).slice(-4)); total++; });
    });
  });
  return total;
}

function _adaKode_(sheet, kode) {
  if (!sheet || sheet.getLastRow() < 1) return false;
  var col = sheet.getRange(1, 1, sheet.getLastRow(), 1).getValues();
  for (var i = 0; i < col.length; i++) { if (String(col[i][0]) === String(kode)) return true; }
  return false;
}

function migrasiPengendalian() {
  if (!_setupAccessOk_()) return { success: false, message: 'Akses ditolak.' };
  var ss = SpreadsheetApp.openById(getSpreadsheetId());
  var log = [];
  var kolomBatal = ['Status', 'Dibatalkan By', 'Dibatalkan At', 'Alasan Batal'];
  [CONFIG.SHEETS.INPUT_PENERIMAAN, CONFIG.SHEETS.INPUT_PENGELUARAN, CONFIG.SHEETS.INPUT_SETORAN].forEach(function(nm) {
    var sh = ss.getSheetByName(nm);
    if (sh) { ensureColumns_(sh, kolomBatal); log.push('🔧 Kolom soft-delete dipastikan → ' + nm); }
    else log.push('⚠️ Sheet tidak ditemukan → ' + nm);
  });

  // FASE 2: kolom tutup buku bermakna (selisih + arsip).
  var shSTB = ss.getSheetByName(CONFIG.SHEETS.SALDO_TUTUP_BUKU);
  if (shSTB) {
    ensureColumns_(shSTB, ['Saldo Tunai Sistem', 'Saldo Bank Sistem', 'Selisih Tunai', 'Selisih Bank', 'Selisih Total', 'Alasan Selisih', 'Arsip File ID', 'Arsip URL', 'Arsip Hash']);
    log.push('🔧 Kolom selisih/arsip dipastikan → ' + CONFIG.SHEETS.SALDO_TUTUP_BUKU);
  }

  // FASE 4: kolom No Bukti + sheet Lampiran.
  [CONFIG.SHEETS.INPUT_PENERIMAAN, CONFIG.SHEETS.INPUT_PENGELUARAN].forEach(function(nm) {
    var sh = ss.getSheetByName(nm);
    if (sh) { ensureColumns_(sh, ['No Bukti']); log.push('🔧 Kolom No Bukti dipastikan → ' + nm); }
  });
  // FASE 5: kolom maker-checker di pengeluaran.
  var shPK5 = ss.getSheetByName(CONFIG.SHEETS.INPUT_PENGELUARAN);
  if (shPK5) { ensureColumns_(shPK5, ['Status Approval', 'Disetujui By', 'Disetujui At']); log.push('🔧 Kolom approval dipastikan → ' + CONFIG.SHEETS.INPUT_PENGELUARAN); }
  if (!ss.getSheetByName(CONFIG.SHEETS.LAMPIRAN)) {
    var shL = ss.insertSheet(CONFIG.SHEETS.LAMPIRAN);
    shL.appendRow(['ID', 'Transaksi ID', 'Tipe', 'Nama File', 'Drive File ID', 'URL', 'Diunggah By', 'Diunggah At', 'Status']);
    log.push('✅ DIBUAT: ' + CONFIG.SHEETS.LAMPIRAN);
  }
  // L2c: kolom pembatalan & koreksi pada Kas Penerobos (soft delete + jejak).
  var shKP6 = ss.getSheetByName(CONFIG.SHEETS.KAS_PENEROBOS);
  if (shKP6) { ensureColumns_(shKP6, ['Dibatalkan By', 'Dibatalkan At', 'Alasan Batal', 'Koreksi Ref']); log.push('🔧 Kolom pembatalan/koreksi dipastikan → ' + CONFIG.SHEETS.KAS_PENEROBOS); }
  // Isi No Bukti untuk data lama (berurutan per periode, urut tanggal+ID).
  try { var bk = _backfillNoBukti_(ss); if (bk) log.push('🔢 No Bukti data lama diisi: ' + bk); } catch(e) { log.push('⚠️ backfill No Bukti gagal: ' + e.message); }

  // FASE 2: jenis khusus "Selisih Kas" di master (untuk baris penyesuaian tutup buku).
  var jm = ss.getSheetByName(CONFIG.SHEETS.PEMASUKAN);
  if (jm && !_adaKode_(jm, 'SELISIH_KAS')) { jm.appendRow(['SELISIH_KAS', 'Selisih Kas', 'Umum', 100, 0, 0, 'Umum', 'Aktif']); log.push('➕ Jenis Selisih Kas → Master Pemasukan'); }
  var jk = ss.getSheetByName(CONFIG.SHEETS.PENGELUARAN);
  if (jk && !_adaKode_(jk, 'SELISIH_KAS')) { jk.appendRow(['SELISIH_KAS', 'Selisih Kas', 'Umum', 'Aktif']); log.push('➕ Jenis Selisih Kas → Master Pengeluaran'); }

  // FASE 5.1: pisah trx.input → trx.input.masuk/keluar di sheet Hak Akses (jika ada).
  try { var hi = _migrasiHakAksesInput_(ss); if (hi) log.push('🔐 Hak Akses: ' + hi); } catch(e) { log.push('⚠️ migrasi Hak Akses input gagal: ' + e.message); }

  Logger.log(log.join('\n'));
  try { SpreadsheetApp.getUi().alert('Migrasi Pengendalian Selesai ✅', log.join('\n'), SpreadsheetApp.getUi().ButtonSet.OK); } catch(e) {}
  return { success: true, log: log };
}

// FASE 5.1: sisipkan baris 'trx.input.masuk'/'trx.input.keluar' ke sheet Hak Akses,
// meniru nilai per-role dari 'trx.input' lama agar perilaku instalasi lama tetap sama.
// Idempoten: hanya menambah bila kedua baris belum ada. Instalasi baru (belum ada sheet
// atau belum ada baris) memakai default ketat dari getDefaultPermMatrix_().
function _migrasiHakAksesInput_(ss) {
  var sheet = ss.getSheetByName(CONFIG.SHEETS.HAK_AKSES);
  if (!sheet || sheet.getLastRow() < 2) return '';
  var rows = sheet.getDataRange().getValues();
  var idx = {};
  for (var i = 1; i < rows.length; i++) idx[String(rows[i][0])] = i;
  if (idx['trx.input.masuk'] != null && idx['trx.input.keluar'] != null) return ''; // sudah ada
  if (idx['trx.input'] == null) return ''; // tak ada legacy → biarkan default ketat
  var legacy = rows[idx['trx.input']];
  var caps = getCapabilities_();
  function labelOf_(code) { for (var j = 0; j < caps.length; j++) if (caps[j].code === code) return caps[j].label; return code; }
  var ditambah = [];
  if (idx['trx.input.masuk'] == null) { sheet.appendRow(['trx.input.masuk', labelOf_('trx.input.masuk')].concat(legacy.slice(2))); ditambah.push('trx.input.masuk'); }
  if (idx['trx.input.keluar'] == null) { sheet.appendRow(['trx.input.keluar', labelOf_('trx.input.keluar')].concat(legacy.slice(2))); ditambah.push('trx.input.keluar'); }
  try { CacheService.getScriptCache().remove('perm_matrix'); } catch(e) {}
  return ditambah.join(' & ') + ' (mengikuti nilai trx.input lama)';
}

// ══════════════════════════════════════════════════════
// SETUP SHEETS — Jalankan SEKALI dari Apps Script Editor
// Menu: Run → setupSheets
// ══════════════════════════════════════════════════════
// Tidy-up (V4): guard fungsi setup/migrasi. Aman untuk bootstrap awal
// (belum ada ADMIN → izinkan), setelah itu hanya ADMIN terverifikasi
// (dijalankan dari editor sebagai pemilik) yang boleh. Pemanggilan via
// google.script.run anonim ditolak (Session kosong).
function _setupAccessOk_() {
  try {
    if (countActiveAdmins_() === 0) return true; // bootstrap pertama kali
    var email = '';
    try { email = Session.getActiveUser().getEmail(); } catch(e) {}
    if (!email) return false;
    var u = getUserByEmail_(email);
    return !!(u && u.role === CONFIG.ROLES.ADMIN && String(u.status || 'Aktif').toLowerCase() !== 'nonaktif');
  } catch(e) { return false; }
}

function setupSheets() {
  if (!_setupAccessOk_()) return { success: false, message: 'Akses ditolak.' };
  var ss = SpreadsheetApp.openById(getSpreadsheetId());
  var log = [];
  var SCHEMA = getSheetSchema_();

  for (var i = 0; i < SCHEMA.length; i++) {
    var s = SCHEMA[i];
    var sheet = ss.getSheetByName(s.name);

    if (!sheet) {
      // Buat sheet baru
      sheet = ss.insertSheet(s.name);
      sheet.appendRow(s.headers);

      // Format header: bold, background biru tua, teks putih
      var headerRange = sheet.getRange(1, 1, 1, s.headers.length);
      headerRange.setFontWeight('bold')
                 .setBackground('#1E40AF')
                 .setFontColor('#FFFFFF')
                 .setHorizontalAlignment('center');
      sheet.setFrozenRows(1);

      // Isi baris sample jika ada
      if (s.sample) {
        sheet.appendRow(s.sample);
        sheet.getRange(2, 1, 1, s.sample.length).setBackground('#EFF6FF');
      }

      // Tambah note di cell A1 jika ada
      if (s.note) {
        sheet.getRange(1, 1).setNote(s.note);
      }

      // Auto-resize kolom
      sheet.autoResizeColumns(1, s.headers.length);

      log.push('✅ DIBUAT: ' + s.name);
    } else {
      // Sheet sudah ada — cek apakah header baris 1 sudah benar
      var existingHeader = sheet.getRange(1, 1, 1, 1).getValue();
      if (!existingHeader) {
        // Header kosong — tambahkan
        sheet.insertRowBefore(1);
        sheet.getRange(1, 1, 1, s.headers.length).setValues([s.headers]);
        var hr = sheet.getRange(1, 1, 1, s.headers.length);
        hr.setFontWeight('bold').setBackground('#1E40AF').setFontColor('#FFFFFF');
        sheet.setFrozenRows(1);
        log.push('🔧 HEADER DITAMBAH: ' + s.name);
      } else {
        log.push('⏭️ SUDAH ADA: ' + s.name);
      }
    }
  }

  // Hapus atau tandai sheet yang tidak dipakai
  var unusedSheets = [
    'Dashboard Kas Tunai',
    'Dashboard Kas Bank',
    'Dashboard Periode',
    'Reconciliation',
    'Summary Bulanan'
  ];
  for (var j = 0; j < unusedSheets.length; j++) {
    var unused = ss.getSheetByName(unusedSheets[j]);
    if (unused) {
      // Rename dengan prefix ARSIP_ agar tidak terhapus datanya
      try {
        unused.setName('ARSIP_' + unusedSheets[j]);
        log.push('📦 DIARSIPKAN: ' + unusedSheets[j] + ' → ARSIP_' + unusedSheets[j]);
      } catch(e) {
        log.push('⚠️ Tidak bisa rename: ' + unusedSheets[j]);
      }
    }
  }

  // Susun urutan sheet sesuai alur kerja
  var sheetOrder = [
    CONFIG.SHEETS.KELOMPOK,
    CONFIG.SHEETS.USER,
    CONFIG.SHEETS.PERIOD,
    CONFIG.SHEETS.PEMASUKAN,
    CONFIG.SHEETS.PENGELUARAN,
    CONFIG.SHEETS.MUSYAWARAH,
    CONFIG.SHEETS.POS_SETORAN,
    CONFIG.SHEETS.ANGGOTA,
    CONFIG.SHEETS.INPUT_PENERIMAAN,
    CONFIG.SHEETS.BUKU_IR,
    CONFIG.SHEETS.INPUT_PENGELUARAN,
    CONFIG.SHEETS.INPUT_SETORAN,
    CONFIG.SHEETS.INPUT_SALDO_BANK,
    CONFIG.SHEETS.SETORAN_DESA,
    CONFIG.SHEETS.BANK_DAILY,
    CONFIG.SHEETS.BANK_PENDING,
    CONFIG.SHEETS.SALDO_TUTUP_BUKU,
    CONFIG.SHEETS.LOG
  ];
  var allSheets = ss.getSheets();
  var pos = 0;
  for (var k = 0; k < sheetOrder.length; k++) {
    var sh = ss.getSheetByName(sheetOrder[k]);
    if (sh) { ss.moveActiveSheet && ss.setActiveSheet(sh) && ss.moveActiveSheet(pos + 1); pos++; }
  }

  // Tampilkan hasil di Logger
  Logger.log('=== SETUP SHEETS SELESAI ===');
  Logger.log(log.join('\n'));
  Logger.log('Total: ' + log.length + ' sheet diproses');

  // Tampilkan alert di browser jika dipanggil dari editor
  try {
    var ui = SpreadsheetApp.getUi();
    ui.alert('Setup Selesai ✅', log.join('\n'), ui.ButtonSet.OK);
  } catch(e) {}

  return { success: true, log: log };
}

// ══════════════════════════════════════════════════════
// MIGRASI POS SETORAN → skema FK eksplisit
// Jalankan SEKALI dari editor: Run → migratePosSetoran
// Mengubah kolom lama (Tipe=manual/auto, Formula=teks bebas)
// menjadi (Sumber Tipe=bukuir/pemasukan/manual, Sumber Ref=key).
// ══════════════════════════════════════════════════════
function migratePosSetoran() {
  if (!_setupAccessOk_()) return { success: false, message: 'Akses ditolak.' };
  var ss = SpreadsheetApp.openById(getSpreadsheetId());
  var sheet = ss.getSheetByName(CONFIG.SHEETS.POS_SETORAN);
  if (!sheet || sheet.getLastRow() < 1) {
    Logger.log('Pos Setoran kosong / tidak ada.');
    return { success: true, log: ['Tidak ada data'] };
  }
  var rows = sheet.getDataRange().getValues();
  var log = [];

  // Tulis header baru (kolom 3 & 4 berubah nama, tambah Pengeluaran Ref)
  var newHeader = ['ID', 'Nama', 'Sumber Tipe', 'Sumber Ref', 'Status', 'Target', 'Pengeluaran Ref'];
  sheet.getRange(1, 1, 1, newHeader.length).setValues([newHeader]);

  for (var i = 1; i < rows.length; i++) {
    if (!rows[i][0]) continue;
    var nama = String(rows[i][1] || '');
    var oldTipe = String(rows[i][2] || '').toLowerCase().trim();
    var oldFormula = String(rows[i][3] || '');

    var sumberTipe = 'manual';
    var sumberRef = '';

    if (oldTipe === 'bukuir' || oldTipe === 'pemasukan' || oldTipe === 'manual') {
      // Sudah skema baru — biarkan
      sumberTipe = oldTipe;
      sumberRef = oldFormula;
    } else if (oldTipe === 'auto') {
      // Coba petakan ke kolom Buku IR berdasar nama/formula lama
      var key = mapPosNamaToIRColName(oldFormula, nama);
      if (key) {
        sumberTipe = 'bukuir';
        sumberRef = key;
      } else {
        sumberTipe = 'manual';
        sumberRef = '';
      }
    }

    sheet.getRange(i + 1, 3).setValue(sumberTipe);
    sheet.getRange(i + 1, 4).setValue(sumberRef);
    log.push(nama + ': ' + oldTipe + '/' + oldFormula + ' → ' + sumberTipe + '/' + (sumberRef || '-'));
  }

  try { CacheService.getScriptCache().remove('master_trx_data'); } catch(e) {}
  Logger.log('=== MIGRASI POS SETORAN ===\n' + log.join('\n'));
  try {
    var ui = SpreadsheetApp.getUi();
    ui.alert('Migrasi Selesai ✅', log.join('\n'), ui.ButtonSet.OK);
  } catch(e) {}
  return { success: true, log: log };
}

// ══════════════════════════════════════════════════════
// PERBAIKI HEADER TABEL → selaraskan dengan kode (normalisasi)
// Jalankan SEKALI dari editor: Run → repairSheetHeaders
// Menimpa baris header (row 1) tiap sheet dengan header kanonik
// dari getSheetSchema_(). TIDAK menyentuh baris data (row >= 2).
// Aman karena data sudah berurutan sesuai kode versi sekarang.
// ══════════════════════════════════════════════════════
function repairSheetHeaders() {
  if (!_setupAccessOk_()) return { success: false, message: 'Akses ditolak.' };
  var ss = SpreadsheetApp.openById(getSpreadsheetId());
  var SCHEMA = getSheetSchema_();
  var log = [];

  for (var i = 0; i < SCHEMA.length; i++) {
    var s = SCHEMA[i];
    var sheet = ss.getSheetByName(s.name);
    if (!sheet) { log.push('⏭️ TIDAK ADA: ' + s.name); continue; }

    var lastCol = sheet.getLastColumn();
    var canonN = s.headers.length;

    // Baca header lama untuk perbandingan/log
    var oldHeader = [];
    if (lastCol > 0) oldHeader = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    var oldStr = oldHeader.join(' | ');
    var newStr = s.headers.join(' | ');

    // Timpa header kanonik
    sheet.getRange(1, 1, 1, canonN).setValues([s.headers]);

    // Kosongkan sel header berlebih (mis. kolom usang yang tersisa)
    if (lastCol > canonN) {
      sheet.getRange(1, canonN + 1, 1, lastCol - canonN).clearContent();
    }

    // Format ulang header + freeze + note
    sheet.getRange(1, 1, 1, canonN)
      .setFontWeight('bold')
      .setBackground('#1E40AF')
      .setFontColor('#FFFFFF')
      .setHorizontalAlignment('center');
    sheet.setFrozenRows(1);
    if (s.note) sheet.getRange(1, 1).setNote(s.note);

    if (oldStr === newStr) {
      log.push('✓ OK: ' + s.name);
    } else {
      log.push('🔧 DIPERBAIKI: ' + s.name + '\n    lama: ' + oldStr + '\n    baru: ' + newStr);
    }
  }

  try { CacheService.getScriptCache().remove('master_trx_data'); } catch(e) {}
  Logger.log('=== PERBAIKI HEADER SELESAI ===\n' + log.join('\n'));
  try {
    var ui = SpreadsheetApp.getUi();
    ui.alert('Perbaiki Header Selesai ✅', log.join('\n'), ui.ButtonSet.OK);
  } catch(e) {}
  return { success: true, log: log };
}

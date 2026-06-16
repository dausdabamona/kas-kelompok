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
    TAGIHAN_PATUNGAN: 'Tagihan Terobosan',
    LOG: 'Activity Log',
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

// ══════════════════════════════════════════════════════
// SETUP SHEETS — Jalankan SEKALI dari Apps Script Editor
// Menu: Run → setupSheets
// ══════════════════════════════════════════════════════
function setupSheets() {
  var ss = SpreadsheetApp.openById(getSpreadsheetId());
  var log = [];

  var SCHEMA = [
    {
      name: CONFIG.SHEETS.KELOMPOK,
      headers: ['ID', 'Nama Kelompok', 'Alamat', 'Ketua', 'Bendahara', 'Tahun Berdiri', 'Keterangan'],
      sample: ['KEL-001', 'Kelompok Jamaah', '', '', '', '', '']
    },
    {
      name: CONFIG.SHEETS.USER,
      headers: ['Email', 'Nama', 'Role', 'Status', 'Created'],
      sample: ['dausdaba@gmail.com', 'Firdaus', 'ADMIN', 'Aktif', new Date()]
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
      headers: ['ID', 'Nama', 'No Telp', 'Alamat', 'Status', 'IR', '1/10 IR', 'Index', 'Infak Daerah'],
      note: 'Status: Aktif / Nonaktif | IR s.d. Infak Daerah dalam Rupiah'
    },
    {
      name: CONFIG.SHEETS.INPUT_PENERIMAAN,
      headers: ['ID', 'Periode ID', 'Jenis ID', 'Anggota ID', 'Tanggal', 'Nominal', 'Sumber Kas', 'Catatan', 'Created By', 'Created At'],
      note: 'Sumber Kas: Tunai / Bank | JANGAN edit manual'
    },
    {
      name: CONFIG.SHEETS.BUKU_IR,
      headers: ['ID', 'Transaksi ID', 'Periode ID', 'Anggota ID', 'Tanggal', 'IR', 'IR 1/10', 'Cicilan', 'Infak Daerah', 'Index', 'Total', 'Created By', 'Created At'],
      note: 'JANGAN edit manual — diisi otomatis dari form Buku IR'
    },
    {
      name: CONFIG.SHEETS.INPUT_PENGELUARAN,
      headers: ['ID', 'Periode ID', 'Jenis ID', 'Tanggal', 'Nominal', 'Sumber Kas', 'Catatan', 'Created By', 'Created At'],
      note: 'Sumber Kas: Tunai / Bank | JANGAN edit manual'
    },
    {
      name: CONFIG.SHEETS.INPUT_SETORAN,
      headers: ['ID', 'Periode ID', 'Tanggal', 'Nominal', 'Arah', 'Created By', 'Created At'],
      note: 'Arah: setor (Tunai→Bank) / tarik (Bank→Tunai) | JANGAN edit manual'
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
      headers: ['ID', 'Periode ID', 'Tanggal Tutup', 'Saldo Tunai Akhir', 'Saldo Bank Akhir', 'Total Kas', 'Status', 'Catatan', 'Created By', 'Created At'],
      note: 'JANGAN edit manual — diisi saat Tutup Buku'
    },
    {
      name: CONFIG.SHEETS.LOG,
      headers: ['Timestamp', 'User', 'Action', 'Detail'],
      note: 'Log otomatis — JANGAN edit manual'
    }
  ];

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

// ──────────────────────────────────────────────────────
// SPREADSHEET — buka sekali per eksekusi
// ──────────────────────────────────────────────────────
var _ss = null;
function getSS_() {
  if (!_ss) _ss = SpreadsheetApp.openById(getSpreadsheetId());
  return _ss;
}

// Konversi nilai tanggal ke string YYYY-MM-DD secara konsisten.
// Diperlukan karena Date object dari spreadsheet tidak bisa di-JSON.stringify.
function toDateStr_(val) {
  if (!val) return '';
  if (val instanceof Date) return Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return String(val);
}

// Buat header-to-index map dari baris header sheet.
// Key dinormalisasi: lowercase, trim, hapus semua spasi dan karakter '/'.
// "Periode ID" → 'periodeid', "1/10 IR" → '110ir', dst.
function headerMap_(headerRow) {
  var map = {};
  for (var i = 0; i < headerRow.length; i++) {
    var key = String(headerRow[i] || '').toLowerCase().trim().replace(/[\s\/]/g, '');
    if (key) map[key] = i;
  }
  return map;
}
// Helper lookup: ambil nilai dari row via headerMap_, fallback ke defaultIdx
function hGet_(row, map, key, defaultIdx) {
  var idx = map[key] !== undefined ? map[key] : defaultIdx;
  return idx >= 0 ? row[idx] : undefined;
}

// Pastikan sheet punya semua kolom header yang diperlukan; tambah di akhir jika belum ada.
// Membuat penambahan kolom (evolusi skema) aman untuk sheet yang sudah berisi data.
function ensureColumns_(sheet, headers) {
  var lastCol = sheet.getLastColumn();
  var existing = lastCol > 0 ? sheet.getRange(1, 1, 1, lastCol).getValues()[0] : [];
  var map = headerMap_(existing);
  var toAdd = [];
  for (var i = 0; i < headers.length; i++) {
    var key = String(headers[i]).toLowerCase().trim().replace(/[\s\/]/g, '');
    if (map[key] === undefined) toAdd.push(headers[i]);
  }
  if (toAdd.length > 0) {
    sheet.getRange(1, lastCol + 1, 1, toAdd.length).setValues([toAdd]);
  }
}

// ──────────────────────────────────────────────────────
// PERIODE
// ──────────────────────────────────────────────────────
function getPeriodeAktif() {
  try {
    var ss = getSS_();
    var sheet = ss.getSheetByName(CONFIG.SHEETS.PERIOD);
    if (!sheet) return null;
    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return null;

    // Temukan posisi kolom dari header row
    var header = data[0];
    var colStatus = -1, colNama = -1, colTglMulai = -1, colTglTutup = -1;
    var colSaldoTunai = -1, colSaldoBank = -1;
    for (var c = 0; c < header.length; c++) {
      var h = String(header[c]).toLowerCase().trim();
      if (h === 'status') colStatus = c;
      else if (h.indexOf('nama') !== -1) { if (colNama < 0) colNama = c; }
      else if (h.indexOf('mulai') !== -1) colTglMulai = c;
      else if (h.indexOf('tutup') !== -1) colTglTutup = c;
      else if (h.indexOf('tunai') !== -1) colSaldoTunai = c;
      else if (h.indexOf('bank') !== -1) colSaldoBank = c;
    }
    // Fallback jika header tidak standar: asumsi ID=0, Nama=1, TglMulai=2, Status=3
    if (colStatus < 0) colStatus = 3;
    if (colNama < 0) colNama = 1;
    if (colTglMulai < 0) colTglMulai = 2;

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][colStatus]).trim() === CONFIG.STATUS.OPEN) {
        var tgl = colTglMulai >= 0 ? data[i][colTglMulai] : '';
        if (tgl instanceof Date) {
          try { tgl = Utilities.formatDate(tgl, Session.getScriptTimeZone(), 'dd/MM/yyyy'); } catch(e) { tgl = String(tgl); }
        }
        return {
          id: String(data[i][0] || ''),
          nama: String(colNama >= 0 ? (data[i][colNama] || '') : ''),
          tanggalMulai: String(tgl || ''),
          status: CONFIG.STATUS.OPEN,
          saldoAwalTunai: colSaldoTunai >= 0 ? (Number(data[i][colSaldoTunai]) || 0) : 0,
          saldoAwalBank: colSaldoBank >= 0 ? (Number(data[i][colSaldoBank]) || 0) : 0
        };
      }
    }
    return null;
  } catch(e) {
    return null;
  }
}

function getAllPeriode() {
  var ss = getSS_();
  var sheet = ss.getSheetByName(CONFIG.SHEETS.PERIOD);
  if (!sheet) return { success: true, data: [] };
  var data = sheet.getDataRange().getValues();
  var result = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][0]) result.push({ id: data[i][0], nama: data[i][1], tanggalMulai: data[i][2], status: data[i][3] });
  }
  return { success: true, data: result };
}

// ──────────────────────────────────────────────────────
// DASHBOARD
// ──────────────────────────────────────────────────────
function getDashboardData() {
  try {
    var auth = checkAuth();
    if (!auth.success) return { success: false, message: auth.message };

    // Cek cache (non-user-specific: saldo + periode, bukan data user)
    var cache = CacheService.getScriptCache();
    var cacheKey = 'dashboard_saldo';
    var cached = cache.get(cacheKey);
    var saldoData = null;
    if (cached) {
      try { saldoData = JSON.parse(cached); } catch(e) {}
    }

    var ss = getSS_();
    var periode = getPeriodeAktif();
    var periodeId = periode ? periode.id : null;

    if (!saldoData) {
      var namaKelompok = 'Kas Kelompok';
      try {
        var sheetK = ss.getSheetByName(CONFIG.SHEETS.KELOMPOK);
        if (sheetK && sheetK.getLastRow() > 1) namaKelompok = sheetK.getRange(2, 2).getValue() || namaKelompok;
      } catch(e) {}
      var saldo = { tunai: 0, bank: 0 };
      try { saldo = calculateSaldo(periodeId, periode); } catch(e) {}
      saldoData = { namaKelompok: namaKelompok, tunai: saldo.tunai, bank: saldo.bank };
      try { cache.put(cacheKey, JSON.stringify(saldoData), 60); } catch(e) {}
    }

    // Ambil jumlah belum dirincikan dari cache buku IR (jika ada, gratis)
    var belumCount = 0;
    try {
      var bukuCache = cache.get('buku_ir_data');
      if (bukuCache) {
        var bd = JSON.parse(bukuCache);
        belumCount = (bd.data && bd.data.belumDirincikan) ? bd.data.belumDirincikan.length : 0;
      }
    } catch(e) {}

    return {
      success: true,
      user: auth.user,
      periode: periode,
      namaKelompok: saldoData.namaKelompok,
      kasTunai: saldoData.tunai,
      kasBank: saldoData.bank,
      totalKas: saldoData.tunai + saldoData.bank,
      belumDirincikanCount: belumCount
    };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function calculateSaldo(periodeId, periode) {
  var ss = getSS_();
  var tunai = (periode && periode.saldoAwalTunai) ? Number(periode.saldoAwalTunai) : 0;
  var bank  = (periode && periode.saldoAwalBank)  ? Number(periode.saldoAwalBank)  : 0;

  var sheetP = ss.getSheetByName(CONFIG.SHEETS.INPUT_PENERIMAAN);
  if (sheetP && sheetP.getLastRow() > 1) {
    var dp = sheetP.getDataRange().getValues();
    var dpH = headerMap_(dp[0]);
    for (var i = 1; i < dp.length; i++) {
      if (!periodeId || String(hGet_(dp[i], dpH, 'periodeid', 1)) === periodeId) {
        var nominal = Number(hGet_(dp[i], dpH, 'nominal', 5)) || 0;
        var sumber  = String(hGet_(dp[i], dpH, 'sumberkas', 6) || '');
        if (sumber === 'Tunai') tunai += nominal;
        else if (sumber === 'Bank') bank += nominal;
      }
    }
  }

  var sheetPK = ss.getSheetByName(CONFIG.SHEETS.INPUT_PENGELUARAN);
  if (sheetPK && sheetPK.getLastRow() > 1) {
    var dpk = sheetPK.getDataRange().getValues();
    var dpkH = headerMap_(dpk[0]);
    for (var i = 1; i < dpk.length; i++) {
      if (!periodeId || String(hGet_(dpk[i], dpkH, 'periodeid', 1)) === periodeId) {
        var nominal = Number(hGet_(dpk[i], dpkH, 'nominal', 4)) || 0;
        var sumber  = String(hGet_(dpk[i], dpkH, 'sumberkas', 5) || '');
        if (sumber === 'Tunai') tunai -= nominal;
        else if (sumber === 'Bank') bank -= nominal;
      }
    }
  }

  var sheetS = ss.getSheetByName(CONFIG.SHEETS.INPUT_SETORAN);
  if (sheetS && sheetS.getLastRow() > 1) {
    var ds = sheetS.getDataRange().getValues();
    var dsH = headerMap_(ds[0]);
    for (var i = 1; i < ds.length; i++) {
      if (!periodeId || String(hGet_(ds[i], dsH, 'periodeid', 1)) === periodeId) {
        var nominal = Number(hGet_(ds[i], dsH, 'nominal', 3)) || 0;
        var arah    = String(hGet_(ds[i], dsH, 'arah', 4) || '');
        if (arah === 'setor') { tunai -= nominal; bank += nominal; }
        else if (arah === 'tarik') { bank -= nominal; tunai += nominal; }
      }
    }
  }

  return { tunai: tunai, bank: bank };
}

// ──────────────────────────────────────────────────────
// PEMERIKSAAN SALDO (sebelum tutup buku)
// Bandingkan saldo sistem dengan saldo fisik/aktual yang diinput
// ──────────────────────────────────────────────────────
function getPemeriksaanSaldo() {
  try {
    var auth = checkAuth();
    if (!auth.success) return { success: false, message: auth.message };
    var periode = getPeriodeAktif();
    if (!periode) return { success: false, message: 'Tidak ada periode aktif' };

    var saldo = calculateSaldo(periode.id, periode);

    // Ambil pemeriksaan terakhir untuk periode ini (jika ada)
    var ss = getSS_();
    var terakhir = null;
    var sheet = ss.getSheetByName(CONFIG.SHEETS.SALDO_TUTUP_BUKU);
    if (sheet && sheet.getLastRow() > 1) {
      var rows = sheet.getDataRange().getValues();
      var h = headerMap_(rows[0]);
      for (var i = rows.length - 1; i >= 1; i--) {
        if (String(hGet_(rows[i], h, 'periodeid', 1)) === String(periode.id)) {
          terakhir = {
            tanggal: toDateStr_(hGet_(rows[i], h, 'tanggaltutup', 2)),
            saldoTunaiAktual: Number(hGet_(rows[i], h, 'saldotunaiakhir', 3)) || 0,
            saldoBankAktual: Number(hGet_(rows[i], h, 'saldobankakhir', 4)) || 0,
            status: String(hGet_(rows[i], h, 'status', 6) || ''),
            catatan: String(hGet_(rows[i], h, 'catatan', 7) || '')
          };
          break;
        }
      }
    }

    return {
      success: true,
      periode: periode,
      saldoSistem: { tunai: saldo.tunai, bank: saldo.bank, total: saldo.tunai + saldo.bank },
      terakhir: terakhir
    };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function submitPemeriksaanSaldo(data) {
  try {
    var auth = requirePerm('saldo.input');
    if (!auth.success) return { success: false, message: auth.message };
    var periode = getPeriodeAktif();
    if (!periode) return { success: false, message: 'Tidak ada periode aktif' };

    var ss = getSS_();
    var sheet = ss.getSheetByName(CONFIG.SHEETS.SALDO_TUTUP_BUKU);
    if (!sheet) {
      sheet = ss.insertSheet(CONFIG.SHEETS.SALDO_TUTUP_BUKU);
      sheet.appendRow(['ID', 'PeriodeID', 'Tanggal Tutup', 'Saldo Tunai Akhir', 'Saldo Bank Akhir', 'Total Kas', 'Status', 'Catatan', 'Created By', 'Created At']);
    }

    var saldoTunai = Number(data.saldoTunaiAktual) || 0;
    var saldoBank = Number(data.saldoBankAktual) || 0;
    var total = saldoTunai + saldoBank;
    // Status: 'Pemeriksaan' = cek saja, 'Tutup' = tutup buku final
    var status = data.tutup ? 'Tutup' : 'Pemeriksaan';

    var id = generateID('SLD');
    sheet.appendRow([id, periode.id, toDateStr_(new Date()), saldoTunai, saldoBank, total,
      status, data.catatan || '', auth.user.email, toDateStr_(new Date())]);

    logActivity(auth.user.email, status === 'Tutup' ? 'TUTUP_BUKU' : 'PEMERIKSAAN_SALDO',
      'Tunai: ' + saldoTunai + ', Bank: ' + saldoBank);

    return { success: true, id: id };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function tutupBuku(data) {
  try {
    var auth = requirePerm('periode.manage');
    if (!auth.success) return { success: false, message: auth.message };
    var periode = getPeriodeAktif();
    if (!periode) return { success: false, message: 'Tidak ada periode aktif yang bisa ditutup' };

    var ss = getSS_();
    var now = new Date();
    var saldoTunai = Number(data.saldoTunaiAktual) || 0;
    var saldoBank  = Number(data.saldoBankAktual)  || 0;
    var total = saldoTunai + saldoBank;

    // Catat ke Saldo Tutup Buku
    var sheetSaldo = ss.getSheetByName(CONFIG.SHEETS.SALDO_TUTUP_BUKU);
    if (!sheetSaldo) {
      sheetSaldo = ss.insertSheet(CONFIG.SHEETS.SALDO_TUTUP_BUKU);
      sheetSaldo.appendRow(['ID', 'Periode ID', 'Tanggal Tutup', 'Saldo Tunai Akhir', 'Saldo Bank Akhir', 'Total Kas', 'Status', 'Catatan', 'Created By', 'Created At']);
    }
    var sldId = generateID('SLD');
    sheetSaldo.appendRow([sldId, periode.id, toDateStr_(now), saldoTunai, saldoBank, total, 'Tutup', data.catatan || '', auth.user.email, toDateStr_(now)]);

    // Update Master Period: ubah Status → CLOSED, isi Tgl Tutup
    var sheetPeriod = ss.getSheetByName(CONFIG.SHEETS.PERIOD);
    if (sheetPeriod && sheetPeriod.getLastRow() > 1) {
      var pRows = sheetPeriod.getDataRange().getValues();
      var pH = headerMap_(pRows[0]);
      var colStatus   = (pH['status']   !== undefined ? pH['status']   : 4) + 1;
      var colTglTutup = (pH['tgltutup'] !== undefined ? pH['tgltutup'] : 3) + 1;
      for (var i = 1; i < pRows.length; i++) {
        if (String(pRows[i][0]) === String(periode.id)) {
          sheetPeriod.getRange(i + 1, colStatus).setValue(CONFIG.STATUS.CLOSED);
          sheetPeriod.getRange(i + 1, colTglTutup).setValue(toDateStr_(now));
          break;
        }
      }
    }

    try {
      var c = CacheService.getScriptCache();
      c.remove('dashboard_saldo');
      c.remove('master_trx_data');
    } catch(e) {}

    logActivity(auth.user.email, 'TUTUP_BUKU', 'Periode: ' + periode.nama + ' | Tunai: ' + saldoTunai + ' Bank: ' + saldoBank);
    return { success: true, id: sldId };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function bukaPeriode(data) {
  try {
    var auth = requirePerm('periode.manage');
    if (!auth.success) return { success: false, message: auth.message };
    if (!data.nama) return { success: false, message: 'Nama periode wajib diisi' };
    if (!data.tglMulai) return { success: false, message: 'Tanggal mulai wajib diisi' };

    var ss = getSS_();
    var sheetPeriod = ss.getSheetByName(CONFIG.SHEETS.PERIOD);
    if (!sheetPeriod) return { success: false, message: 'Sheet Master Period tidak ditemukan' };

    // Validasi: tidak boleh ada periode OPEN lain
    var pRows = sheetPeriod.getDataRange().getValues();
    var pH = headerMap_(pRows[0]);
    for (var i = 1; i < pRows.length; i++) {
      if (String(hGet_(pRows[i], pH, 'status', 4)).trim() === CONFIG.STATUS.OPEN) {
        return { success: false, message: 'Masih ada periode OPEN. Tutup periode aktif terlebih dahulu.' };
      }
    }

    var id = generateID('PER');
    sheetPeriod.appendRow([
      id,
      data.nama,
      toDateStr_(new Date(data.tglMulai)),
      '',
      CONFIG.STATUS.OPEN,
      Number(data.saldoAwalTunai) || 0,
      Number(data.saldoAwalBank)  || 0,
      data.catatan || ''
    ]);

    try { CacheService.getScriptCache().remove('dashboard_saldo'); } catch(e) {}
    logActivity(auth.user.email, 'BUKA_PERIODE', 'Periode: ' + data.nama);
    return { success: true, id: id };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

// ──────────────────────────────────────────────────────
// TRANSAKSI
// ──────────────────────────────────────────────────────
function submitTransaksi(data) {
  try {
    var auth = checkAuth();
    if (!auth.success) return { success: false, message: auth.message };
    var ss = getSS_();
    var periode = getPeriodeAktif();
    if (!periode) return { success: false, message: 'Tidak ada periode aktif' };

    var id = generateID('TRX');
    var now = new Date();
    var tgl = toDateStr_(data.tanggal ? new Date(data.tanggal) : now);

    // Validasi FK: jenisId harus ada di master
    if (data.tipe === 'masuk' || data.tipe === 'keluar') {
      var masterSheet = ss.getSheetByName(data.tipe === 'masuk' ? CONFIG.SHEETS.PEMASUKAN : CONFIG.SHEETS.PENGELUARAN);
      if (masterSheet) {
        var masterRows = masterSheet.getDataRange().getValues();
        var validJenis = false;
        for (var mi = 1; mi < masterRows.length; mi++) {
          if (String(masterRows[mi][0]) === String(data.jenisId)) { validJenis = true; break; }
        }
        if (!validJenis) return { success: false, message: 'Jenis tidak ditemukan di master data.' };
      }
    }

    if (data.tipe === 'masuk') {
      // Penerobos: tulis ke KAS_PENEROBOS dulu (belum masuk kas utama)
      if (auth.user.role === CONFIG.ROLES.PENEROBOS) {
        var kpSheet = ss.getSheetByName(CONFIG.SHEETS.KAS_PENEROBOS);
        if (!kpSheet) {
          kpSheet = ss.insertSheet(CONFIG.SHEETS.KAS_PENEROBOS);
          kpSheet.appendRow(['ID', 'Periode ID', 'Tanggal', 'Jenis ID', 'Anggota ID', 'Nominal', 'Sumber Kas', 'Catatan', 'Penerobos Email', 'Status', 'Serah Terima ID', 'Created At']);
        }
        kpSheet.appendRow([id, periode.id, tgl, data.jenisId, data.anggotaId || '', Number(data.nominal), data.sumberKas || 'Tunai', data.catatan || '', auth.user.email, 'Aktif', '', toDateStr_(now)]);
        logActivity(auth.user.email, 'KAS_PENEROBOS_VIA_TRX', 'Nominal: ' + data.nominal);
        try { CacheService.getScriptCache().remove('dashboard_saldo'); } catch(e) {}
        return { success: true, id: id, viaPenerobos: true };
      }
      var sheet = ss.getSheetByName(CONFIG.SHEETS.INPUT_PENERIMAAN);
      if (!sheet) return { success: false, message: 'Sheet penerimaan tidak ditemukan' };
      sheet.appendRow([id, periode.id, data.jenisId, data.anggotaId || '', tgl, Number(data.nominal), data.sumberKas, data.catatan || '', auth.user.email, toDateStr_(now)]);
      logActivity(auth.user.email, 'PEMASUKAN', 'Nominal: ' + data.nominal);
    } else if (data.tipe === 'keluar') {
      var sheet = ss.getSheetByName(CONFIG.SHEETS.INPUT_PENGELUARAN);
      if (!sheet) return { success: false, message: 'Sheet pengeluaran tidak ditemukan' };
      sheet.appendRow([id, periode.id, data.jenisId, tgl, Number(data.nominal), data.sumberKas, data.catatan || '', auth.user.email, toDateStr_(now)]);
      logActivity(auth.user.email, 'PENGELUARAN', 'Nominal: ' + data.nominal);
    } else if (data.tipe === 'mutasi') {
      var sheet = ss.getSheetByName(CONFIG.SHEETS.INPUT_SETORAN);
      if (!sheet) {
        sheet = ss.insertSheet(CONFIG.SHEETS.INPUT_SETORAN);
        sheet.appendRow(['ID', 'PeriodeID', 'Tanggal', 'Nominal', 'Arah', 'CreatedBy', 'CreatedAt']);
      }
      sheet.appendRow([id, periode.id, tgl, Number(data.nominal), data.arah, auth.user.email, toDateStr_(now)]);
      logActivity(auth.user.email, 'MUTASI', 'Arah: ' + data.arah + ' Nominal: ' + data.nominal);
    }

    // Invalidate saldo cache setiap ada transaksi baru
    try { CacheService.getScriptCache().remove('dashboard_saldo'); } catch(e) {}
    return { success: true, id: id };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function updateTransaksi(data) {
  try {
    var cap = (data.sumberKas === 'Tunai') ? 'trx.edit.tunai' : 'trx.edit.bank';
    var auth = requirePerm(cap);
    if (!auth.success) return { success: false, message: auth.message };
    var periode = getPeriodeAktif();
    if (!periode) return { success: false, message: 'Tidak ada periode aktif' };
    if (periode.status !== CONFIG.STATUS.OPEN) return { success: false, message: 'Periode sudah ditutup, tidak bisa mengedit transaksi' };

    var ss = getSS_();
    var tgl = toDateStr_(data.tanggal ? new Date(data.tanggal) : new Date());

    if (data.tipe === 'masuk') {
      var sheet = ss.getSheetByName(CONFIG.SHEETS.INPUT_PENERIMAAN);
      if (!sheet) return { success: false, message: 'Sheet tidak ditemukan' };
      var rows = sheet.getDataRange().getValues();
      var h = headerMap_(rows[0]);
      for (var i = 1; i < rows.length; i++) {
        if (String(hGet_(rows[i], h, 'id', 0)) === String(data.id)) {
          var colJenis   = (h['jenisid']    !== undefined ? h['jenisid']    : 2) + 1;
          var colAnggota = (h['anggotaid']  !== undefined ? h['anggotaid']  : 3) + 1;
          var colTgl     = (h['tanggal']    !== undefined ? h['tanggal']    : 4) + 1;
          var colNom     = (h['nominal']    !== undefined ? h['nominal']    : 5) + 1;
          var colSumber  = (h['sumberkas']  !== undefined ? h['sumberkas']  : 6) + 1;
          var colCat     = (h['catatan']    !== undefined ? h['catatan']    : 7) + 1;
          sheet.getRange(i + 1, colJenis).setValue(data.jenisId);
          sheet.getRange(i + 1, colAnggota).setValue(data.anggotaId || '');
          sheet.getRange(i + 1, colTgl).setValue(tgl);
          sheet.getRange(i + 1, colNom).setValue(Number(data.nominal) || 0);
          sheet.getRange(i + 1, colSumber).setValue(data.sumberKas);
          sheet.getRange(i + 1, colCat).setValue(data.catatan || '');
          try { CacheService.getScriptCache().remove('dashboard_saldo'); } catch(e) {}
          logActivity(auth.user.email, 'EDIT_PEMASUKAN', 'ID: ' + data.id);
          return { success: true };
        }
      }
      return { success: false, message: 'Transaksi tidak ditemukan' };
    } else if (data.tipe === 'keluar') {
      var sheet = ss.getSheetByName(CONFIG.SHEETS.INPUT_PENGELUARAN);
      if (!sheet) return { success: false, message: 'Sheet tidak ditemukan' };
      var rows = sheet.getDataRange().getValues();
      var h = headerMap_(rows[0]);
      for (var i = 1; i < rows.length; i++) {
        if (String(hGet_(rows[i], h, 'id', 0)) === String(data.id)) {
          var colJenis  = (h['jenisid']   !== undefined ? h['jenisid']   : 2) + 1;
          var colTgl    = (h['tanggal']   !== undefined ? h['tanggal']   : 3) + 1;
          var colNom    = (h['nominal']   !== undefined ? h['nominal']   : 4) + 1;
          var colSumber = (h['sumberkas'] !== undefined ? h['sumberkas'] : 5) + 1;
          var colCat    = (h['catatan']   !== undefined ? h['catatan']   : 6) + 1;
          sheet.getRange(i + 1, colJenis).setValue(data.jenisId);
          sheet.getRange(i + 1, colTgl).setValue(tgl);
          sheet.getRange(i + 1, colNom).setValue(Number(data.nominal) || 0);
          sheet.getRange(i + 1, colSumber).setValue(data.sumberKas);
          sheet.getRange(i + 1, colCat).setValue(data.catatan || '');
          try { CacheService.getScriptCache().remove('dashboard_saldo'); } catch(e) {}
          logActivity(auth.user.email, 'EDIT_PENGELUARAN', 'ID: ' + data.id);
          return { success: true };
        }
      }
      return { success: false, message: 'Transaksi tidak ditemukan' };
    }
    return { success: false, message: 'Tipe transaksi tidak valid' };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function deleteTransaksi(data) {
  try {
    var cap = (data.sumberKas === 'Tunai') ? 'trx.edit.tunai' : 'trx.edit.bank';
    var auth = requirePerm(cap);
    if (!auth.success) return { success: false, message: auth.message };
    var periode = getPeriodeAktif();
    if (!periode) return { success: false, message: 'Tidak ada periode aktif' };
    if (periode.status !== CONFIG.STATUS.OPEN) return { success: false, message: 'Periode sudah ditutup, tidak bisa menghapus transaksi' };

    var ss = getSS_();
    var sheetName = (data.tipe === 'masuk') ? CONFIG.SHEETS.INPUT_PENERIMAAN : CONFIG.SHEETS.INPUT_PENGELUARAN;
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) return { success: false, message: 'Sheet tidak ditemukan' };
    var rows = sheet.getDataRange().getValues();
    var h = headerMap_(rows[0]);
    for (var i = 1; i < rows.length; i++) {
      if (String(hGet_(rows[i], h, 'id', 0)) === String(data.id)) {
        sheet.deleteRow(i + 1);
        try { CacheService.getScriptCache().remove('dashboard_saldo'); } catch(e) {}
        logActivity(auth.user.email, 'HAPUS_' + (data.tipe === 'masuk' ? 'PEMASUKAN' : 'PENGELUARAN'), 'ID: ' + data.id);
        return { success: true };
      }
    }
    return { success: false, message: 'Transaksi tidak ditemukan' };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

// ──────────────────────────────────────────────────────
// MASTER DATA
// ──────────────────────────────────────────────────────
// MASTER DATA — satu fungsi untuk semua, dengan cache
// ──────────────────────────────────────────────────────
function getTransaksiMasterData() {
  try {
    var auth = checkAuth();
    if (!auth.success) return { success: false, message: auth.message };
    var cache = CacheService.getScriptCache();
    var cacheKey = 'master_trx_data';
    var cached = cache.get(cacheKey);
    if (cached) {
      try {
        var parsed = JSON.parse(cached);
        parsed.success = true;
        return parsed;
      } catch(e) {}
    }
    var ss = getSS_();
    // Pemasukan
    // Kolom: A=Kode, B=Nama, C=Kategori, D=%Kelompok, E=%Desa, F=%Daerah, G=InputTipe, H=Status
    var pemasukan = [];
    var sheetP = ss.getSheetByName(CONFIG.SHEETS.PEMASUKAN);
    if (sheetP) {
      var rp = sheetP.getDataRange().getValues();
      var header = rp[0] || [];
      // Cari kolom Status secara dinamis (toleran terhadap urutan kolom berbeda)
      var statusColP = -1;
      for (var c = 0; c < header.length; c++) {
        if (String(header[c]).toLowerCase().trim() === 'status') { statusColP = c; break; }
      }
      for (var i = 1; i < rp.length; i++) {
        if (!rp[i][0]) continue;
        var statusVal = statusColP >= 0 ? String(rp[i][statusColP] || '') : 'Aktif';
        pemasukan.push({
          id: String(rp[i][0]),
          nama: String(rp[i][1] || ''),
          kategori: String(rp[i][2] || ''),
          status: statusVal
        });
      }
    }
    // Pengeluaran — cari kolom Status secara dinamis
    var pengeluaran = [];
    var sheetPK = ss.getSheetByName(CONFIG.SHEETS.PENGELUARAN);
    if (sheetPK) {
      var rpk = sheetPK.getDataRange().getValues();
      var headerPK = rpk[0] || [];
      var statusColPK = -1;
      for (var c = 0; c < headerPK.length; c++) {
        if (String(headerPK[c]).toLowerCase().trim() === 'status') { statusColPK = c; break; }
      }
      for (var i = 1; i < rpk.length; i++) {
        if (!rpk[i][0]) continue;
        var statusVal = statusColPK >= 0 ? String(rpk[i][statusColPK] || '') : 'Aktif';
        pengeluaran.push({
          id: String(rpk[i][0]),
          nama: String(rpk[i][1] || ''),
          kategori: String(rpk[i][2] || ''),
          status: statusVal
        });
      }
    }
    // Anggota
    var anggota = [];
    var sheetA = ss.getSheetByName(CONFIG.SHEETS.ANGGOTA);
    if (sheetA) {
      var ra = sheetA.getDataRange().getValues();
      for (var i = 1; i < ra.length; i++) {
        if (ra[i][0]) anggota.push({
          id: String(ra[i][0]), nama: String(ra[i][1]), noTelp: String(ra[i][2] || ''),
          alamat: String(ra[i][3] || ''), status: String(ra[i][4] || ''),
          ir: Number(ra[i][5]) || 0, ir10: Number(ra[i][6]) || 0,
          index: Number(ra[i][7]) || 0, infakDaerah: Number(ra[i][8]) || 0
        });
      }
    }
    var result = { pemasukan: pemasukan, pengeluaran: pengeluaran, anggota: anggota };
    try { cache.put(cacheKey, JSON.stringify(result), 300); } catch(e) {}
    result.success = true;
    return result;
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function getMasterPemasukan() {
  try {
    var ss = getSS_();
    var sheet = ss.getSheetByName(CONFIG.SHEETS.PEMASUKAN);
    if (!sheet) return { success: true, data: [] };
    var rows = sheet.getDataRange().getValues();
    var header = rows[0] || [];
    var statusCol = 3; // default kolom D
    for (var c = 0; c < header.length; c++) {
      if (String(header[c]).toLowerCase().trim() === 'status') { statusCol = c; break; }
    }
    var result = [];
    for (var i = 1; i < rows.length; i++) {
      if (rows[i][0]) result.push({ id: String(rows[i][0]), nama: String(rows[i][1] || ''), kategori: String(rows[i][2] || ''), status: String(rows[i][statusCol] || '') });
    }
    return { success: true, data: result };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function getMasterPengeluaran() {
  try {
    var ss = getSS_();
    var sheet = ss.getSheetByName(CONFIG.SHEETS.PENGELUARAN);
    if (!sheet) return { success: true, data: [] };
    var rows = sheet.getDataRange().getValues();
    var result = [];
    for (var i = 1; i < rows.length; i++) {
      if (rows[i][0]) result.push({ id: rows[i][0], nama: rows[i][1], kategori: rows[i][2], status: rows[i][3] });
    }
    return { success: true, data: result };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function getAnggota() {
  try {
    var ss = getSS_();
    var sheet = ss.getSheetByName(CONFIG.SHEETS.ANGGOTA);
    if (!sheet) return { success: true, data: [] };
    var rows = sheet.getDataRange().getValues();
    var result = [];
    for (var i = 1; i < rows.length; i++) {
      if (rows[i][0]) result.push({
        id: rows[i][0], nama: rows[i][1], noTelp: rows[i][2],
        alamat: rows[i][3], status: rows[i][4],
        ir: Number(rows[i][5]) || 0,
        ir10: Number(rows[i][6]) || 0,
        index: Number(rows[i][7]) || 0,
        infakDaerah: Number(rows[i][8]) || 0
      });
    }
    return { success: true, data: result };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function addAnggota(data) {
  try {
    var auth = requirePerm('jamaah.add');
    if (!auth.success) return { success: false, message: auth.message };
    var ss = getSS_();
    var sheet = ss.getSheetByName(CONFIG.SHEETS.ANGGOTA);
    if (!sheet) {
      sheet = ss.insertSheet(CONFIG.SHEETS.ANGGOTA);
      sheet.appendRow(['ID', 'Nama', 'NoTelp', 'Alamat', 'Status', 'IR', '1/10 IR', 'Index', 'Infak Daerah']);
    }
    var id = generateID('ANG');
    sheet.appendRow([id, data.nama, data.noTelp || '', data.alamat || '', data.status || 'Aktif',
      Number(data.ir) || 0, Number(data.ir10) || 0, Number(data.index) || 0, Number(data.infakDaerah) || 0]);
    try { CacheService.getScriptCache().remove('master_trx_data'); } catch(e) {}
    return { success: true, id: id };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function updateAnggota(data) {
  try {
    var auth = requirePerm('jamaah.edit');
    if (!auth.success) return { success: false, message: auth.message };
    var ss = getSS_();
    var sheet = ss.getSheetByName(CONFIG.SHEETS.ANGGOTA);
    if (!sheet) return { success: false, message: 'Sheet tidak ditemukan' };
    var rows = sheet.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      if (rows[i][0] === data.id) {
        sheet.getRange(i + 1, 2, 1, 8).setValues([[data.nama, data.noTelp || '', data.alamat || '',
          data.status || 'Aktif', Number(data.ir) || 0, Number(data.ir10) || 0, Number(data.index) || 0, Number(data.infakDaerah) || 0]]);
    try { CacheService.getScriptCache().remove('master_trx_data'); } catch(e) {}
        return { success: true };
      }
    }
    return { success: false, message: 'Anggota tidak ditemukan' };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function deleteAnggota(id) {
  try {
    var auth = requirePerm('jamaah.delete');
    if (!auth.success) return { success: false, message: auth.message };
    var ss = getSS_();
    var sheet = ss.getSheetByName(CONFIG.SHEETS.ANGGOTA);
    if (!sheet) return { success: false, message: 'Sheet tidak ditemukan' };
    var rows = sheet.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      if (rows[i][0] === id) { sheet.deleteRow(i + 1); return { success: true }; }
    }
    return { success: false, message: 'Anggota tidak ditemukan' };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

// ──────────────────────────────────────────────────────
// PATUNGAN / MUSYAWARAH PEMBIAYAAN
// ──────────────────────────────────────────────────────

function getPatunganList() {
  try {
    var auth = checkAuth();
    if (!auth.success) return { success: false, message: auth.message };
    var ss = getSS_();
    var sheet = ss.getSheetByName(CONFIG.SHEETS.PATUNGAN);
    if (!sheet) return { success: true, data: [] };
    var rows = sheet.getDataRange().getValues();
    var h = headerMap_(rows[0]);
    var result = [];
    for (var i = 1; i < rows.length; i++) {
      if (!hGet_(rows[i], h, 'id', 0)) continue;
      // load grade config (stored as JSON in GradeConfig column)
      var gradeStr = String(hGet_(rows[i], h, 'gradeconfig', 5) || '{}');
      var gradeConfig = {};
      try { gradeConfig = JSON.parse(gradeStr); } catch(e) {}
      result.push({
        id: String(hGet_(rows[i], h, 'id', 0)),
        nama: String(hGet_(rows[i], h, 'nama', 1) || ''),
        tanggal: toDateStr_(hGet_(rows[i], h, 'tanggal', 2)),
        periodeId: String(hGet_(rows[i], h, 'periodeid', 3) || ''),
        status: String(hGet_(rows[i], h, 'status', 4) || 'Aktif'),
        gradeConfig: gradeConfig,
        catatan: String(hGet_(rows[i], h, 'catatan', 6) || ''),
        createdBy: String(hGet_(rows[i], h, 'createdby', 7) || '')
      });
    }
    return { success: true, data: result };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function addPatungan(data) {
  try {
    var auth = requirePerm('terobosan.create');
    if (!auth.success) return { success: false, message: auth.message };
    if (!data.nama) return { success: false, message: 'Nama patungan wajib diisi' };
    if (!data.gradeConfig || Object.keys(data.gradeConfig).length === 0)
      return { success: false, message: 'Konfigurasi grade wajib diisi' };

    var ss = getSS_();
    var sheet = ss.getSheetByName(CONFIG.SHEETS.PATUNGAN);
    if (!sheet) {
      sheet = ss.insertSheet(CONFIG.SHEETS.PATUNGAN);
      sheet.appendRow(['ID', 'Nama', 'Tanggal', 'PeriodeID', 'Status', 'GradeConfig', 'Catatan', 'CreatedBy', 'CreatedAt']);
    }
    ensureColumns_(sheet, ['ID', 'Nama', 'Tanggal', 'PeriodeID', 'Status', 'GradeConfig', 'Catatan', 'CreatedBy', 'CreatedAt']);

    var periode = getPeriodeAktif();
    var id = generateID('PAT');
    var now = new Date();
    sheet.appendRow([
      id, data.nama,
      data.tanggal || toDateStr_(now),
      periode ? periode.id : '',
      'Aktif',
      JSON.stringify(data.gradeConfig),
      data.catatan || '',
      auth.user.email,
      toDateStr_(now)
    ]);

    // Auto-generate tagihan per anggota berdasarkan grade
    var result = generateTagihanPatungan_(ss, id, data.gradeConfig, auth.user.email);
    return { success: true, id: id, tagihanCount: result.count, total: result.total };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function generateTagihanPatungan_(ss, patunganId, gradeConfig, createdBy) {
  var sheetA = ss.getSheetByName(CONFIG.SHEETS.ANGGOTA);
  var sheetT = ss.getSheetByName(CONFIG.SHEETS.TAGIHAN_PATUNGAN);
  if (!sheetT) {
    sheetT = ss.insertSheet(CONFIG.SHEETS.TAGIHAN_PATUNGAN);
    sheetT.appendRow(['ID', 'PatunganID', 'AnggotaID', 'AnggotaNama', 'Grade', 'Nominal', 'StatusBayar', 'TanggalBayar', 'Catatan', 'CreatedBy', 'CreatedAt']);
  }
  ensureColumns_(sheetT, ['ID', 'PatunganID', 'AnggotaID', 'AnggotaNama', 'Grade', 'Nominal', 'StatusBayar', 'TanggalBayar', 'Catatan', 'CreatedBy', 'CreatedAt']);

  var count = 0, total = 0;
  if (!sheetA) return { count: 0, total: 0 };
  var rowsA = sheetA.getDataRange().getValues();
  var hA = headerMap_(rowsA[0]);
  var now = new Date();

  for (var i = 1; i < rowsA.length; i++) {
    var aid = String(rowsA[i][0] || '');
    if (!aid) continue;
    var status = String(hGet_(rowsA[i], hA, 'status', 4) || 'Aktif');
    if (status !== 'Aktif') continue;
    var grade = String(hGet_(rowsA[i], hA, 'grade', 9) || '').toUpperCase().trim();
    if (!grade || !gradeConfig[grade]) continue;
    var nominal = Number(gradeConfig[grade]) || 0;
    if (nominal <= 0) continue;
    var aNama = String(hGet_(rowsA[i], hA, 'nama', 1) || '');
    var tid = generateID('TGH');
    sheetT.appendRow([tid, patunganId, aid, aNama, grade, nominal, 'Belum', '', '', createdBy, toDateStr_(now)]);
    count++;
    total += nominal;
  }
  return { count: count, total: total };
}

function getTagihanPatungan(patunganId) {
  try {
    var auth = checkAuth();
    if (!auth.success) return { success: false, message: auth.message };
    var ss = getSS_();

    // Load patungan header
    var sheetP = ss.getSheetByName(CONFIG.SHEETS.PATUNGAN);
    var patungan = null;
    if (sheetP) {
      var rp = sheetP.getDataRange().getValues();
      var hp = headerMap_(rp[0]);
      for (var i = 1; i < rp.length; i++) {
        if (String(hGet_(rp[i], hp, 'id', 0)) === patunganId) {
          var gcStr = String(hGet_(rp[i], hp, 'gradeconfig', 5) || '{}');
          var gc = {}; try { gc = JSON.parse(gcStr); } catch(e) {}
          patungan = {
            id: patunganId,
            nama: String(hGet_(rp[i], hp, 'nama', 1) || ''),
            tanggal: toDateStr_(hGet_(rp[i], hp, 'tanggal', 2)),
            status: String(hGet_(rp[i], hp, 'status', 4) || ''),
            gradeConfig: gc,
            catatan: String(hGet_(rp[i], hp, 'catatan', 6) || '')
          };
          break;
        }
      }
    }
    if (!patungan) return { success: false, message: 'Patungan tidak ditemukan' };

    var sheetT = ss.getSheetByName(CONFIG.SHEETS.TAGIHAN_PATUNGAN);
    var tagihan = [];
    var totalTarget = 0, totalLunas = 0, totalBelum = 0;
    if (sheetT) {
      var rt = sheetT.getDataRange().getValues();
      var ht = headerMap_(rt[0]);
      for (var i = 1; i < rt.length; i++) {
        if (String(hGet_(rt[i], ht, 'patunganid', 1)) !== patunganId) continue;
        var nominal = Number(hGet_(rt[i], ht, 'nominal', 5)) || 0;
        var statusBayar = String(hGet_(rt[i], ht, 'statusbayar', 6) || 'Belum');
        totalTarget += nominal;
        if (statusBayar === 'Lunas') totalLunas += nominal;
        else totalBelum += nominal;
        tagihan.push({
          id: String(hGet_(rt[i], ht, 'id', 0)),
          anggotaId: String(hGet_(rt[i], ht, 'anggotaid', 2) || ''),
          anggotaNama: String(hGet_(rt[i], ht, 'anggotanama', 3) || ''),
          grade: String(hGet_(rt[i], ht, 'grade', 4) || ''),
          nominal: nominal,
          statusBayar: statusBayar,
          tanggalBayar: toDateStr_(hGet_(rt[i], ht, 'tanggalbayar', 7)),
          catatan: String(hGet_(rt[i], ht, 'catatan', 8) || '')
        });
      }
    }
    tagihan.sort(function(a, b) { return a.grade < b.grade ? -1 : a.grade > b.grade ? 1 : a.anggotaNama.localeCompare(b.anggotaNama); });
    return { success: true, patungan: patungan, tagihan: tagihan, totalTarget: totalTarget, totalLunas: totalLunas, totalBelum: totalBelum };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function bayarTagihanPatungan(tagihanId, catatan) {
  try {
    var auth = requirePerm('terobosan.bayar');
    if (!auth.success) return { success: false, message: auth.message };
    var ss = getSS_();
    var sheet = ss.getSheetByName(CONFIG.SHEETS.TAGIHAN_PATUNGAN);
    if (!sheet) return { success: false, message: 'Sheet penerobosan tidak ditemukan' };
    var rows = sheet.getDataRange().getValues();
    var h = headerMap_(rows[0]);
    var today = toDateStr_(new Date());
    for (var i = 1; i < rows.length; i++) {
      if (String(hGet_(rows[i], h, 'id', 0)) === tagihanId) {
        var colStatus = (h['statusbayar'] !== undefined ? h['statusbayar'] : 6) + 1;
        var colTgl = (h['tanggalbayar'] !== undefined ? h['tanggalbayar'] : 7) + 1;
        var colCat = (h['catatan'] !== undefined ? h['catatan'] : 8) + 1;
        sheet.getRange(i + 1, colStatus).setValue('Lunas');
        sheet.getRange(i + 1, colTgl).setValue(today);
        sheet.getRange(i + 1, colCat).setValue(catatan || '');
        return { success: true };
      }
    }
    return { success: false, message: 'Penerobosan tidak ditemukan' };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function batalBayarTagihanPatungan(tagihanId) {
  try {
    var auth = requirePerm('terobosan.batal');
    if (!auth.success) return { success: false, message: auth.message };
    var ss = getSS_();
    var sheet = ss.getSheetByName(CONFIG.SHEETS.TAGIHAN_PATUNGAN);
    if (!sheet) return { success: false, message: 'Sheet penerobosan tidak ditemukan' };
    var rows = sheet.getDataRange().getValues();
    var h = headerMap_(rows[0]);
    for (var i = 1; i < rows.length; i++) {
      if (String(hGet_(rows[i], h, 'id', 0)) === tagihanId) {
        var colStatus = (h['statusbayar'] !== undefined ? h['statusbayar'] : 6) + 1;
        var colTgl = (h['tanggalbayar'] !== undefined ? h['tanggalbayar'] : 7) + 1;
        sheet.getRange(i + 1, colStatus).setValue('Belum');
        sheet.getRange(i + 1, colTgl).setValue('');
        return { success: true };
      }
    }
    return { success: false, message: 'Penerobosan tidak ditemukan' };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function updateAnggotaGrade(anggotaId, grade) {
  try {
    var auth = requirePerm('grade.edit');
    if (!auth.success) return { success: false, message: auth.message };
    var ss = getSS_();
    var sheet = ss.getSheetByName(CONFIG.SHEETS.ANGGOTA);
    if (!sheet) return { success: false, message: 'Sheet Anggota tidak ditemukan' };
    ensureColumns_(sheet, ['Grade']);
    var rows = sheet.getDataRange().getValues();
    var h = headerMap_(rows[0]);
    var colGrade = (h['grade'] !== undefined ? h['grade'] : -1) + 1;
    if (colGrade <= 0) return { success: false, message: 'Kolom Grade tidak ditemukan' };
    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === anggotaId) {
        sheet.getRange(i + 1, colGrade).setValue(grade);
        try { CacheService.getScriptCache().remove('master_trx_data'); } catch(e) {}
        return { success: true };
      }
    }
    return { success: false, message: 'Anggota tidak ditemukan' };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function getAnggotaWithGrade() {
  try {
    var auth = checkAuth();
    if (!auth.success) return { success: false, message: auth.message };
    var ss = getSS_();
    var sheet = ss.getSheetByName(CONFIG.SHEETS.ANGGOTA);
    if (!sheet) return { success: true, data: [] };
    ensureColumns_(sheet, ['Grade']);
    var rows = sheet.getDataRange().getValues();
    var h = headerMap_(rows[0]);
    var result = [];
    for (var i = 1; i < rows.length; i++) {
      if (!rows[i][0]) continue;
      result.push({
        id: String(rows[i][0]),
        nama: String(hGet_(rows[i], h, 'nama', 1) || ''),
        status: String(hGet_(rows[i], h, 'status', 4) || 'Aktif'),
        grade: String(hGet_(rows[i], h, 'grade', 9) || '')
      });
    }
    return { success: true, data: result };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

// ──────────────────────────────────────────────────────
// BUKU IR
// ──────────────────────────────────────────────────────
function getBukuIRData() {
  try {
    var auth = checkAuth();
    if (!auth.success) return { success: false, message: auth.message };

    var cache = CacheService.getScriptCache();
    var cacheKey = 'buku_ir_data';
    var cached = cache.get(cacheKey);
    if (cached) {
      try { return JSON.parse(cached); } catch(e) {}
    }

    var ss = getSS_();
    var periode = getPeriodeAktif();
    var periodeId = periode ? periode.id : null;

    var sheetP = ss.getSheetByName(CONFIG.SHEETS.INPUT_PENERIMAAN);
    var sheetIR = ss.getSheetByName(CONFIG.SHEETS.BUKU_IR);
    var sheetMaster = ss.getSheetByName(CONFIG.SHEETS.PEMASUKAN);
    var sheetAnggota = ss.getSheetByName(CONFIG.SHEETS.ANGGOTA);

    if (!sheetP) return { success: true, data: { belumDirincikan: [], sudahDirincikan: [] } };

    var bukuIRIds = [];
    if (sheetMaster) {
      var masterRows = sheetMaster.getDataRange().getValues();
      for (var i = 1; i < masterRows.length; i++) {
        if (masterRows[i][2] === 'Buku IR') bukuIRIds.push(masterRows[i][0]);
      }
    }

    var anggotaMap = {};
    if (sheetAnggota) {
      var angRows = sheetAnggota.getDataRange().getValues();
      for (var i = 1; i < angRows.length; i++) {
        if (angRows[i][0]) anggotaMap[angRows[i][0]] = {
          nama: angRows[i][1], ir: Number(angRows[i][5]) || 0,
          ir10: Number(angRows[i][6]) || 0, index: Number(angRows[i][7]) || 0,
          infakDaerah: Number(angRows[i][8]) || 0
        };
      }
    }

    // Kumpulkan rincian yang sudah ada per transaksiId (untuk edit)
    var rincianMap = {};
    if (sheetIR && sheetIR.getLastRow() > 1) {
      var irRows = sheetIR.getDataRange().getValues();
      var irH = headerMap_(irRows[0]);
      for (var i = 1; i < irRows.length; i++) {
        var trxId = String(hGet_(irRows[i], irH, 'transaksiid', 1) || '');
        if (!trxId) continue;
        rincianMap[trxId] = {
          rincianId: String(hGet_(irRows[i], irH, 'id', 0) || ''),
          ir: Number(hGet_(irRows[i], irH, 'ir', 5)) || 0,
          ir10: Number(hGet_(irRows[i], irH, 'ir10', 6)) || 0,
          cicilan: Number(hGet_(irRows[i], irH, 'cicilan', 7)) || 0,
          infakDaerah: Number(hGet_(irRows[i], irH, 'infakdaerah', 8)) || 0,
          index: Number(hGet_(irRows[i], irH, 'index', 9)) || 0
        };
      }
    }

    var belum = [], sudah = [];
    if (sheetP.getLastRow() > 1) {
      var pRows = sheetP.getDataRange().getValues();
      var pHdr = headerMap_(pRows[0]);
      var c_id = pHdr['id'] !== undefined ? pHdr['id'] : 0;
      var c_pid = pHdr['periodeid'] !== undefined ? pHdr['periodeid'] : 1;
      var c_jid = pHdr['jenisid'] !== undefined ? pHdr['jenisid'] : 2;
      var c_aid = pHdr['anggotaid'] !== undefined ? pHdr['anggotaid'] : 3;
      var c_tgl = pHdr['tanggal'] !== undefined ? pHdr['tanggal'] : 4;
      var c_nom = pHdr['nominal'] !== undefined ? pHdr['nominal'] : 5;
      var c_kas = pHdr['sumberkas'] !== undefined ? pHdr['sumberkas'] : 6;
      var c_cat = pHdr['catatan'] !== undefined ? pHdr['catatan'] : 7;
      for (var i = 1; i < pRows.length; i++) {
        var row = pRows[i];
        if (!row[c_id]) continue;
        if (periodeId && row[c_pid] !== periodeId) continue;
        if (bukuIRIds.indexOf(row[c_jid]) === -1) continue;
        var trxIdStr = String(row[c_id]);
        var item = {
          id: trxIdStr, periodeId: String(row[c_pid]), jenisId: String(row[c_jid]),
          anggotaId: String(row[c_aid] || ''),
          tanggal: toDateStr_(row[c_tgl]),
          nominal: Number(row[c_nom]) || 0,
          sumberKas: String(row[c_kas] || ''),
          catatan: String(row[c_cat] || ''),
          anggota: anggotaMap[String(row[c_aid])] || null
        };
        if (rincianMap[trxIdStr]) {
          item.rincian = rincianMap[trxIdStr]; // sertakan breakdown untuk edit
          sudah.push(item);
        } else {
          belum.push(item);
        }
      }
    }

    var result = { success: true, data: { belumDirincikan: belum, sudahDirincikan: sudah } };
    try { cache.put(cacheKey, JSON.stringify(result), 90); } catch(e) {}
    return result;
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function submitRincianIR(data) {
  try {
    var auth = checkAuth();
    if (!auth.success) return { success: false, message: auth.message };
    var ss = getSS_();
    var sheet = ss.getSheetByName(CONFIG.SHEETS.BUKU_IR);
    if (!sheet) {
      sheet = ss.insertSheet(CONFIG.SHEETS.BUKU_IR);
      // Kolom Total dihapus — nilai derived, dihitung saat read (IR+IR10+Cicilan+InfakDaerah+Index)
      sheet.appendRow(['ID', 'TransaksiID', 'PeriodeID', 'AnggotaID', 'Tanggal', 'IR', 'IR10', 'Cicilan', 'InfakDaerah', 'Index', 'CreatedBy', 'CreatedAt']);
    }
    // Validasi: transaksiId harus ada di Input Penerimaan
    var sheetP = ss.getSheetByName(CONFIG.SHEETS.INPUT_PENERIMAAN);
    if (sheetP) {
      var pRows = sheetP.getDataRange().getValues();
      var validTrx = false;
      for (var vi = 1; vi < pRows.length; vi++) {
        if (String(pRows[vi][0]) === String(data.transaksiId)) { validTrx = true; break; }
      }
      if (!validTrx) return { success: false, message: 'Transaksi tidak ditemukan.' };
    }
    var ir = Number(data.ir) || 0;
    var ir10 = Number(data.ir10) || 0;
    var cicilan = Number(data.cicilan) || 0;
    var infakDaerah = Number(data.infakDaerah) || 0;
    var index = Number(data.index) || 0;

    // Cek apakah transaksi ini sudah pernah dirincikan → update, bukan tambah baru.
    // Edit hanya boleh selama periode masih OPEN (getPeriodeAktif memfilter ke periode aktif).
    var irRows = sheet.getDataRange().getValues();
    var irH = headerMap_(irRows[0]);
    var col_trx = irH['transaksiid'] !== undefined ? irH['transaksiid'] : 1;
    var col_ir = (irH['ir'] !== undefined ? irH['ir'] : 5) + 1;
    var existingRow = -1;
    for (var ri = 1; ri < irRows.length; ri++) {
      if (String(irRows[ri][col_trx]) === String(data.transaksiId)) { existingRow = ri + 1; break; }
    }

    var id;
    if (existingRow > 0) {
      // Update kolom IR..Index pada baris yang ada (5 kolom berurutan)
      sheet.getRange(existingRow, col_ir, 1, 5).setValues([[ir, ir10, cicilan, infakDaerah, index]]);
      var col_upd = irH['createdat'] !== undefined ? irH['createdat'] : -1;
      if (col_upd >= 0) sheet.getRange(existingRow, col_upd + 1).setValue(toDateStr_(new Date()));
      id = String(irRows[existingRow - 1][irH['id'] !== undefined ? irH['id'] : 0]);
      logActivity(auth.user.email, 'RINCIAN_EDIT', 'Transaksi: ' + data.transaksiId);
    } else {
      id = generateID('IR');
      sheet.appendRow([id, data.transaksiId, data.periodeId, data.anggotaId, toDateStr_(data.tanggal),
        ir, ir10, cicilan, infakDaerah, index, auth.user.email, toDateStr_(new Date())]);
    }
    try { var c = CacheService.getScriptCache(); c.remove('master_trx_data'); c.remove('buku_ir_data'); c.remove('dashboard_saldo'); } catch(e) {}
    return { success: true, id: id, updated: existingRow > 0 };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

// ──────────────────────────────────────────────────────
// REKAPITULASI
// ──────────────────────────────────────────────────────
function getRekapitulasiData() {
  try {
    var auth = checkAuth();
    if (!auth.success) return { success: false, message: auth.message };
    var ss = getSS_();
    var periode = getPeriodeAktif();
    var periodeId = periode ? periode.id : null;

    var sheetP = ss.getSheetByName(CONFIG.SHEETS.INPUT_PENERIMAAN);
    var sheetPK = ss.getSheetByName(CONFIG.SHEETS.INPUT_PENGELUARAN);
    var sheetMasterP = ss.getSheetByName(CONFIG.SHEETS.PEMASUKAN);
    var sheetMasterPK = ss.getSheetByName(CONFIG.SHEETS.PENGELUARAN);

    var masterPMap = {}, masterPKMap = {};
    if (sheetMasterP) {
      var mp = sheetMasterP.getDataRange().getValues();
      for (var i = 1; i < mp.length; i++) masterPMap[mp[i][0]] = mp[i][1];
    }
    if (sheetMasterPK) {
      var mpk = sheetMasterPK.getDataRange().getValues();
      for (var i = 1; i < mpk.length; i++) masterPKMap[mpk[i][0]] = mpk[i][1];
    }

    var pemasukan = [], pengeluaran = [];
    var totalPTunai = 0, totalPBank = 0, totalPKTunai = 0, totalPKBank = 0;

    // Build anggota name map for display
    var anggotaNameMap = {};
    var sheetAng = ss.getSheetByName(CONFIG.SHEETS.ANGGOTA);
    if (sheetAng) {
      var angRows = sheetAng.getDataRange().getValues();
      for (var i = 1; i < angRows.length; i++) {
        if (angRows[i][0]) anggotaNameMap[String(angRows[i][0])] = String(angRows[i][1] || '');
      }
    }

    if (sheetP && sheetP.getLastRow() > 1) {
      var dp = sheetP.getDataRange().getValues();
      var dpH = headerMap_(dp[0]);
      for (var i = 1; i < dp.length; i++) {
        if (!hGet_(dp[i], dpH, 'id', 0)) continue;
        if (periodeId && String(hGet_(dp[i], dpH, 'periodeid', 1)) !== periodeId) continue;
        var nominal = Number(hGet_(dp[i], dpH, 'nominal', 5)) || 0;
        var sumber  = String(hGet_(dp[i], dpH, 'sumberkas', 6) || '');
        var jenisId = String(hGet_(dp[i], dpH, 'jenisid', 2) || '');
        var anggotaId = String(hGet_(dp[i], dpH, 'anggotaid', 3) || '');
        var catatan = String(hGet_(dp[i], dpH, 'catatan', 7) || '');
        pemasukan.push({
          id: String(hGet_(dp[i], dpH, 'id', 0)),
          jenisId: jenisId,
          jenis: masterPMap[jenisId] || jenisId,
          tanggal: toDateStr_(hGet_(dp[i], dpH, 'tanggal', 4)),
          nominal: nominal, sumber: sumber,
          anggotaId: anggotaId,
          anggota: anggotaNameMap[anggotaId] || '',
          catatan: catatan
        });
        if (sumber === 'Tunai') totalPTunai += nominal;
        else if (sumber === 'Bank') totalPBank += nominal;
      }
    }

    if (sheetPK && sheetPK.getLastRow() > 1) {
      var dpk = sheetPK.getDataRange().getValues();
      var dpkH = headerMap_(dpk[0]);
      for (var i = 1; i < dpk.length; i++) {
        if (!hGet_(dpk[i], dpkH, 'id', 0)) continue;
        if (periodeId && String(hGet_(dpk[i], dpkH, 'periodeid', 1)) !== periodeId) continue;
        var nominal = Number(hGet_(dpk[i], dpkH, 'nominal', 4)) || 0;
        var sumber  = String(hGet_(dpk[i], dpkH, 'sumberkas', 5) || '');
        var jenisId = String(hGet_(dpk[i], dpkH, 'jenisid', 2) || '');
        var catatan = String(hGet_(dpk[i], dpkH, 'catatan', 6) || '');
        pengeluaran.push({
          id: String(hGet_(dpk[i], dpkH, 'id', 0)),
          jenisId: jenisId,
          jenis: masterPKMap[jenisId] || jenisId,
          tanggal: toDateStr_(hGet_(dpk[i], dpkH, 'tanggal', 3)),
          nominal: nominal, sumber: sumber,
          catatan: catatan
        });
        if (sumber === 'Tunai') totalPKTunai += nominal;
        else if (sumber === 'Bank') totalPKBank += nominal;
      }
    }

    var saldo = calculateSaldo(periodeId, periode);

    return {
      success: true,
      periode: periode,
      pemasukan: pemasukan,
      pengeluaran: pengeluaran,
      totalPemasukanTunai: totalPTunai,
      totalPemasukanBank: totalPBank,
      totalPengeluaranTunai: totalPKTunai,
      totalPengeluaranBank: totalPKBank,
      saldoTunai: saldo.tunai,
      saldoBank: saldo.bank
    };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

// ──────────────────────────────────────────────────────
// SETTING — POS PEMASUKAN
// ──────────────────────────────────────────────────────
function addPosPemasukan(data) {
  try {
    var auth = requirePerm('master.manage');
    if (!auth.success) return { success: false, message: auth.message };
    var ss = getSS_();
    var sheet = ss.getSheetByName(CONFIG.SHEETS.PEMASUKAN);
    if (!sheet) {
      sheet = ss.insertSheet(CONFIG.SHEETS.PEMASUKAN);
      sheet.appendRow(['ID', 'Nama', 'Kategori', 'Status']);
    }
    var id = generateID('PMS');
    sheet.appendRow([id, data.nama, data.kategori || 'Umum', data.status || 'Aktif']);
    try { CacheService.getScriptCache().remove('master_trx_data'); } catch(e) {}
    return { success: true, id: id };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function updatePosPemasukan(data) {
  try {
    var auth = requirePerm('master.manage');
    if (!auth.success) return { success: false, message: auth.message };
    var ss = getSS_();
    var sheet = ss.getSheetByName(CONFIG.SHEETS.PEMASUKAN);
    if (!sheet) return { success: false, message: 'Sheet tidak ditemukan' };
    var rows = sheet.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      if (rows[i][0] === data.id) {
        sheet.getRange(i + 1, 2, 1, 3).setValues([[data.nama, data.kategori, data.status]]);
    try { CacheService.getScriptCache().remove('master_trx_data'); } catch(e) {}
        return { success: true };
      }
    }
    return { success: false, message: 'Data tidak ditemukan' };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function deletePosPemasukan(id) {
  try {
    var auth = requirePerm('master.manage');
    if (!auth.success) return { success: false, message: auth.message };
    var ss = getSS_();
    var sheet = ss.getSheetByName(CONFIG.SHEETS.PEMASUKAN);
    if (!sheet) return { success: false, message: 'Sheet tidak ditemukan' };
    var rows = sheet.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      if (rows[i][0] === id) { sheet.deleteRow(i + 1); return { success: true }; }
    }
    return { success: false, message: 'Data tidak ditemukan' };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

// ──────────────────────────────────────────────────────
// SETTING — POS PENGELUARAN
// ──────────────────────────────────────────────────────
function addPosPengeluaran(data) {
  try {
    var auth = requirePerm('master.manage');
    if (!auth.success) return { success: false, message: auth.message };
    var ss = getSS_();
    var sheet = ss.getSheetByName(CONFIG.SHEETS.PENGELUARAN);
    if (!sheet) {
      sheet = ss.insertSheet(CONFIG.SHEETS.PENGELUARAN);
      sheet.appendRow(['ID', 'Nama', 'Kategori', 'Status']);
    }
    var id = generateID('PNK');
    sheet.appendRow([id, data.nama, data.kategori || 'Umum', data.status || 'Aktif']);
    try { CacheService.getScriptCache().remove('master_trx_data'); } catch(e) {}
    return { success: true, id: id };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function updatePosPengeluaran(data) {
  try {
    var auth = requirePerm('master.manage');
    if (!auth.success) return { success: false, message: auth.message };
    var ss = getSS_();
    var sheet = ss.getSheetByName(CONFIG.SHEETS.PENGELUARAN);
    if (!sheet) return { success: false, message: 'Sheet tidak ditemukan' };
    var rows = sheet.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      if (rows[i][0] === data.id) {
        sheet.getRange(i + 1, 2, 1, 3).setValues([[data.nama, data.kategori, data.status]]);
    try { CacheService.getScriptCache().remove('master_trx_data'); } catch(e) {}
        return { success: true };
      }
    }
    return { success: false, message: 'Data tidak ditemukan' };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function deletePosPengeluaran(id) {
  try {
    var auth = requirePerm('master.manage');
    if (!auth.success) return { success: false, message: auth.message };
    var ss = getSS_();
    var sheet = ss.getSheetByName(CONFIG.SHEETS.PENGELUARAN);
    if (!sheet) return { success: false, message: 'Sheet tidak ditemukan' };
    var rows = sheet.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      if (rows[i][0] === id) { sheet.deleteRow(i + 1); return { success: true }; }
    }
    return { success: false, message: 'Data tidak ditemukan' };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

// ──────────────────────────────────────────────────────
// SETTING — POS SETORAN
// ──────────────────────────────────────────────────────
function getPosSetoranAll() {
  try {
    var ss = getSS_();
    var sheet = ss.getSheetByName(CONFIG.SHEETS.POS_SETORAN);
    if (!sheet) return { success: true, data: [], pemasukanList: [] };
    var rows = sheet.getDataRange().getValues();
    var h = headerMap_(rows[0]);
    var result = [];
    for (var i = 1; i < rows.length; i++) {
      if (!rows[i][0]) continue;
      result.push({
        id: rows[i][0],
        nama: String(hGet_(rows[i], h, 'nama', 1) || ''),
        sumberTipe: String(hGet_(rows[i], h, 'sumbertipe', 2) || 'manual'),
        sumberRef: String(hGet_(rows[i], h, 'sumberref', 3) || ''),
        status: String(hGet_(rows[i], h, 'status', 4) || 'Aktif'),
        target: Number(hGet_(rows[i], h, 'target', 5)) || 0,
        pengeluaranRef: String(hGet_(rows[i], h, 'pengeluaranref', 6) || '')
      });
    }
    // Daftar jenis pemasukan (untuk dropdown Sumber Ref tipe 'pemasukan')
    var pemasukanList = [];
    var sheetP = ss.getSheetByName(CONFIG.SHEETS.PEMASUKAN);
    if (sheetP && sheetP.getLastRow() > 1) {
      var pRows = sheetP.getDataRange().getValues();
      for (var j = 1; j < pRows.length; j++) {
        if (!pRows[j][0]) continue;
        pemasukanList.push({ kode: String(pRows[j][0]), nama: String(pRows[j][1] || '') });
      }
    }
    // Daftar jenis pengeluaran (untuk dropdown Pengeluaran Ref)
    var pengeluaranList = [];
    var sheetPK = ss.getSheetByName(CONFIG.SHEETS.PENGELUARAN);
    if (sheetPK && sheetPK.getLastRow() > 1) {
      var pkRows = sheetPK.getDataRange().getValues();
      for (var j = 1; j < pkRows.length; j++) {
        if (!pkRows[j][0]) continue;
        pengeluaranList.push({ kode: String(pkRows[j][0]), nama: String(pkRows[j][1] || '') });
      }
    }
    return { success: true, data: result, pemasukanList: pemasukanList, pengeluaranList: pengeluaranList };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function addPosSetoran(data) {
  try {
    var auth = requirePerm('master.manage');
    if (!auth.success) return { success: false, message: auth.message };
    var ss = getSS_();
    var sheet = ss.getSheetByName(CONFIG.SHEETS.POS_SETORAN);
    if (!sheet) {
      sheet = ss.insertSheet(CONFIG.SHEETS.POS_SETORAN);
      sheet.appendRow(['ID', 'Nama', 'Sumber Tipe', 'Sumber Ref', 'Status', 'Target', 'Pengeluaran Ref']);
    }
    ensureColumns_(sheet, ['Pengeluaran Ref']);
    var id = generateID('PST');
    // Sumber Ref hanya relevan untuk tipe bukuir/pemasukan
    var sumberTipe = data.sumberTipe || 'manual';
    var sumberRef = (sumberTipe === 'manual') ? '' : (data.sumberRef || '');
    sheet.appendRow([id, data.nama, sumberTipe, sumberRef, data.status || 'Aktif', Number(data.target) || 0, data.pengeluaranRef || '']);
    try { CacheService.getScriptCache().remove('master_trx_data'); } catch(e) {}
    return { success: true, id: id };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function updatePosSetoran(data) {
  try {
    var auth = requirePerm('master.manage');
    if (!auth.success) return { success: false, message: auth.message };
    var ss = getSS_();
    var sheet = ss.getSheetByName(CONFIG.SHEETS.POS_SETORAN);
    if (!sheet) return { success: false, message: 'Sheet tidak ditemukan' };
    ensureColumns_(sheet, ['Pengeluaran Ref']);
    var rows = sheet.getDataRange().getValues();
    var h = headerMap_(rows[0]);
    var colPK = h['pengeluaranref'] !== undefined ? h['pengeluaranref'] : 6;
    var sumberTipe = data.sumberTipe || 'manual';
    var sumberRef = (sumberTipe === 'manual') ? '' : (data.sumberRef || '');
    for (var i = 1; i < rows.length; i++) {
      if (rows[i][0] === data.id) {
        sheet.getRange(i + 1, 2, 1, 5).setValues([[data.nama, sumberTipe, sumberRef, data.status, Number(data.target) || 0]]);
        sheet.getRange(i + 1, colPK + 1).setValue(data.pengeluaranRef || '');
    try { CacheService.getScriptCache().remove('master_trx_data'); } catch(e) {}
        return { success: true };
      }
    }
    return { success: false, message: 'Data tidak ditemukan' };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function deletePosSetoran(id) {
  try {
    var auth = requirePerm('master.manage');
    if (!auth.success) return { success: false, message: auth.message };
    var ss = getSS_();
    var sheet = ss.getSheetByName(CONFIG.SHEETS.POS_SETORAN);
    if (!sheet) return { success: false, message: 'Sheet tidak ditemukan' };
    var rows = sheet.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      if (rows[i][0] === id) { sheet.deleteRow(i + 1); return { success: true }; }
    }
    return { success: false, message: 'Data tidak ditemukan' };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

// ──────────────────────────────────────────────────────
// SETTING — MUSYAWAROH
// ──────────────────────────────────────────────────────
function getMusyawaroh() {
  try {
    var ss = getSS_();
    var sheet = ss.getSheetByName(CONFIG.SHEETS.MUSYAWARAH);
    if (!sheet) return { success: true, data: [] };
    var rows = sheet.getDataRange().getValues();
    var result = [];
    for (var i = 1; i < rows.length; i++) {
      if (rows[i][0]) result.push({ id: rows[i][0], nama: rows[i][1], nilai: Number(rows[i][2]) || 0, keterangan: rows[i][3] });
    }
    try { CacheService.getScriptCache().remove('master_trx_data'); } catch(e) {}
    return { success: true, data: result };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function addMusyawaroh(data) {
  try {
    var auth = requirePerm('master.manage');
    if (!auth.success) return { success: false, message: auth.message };
    var ss = getSS_();
    var sheet = ss.getSheetByName(CONFIG.SHEETS.MUSYAWARAH);
    if (!sheet) {
      sheet = ss.insertSheet(CONFIG.SHEETS.MUSYAWARAH);
      sheet.appendRow(['ID', 'Nama', 'Nilai', 'Keterangan']);
    }
    var id = generateID('MSY');
    sheet.appendRow([id, data.nama, Number(data.nilai) || 0, data.keterangan || '']);
    try { CacheService.getScriptCache().remove('master_trx_data'); } catch(e) {}
    return { success: true, id: id };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function updateMusyawaroh(data) {
  try {
    var auth = requirePerm('master.manage');
    if (!auth.success) return { success: false, message: auth.message };
    var ss = getSS_();
    var sheet = ss.getSheetByName(CONFIG.SHEETS.MUSYAWARAH);
    if (!sheet) return { success: false, message: 'Sheet tidak ditemukan' };
    var rows = sheet.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      if (rows[i][0] === data.id) {
        sheet.getRange(i + 1, 2, 1, 3).setValues([[data.nama, Number(data.nilai) || 0, data.keterangan || '']]);
    try { CacheService.getScriptCache().remove('master_trx_data'); } catch(e) {}
        return { success: true };
      }
    }
    return { success: false, message: 'Data tidak ditemukan' };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function deleteMusyawaroh(id) {
  try {
    var auth = requirePerm('master.manage');
    if (!auth.success) return { success: false, message: auth.message };
    var ss = getSS_();
    var sheet = ss.getSheetByName(CONFIG.SHEETS.MUSYAWARAH);
    if (!sheet) return { success: false, message: 'Sheet tidak ditemukan' };
    var rows = sheet.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      if (rows[i][0] === id) { sheet.deleteRow(i + 1); return { success: true }; }
    }
    return { success: false, message: 'Data tidak ditemukan' };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

// ──────────────────────────────────────────────────────
// SETORAN DESA
// ──────────────────────────────────────────────────────
function getRekapSetoran() {
  try {
    var auth = checkAuth();
    if (!auth.success) return { success: false, message: auth.message };
    var ss = getSS_();
    var sheetPos = ss.getSheetByName(CONFIG.SHEETS.POS_SETORAN);
    var sheetSetoran = ss.getSheetByName(CONFIG.SHEETS.SETORAN_DESA);
    if (!sheetPos) return { success: true, data: [] };

    var posRows = sheetPos.getDataRange().getValues();
    var periode = getPeriodeAktif();
    var periodeId = periode ? periode.id : null;
    var setoranMap = {};
    if (sheetSetoran && sheetSetoran.getLastRow() > 1) {
      var sRows = sheetSetoran.getDataRange().getValues();
      var sHdr = headerMap_(sRows[0]);
      var s_posid = sHdr['posid'] !== undefined ? sHdr['posid'] : 1;
      var s_perid = sHdr['periodeid'] !== undefined ? sHdr['periodeid'] : 2;
      var s_real = sHdr['realisasi'] !== undefined ? sHdr['realisasi'] : 3;
      var s_cat = sHdr['catatan'] !== undefined ? sHdr['catatan'] : 4;
      var s_kas = sHdr['sumberkas'] !== undefined ? sHdr['sumberkas'] : -1;
      for (var i = 1; i < sRows.length; i++) {
        if (!sRows[i][0]) continue;
        // Filter per periode aktif jika ada PeriodeID
        if (periodeId && sRows[i][s_perid] && String(sRows[i][s_perid]) !== periodeId) continue;
        setoranMap[String(sRows[i][s_posid])] = {
          id: sRows[i][0], realisasi: Number(sRows[i][s_real]) || 0, catatan: String(sRows[i][s_cat] || ''),
          sumberKas: s_kas >= 0 ? String(sRows[i][s_kas] || '') : ''
        };
      }
    }

    // Baca BUKU_IR untuk pos setoran sumber 'bukuir' (jumlah per kolom)
    var irData = [], irColMap = {};
    var sheetIR = ss.getSheetByName(CONFIG.SHEETS.BUKU_IR);
    if (sheetIR && sheetIR.getLastRow() > 1) {
      var irRows = sheetIR.getDataRange().getValues();
      irColMap = headerMap_(irRows[0]);
      var ir_perid = irColMap['periodeid'] !== undefined ? irColMap['periodeid'] : 2;
      for (var j = 1; j < irRows.length; j++) {
        if (!irRows[j][0]) continue;
        if (periodeId && String(irRows[j][ir_perid]) !== periodeId) continue;
        irData.push(irRows[j]);
      }
    }

    // Untuk sumber 'pemasukan': total penerimaan per jenisId + % Desa dari Master Pemasukan
    var masukPerJenis = {}; // jenisId → total nominal periode aktif
    var sheetPmsk = ss.getSheetByName(CONFIG.SHEETS.INPUT_PENERIMAAN);
    if (sheetPmsk && sheetPmsk.getLastRow() > 1) {
      var pmRows = sheetPmsk.getDataRange().getValues();
      var pmH = headerMap_(pmRows[0]);
      for (var k = 1; k < pmRows.length; k++) {
        if (!hGet_(pmRows[k], pmH, 'id', 0)) continue;
        if (periodeId && String(hGet_(pmRows[k], pmH, 'periodeid', 1)) !== periodeId) continue;
        var jid = String(hGet_(pmRows[k], pmH, 'jenisid', 2) || '');
        if (!jid) continue;
        masukPerJenis[jid] = (masukPerJenis[jid] || 0) + (Number(hGet_(pmRows[k], pmH, 'nominal', 5)) || 0);
      }
    }
    var pctDesaMap = {}; // jenisKode → % Desa
    var sheetMP = ss.getSheetByName(CONFIG.SHEETS.PEMASUKAN);
    if (sheetMP && sheetMP.getLastRow() > 1) {
      var mpRows = sheetMP.getDataRange().getValues();
      var mpH = headerMap_(mpRows[0]);
      for (var m = 1; m < mpRows.length; m++) {
        var kode = String(hGet_(mpRows[m], mpH, 'kode', 0) || hGet_(mpRows[m], mpH, 'id', 0) || '');
        if (!kode) continue;
        pctDesaMap[kode] = Number(hGet_(mpRows[m], mpH, '%desa', 4)) || 0;
      }
    }

    var result = [];
    var posH = headerMap_(posRows[0]);
    for (var i = 1; i < posRows.length; i++) {
      if (!posRows[i][0]) continue;
      var posStatus = posH['status'] !== undefined ? posRows[i][posH['status']] : posRows[i][4];
      if (String(posStatus) !== 'Aktif') continue;
      var posId = posRows[i][0];
      var posNama = posRows[i][posH['nama'] !== undefined ? posH['nama'] : 1];
      var sumberTipe = String(posRows[i][posH['sumbertipe'] !== undefined ? posH['sumbertipe'] : 2] || 'manual').toLowerCase();
      var sumberRef = String(posRows[i][posH['sumberref'] !== undefined ? posH['sumberref'] : 3] || '');
      var pengeluaranRef = String(posRows[i][posH['pengeluaranref'] !== undefined ? posH['pengeluaranref'] : 6] || '');

      // Hitung target sesuai sumber data (FK eksplisit — tidak ada tebak nama)
      var target = 0;
      var sumberKet = '';
      if (sumberTipe === 'bukuir') {
        // Per kolom — sesuai pilihan dropdown di tabel Pos Setoran (Sumber Ref)
        var colIdx = mapPosNamaToBukuIRCol(sumberRef, sumberRef, irColMap);
        if (colIdx >= 0) {
          for (var r = 0; r < irData.length; r++) target += Number(irData[r][colIdx]) || 0;
        }
        sumberKet = 'Buku IR: ' + sumberRef;
      } else if (sumberTipe === 'pemasukan') {
        var totalMasuk = masukPerJenis[sumberRef] || 0;
        var pct = pctDesaMap[sumberRef];
        // Jika % Desa diset, ambil porsinya; jika tidak, ambil total penuh
        target = (pct && pct > 0) ? Math.round(totalMasuk * pct / 100) : totalMasuk;
        sumberKet = 'Pemasukan' + (pct ? ' (' + pct + '% Desa)' : '');
      } else {
        target = Number(posRows[i][posH['target'] !== undefined ? posH['target'] : 5]) || 0;
        sumberKet = 'Manual';
      }

      var realisasi = setoranMap[posId] ? setoranMap[posId].realisasi : 0;
      var sisa = target - realisasi;
      var pctReal = target > 0 ? Math.round((realisasi / target) * 100) : (realisasi > 0 ? 100 : 0);
      result.push({
        id: posId, nama: posNama,
        sumberTipe: sumberTipe, sumberRef: sumberRef, sumberKet: sumberKet,
        pengeluaranRef: pengeluaranRef,
        target: target, realisasi: realisasi, sisa: sisa, persen: pctReal,
        status: pctReal >= 100 ? 'Lunas' : 'Belum Lunas',
        catatan: setoranMap[posId] ? setoranMap[posId].catatan : '',
        sumberKas: setoranMap[posId] ? (setoranMap[posId].sumberKas || 'Tunai') : 'Tunai',
        isAuto: sumberTipe !== 'manual'
      });
    }
    return { success: true, data: result };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function submitRealisasiSetoran(data) {
  try {
    var auth = requirePerm('setoran.realisasi');
    if (!auth.success) return { success: false, message: auth.message };
    var ss = getSS_();
    var periode = getPeriodeAktif();
    if (!periode) return { success: false, message: 'Tidak ada periode aktif' };
    var sheet = ss.getSheetByName(CONFIG.SHEETS.SETORAN_DESA);
    if (!sheet) {
      sheet = ss.insertSheet(CONFIG.SHEETS.SETORAN_DESA);
      sheet.appendRow(['ID', 'PosID', 'PeriodeID', 'Realisasi', 'Catatan', 'Sumber Kas', 'Pengeluaran ID', 'Updated By', 'Updated At']);
    }
    // Pastikan kolom baru ada (evolusi skema aman)
    ensureColumns_(sheet, ['Sumber Kas', 'Pengeluaran ID', 'Updated By', 'Updated At']);

    var realisasi = Number(data.realisasi) || 0;
    var sumberKas = data.sumberKas === 'Bank' ? 'Bank' : 'Tunai';

    // Cari nama pos + pengeluaran ref untuk mencatat pengeluaran kas
    var posNama = '', pengeluaranRef = '';
    var sheetPos = ss.getSheetByName(CONFIG.SHEETS.POS_SETORAN);
    if (sheetPos && sheetPos.getLastRow() > 1) {
      var posRows = sheetPos.getDataRange().getValues();
      var posH = headerMap_(posRows[0]);
      for (var p = 1; p < posRows.length; p++) {
        if (String(posRows[p][0]) === String(data.posId)) {
          posNama = String(hGet_(posRows[p], posH, 'nama', 1) || '');
          pengeluaranRef = String(hGet_(posRows[p], posH, 'pengeluaranref', 6) || '');
          break;
        }
      }
    }

    var rows = sheet.getDataRange().getValues();
    var hdr = headerMap_(rows[0]);
    var c_posid = hdr['posid'] !== undefined ? hdr['posid'] : 1;
    var c_perid = hdr['periodeid'] !== undefined ? hdr['periodeid'] : 2;
    var c_real = hdr['realisasi'] !== undefined ? hdr['realisasi'] : 3;
    var c_cat = hdr['catatan'] !== undefined ? hdr['catatan'] : 4;
    var c_kas = hdr['sumberkas'] !== undefined ? hdr['sumberkas'] : 5;
    var c_pkid = hdr['pengeluaranid'] !== undefined ? hdr['pengeluaranid'] : 6;
    var c_uby = hdr['updatedby'] !== undefined ? hdr['updatedby'] : 7;
    var c_uat = hdr['updatedat'] !== undefined ? hdr['updatedat'] : 8;

    var found = false, existingPkId = '', rowNum = -1;
    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][c_posid]) === String(data.posId) && String(rows[i][c_perid]) === String(periode.id)) {
        existingPkId = String(rows[i][c_pkid] || '');
        rowNum = i + 1;
        found = true;
        break;
      }
    }

    // Catat realisasi sebagai pengeluaran kas (single source of truth untuk saldo).
    // Hanya jika pos punya Pengeluaran Ref (FK) dan realisasi > 0.
    var pkId = existingPkId;
    if (pengeluaranRef) {
      pkId = upsertSetoranPengeluaran_(existingPkId, {
        jenisId: pengeluaranRef, periodeId: periode.id, nominal: realisasi,
        sumberKas: sumberKas, catatan: 'Setoran Desa: ' + (posNama || data.posId) + (data.catatan ? ' — ' + data.catatan : ''),
        email: auth.user.email
      });
    }

    if (found) {
      sheet.getRange(rowNum, c_real + 1).setValue(realisasi);
      sheet.getRange(rowNum, c_cat + 1).setValue(data.catatan || '');
      sheet.getRange(rowNum, c_kas + 1).setValue(sumberKas);
      sheet.getRange(rowNum, c_pkid + 1).setValue(pkId || '');
      sheet.getRange(rowNum, c_uby + 1).setValue(auth.user.email);
      sheet.getRange(rowNum, c_uat + 1).setValue(toDateStr_(new Date()));
    } else {
      var id = generateID('STR');
      sheet.appendRow([id, data.posId, periode.id, realisasi, data.catatan || '', sumberKas, pkId || '', auth.user.email, toDateStr_(new Date())]);
    }
    try {
      var cache2 = CacheService.getScriptCache();
      cache2.remove('master_trx_data');
      cache2.remove('laporan_setoran_' + periode.id);
      cache2.remove('dashboard_saldo');
    } catch(e) {}
    logActivity(auth.user.email, 'SETORAN_DESA', (posNama || data.posId) + ': ' + realisasi);
    return { success: true };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

// Buat atau perbarui satu baris Input Pengeluaran yang mewakili realisasi setoran desa.
// Mengembalikan ID pengeluaran. Jika nominal 0, baris dikosongkan nominalnya (tetap ada untuk audit).
function upsertSetoranPengeluaran_(existingId, info) {
  var ss = getSS_();
  var sheet = ss.getSheetByName(CONFIG.SHEETS.INPUT_PENGELUARAN);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEETS.INPUT_PENGELUARAN);
    sheet.appendRow(['ID', 'Periode ID', 'Jenis ID', 'Tanggal', 'Nominal', 'Sumber Kas', 'Catatan', 'Created By', 'Created At']);
  }
  var rows = sheet.getDataRange().getValues();
  var h = headerMap_(rows[0]);
  var c_id = h['id'] !== undefined ? h['id'] : 0;
  var c_jid = h['jenisid'] !== undefined ? h['jenisid'] : 2;
  var c_tgl = h['tanggal'] !== undefined ? h['tanggal'] : 3;
  var c_nom = h['nominal'] !== undefined ? h['nominal'] : 4;
  var c_kas = h['sumberkas'] !== undefined ? h['sumberkas'] : 5;
  var c_cat = h['catatan'] !== undefined ? h['catatan'] : 6;

  if (existingId) {
    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][c_id]) === String(existingId)) {
        sheet.getRange(i + 1, c_jid + 1).setValue(info.jenisId);
        sheet.getRange(i + 1, c_nom + 1).setValue(info.nominal);
        sheet.getRange(i + 1, c_kas + 1).setValue(info.sumberKas);
        sheet.getRange(i + 1, c_cat + 1).setValue(info.catatan);
        sheet.getRange(i + 1, c_tgl + 1).setValue(toDateStr_(new Date()));
        return existingId;
      }
    }
  }
  // Tidak ditemukan / belum ada → buat baru
  var newId = generateID('PNK');
  sheet.appendRow([newId, info.periodeId, info.jenisId, toDateStr_(new Date()),
    info.nominal, info.sumberKas, info.catatan, info.email, toDateStr_(new Date())]);
  return newId;
}

// ──────────────────────────────────────────────────────
// PDF
// ──────────────────────────────────────────────────────
function generatePDF(periodeId) {
  try {
    var auth = checkAuth();
    if (!auth.success) return { success: false, message: auth.message };
    var rekap = getRekapitulasiData();
    if (!rekap.success) return rekap;

    // Lampirkan data tambahan (best-effort; jika gagal, section dilewati)
    try { var ls = getLaporanSetoran(); if (ls && ls.success) rekap.setoran = ls; } catch(e) {}
    try { var pb = getPembelaanData(); if (pb && pb.success) rekap.pembelaan = pb; } catch(e) {}
    try {
      var pl = getPatunganList();
      if (pl && pl.success) {
        var terobosan = [];
        (pl.data || []).forEach(function(pat) {
          var t = getTagihanPatungan(pat.id);
          if (t && t.success) terobosan.push(t);
        });
        rekap.terobosan = terobosan;
      }
    } catch(e) {}
    try {
      var ss2 = getSS_();
      var stList = getSerahTerimaForPDF_(ss2, (rekap.periode && rekap.periode.id) ? rekap.periode.id : '');
      if (stList && stList.length > 0) rekap.serahTerima = stList;
    } catch(e) {}

    var html = buildPDFHTML(rekap);
    // Kirim HTML string ke client — browser akan render langsung via window.open
    return { success: true, html: html };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function buildPDFHTML(data) {
  var p = data.periode || {};
  var isFinal = p.status === 'Tutup';
  var statusLabel = isFinal ? 'LAPORAN FINAL' : 'LAPORAN INTERIM';
  var statusColor = isFinal ? '#166534' : '#92400e';
  var statusBg   = isFinal ? '#dcfce7' : '#fef3c7';
  var cetakTgl = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd MMMM yyyy HH:mm');

  // Semua style inline — Google Drive viewer memblokir <style> tag
  var S = {
    body:    'font-family:Arial,sans-serif;font-size:12px;color:#111;margin:24px;',
    h1:      'font-size:18px;margin:0 0 4px;',
    h2:      'font-size:14px;margin:20px 0 6px;border-bottom:2px solid #333;padding-bottom:4px;',
    sub:     'color:#555;font-size:12px;margin-bottom:16px;',
    badge:   'display:inline-block;padding:3px 10px;border-radius:4px;font-weight:bold;font-size:11px;background:' + statusBg + ';color:' + statusColor + ';',
    tbl:     'border-collapse:collapse;width:100%;margin-bottom:8px;font-size:11px;',
    th:      'background:#f3f4f6;text-align:left;padding:5px 8px;border:1px solid #d1d5db;font-size:11px;',
    thR:     'background:#f3f4f6;text-align:right;padding:5px 8px;border:1px solid #d1d5db;font-size:11px;',
    td:      'padding:4px 8px;border:1px solid #d1d5db;vertical-align:top;',
    tdR:     'padding:4px 8px;border:1px solid #d1d5db;text-align:right;vertical-align:top;',
    tdE:     'padding:4px 8px;border:1px solid #d1d5db;text-align:center;color:#888;',
    tot:     'padding:4px 8px;border:1px solid #d1d5db;font-weight:bold;background:#e5e7eb;',
    totR:    'padding:4px 8px;border:1px solid #d1d5db;font-weight:bold;background:#e5e7eb;text-align:right;',
    cfSec:   'padding:5px 8px;border:1px solid #d1d5db;background:#f3f4f6;font-weight:bold;',
    cfInd:   'padding:4px 8px 4px 24px;border:1px solid #d1d5db;',
    cfIndR:  'padding:4px 8px;border:1px solid #d1d5db;text-align:right;',
    cfSaldo: 'padding:5px 8px;border:1px solid #d1d5db;background:#dbeafe;font-weight:bold;',
    cfSaldoR:'padding:5px 8px;border:1px solid #d1d5db;background:#dbeafe;font-weight:bold;text-align:right;',
    cfTotal: 'padding:5px 8px;border:1px solid #d1d5db;background:#1e3a5f;color:#fff;font-weight:bold;',
    cfTotalR:'padding:5px 8px;border:1px solid #d1d5db;background:#1e3a5f;color:#fff;font-weight:bold;text-align:right;',
    foot:    'margin-top:32px;font-size:11px;color:#888;'
  };

  var html = '<!DOCTYPE html><html><head><meta charset="UTF-8"><title>Laporan Kas</title></head><body style="' + S.body + '">';

  // Header
  html += '<div style="display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:12px;">' +
    '<div><h1 style="' + S.h1 + '">Laporan Keuangan Kas Kelompok</h1>' +
    '<div style="' + S.sub + '">Periode: <strong>' + esc_(p.nama || '-') + '</strong> &nbsp;|&nbsp; Dicetak: ' + cetakTgl + '</div></div>' +
    '<div><span style="' + S.badge + '">' + statusLabel + '</span></div>' +
    '</div>';

  // ── A. Pemasukan ──
  var pRows = (data.pemasukan || []).slice().sort(function(a, b) { return (a.tanggal||'') < (b.tanggal||'') ? -1 : 1; });
  html += '<h2 style="' + S.h2 + '">A. Pemasukan</h2>';
  html += '<table style="' + S.tbl + '"><tr>' +
    '<th style="' + S.th + '">#</th>' +
    '<th style="' + S.th + '">Tanggal</th>' +
    '<th style="' + S.th + '">Jenis</th>' +
    '<th style="' + S.th + '">Anggota</th>' +
    '<th style="' + S.th + '">Keterangan</th>' +
    '<th style="' + S.th + '">Sumber Kas</th>' +
    '<th style="' + S.thR + '">Nominal</th>' +
    '</tr>';
  var totalPTunai = 0, totalPBank = 0;
  pRows.forEach(function(row, idx) {
    var bg = idx % 2 === 0 ? '' : 'background:#f9fafb;';
    var td = 'padding:4px 8px;border:1px solid #d1d5db;vertical-align:top;' + bg;
    var tdR = 'padding:4px 8px;border:1px solid #d1d5db;text-align:right;vertical-align:top;' + bg;
    html += '<tr>' +
      '<td style="' + td + '">' + (idx+1) + '</td>' +
      '<td style="' + td + '">' + fmtTanggal(row.tanggal) + '</td>' +
      '<td style="' + td + '">' + esc_(row.jenis) + '</td>' +
      '<td style="' + td + '">' + esc_(row.anggota || '') + '</td>' +
      '<td style="' + td + '">' + esc_(row.catatan || '') + '</td>' +
      '<td style="' + td + '">' + esc_(row.sumber) + '</td>' +
      '<td style="' + tdR + '">' + fmtRp(row.nominal) + '</td>' +
      '</tr>';
    if (row.sumber === 'Tunai') totalPTunai += row.nominal;
    else if (row.sumber === 'Bank') totalPBank += row.nominal;
  });
  if (pRows.length === 0) html += '<tr><td colspan="7" style="' + S.tdE + '">Belum ada data pemasukan</td></tr>';
  html += '<tr><td colspan="5" style="' + S.tot + '"></td><td style="' + S.tot + '">Subtotal Tunai</td><td style="' + S.totR + '">' + fmtRp(totalPTunai) + '</td></tr>';
  html += '<tr><td colspan="5" style="' + S.tot + '"></td><td style="' + S.tot + '">Subtotal Bank</td><td style="' + S.totR + '">' + fmtRp(totalPBank) + '</td></tr>';
  html += '<tr><td colspan="5" style="' + S.tot + '"></td><td style="' + S.tot + '">TOTAL PEMASUKAN</td><td style="' + S.totR + '">' + fmtRp(totalPTunai + totalPBank) + '</td></tr>';
  html += '</table>';

  // ── B. Pengeluaran ──
  var pkRows = (data.pengeluaran || []).slice().sort(function(a, b) { return (a.tanggal||'') < (b.tanggal||'') ? -1 : 1; });
  html += '<h2 style="' + S.h2 + '">B. Pengeluaran</h2>';
  html += '<table style="' + S.tbl + '"><tr>' +
    '<th style="' + S.th + '">#</th>' +
    '<th style="' + S.th + '">Tanggal</th>' +
    '<th style="' + S.th + '">Jenis</th>' +
    '<th style="' + S.th + '">Keterangan</th>' +
    '<th style="' + S.th + '">Sumber Kas</th>' +
    '<th style="' + S.thR + '">Nominal</th>' +
    '</tr>';
  var totalPKTunai = 0, totalPKBank = 0;
  pkRows.forEach(function(row, idx) {
    var bg = idx % 2 === 0 ? '' : 'background:#f9fafb;';
    var td = 'padding:4px 8px;border:1px solid #d1d5db;vertical-align:top;' + bg;
    var tdR = 'padding:4px 8px;border:1px solid #d1d5db;text-align:right;vertical-align:top;' + bg;
    html += '<tr>' +
      '<td style="' + td + '">' + (idx+1) + '</td>' +
      '<td style="' + td + '">' + fmtTanggal(row.tanggal) + '</td>' +
      '<td style="' + td + '">' + esc_(row.jenis) + '</td>' +
      '<td style="' + td + '">' + esc_(row.catatan || '') + '</td>' +
      '<td style="' + td + '">' + esc_(row.sumber) + '</td>' +
      '<td style="' + tdR + '">' + fmtRp(row.nominal) + '</td>' +
      '</tr>';
    if (row.sumber === 'Tunai') totalPKTunai += row.nominal;
    else if (row.sumber === 'Bank') totalPKBank += row.nominal;
  });
  if (pkRows.length === 0) html += '<tr><td colspan="6" style="' + S.tdE + '">Belum ada data pengeluaran</td></tr>';
  html += '<tr><td colspan="4" style="' + S.tot + '"></td><td style="' + S.tot + '">Subtotal Tunai</td><td style="' + S.totR + '">' + fmtRp(totalPKTunai) + '</td></tr>';
  html += '<tr><td colspan="4" style="' + S.tot + '"></td><td style="' + S.tot + '">Subtotal Bank</td><td style="' + S.totR + '">' + fmtRp(totalPKBank) + '</td></tr>';
  html += '<tr><td colspan="4" style="' + S.tot + '"></td><td style="' + S.tot + '">TOTAL PENGELUARAN</td><td style="' + S.totR + '">' + fmtRp(totalPKTunai + totalPKBank) + '</td></tr>';
  html += '</table>';

  // ── C. Cash Flow ──
  var saldoAwalTunai  = Number(p.saldoAwalTunai) || 0;
  var saldoAwalBank   = Number(p.saldoAwalBank)  || 0;
  var saldoAkhirTunai = saldoAwalTunai + totalPTunai - totalPKTunai;
  var saldoAkhirBank  = saldoAwalBank  + totalPBank  - totalPKBank;
  var grandTotal = saldoAkhirTunai + saldoAkhirBank;

  html += '<h2 style="' + S.h2 + '">C. Arus Kas (Cash Flow)</h2>';
  html += '<table style="' + S.tbl + '">' +
    '<tr><th style="' + S.th + 'width:60%">Keterangan</th><th style="' + S.thR + 'width:20%">Tunai</th><th style="' + S.thR + 'width:20%">Bank</th></tr>' +

    '<tr><td colspan="3" style="' + S.cfSec + '">Saldo Awal Periode</td></tr>' +
    '<tr><td style="' + S.cfInd + '">Saldo Awal</td><td style="' + S.cfIndR + '">' + fmtRp(saldoAwalTunai) + '</td><td style="' + S.cfIndR + '">' + fmtRp(saldoAwalBank) + '</td></tr>' +

    '<tr><td colspan="3" style="' + S.cfSec + '">(+) Pemasukan</td></tr>' +
    '<tr><td style="' + S.cfInd + '">Total Pemasukan</td><td style="' + S.cfIndR + '">' + fmtRp(totalPTunai) + '</td><td style="' + S.cfIndR + '">' + fmtRp(totalPBank) + '</td></tr>' +

    '<tr><td colspan="3" style="' + S.cfSec + '">(-) Pengeluaran</td></tr>' +
    '<tr><td style="' + S.cfInd + '">Total Pengeluaran</td><td style="' + S.cfIndR + '">(' + fmtRp(totalPKTunai) + ')</td><td style="' + S.cfIndR + '">(' + fmtRp(totalPKBank) + ')</td></tr>' +

    '<tr><td style="' + S.cfSaldo + '">Saldo Akhir Tunai / Bank</td><td style="' + S.cfSaldoR + '">' + fmtRp(saldoAkhirTunai) + '</td><td style="' + S.cfSaldoR + '">' + fmtRp(saldoAkhirBank) + '</td></tr>' +
    '<tr><td style="' + S.cfTotal + '">TOTAL KAS</td><td style="' + S.cfTotalR + '" colspan="2">' + fmtRp(grandTotal) + '</td></tr>' +
    '</table>';

  // ── D. Setoran ke Desa ──
  if (data.setoran && data.setoran.data) {
    var sd = data.setoran;
    var sdRows = sd.data || [];
    html += '<h2 style="' + S.h2 + '">D. Setoran ke Desa</h2>';
    html += '<table style="' + S.tbl + '"><tr>' +
      '<th style="' + S.th + '">#</th>' +
      '<th style="' + S.th + '">Pos / Jenis</th>' +
      '<th style="' + S.thR + '">Kewajiban</th>' +
      '<th style="' + S.thR + '">Sudah Setor</th>' +
      '<th style="' + S.thR + '">Sisa</th>' +
      '<th style="' + S.th + '">Status</th>' +
      '</tr>';
    sdRows.forEach(function(row, idx) {
      var bg = idx % 2 === 0 ? '' : 'background:#f9fafb;';
      var td = 'padding:4px 8px;border:1px solid #d1d5db;vertical-align:top;' + bg;
      var tdR = 'padding:4px 8px;border:1px solid #d1d5db;text-align:right;vertical-align:top;' + bg;
      html += '<tr>' +
        '<td style="' + td + '">' + (idx+1) + '</td>' +
        '<td style="' + td + '">' + esc_(row.nama) + '</td>' +
        '<td style="' + tdR + '">' + fmtRp(row.kewajiban) + '</td>' +
        '<td style="' + tdR + '">' + fmtRp(row.sudahSetor) + '</td>' +
        '<td style="' + tdR + '">' + fmtRp(row.sisa) + '</td>' +
        '<td style="' + td + '">' + esc_(row.status) + '</td>' +
        '</tr>';
    });
    if (sdRows.length === 0) html += '<tr><td colspan="6" style="' + S.tdE + '">Belum ada data setoran desa</td></tr>';
    var sm = sd.summary || {};
    html += '<tr><td colspan="2" style="' + S.tot + '">TOTAL</td>' +
      '<td style="' + S.totR + '">' + fmtRp(sm.totalTarget) + '</td>' +
      '<td style="' + S.totR + '">' + fmtRp(sm.totalSudahSetor) + '</td>' +
      '<td style="' + S.totR + '">' + fmtRp(sm.totalSisa) + '</td>' +
      '<td style="' + S.tot + '"></td></tr>';
    html += '</table>';
  }

  // ── E. Target Pembelaan ──
  if (data.pembelaan) {
    var pb = data.pembelaan;
    var belum = pb.belum || [];
    var lunas = pb.lunas || [];
    var all = lunas.concat(belum);
    html += '<h2 style="' + S.h2 + '">E. Target Pembelaan</h2>';
    html += '<table style="' + S.tbl + '"><tr>' +
      '<th style="' + S.th + '">#</th>' +
      '<th style="' + S.th + '">Jamaah</th>' +
      '<th style="' + S.thR + '">Nominal</th>' +
      '<th style="' + S.th + '">Tgl Janji</th>' +
      '<th style="' + S.th + '">Status</th>' +
      '</tr>';
    var totPbLunas = 0, totPbBelum = 0;
    all.forEach(function(row, idx) {
      var bg = idx % 2 === 0 ? '' : 'background:#f9fafb;';
      var td = 'padding:4px 8px;border:1px solid #d1d5db;vertical-align:top;' + bg;
      var tdR = 'padding:4px 8px;border:1px solid #d1d5db;text-align:right;vertical-align:top;' + bg;
      var isLunas = row.status === 'Lunas';
      if (isLunas) totPbLunas += Number(row.nominal) || 0;
      else totPbBelum += Number(row.nominal) || 0;
      html += '<tr>' +
        '<td style="' + td + '">' + (idx+1) + '</td>' +
        '<td style="' + td + '">' + esc_(row.anggotaNama) + '</td>' +
        '<td style="' + tdR + '">' + fmtRp(row.nominal) + '</td>' +
        '<td style="' + td + '">' + (row.tanggalJanji ? fmtTanggal(row.tanggalJanji) : '-') + '</td>' +
        '<td style="' + td + '">' + esc_(row.status) + '</td>' +
        '</tr>';
    });
    if (all.length === 0) html += '<tr><td colspan="5" style="' + S.tdE + '">Belum ada data pembelaan</td></tr>';
    html += '<tr><td colspan="2" style="' + S.tot + '">Sudah Lunas</td><td style="' + S.totR + '">' + fmtRp(totPbLunas) + '</td><td colspan="2" style="' + S.tot + '"></td></tr>';
    html += '<tr><td colspan="2" style="' + S.tot + '">Belum Lunas</td><td style="' + S.totR + '">' + fmtRp(totPbBelum) + '</td><td colspan="2" style="' + S.tot + '"></td></tr>';
    html += '<tr><td colspan="2" style="' + S.tot + '">TOTAL TARGET</td><td style="' + S.totR + '">' + fmtRp(totPbLunas + totPbBelum) + '</td><td colspan="2" style="' + S.tot + '"></td></tr>';
    html += '</table>';
  }

  // ── F. Hasil Musyawarah / Terobosan Kelompok ──
  if (data.terobosan && data.terobosan.length > 0) {
    html += '<h2 style="' + S.h2 + '">F. Hasil Musyawarah / Terobosan Kelompok</h2>';
    data.terobosan.forEach(function(t) {
      var pat = t.patungan || {};
      var tag = t.tagihan || [];
      html += '<div style="font-weight:bold;font-size:12px;margin:12px 0 4px;">' + esc_(pat.nama) +
        ' <span style="font-weight:normal;color:#666;">(' + (pat.tanggal ? fmtTanggal(pat.tanggal) : '-') + ')</span></div>';
      html += '<table style="' + S.tbl + '"><tr>' +
        '<th style="' + S.th + '">#</th>' +
        '<th style="' + S.th + '">Jamaah</th>' +
        '<th style="' + S.th + '">Grade</th>' +
        '<th style="' + S.thR + '">Nominal</th>' +
        '<th style="' + S.th + '">Status</th>' +
        '</tr>';
      tag.forEach(function(row, idx) {
        var bg = idx % 2 === 0 ? '' : 'background:#f9fafb;';
        var td = 'padding:4px 8px;border:1px solid #d1d5db;vertical-align:top;' + bg;
        var tdR = 'padding:4px 8px;border:1px solid #d1d5db;text-align:right;vertical-align:top;' + bg;
        html += '<tr>' +
          '<td style="' + td + '">' + (idx+1) + '</td>' +
          '<td style="' + td + '">' + esc_(row.anggotaNama) + '</td>' +
          '<td style="' + td + '">' + esc_(row.grade) + '</td>' +
          '<td style="' + tdR + '">' + fmtRp(row.nominal) + '</td>' +
          '<td style="' + td + '">' + (row.statusBayar === 'Lunas' ? 'Lunas' : 'Belum') + '</td>' +
          '</tr>';
      });
      if (tag.length === 0) html += '<tr><td colspan="5" style="' + S.tdE + '">Belum ada penerobosan</td></tr>';
      html += '<tr><td colspan="3" style="' + S.tot + '">Target</td><td style="' + S.totR + '">' + fmtRp(t.totalTarget) + '</td><td style="' + S.tot + '"></td></tr>';
      html += '<tr><td colspan="3" style="' + S.tot + '">Sudah Bayar</td><td style="' + S.totR + '">' + fmtRp(t.totalLunas) + '</td><td style="' + S.tot + '"></td></tr>';
      html += '<tr><td colspan="3" style="' + S.tot + '">Belum Bayar</td><td style="' + S.totR + '">' + fmtRp(t.totalBelum) + '</td><td style="' + S.tot + '"></td></tr>';
      html += '</table>';
    });
  }

  // ── G. Serah Terima Penerobos ──
  if (data.serahTerima && data.serahTerima.length > 0) {
    html += '<h2 style="' + S.h2 + '">G. Serah Terima Penerobos</h2>';
    html += '<table style="' + S.tbl + '"><tr>' +
      '<th style="' + S.th + '">Tanggal</th>' +
      '<th style="' + S.th + '">Dari (Penerobos)</th>' +
      '<th style="' + S.th + '">Dikonfirmasi Oleh</th>' +
      '<th style="' + S.th + '">Tujuan</th>' +
      '<th style="' + S.thR + '">Total</th>' +
      '</tr>';
    var totalST = 0;
    data.serahTerima.forEach(function(st, idx) {
      var bg = idx % 2 === 0 ? '' : 'background:#f9fafb;';
      var td = 'padding:4px 8px;border:1px solid #d1d5db;vertical-align:top;' + bg;
      var tdR = 'padding:4px 8px;border:1px solid #d1d5db;text-align:right;vertical-align:top;' + bg;
      html += '<tr>' +
        '<td style="' + td + '">' + (st.tanggal ? fmtTanggal(st.tanggal) : '-') + '</td>' +
        '<td style="' + td + '">' + esc_(st.penerobosEmail) + '</td>' +
        '<td style="' + td + '">' + esc_(st.dikonfirmasiBy) + '</td>' +
        '<td style="' + td + '">' + esc_(st.sumberTujuan) + '</td>' +
        '<td style="' + tdR + '">' + fmtRp(st.total) + '</td>' +
        '</tr>';
      // Rincian items
      if (st.items && st.items.length > 0) {
        var itemStr = st.items.map(function(it) { return esc_(it.jenisNama) + ': ' + fmtRp(it.nominal); }).join(' &nbsp;|&nbsp; ');
        html += '<tr><td colspan="5" style="padding:2px 8px 6px 20px;border:1px solid #d1d5db;font-size:11px;color:#555;">' + itemStr + '</td></tr>';
      }
      totalST += st.total;
    });
    html += '<tr><td colspan="4" style="' + S.tot + '">Total Serah Terima</td><td style="' + S.totR + '">' + fmtRp(totalST) + '</td></tr>';
    html += '</table>';
  }

  html += '<div style="' + S.foot + '">Laporan ini dibuat otomatis oleh sistem Kas Kelompok pada ' + cetakTgl + '.</div>';
  html += '</body></html>';
  return html;
}

function esc_(s) {
  if (!s) return '-';
  return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
}

// ──────────────────────────────────────────────────────
// TEST FUNCTIONS
// ──────────────────────────────────────────────────────
function testDashboard() {
  var r = getDashboardData();
  Logger.log(JSON.stringify(r));
}

function testFindPeriode() {
  var r = getPeriodeAktif();
  Logger.log(JSON.stringify(r));
}

function testBukuIR() {
  var r = getBukuIRData();
  Logger.log(JSON.stringify(r));
}

// ──────────────────────────────────────────────────────
// REKONSILIASI BANK — TOP-LEVEL SCOPE (BUG FIXED)
// ──────────────────────────────────────────────────────
function getBankDaily(periodeId) {
  try {
    var auth = checkAuth();
    if (!auth.success) return { success: false, message: auth.message };
    var ss = getSS_();
    var sheet = ss.getSheetByName(CONFIG.SHEETS.BANK_DAILY);
    if (!sheet) return { success: true, data: [] };
    var rows = sheet.getDataRange().getValues();
    var h = headerMap_(rows[0]);
    var result = [];
    for (var i = 1; i < rows.length; i++) {
      if (!hGet_(rows[i], h, 'id', 0)) continue;
      if (periodeId && String(hGet_(rows[i], h, 'periodeid', 1)) !== periodeId) continue;
      result.push({
        id: String(hGet_(rows[i], h, 'id', 0)),
        periodeId: String(hGet_(rows[i], h, 'periodeid', 1)),
        tanggal: toDateStr_(hGet_(rows[i], h, 'tanggal', 2)),
        saldoAwal: Number(hGet_(rows[i], h, 'saldoawal', 3)) || 0,
        pemasukan: Number(hGet_(rows[i], h, 'pemasukan', 4)) || 0,
        pengeluaran: Number(hGet_(rows[i], h, 'pengeluaran', 5)) || 0,
        saldoAkhirTeoritis: Number(hGet_(rows[i], h, 'saldoakhirteoritis', 6)) || 0,
        saldoAkhirActual: Number(hGet_(rows[i], h, 'saldoakhiractual', 7)) || 0,
        selisih: Number(hGet_(rows[i], h, 'selisih', 8)) || 0,
        status: String(hGet_(rows[i], h, 'status', 9) || ''),
        catatan: String(hGet_(rows[i], h, 'catatan', 10) || '')
      });
    }
    return { success: true, data: result };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function addBankTransaction(data) {
  try {
    var auth = requirePerm('bank.input');
    if (!auth.success) return { success: false, message: auth.message };
    var ss = getSS_();
    var sheet = ss.getSheetByName(CONFIG.SHEETS.BANK_DAILY);
    if (!sheet) {
      sheet = ss.insertSheet(CONFIG.SHEETS.BANK_DAILY);
      sheet.appendRow(['ID', 'PeriodeID', 'Tanggal', 'Saldo Awal', 'Pemasukan', 'Pengeluaran', 'Saldo Akhir Teoritis', 'Saldo Akhir Actual', 'Selisih', 'Status', 'Catatan', 'Last Updated']);
    }
    var id = generateID('BNK');
    var saldoAkhirTeoritis = (Number(data.saldoAwal) || 0) + (Number(data.pemasukan) || 0) - (Number(data.pengeluaran) || 0);
    var selisih = (Number(data.saldoAkhirActual) || 0) - saldoAkhirTeoritis;
    sheet.appendRow([id, data.periodeId, toDateStr_(data.tanggal),
      Number(data.saldoAwal) || 0, Number(data.pemasukan) || 0, Number(data.pengeluaran) || 0,
      saldoAkhirTeoritis, Number(data.saldoAkhirActual) || 0, selisih,
      selisih === 0 ? 'Balance' : 'Selisih', data.catatan || '', toDateStr_(new Date())]);
    try { CacheService.getScriptCache().remove('master_trx_data'); } catch(e) {}
    return { success: true, id: id };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function updateBankDaily(data) {
  try {
    var auth = requirePerm('bank.manage');
    if (!auth.success) return { success: false, message: auth.message };
    var ss = getSS_();
    var sheet = ss.getSheetByName(CONFIG.SHEETS.BANK_DAILY);
    if (!sheet) return { success: false, message: 'Sheet tidak ditemukan' };
    var rows = sheet.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      if (rows[i][0] === data.id) {
        var saldoAkhirTeoritis = (Number(data.saldoAwal) || 0) + (Number(data.pemasukan) || 0) - (Number(data.pengeluaran) || 0);
        var selisih = (Number(data.saldoAkhirActual) || 0) - saldoAkhirTeoritis;
        sheet.getRange(i + 1, 4, 1, 9).setValues([[
          Number(data.saldoAwal) || 0, Number(data.pemasukan) || 0, Number(data.pengeluaran) || 0,
          saldoAkhirTeoritis, Number(data.saldoAkhirActual) || 0, selisih,
          selisih === 0 ? 'Balance' : 'Selisih', data.catatan || '', toDateStr_(new Date())
        ]]);
    try { CacheService.getScriptCache().remove('master_trx_data'); } catch(e) {}
        return { success: true };
      }
    }
    return { success: false, message: 'Data tidak ditemukan' };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function getBankPending(periodeId) {
  try {
    var auth = checkAuth();
    if (!auth.success) return { success: false, message: auth.message };
    var ss = getSS_();
    var sheet = ss.getSheetByName(CONFIG.SHEETS.BANK_PENDING);
    if (!sheet) return { success: true, data: [] };
    var rows = sheet.getDataRange().getValues();
    var h = headerMap_(rows[0]);
    var result = [];
    for (var i = 1; i < rows.length; i++) {
      if (!hGet_(rows[i], h, 'id', 0)) continue;
      if (periodeId && String(hGet_(rows[i], h, 'periodeid', 1)) !== periodeId) continue;
      result.push({
        id: String(hGet_(rows[i], h, 'id', 0)),
        periodeId: String(hGet_(rows[i], h, 'periodeid', 1)),
        tanggal: toDateStr_(hGet_(rows[i], h, 'tanggal', 2)),
        keterangan: String(hGet_(rows[i], h, 'keterangan', 3) || ''),
        nominal: Number(hGet_(rows[i], h, 'nominal', 4)) || 0,
        sumber: String(hGet_(rows[i], h, 'sumber', 5) || ''),
        status: String(hGet_(rows[i], h, 'status', 6) || ''),
        tanggalFound: toDateStr_(hGet_(rows[i], h, 'tanggalfound', 7)),
        catatan: String(hGet_(rows[i], h, 'catatan', 8) || '')
      });
    }
    return { success: true, data: result };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function addManualBankIncome(data) {
  try {
    var auth = requirePerm('bank.manage');
    if (!auth.success) return { success: false, message: auth.message };
    var result = submitTransaksi({
      tipe: 'masuk', jenisId: data.jenisId, nominal: data.nominal,
      sumberKas: 'Bank', tanggal: data.tanggal, catatan: data.catatan
    });
    return result;
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function addPendingTransaction(data) {
  try {
    var auth = requirePerm('bank.manage');
    if (!auth.success) return { success: false, message: auth.message };
    var ss = getSS_();
    var sheet = ss.getSheetByName(CONFIG.SHEETS.BANK_PENDING);
    if (!sheet) {
      sheet = ss.insertSheet(CONFIG.SHEETS.BANK_PENDING);
      sheet.appendRow(['ID', 'PeriodeID', 'Tanggal', 'Keterangan', 'Nominal', 'Sumber', 'Status', 'Tanggal Found', 'Catatan', 'Last Updated']);
    }
    var periode = getPeriodeAktif();
    if (!periode) return { success: false, message: 'Tidak ada periode aktif' };
    var id = generateID('PND');
    sheet.appendRow([id, periode.id, toDateStr_(data.tanggal), data.keterangan, Number(data.nominal) || 0,
      data.sumber || 'Manual', 'Pending', '', data.catatan || '', toDateStr_(new Date())]);
    try { CacheService.getScriptCache().remove('master_trx_data'); } catch(e) {}
    return { success: true, id: id };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function getRekonsiliasiData() {
  try {
    var auth = checkAuth();
    if (!auth.success) return { success: false, message: auth.message };
    // FIX BUG 2: getPeriodeAktif() returns object directly, not {success, data} wrapper
    var periode = getPeriodeAktif();
    if (!periode) return { success: false, message: 'Tidak ada periode aktif' };

    var periodeId = periode.id;
    var bankDaily = getBankDaily(periodeId);
    var bankPending = getBankPending(periodeId);
    var saldo = calculateSaldo(periodeId, periode);

    var totalPending = 0;
    (bankPending.data || []).forEach(function(p) {
      if (p.status === 'Pending') totalPending += p.nominal;
    });

    var lastDaily = null;
    if (bankDaily.data && bankDaily.data.length > 0) {
      lastDaily = bankDaily.data[bankDaily.data.length - 1];
    }

    return {
      success: true,
      periode: periode,
      saldoSystem: saldo.bank,
      saldoBankAktual: lastDaily ? lastDaily.saldoAkhirActual : 0,
      selisih: lastDaily ? (lastDaily.saldoAkhirActual - saldo.bank) : 0,
      totalPending: totalPending,
      bankDaily: bankDaily.data || [],
      bankPending: bankPending.data || []
    };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function submitRekonsiliasiBank(data) {
  try {
    var auth = requirePerm('bank.manage');
    if (!auth.success) return { success: false, message: auth.message };
    var result = addBankTransaction(data);
    if (!result.success) return result;
    logActivity(auth.user.email, 'REKONSILIASI', 'Saldo actual: ' + data.saldoAkhirActual);
    try { CacheService.getScriptCache().remove('master_trx_data'); } catch(e) {}
    return { success: true };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function findPendingTransactions(keyword) {
  try {
    var auth = checkAuth();
    if (!auth.success) return { success: false, message: auth.message };
    var periode = getPeriodeAktif();
    var periodeId = periode ? periode.id : null;
    var bankPending = getBankPending(periodeId);
    var results = (bankPending.data || []).filter(function(p) {
      return p.status === 'Pending' && (!keyword || p.keterangan.toLowerCase().indexOf(keyword.toLowerCase()) !== -1);
    });
    try { CacheService.getScriptCache().remove('master_trx_data'); } catch(e) {}
    return { success: true, data: results };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function updatePendingStatus(id, status, catatan) {
  try {
    var auth = requirePerm('bank.manage');
    if (!auth.success) return { success: false, message: auth.message };
    var ss = getSS_();
    var sheet = ss.getSheetByName(CONFIG.SHEETS.BANK_PENDING);
    if (!sheet) return { success: false, message: 'Sheet tidak ditemukan' };
    var rows = sheet.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      if (rows[i][0] === id) {
        sheet.getRange(i + 1, 7, 1, 4).setValues([[
          status,
          status === 'Found' ? toDateStr_(new Date()) : toDateStr_(rows[i][7]),
          catatan || rows[i][8],
          toDateStr_(new Date())
        ]]);
    try { CacheService.getScriptCache().remove('master_trx_data'); } catch(e) {}
        return { success: true };
      }
    }
    return { success: false, message: 'Data tidak ditemukan' };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

// ──────────────────────────────────────────────────────
// LAPORAN SETORAN
// ──────────────────────────────────────────────────────

function hitungJumlahBulan(tglMulai) {
  try {
    if (!tglMulai) return 1;
    var mulai;
    if (tglMulai instanceof Date) {
      mulai = tglMulai;
    } else {
      // Coba parse dd/MM/yyyy atau ISO
      var s = String(tglMulai).trim();
      if (s.indexOf('/') !== -1) {
        var parts = s.split('/');
        if (parts.length === 3) {
          mulai = new Date(Number(parts[2]), Number(parts[1]) - 1, Number(parts[0]));
        }
      } else {
        mulai = new Date(s);
      }
    }
    if (!mulai || isNaN(mulai.getTime())) return 1;
    var now = new Date();
    var bulan = (now.getFullYear() - mulai.getFullYear()) * 12 + (now.getMonth() - mulai.getMonth());
    // Jika hari sekarang >= hari mulai, hitung bulan ini juga
    if (now.getDate() >= mulai.getDate()) bulan += 1;
    return bulan < 1 ? 1 : bulan;
  } catch(e) {
    return 1;
  }
}

// Kembalikan nama kolom header BUKU_IR yang sesuai dengan pos setoran.
// Menggunakan nama kolom (bukan index numerik hardcoded) agar tidak rapuh.
// PENTING: nama pos seperti "infak jumatan" TIDAK boleh dipetakan ke infakdaerah —
// hanya nama yang secara eksplisit merujuk kolom Buku IR yang diizinkan.
function mapPosNamaToIRColName(formula, nama) {
  var src = String(formula || nama || '').toLowerCase().trim();
  var clean = src.replace(/[\s\/]/g, ''); // normalisasi sama seperti headerMap_

  // IR — harus exact, hindari false-positive pada "cicilan" dsb
  if (src === 'ir' || clean === 'ir' || src === 'col5' || src === '5') return 'ir';

  // 1/10 IR — berbagai penulisan; clean '1/10ir' → '110ir', 'ir10', '110 ir'
  if (src.indexOf('1/10') !== -1 || clean === 'ir10' || clean === '110ir' || src === 'col6' || src === '6') return 'ir10';

  // Cicilan
  if (src.indexOf('cicilan') !== -1 || src === 'col7' || src === '7') return 'cicilan';

  // Infak Daerah — HARUS spesifik; "infak jumatan", "infak romadhon" TIDAK termasuk
  if (clean === 'infakdaerah' || src === 'infak daerah' || src === 'col8' || src === '8') return 'infakdaerah';

  // Index
  if (src.indexOf('index') !== -1 || src === 'col9' || src === '9') return 'index';

  return '';
}

// Wrapper: return index numerik via irColMap.
// Menangani dua kemungkinan header: 'IR10' (→ key 'ir10') atau '1/10 IR' (→ key '110ir').
function mapPosNamaToBukuIRCol(formula, nama, irColMap) {
  var colName = mapPosNamaToIRColName(formula, nama);
  if (!colName || !irColMap) return -1;
  if (irColMap[colName] !== undefined) return irColMap[colName];
  // Fallback alias: headerMap_ pada '1/10 IR' → '110ir', sedangkan kita return 'ir10' (dan sebaliknya)
  var alias = { 'ir10': '110ir', '110ir': 'ir10' };
  var alt = alias[colName];
  if (alt && irColMap[alt] !== undefined) return irColMap[alt];
  return -1;
}

function getLaporanSetoran() {
  try {
    var auth = checkAuth();
    if (!auth.success) return { success: false, message: auth.message };

    var periode = getPeriodeAktif();
    if (!periode) return { success: false, message: 'Tidak ada periode aktif' };
    var periodeId = periode.id;

    var cache = CacheService.getScriptCache();
    var cacheKey = 'laporan_setoran_' + periodeId;
    var cached = cache.get(cacheKey);
    if (cached) {
      try {
        var parsed = JSON.parse(cached);
        parsed.success = true;
        return parsed;
      } catch(e) {}
    }

    var ss = getSS_();

    // ── 1. Master Pemasukan: ambil %Desa & %Daerah per jenis ──
    var sheetMaster = ss.getSheetByName(CONFIG.SHEETS.PEMASUKAN);
    var jenisList = [];          // {kode, nama, kategori, pctDesa, pctDaerah}
    var jenisMap = {};           // kode → jenis
    if (sheetMaster && sheetMaster.getLastRow() > 1) {
      var mRows = sheetMaster.getDataRange().getValues();
      var mH = headerMap_(mRows[0]);
      for (var i = 1; i < mRows.length; i++) {
        var kode = String(hGet_(mRows[i], mH, 'kode', 0) || '');
        if (!kode) continue;
        var statusJ = String(hGet_(mRows[i], mH, 'status', 7) || '').trim().toLowerCase();
        if (statusJ === 'nonaktif') continue;
        var jenis = {
          kode: kode,
          nama: String(hGet_(mRows[i], mH, 'namapemasukan', 1) || hGet_(mRows[i], mH, 'nama', 1) || ''),
          kategori: String(hGet_(mRows[i], mH, 'kategori', 2) || ''),
          pctDesa: Number(hGet_(mRows[i], mH, '%desa', 4)) || 0,
          pctDaerah: Number(hGet_(mRows[i], mH, '%daerah', 5)) || 0
        };
        jenisList.push(jenis);
        jenisMap[kode] = jenis;
      }
    }

    // ── 2. Input Penerimaan: jumlahkan nominal per jenis untuk periode aktif ──
    var totalMasukPerJenis = {}; // kode → total nominal
    var sheetPmsk = ss.getSheetByName(CONFIG.SHEETS.INPUT_PENERIMAAN);
    if (sheetPmsk && sheetPmsk.getLastRow() > 1) {
      var pRows = sheetPmsk.getDataRange().getValues();
      var pH = headerMap_(pRows[0]);
      for (var i = 1; i < pRows.length; i++) {
        if (!hGet_(pRows[i], pH, 'id', 0)) continue;
        if (String(hGet_(pRows[i], pH, 'periodeid', 1)) !== periodeId) continue;
        var jid = String(hGet_(pRows[i], pH, 'jenisid', 2) || '');
        var nom = Number(hGet_(pRows[i], pH, 'nominal', 5)) || 0;
        totalMasukPerJenis[jid] = (totalMasukPerJenis[jid] || 0) + nom;
      }
    }

    // ── 3. Setoran Desa: realisasi per jenis (disimpan di kolom PosID) per periode ──
    var sheetSetoran = ss.getSheetByName(CONFIG.SHEETS.SETORAN_DESA);
    var setoranMap = {};        // kode → realisasi
    var setoranTerisi = {};     // kode → true jika sudah pernah diinput manual
    if (sheetSetoran && sheetSetoran.getLastRow() > 1) {
      var sRows = sheetSetoran.getDataRange().getValues();
      var sHdrL = headerMap_(sRows[0]);
      var sp_posid = sHdrL['posid'] !== undefined ? sHdrL['posid'] : 1;
      var sp_perid = sHdrL['periodeid'] !== undefined ? sHdrL['periodeid'] : 2;
      var sp_real  = sHdrL['realisasi'] !== undefined ? sHdrL['realisasi'] : 3;
      for (var i = 1; i < sRows.length; i++) {
        if (!sRows[i][0]) continue;
        if (sRows[i][sp_perid] && String(sRows[i][sp_perid]) !== periodeId) continue;
        setoranMap[String(sRows[i][sp_posid])] = Number(sRows[i][sp_real]) || 0;
        setoranTerisi[String(sRows[i][sp_posid])] = true;
      }
    }

    // ── 3b. Target musyawaroh (tarif × jamaah aktif × bulan) — kewajiban setor ──
    var jumlahBulan = hitungJumlahBulan(periode.tanggalMulai);
    var sheetAnggota = ss.getSheetByName(CONFIG.SHEETS.ANGGOTA);
    var jumlahJamaahAktif = 0;
    if (sheetAnggota && sheetAnggota.getLastRow() > 1) {
      var angRows = sheetAnggota.getDataRange().getValues();
      var angH = headerMap_(angRows[0]);
      for (var i = 1; i < angRows.length; i++) {
        if (!hGet_(angRows[i], angH, 'id', 0)) continue;
        var statusAng = String(hGet_(angRows[i], angH, 'status', 4) || '').trim().toLowerCase();
        if (statusAng !== 'nonaktif') jumlahJamaahAktif++;
      }
    }
    var sheetMsy = ss.getSheetByName(CONFIG.SHEETS.MUSYAWARAH);
    var musyawarahMap = {};
    if (sheetMsy && sheetMsy.getLastRow() > 1) {
      var msyRows = sheetMsy.getDataRange().getValues();
      var msyH = headerMap_(msyRows[0]);
      for (var i = 1; i < msyRows.length; i++) {
        if (!hGet_(msyRows[i], msyH, 'id', 0)) continue;
        var msyNama = String(hGet_(msyRows[i], msyH, 'nama', 1) || '').toLowerCase().trim();
        if (msyNama) musyawarahMap[msyNama] = Number(hGet_(msyRows[i], msyH, 'nilai', 2)) || 0;
      }
    }

    // ── 4. Hitung jatah Desa per jenis ──
    var totalJatahDesa = 0, totalJatahDaerah = 0, totalSudahSetor = 0, totalTargetMsy = 0, totalTopUp = 0;
    var dataPos = [];
    for (var j = 0; j < jenisList.length; j++) {
      var jn = jenisList[j];
      if (jn.pctDesa <= 0 && jn.pctDaerah <= 0) continue;

      var totalMasuk = totalMasukPerJenis[jn.kode] || 0;
      var jatahDesa = Math.round(totalMasuk * jn.pctDesa / 100);
      var jatahDaerah = Math.round(totalMasuk * jn.pctDaerah / 100);

      // Target musyawaroh (kewajiban) — cocokkan nama jenis ke master musyawaroh
      var tarif = 0;
      var namaLower = jn.nama.toLowerCase().trim();
      if (musyawarahMap[namaLower] !== undefined) {
        tarif = musyawarahMap[namaLower];
      } else {
        for (var mk in musyawarahMap) {
          if (namaLower.indexOf(mk) !== -1 || mk.indexOf(namaLower) !== -1) { tarif = musyawarahMap[mk]; break; }
        }
      }
      var targetMusyawaroh = tarif * jumlahJamaahAktif * jumlahBulan;

      // Kewajiban setor ke desa = target musyawaroh jika ada, jika tidak pakai jatah desa
      var kewajiban = targetMusyawaroh > 0 ? targetMusyawaroh : jatahDesa;
      // Kekurangan dari pemasukan yang harus ditambah dari kas
      var topUpKas = Math.max(0, kewajiban - jatahDesa);

      // Realisasi: jika belum pernah diinput manual, sarankan = kewajiban (bisa diubah)
      var sudahSetor = setoranTerisi[jn.kode] ? setoranMap[jn.kode] : 0;
      var sisa = kewajiban - sudahSetor;
      var persen = kewajiban > 0 ? Math.round((sudahSetor / kewajiban) * 100 * 10) / 10 : 0;

      var status, tindakLanjut;
      if (kewajiban === 0) {
        status = 'Belum Ada';
        tindakLanjut = 'Belum ada pemasukan / target untuk jenis ini.';
      } else if (sudahSetor >= kewajiban) {
        status = 'Lunas';
        tindakLanjut = 'Setoran ke desa sudah lunas.';
      } else if (sudahSetor > 0) {
        status = 'Sebagian';
        tindakLanjut = 'Sudah disetor ' + fmtRp(sudahSetor) + '. Sisa ' + fmtRp(sisa) +
          (topUpKas > 0 ? ' (termasuk ' + fmtRp(topUpKas) + ' tambahan dari kas)' : '') + '.';
      } else {
        status = 'Belum Setor';
        tindakLanjut = 'Wajib setor ' + fmtRp(kewajiban) +
          (topUpKas > 0 ? '. Pemasukan ' + fmtRp(jatahDesa) + ', kurang ' + fmtRp(topUpKas) + ' ditambah dari kas.' : '.');
      }

      totalJatahDesa += jatahDesa;
      totalJatahDaerah += jatahDaerah;
      totalSudahSetor += sudahSetor;
      totalTargetMsy += targetMusyawaroh;
      totalTopUp += topUpKas;

      dataPos.push({
        id: jn.kode,
        nama: jn.nama,
        kategori: jn.kategori,
        pctDesa: jn.pctDesa,
        pctDaerah: jn.pctDaerah,
        totalMasuk: totalMasuk,
        jatahDesa: jatahDesa,
        jatahDaerah: jatahDaerah,
        targetMusyawaroh: targetMusyawaroh,
        kewajiban: kewajiban,
        topUpKas: topUpKas,
        target: kewajiban,        // kompat frontend lama
        terkumpul: jatahDesa,
        sudahSetor: sudahSetor,
        sudahDiinput: !!setoranTerisi[jn.kode],
        sisa: sisa,
        persen: persen,
        status: status,
        tindakLanjut: tindakLanjut
      });
    }

    var result = {
      periode: periode,
      summary: {
        totalTarget: totalTargetMsy > 0 ? totalTargetMsy : totalJatahDesa,
        totalTerkumpul: totalJatahDesa,
        totalJatahDesa: totalJatahDesa,
        totalJatahDaerah: totalJatahDaerah,
        totalTargetMusyawaroh: totalTargetMsy,
        totalSudahSetor: totalSudahSetor,
        totalTopUpKas: totalTopUp,
        totalSisa: (totalTargetMsy > 0 ? totalTargetMsy : totalJatahDesa) - totalSudahSetor,
        jumlahJamaahAktif: jumlahJamaahAktif,
        jumlahBulan: jumlahBulan
      },
      data: dataPos
    };

    try { cache.put(cacheKey, JSON.stringify(result), 90); } catch(e) {}
    result.success = true;
    return result;
  } catch(e) {
    return { success: false, message: e.message };
  }
}

// ──────────────────────────────────────────────────────
// PENEROBOS — Tagihan Infak per Jamaah
// ──────────────────────────────────────────────────────

// Ringkasan semua pos pemasukan (non-BukuIR) dan status bayar jamaah.
// Digunakan oleh halaman dashboard Penerobos.
function getTagihanPenerobos() {
  try {
    var auth = checkAuth();
    if (!auth.success) return { success: false, message: auth.message };

    var ss = getSS_();
    var periode = getPeriodeAktif();
    if (!periode) return { success: false, message: 'Tidak ada periode aktif' };
    var periodeId = periode.id;

    // Ambil semua jenis pemasukan aktif
    var sheetMaster = ss.getSheetByName(CONFIG.SHEETS.PEMASUKAN);
    var jenisAktif = [];
    if (sheetMaster && sheetMaster.getLastRow() > 1) {
      var mRows = sheetMaster.getDataRange().getValues();
      var mH = headerMap_(mRows[0]);
      for (var i = 1; i < mRows.length; i++) {
        if (!hGet_(mRows[i], mH, 'id', 0) && !hGet_(mRows[i], mH, 'kode', 0)) continue;
        var id = String(hGet_(mRows[i], mH, 'kode', 0) || hGet_(mRows[i], mH, 'id', 0) || '');
        var nama = String(hGet_(mRows[i], mH, 'namapemasukan', 1) || hGet_(mRows[i], mH, 'nama', 1) || '');
        var kategori = String(hGet_(mRows[i], mH, 'kategori', 2) || '');
        var statusP = String(hGet_(mRows[i], mH, 'status', 7) || hGet_(mRows[i], mH, 'status', 3) || '').toLowerCase();
        if (statusP === 'nonaktif') continue;
        if (!id || !nama) continue;
        jenisAktif.push({ id: id, nama: nama, kategori: kategori });
      }
    }

    // Ambil semua anggota aktif
    var sheetAnggota = ss.getSheetByName(CONFIG.SHEETS.ANGGOTA);
    var anggotaList = [];
    if (sheetAnggota && sheetAnggota.getLastRow() > 1) {
      var angRows = sheetAnggota.getDataRange().getValues();
      var angH = headerMap_(angRows[0]);
      for (var i = 1; i < angRows.length; i++) {
        var angId = String(hGet_(angRows[i], angH, 'id', 0) || '');
        if (!angId) continue;
        var statusAng = String(hGet_(angRows[i], angH, 'status', 4) || '').toLowerCase();
        if (statusAng === 'nonaktif') continue;
        anggotaList.push({
          id: angId,
          nama: String(hGet_(angRows[i], angH, 'nama', 1) || ''),
          noTelp: String(hGet_(angRows[i], angH, 'notelp', 2) || '')
        });
      }
    }
    var totalAnggota = anggotaList.length;

    // Baca semua penerimaan periode ini — bangun map jenisId → Set(anggotaId) + nominal
    var bayarMap = {}; // jenisId → { [anggotaId]: nominal }
    var sheetPmsk = ss.getSheetByName(CONFIG.SHEETS.INPUT_PENERIMAAN);
    if (sheetPmsk && sheetPmsk.getLastRow() > 1) {
      var pRows = sheetPmsk.getDataRange().getValues();
      var pH = headerMap_(pRows[0]);
      for (var i = 1; i < pRows.length; i++) {
        if (!hGet_(pRows[i], pH, 'id', 0)) continue;
        if (String(hGet_(pRows[i], pH, 'periodeid', 1)) !== periodeId) continue;
        var jId = String(hGet_(pRows[i], pH, 'jenisid', 2) || '');
        var aId = String(hGet_(pRows[i], pH, 'anggotaid', 3) || '');
        var nom = Number(hGet_(pRows[i], pH, 'nominal', 5)) || 0;
        if (!jId) continue;
        if (!bayarMap[jId]) bayarMap[jId] = {};
        if (aId) bayarMap[jId][aId] = (bayarMap[jId][aId] || 0) + nom;
      }
    }

    // Per jenis: hitung belum bayar
    var summary = jenisAktif.map(function(j) {
      var paid = bayarMap[j.id] || {};
      var sudahBayar = [], belumBayar = [];
      anggotaList.forEach(function(a) {
        if (paid[a.id]) {
          sudahBayar.push({ id: a.id, nama: a.nama, noTelp: a.noTelp, jumlah: paid[a.id] });
        } else {
          belumBayar.push({ id: a.id, nama: a.nama, noTelp: a.noTelp });
        }
      });
      return {
        id: j.id, nama: j.nama, kategori: j.kategori,
        totalAnggota: totalAnggota,
        sudah: sudahBayar.length,
        belum: belumBayar.length,
        belumBayar: belumBayar,
        sudahBayar: sudahBayar
      };
    });

    return { success: true, periode: periode, summary: summary };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

// Rekap Buku IR yang belum diserahkan untuk periode aktif
// Dipakai tab "Buku IR" di dashboard Penerobos
function getBukuIRBelumSerah() {
  try {
    var auth = checkAuth();
    if (!auth.success) return { success: false, message: auth.message };
    var ss = getSS_();
    var periode = getPeriodeAktif();
    if (!periode) return { success: false, message: 'Tidak ada periode aktif' };
    var periodeId = periode.id;

    // ID jenis yang kategorinya Buku IR
    var bukuIRIds = [];
    var sheetMaster = ss.getSheetByName(CONFIG.SHEETS.PEMASUKAN);
    if (sheetMaster) {
      var mRows = sheetMaster.getDataRange().getValues();
      for (var i = 1; i < mRows.length; i++) {
        if (String(mRows[i][2]).toLowerCase() === 'buku ir') bukuIRIds.push(String(mRows[i][0]));
      }
    }

    // Siapa yang sudah punya rincian di BUKU_IR
    var sudahRincianSet = {};
    var sheetIR = ss.getSheetByName(CONFIG.SHEETS.BUKU_IR);
    if (sheetIR && sheetIR.getLastRow() > 1) {
      var irRows = sheetIR.getDataRange().getValues();
      var irH = headerMap_(irRows[0]);
      for (var i = 1; i < irRows.length; i++) {
        if (!irRows[i][0]) continue;
        if (periodeId && String(hGet_(irRows[i], irH, 'periodeid', 2)) !== periodeId) continue;
        var tId = String(hGet_(irRows[i], irH, 'transaksiid', 1) || '');
        if (tId) sudahRincianSet[tId] = true;
      }
    }

    // Penerimaan jenis Buku IR periode ini — pisah belum/sudah rincian
    var belumSerah = [], sudahSerah = [];
    var sheetPmsk = ss.getSheetByName(CONFIG.SHEETS.INPUT_PENERIMAAN);
    var anggotaMap = {};
    var sheetAng = ss.getSheetByName(CONFIG.SHEETS.ANGGOTA);
    if (sheetAng) {
      var aRows = sheetAng.getDataRange().getValues();
      for (var i = 1; i < aRows.length; i++) {
        if (aRows[i][0]) anggotaMap[String(aRows[i][0])] = { nama: String(aRows[i][1] || ''), noTelp: String(aRows[i][2] || '') };
      }
    }
    if (sheetPmsk && sheetPmsk.getLastRow() > 1) {
      var pRows = sheetPmsk.getDataRange().getValues();
      var pH = headerMap_(pRows[0]);
      for (var i = 1; i < pRows.length; i++) {
        if (!hGet_(pRows[i], pH, 'id', 0)) continue;
        if (String(hGet_(pRows[i], pH, 'periodeid', 1)) !== periodeId) continue;
        if (bukuIRIds.indexOf(String(hGet_(pRows[i], pH, 'jenisid', 2))) === -1) continue;
        var trxId = String(hGet_(pRows[i], pH, 'id', 0));
        var aId = String(hGet_(pRows[i], pH, 'anggotaid', 3) || '');
        var ang = anggotaMap[aId] || { nama: aId, noTelp: '' };
        var entry = {
          trxId: trxId, anggotaId: aId, nama: ang.nama, noTelp: ang.noTelp,
          nominal: Number(hGet_(pRows[i], pH, 'nominal', 5)) || 0,
          tanggal: toDateStr_(hGet_(pRows[i], pH, 'tanggal', 4))
        };
        if (sudahRincianSet[trxId]) sudahSerah.push(entry);
        else belumSerah.push(entry);
      }
    }
    belumSerah.sort(function(a,b){ return a.nama.localeCompare(b.nama); });
    sudahSerah.sort(function(a,b){ return a.nama.localeCompare(b.nama); });
    return { success: true, belumSerah: belumSerah, sudahSerah: sudahSerah };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

// ──────────────────────────────────────────────────────
// PEMBELAAN — Kesanggupan bayar dari jamaah
// ──────────────────────────────────────────────────────
function getPembelaanData() {
  try {
    var auth = checkAuth();
    if (!auth.success) return { success: false, message: auth.message };
    var ss = getSS_();
    var periode = getPeriodeAktif();
    if (!periode) return { success: false, message: 'Tidak ada periode aktif' };

    // Daftar anggota aktif untuk dropdown
    var anggotaList = [];
    var sheetAng = ss.getSheetByName(CONFIG.SHEETS.ANGGOTA);
    if (sheetAng && sheetAng.getLastRow() > 1) {
      var aRows = sheetAng.getDataRange().getValues();
      var aH = headerMap_(aRows[0]);
      for (var i = 1; i < aRows.length; i++) {
        var aid = String(hGet_(aRows[i], aH, 'id', 0) || '');
        if (!aid) continue;
        var st = String(hGet_(aRows[i], aH, 'status', 4) || '').toLowerCase();
        if (st === 'nonaktif') continue;
        anggotaList.push({ id: aid, nama: String(hGet_(aRows[i], aH, 'nama', 1) || ''), noTelp: String(hGet_(aRows[i], aH, 'notelp', 2) || '') });
      }
    }

    // Data pembelaan periode ini
    var records = [];
    var sheet = ss.getSheetByName(CONFIG.SHEETS.PEMBELAAN);
    if (sheet && sheet.getLastRow() > 1) {
      var rows = sheet.getDataRange().getValues();
      var h = headerMap_(rows[0]);
      for (var i = 1; i < rows.length; i++) {
        if (!hGet_(rows[i], h, 'id', 0)) continue;
        if (String(hGet_(rows[i], h, 'periodeid', 2)) !== String(periode.id)) continue;
        records.push({
          id: String(hGet_(rows[i], h, 'id', 0)),
          anggotaId: String(hGet_(rows[i], h, 'anggotaid', 1) || ''),
          anggotaNama: String(hGet_(rows[i], h, 'anggotnama', 3) || hGet_(rows[i], h, 'anggotnama', 3) || ''),
          nominal: Number(hGet_(rows[i], h, 'nominalsanggup', 4)) || 0,
          tanggalSanggup: toDateStr_(hGet_(rows[i], h, 'tanggalsanggup', 5)),
          tanggalJanji: toDateStr_(hGet_(rows[i], h, 'tanggaljanji', 6)),
          catatan: String(hGet_(rows[i], h, 'catatan', 7) || ''),
          status: String(hGet_(rows[i], h, 'status', 8) || 'Belum'),
          lunasAt: toDateStr_(hGet_(rows[i], h, 'lunasat', 9))
        });
      }
    }

    // Pisah belum/lunas, urutkan berdasar tanggal janji terdekat
    var belum = records.filter(function(r){ return r.status !== 'Lunas'; });
    var lunas = records.filter(function(r){ return r.status === 'Lunas'; });
    belum.sort(function(a,b){ return (a.tanggalJanji||'').localeCompare(b.tanggalJanji||''); });

    return { success: true, periode: periode, anggotaList: anggotaList, belum: belum, lunas: lunas };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function submitKesanggupanPembelaan(data) {
  try {
    var auth = checkAuth();
    if (!auth.success) return { success: false, message: auth.message };
    var ss = getSS_();
    var periode = getPeriodeAktif();
    if (!periode) return { success: false, message: 'Tidak ada periode aktif' };

    // Cari nama anggota
    var angNama = data.anggotaId;
    var sheetAng = ss.getSheetByName(CONFIG.SHEETS.ANGGOTA);
    if (sheetAng) {
      var aRows = sheetAng.getDataRange().getValues();
      for (var i = 1; i < aRows.length; i++) {
        if (String(aRows[i][0]) === String(data.anggotaId)) { angNama = String(aRows[i][1] || data.anggotaId); break; }
      }
    }

    var sheet = ss.getSheetByName(CONFIG.SHEETS.PEMBELAAN);
    if (!sheet) {
      sheet = ss.insertSheet(CONFIG.SHEETS.PEMBELAAN);
      sheet.appendRow(['ID','AnggotaID','PeriodeID','AnggotaNama','NominalSanggup','TanggalSanggup','TanggalJanji','Catatan','Status','LunasAt','CreatedBy','CreatedAt']);
    }
    var id = generateID('PBL');
    sheet.appendRow([id, data.anggotaId, periode.id, angNama,
      Number(data.nominal) || 0,
      toDateStr_(new Date()), toDateStr_(data.tanggalJanji),
      data.catatan || '', 'Belum', '', auth.user.email, toDateStr_(new Date())]);
    logActivity(auth.user.email, 'PEMBELAAN_INPUT', angNama + ' sanggup ' + data.nominal);
    return { success: true, id: id };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function updateStatusPembelaan(id, status) {
  try {
    var auth = checkAuth();
    if (!auth.success) return { success: false, message: auth.message };
    var ss = getSS_();
    var sheet = ss.getSheetByName(CONFIG.SHEETS.PEMBELAAN);
    if (!sheet) return { success: false, message: 'Sheet tidak ditemukan' };
    var rows = sheet.getDataRange().getValues();
    var h = headerMap_(rows[0]);
    var colStatus = (h['status'] !== undefined ? h['status'] : 8) + 1;
    var colLunas = (h['lunasat'] !== undefined ? h['lunasat'] : 9) + 1;
    for (var i = 1; i < rows.length; i++) {
      if (String(hGet_(rows[i], h, 'id', 0)) === String(id)) {
        sheet.getRange(i + 1, colStatus).setValue(status);
        if (status === 'Lunas') sheet.getRange(i + 1, colLunas).setValue(toDateStr_(new Date()));
        logActivity(auth.user.email, 'PEMBELAAN_UPDATE', 'ID: ' + id + ' → ' + status);
        return { success: true };
      }
    }
    return { success: false, message: 'Data tidak ditemukan' };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function getJamaahBelumBayar(jenisId) {
  try {
    var auth = checkAuth();
    if (!auth.success) return { success: false, message: auth.message };

    var ss = getSS_();
    var periode = getPeriodeAktif();
    if (!periode) return { success: false, message: 'Tidak ada periode aktif' };
    var periodeId = periode.id;

    // Cari nama jenis dari Master Pemasukan (untuk judul)
    var jenisNama = String(jenisId);
    var sheetMaster = ss.getSheetByName(CONFIG.SHEETS.PEMASUKAN);
    if (sheetMaster && sheetMaster.getLastRow() > 1) {
      var mRows = sheetMaster.getDataRange().getValues();
      var mH = headerMap_(mRows[0]);
      for (var i = 1; i < mRows.length; i++) {
        if (String(hGet_(mRows[i], mH, 'kode', 0)) === String(jenisId)) {
          jenisNama = String(hGet_(mRows[i], mH, 'namapemasukan', 1) || hGet_(mRows[i], mH, 'nama', 1) || jenisId);
          break;
        }
      }
    }

    // Ambil semua anggota aktif
    var sheetAnggota = ss.getSheetByName(CONFIG.SHEETS.ANGGOTA);
    var anggotaMap = {};
    if (sheetAnggota && sheetAnggota.getLastRow() > 1) {
      var angRows = sheetAnggota.getDataRange().getValues();
      var angHdr = headerMap_(angRows[0]);
      for (var i = 1; i < angRows.length; i++) {
        if (!hGet_(angRows[i], angHdr, 'id', 0)) continue;
        var statusAng = String(hGet_(angRows[i], angHdr, 'status', 4) || '').trim().toLowerCase();
        if (statusAng !== 'nonaktif') {
          var angId = String(hGet_(angRows[i], angHdr, 'id', 0));
          anggotaMap[angId] = {
            id: angId,
            nama: String(hGet_(angRows[i], angHdr, 'nama', 1) || ''),
            noTelp: String(hGet_(angRows[i], angHdr, 'notelp', 2) || '')
          };
        }
      }
    }

    // Cek siapa yang sudah membayar jenis ini di Input Penerimaan periode aktif
    var sudahBayarMap = {}; // anggotaId → total nominal
    var sheetPmsk = ss.getSheetByName(CONFIG.SHEETS.INPUT_PENERIMAAN);
    if (sheetPmsk && sheetPmsk.getLastRow() > 1) {
      var pRows = sheetPmsk.getDataRange().getValues();
      var pH = headerMap_(pRows[0]);
      for (var i = 1; i < pRows.length; i++) {
        if (!hGet_(pRows[i], pH, 'id', 0)) continue;
        if (String(hGet_(pRows[i], pH, 'periodeid', 1)) !== periodeId) continue;
        if (String(hGet_(pRows[i], pH, 'jenisid', 2)) !== String(jenisId)) continue;
        var aId = String(hGet_(pRows[i], pH, 'anggotaid', 3) || '');
        if (!aId) continue;
        var nom = Number(hGet_(pRows[i], pH, 'nominal', 5)) || 0;
        sudahBayarMap[aId] = (sudahBayarMap[aId] || 0) + nom;
      }
    }

    var belumBayar = [], sudahBayar = [];
    for (var aid in anggotaMap) {
      var ang = anggotaMap[aid];
      if (sudahBayarMap[aid]) {
        sudahBayar.push({ id: aid, nama: ang.nama, noTelp: ang.noTelp, jumlahBayar: sudahBayarMap[aid] });
      } else {
        belumBayar.push({ id: aid, nama: ang.nama, noTelp: ang.noTelp, targetBayar: 0 });
      }
    }

    belumBayar.sort(function(a, b) { return a.nama.localeCompare(b.nama); });
    sudahBayar.sort(function(a, b) { return a.nama.localeCompare(b.nama); });

    return {
      success: true,
      posNama: jenisNama,
      periode: periode,
      belumBayar: belumBayar,
      sudahBayar: sudahBayar,
      total: {
        belum: belumBayar.length,
        sudah: sudahBayar.length,
        totalAnggota: belumBayar.length + sudahBayar.length
      }
    };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

// ──────────────────────────────────────────────────────
// KAS PENEROBOS & SERAH TERIMA
// ──────────────────────────────────────────────────────
function submitKasPenerobos(data) {
  try {
    var auth = requirePerm('penerobos.input');
    if (!auth.success) return { success: false, message: auth.message };
    var ss = getSS_();
    var periode = getPeriodeAktif();
    if (!periode) return { success: false, message: 'Tidak ada periode aktif' };

    // Validasi FK jenis
    var sheetMaster = ss.getSheetByName(CONFIG.SHEETS.PEMASUKAN);
    if (sheetMaster) {
      var mrows = sheetMaster.getDataRange().getValues();
      var valid = false;
      for (var mi = 1; mi < mrows.length; mi++) {
        if (String(mrows[mi][0]) === String(data.jenisId)) { valid = true; break; }
      }
      if (!valid) return { success: false, message: 'Jenis tidak ditemukan di master data' };
    }

    var sheet = ss.getSheetByName(CONFIG.SHEETS.KAS_PENEROBOS);
    if (!sheet) {
      sheet = ss.insertSheet(CONFIG.SHEETS.KAS_PENEROBOS);
      sheet.appendRow(['ID', 'Periode ID', 'Tanggal', 'Jenis ID', 'Anggota ID', 'Nominal', 'Sumber Kas', 'Catatan', 'Penerobos Email', 'Status', 'Serah Terima ID', 'Created At']);
    }
    var id = generateID('KP');
    var tgl = toDateStr_(data.tanggal ? new Date(data.tanggal) : new Date());
    sheet.appendRow([id, periode.id, tgl, data.jenisId, data.anggotaId || '', Number(data.nominal) || 0, data.sumberKas || 'Tunai', data.catatan || '', auth.user.email, 'Aktif', '', toDateStr_(new Date())]);
    logActivity(auth.user.email, 'KAS_PENEROBOS', 'Nominal: ' + data.nominal);
    return { success: true, id: id };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function getKasPenerobos() {
  try {
    var auth = requirePerm('penerobos.input');
    if (!auth.success) return { success: false, message: auth.message };
    var ss = getSS_();
    var periode = getPeriodeAktif();
    var periodeId = periode ? periode.id : null;

    var sheet = ss.getSheetByName(CONFIG.SHEETS.KAS_PENEROBOS);
    if (!sheet || sheet.getLastRow() < 2) return { success: true, data: [], periode: periode };

    // Build jenis map
    var jenisMap = {};
    var sheetM = ss.getSheetByName(CONFIG.SHEETS.PEMASUKAN);
    if (sheetM) {
      var mr = sheetM.getDataRange().getValues();
      for (var i = 1; i < mr.length; i++) if (mr[i][0]) jenisMap[String(mr[i][0])] = String(mr[i][1] || '');
    }
    // Build anggota map
    var angMap = {};
    var sheetA = ss.getSheetByName(CONFIG.SHEETS.ANGGOTA);
    if (sheetA) {
      var ar = sheetA.getDataRange().getValues();
      for (var i = 1; i < ar.length; i++) if (ar[i][0]) angMap[String(ar[i][0])] = String(ar[i][1] || '');
    }

    var rows = sheet.getDataRange().getValues();
    var h = headerMap_(rows[0]);
    var isAdmin = auth.user.role === CONFIG.ROLES.ADMIN;
    var result = [];
    for (var i = 1; i < rows.length; i++) {
      if (!hGet_(rows[i], h, 'id', 0)) continue;
      if (periodeId && String(hGet_(rows[i], h, 'periodeid', 1)) !== periodeId) continue;
      var email = String(hGet_(rows[i], h, 'penerobosemail', 8) || '');
      if (!isAdmin && email !== auth.user.email) continue;
      var jenisId = String(hGet_(rows[i], h, 'jenisid', 3) || '');
      var anggotaId = String(hGet_(rows[i], h, 'anggotaid', 4) || '');
      result.push({
        id: String(hGet_(rows[i], h, 'id', 0)),
        periodeId: String(hGet_(rows[i], h, 'periodeid', 1) || ''),
        tanggal: toDateStr_(hGet_(rows[i], h, 'tanggal', 2)),
        jenisId: jenisId,
        jenis: jenisMap[jenisId] || jenisId,
        anggotaId: anggotaId,
        anggota: angMap[anggotaId] || '',
        nominal: Number(hGet_(rows[i], h, 'nominal', 5)) || 0,
        sumberKas: String(hGet_(rows[i], h, 'sumberkas', 6) || 'Tunai'),
        catatan: String(hGet_(rows[i], h, 'catatan', 7) || ''),
        penerobosEmail: email,
        status: String(hGet_(rows[i], h, 'status', 9) || 'Aktif'),
        serahTerimaId: String(hGet_(rows[i], h, 'serahterimaid', 10) || '')
      });
    }
    result.sort(function(a, b) { return b.tanggal.localeCompare(a.tanggal); });
    return { success: true, data: result, periode: periode };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function buatSerahTerima(data) {
  try {
    var auth = requirePerm('penerobos.input');
    if (!auth.success) return { success: false, message: auth.message };
    if (!data.itemIds || data.itemIds.length === 0) return { success: false, message: 'Pilih minimal 1 item kas' };
    if (!data.sumberTujuan) return { success: false, message: 'Tentukan tujuan (Tunai/Bank)' };

    var ss = getSS_();
    var sheet = ss.getSheetByName(CONFIG.SHEETS.KAS_PENEROBOS);
    if (!sheet) return { success: false, message: 'Sheet Kas Penerobos tidak ditemukan' };

    var rows = sheet.getDataRange().getValues();
    var h = headerMap_(rows[0]);
    var isAdmin = auth.user.role === CONFIG.ROLES.ADMIN;
    var colStatus = (h['status'] !== undefined ? h['status'] : 9) + 1;
    var colSerahId = (h['serahterimaid'] !== undefined ? h['serahterimaid'] : 10) + 1;
    var periodeId = null;
    var total = 0;
    var matchedRows = [];

    // Validasi semua item
    for (var ii = 0; ii < data.itemIds.length; ii++) {
      var targetId = String(data.itemIds[ii]);
      var found = false;
      for (var i = 1; i < rows.length; i++) {
        if (String(hGet_(rows[i], h, 'id', 0)) !== targetId) continue;
        var email = String(hGet_(rows[i], h, 'penerobosemail', 8) || '');
        if (!isAdmin && email !== auth.user.email) return { success: false, message: 'Item bukan milik Anda: ' + targetId };
        var status = String(hGet_(rows[i], h, 'status', 9) || '');
        if (status !== 'Aktif') return { success: false, message: 'Item sudah diserahkan: ' + targetId };
        var pid = String(hGet_(rows[i], h, 'periodeid', 1) || '');
        if (!periodeId) periodeId = pid;
        total += Number(hGet_(rows[i], h, 'nominal', 5)) || 0;
        matchedRows.push({ rowIdx: i + 1 });
        found = true;
        break;
      }
      if (!found) return { success: false, message: 'Item tidak ditemukan: ' + targetId };
    }

    // Buat Serah Terima
    var stSheet = ss.getSheetByName(CONFIG.SHEETS.SERAH_TERIMA);
    if (!stSheet) {
      stSheet = ss.insertSheet(CONFIG.SHEETS.SERAH_TERIMA);
      stSheet.appendRow(['ID', 'Periode ID', 'Tanggal Serah', 'Penerobos Email', 'Total', 'Sumber Tujuan', 'Catatan', 'Status', 'Dikonfirmasi By', 'Dikonfirmasi At', 'Penerimaan ID', 'Created At']);
    }
    var stId = generateID('ST');
    var now = toDateStr_(new Date());
    stSheet.appendRow([stId, periodeId, now, auth.user.email, total, data.sumberTujuan, data.catatan || '', 'Menunggu', '', '', '', now]);

    // Update tiap item
    for (var j = 0; j < matchedRows.length; j++) {
      sheet.getRange(matchedRows[j].rowIdx, colStatus).setValue('Diserahkan');
      sheet.getRange(matchedRows[j].rowIdx, colSerahId).setValue(stId);
    }

    logActivity(auth.user.email, 'BUAT_SERAH_TERIMA', 'ID: ' + stId + ' Total: ' + total);
    return { success: true, id: stId, total: total };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function getSerahTerimaList() {
  try {
    var auth = checkAuth();
    if (!auth.success) return { success: false, message: auth.message };
    var ss = getSS_();
    var periode = getPeriodeAktif();
    var periodeId = periode ? periode.id : null;

    var sheet = ss.getSheetByName(CONFIG.SHEETS.SERAH_TERIMA);
    if (!sheet || sheet.getLastRow() < 2) return { success: true, data: [], periode: periode };

    var rows = sheet.getDataRange().getValues();
    var h = headerMap_(rows[0]);
    var role = auth.user.role;
    var isB1 = role === CONFIG.ROLES.BENDAHARA_1 || role === CONFIG.ROLES.ADMIN;
    var isB2 = role === CONFIG.ROLES.BENDAHARA_2 || role === CONFIG.ROLES.ADMIN;
    var isPenerobos = role === CONFIG.ROLES.PENEROBOS;
    var result = [];

    for (var i = 1; i < rows.length; i++) {
      if (!hGet_(rows[i], h, 'id', 0)) continue;
      if (periodeId && String(hGet_(rows[i], h, 'periodeid', 1)) !== periodeId) continue;
      var sumberTujuan = String(hGet_(rows[i], h, 'sumbertujuan', 5) || '');
      var email = String(hGet_(rows[i], h, 'penerobosemail', 3) || '');
      // Filter: Penerobos hanya lihat milik sendiri; B1 lihat Tunai; B2 lihat Bank; Admin lihat semua
      if (isPenerobos && email !== auth.user.email) continue;
      if (!isPenerobos && role !== CONFIG.ROLES.ADMIN) {
        if (role === CONFIG.ROLES.BENDAHARA_1 && sumberTujuan !== 'Tunai') continue;
        if (role === CONFIG.ROLES.BENDAHARA_2 && sumberTujuan !== 'Bank') continue;
      }
      result.push({
        id: String(hGet_(rows[i], h, 'id', 0)),
        periodeId: String(hGet_(rows[i], h, 'periodeid', 1) || ''),
        tanggalSerah: toDateStr_(hGet_(rows[i], h, 'tanggalserah', 2)),
        penerobosEmail: email,
        total: Number(hGet_(rows[i], h, 'total', 4)) || 0,
        sumberTujuan: sumberTujuan,
        catatan: String(hGet_(rows[i], h, 'catatan', 6) || ''),
        status: String(hGet_(rows[i], h, 'status', 7) || 'Menunggu'),
        dikonfirmasiBy: String(hGet_(rows[i], h, 'dikonfirmasiby', 8) || ''),
        dikonfirmasiAt: toDateStr_(hGet_(rows[i], h, 'dikonfirmasiat', 9)),
        penerimaanId: String(hGet_(rows[i], h, 'penerimaanid', 10) || '')
      });
    }
    result.sort(function(a, b) { return b.tanggalSerah.localeCompare(a.tanggalSerah); });
    return { success: true, data: result, periode: periode };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function konfirmasiSerahTerima(serahTerimaId) {
  try {
    var auth = requirePerm('serahterima.konfirmasi');
    if (!auth.success) return { success: false, message: auth.message };
    var ss = getSS_();
    var periode = getPeriodeAktif();
    if (!periode) return { success: false, message: 'Tidak ada periode aktif' };

    var stSheet = ss.getSheetByName(CONFIG.SHEETS.SERAH_TERIMA);
    if (!stSheet) return { success: false, message: 'Sheet Serah Terima tidak ditemukan' };

    var rows = stSheet.getDataRange().getValues();
    var h = headerMap_(rows[0]);
    var stRow = -1;
    var stData = null;
    for (var i = 1; i < rows.length; i++) {
      if (String(hGet_(rows[i], h, 'id', 0)) === serahTerimaId) {
        stRow = i + 1;
        stData = {
          status: String(hGet_(rows[i], h, 'status', 7) || ''),
          total: Number(hGet_(rows[i], h, 'total', 4)) || 0,
          sumberTujuan: String(hGet_(rows[i], h, 'sumbertujuan', 5) || ''),
          penerobosEmail: String(hGet_(rows[i], h, 'penerobosemail', 3) || ''),
          periodeId: String(hGet_(rows[i], h, 'periodeid', 1) || '')
        };
        break;
      }
    }
    if (!stData) return { success: false, message: 'Serah Terima tidak ditemukan' };
    if (stData.status !== 'Menunggu') return { success: false, message: 'Serah Terima sudah dikonfirmasi' };

    // Pastikan Bendahara konfirmasi sesuai tujuan
    var role = auth.user.role;
    if (role === CONFIG.ROLES.BENDAHARA_1 && stData.sumberTujuan !== 'Tunai')
      return { success: false, message: 'Anda hanya bisa konfirmasi serah terima Tunai' };
    if (role === CONFIG.ROLES.BENDAHARA_2 && stData.sumberTujuan !== 'Bank')
      return { success: false, message: 'Anda hanya bisa konfirmasi serah terima Bank' };

    // Ambil semua item KAS_PENEROBOS untuk serah terima ini
    var kpSheet = ss.getSheetByName(CONFIG.SHEETS.KAS_PENEROBOS);
    if (!kpSheet) return { success: false, message: 'Sheet Kas Penerobos tidak ditemukan' };
    var kpRows = kpSheet.getDataRange().getValues();
    var kpH = headerMap_(kpRows[0]);
    var items = [];
    for (var ki = 1; ki < kpRows.length; ki++) {
      if (String(hGet_(kpRows[ki], kpH, 'serahterimaid', 10)) === serahTerimaId) {
        items.push({
          jenisId:  String(hGet_(kpRows[ki], kpH, 'jenisid', 3) || ''),
          anggotaId: String(hGet_(kpRows[ki], kpH, 'anggotaid', 4) || ''),
          nominal:  Number(hGet_(kpRows[ki], kpH, 'nominal', 5)) || 0,
          catatan:  String(hGet_(kpRows[ki], kpH, 'catatan', 7) || '')
        });
      }
    }
    if (items.length === 0) return { success: false, message: 'Tidak ada item kas penerobos untuk serah terima ini' };

    // Buat INPUT_PENERIMAAN per item
    var sheetP = ss.getSheetByName(CONFIG.SHEETS.INPUT_PENERIMAAN);
    if (!sheetP) return { success: false, message: 'Sheet penerimaan tidak ditemukan' };
    var now = toDateStr_(new Date());
    var penIds = [];
    for (var ji = 0; ji < items.length; ji++) {
      var penId = generateID('TRX');
      var catatanEntry = 'Serah Terima dari: ' + stData.penerobosEmail + (items[ji].catatan ? ' | ' + items[ji].catatan : '');
      sheetP.appendRow([penId, stData.periodeId, items[ji].jenisId, items[ji].anggotaId,
        now, items[ji].nominal, stData.sumberTujuan, catatanEntry, auth.user.email, now]);
      penIds.push(penId);
    }

    // Update Serah Terima
    var colStatus = (h['status'] !== undefined ? h['status'] : 7) + 1;
    var colKonfBy = (h['dikonfirmasiby'] !== undefined ? h['dikonfirmasiby'] : 8) + 1;
    var colKonfAt = (h['dikonfirmasiat'] !== undefined ? h['dikonfirmasiat'] : 9) + 1;
    var colPenId  = (h['penerimaanid'] !== undefined ? h['penerimaanid'] : 10) + 1;
    stSheet.getRange(stRow, colStatus).setValue('Dikonfirmasi');
    stSheet.getRange(stRow, colKonfBy).setValue(auth.user.email);
    stSheet.getRange(stRow, colKonfAt).setValue(now);
    stSheet.getRange(stRow, colPenId).setValue(penIds.join(','));

    try { CacheService.getScriptCache().remove('dashboard_saldo'); } catch(e) {}
    logActivity(auth.user.email, 'KONFIRMASI_SERAH_TERIMA', 'ID: ' + serahTerimaId + ' Items: ' + items.length + ' Nominal: ' + stData.total);
    return { success: true, penerimaanIds: penIds, count: items.length };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function getSerahTerimaDetail(serahTerimaId) {
  try {
    var auth = checkAuth();
    if (!auth.success) return { success: false, message: auth.message };
    var ss = getSS_();

    var stSheet = ss.getSheetByName(CONFIG.SHEETS.SERAH_TERIMA);
    if (!stSheet) return { success: false, message: 'Sheet tidak ditemukan' };
    var rows = stSheet.getDataRange().getValues();
    var h = headerMap_(rows[0]);
    var stData = null;
    for (var i = 1; i < rows.length; i++) {
      if (String(hGet_(rows[i], h, 'id', 0)) === serahTerimaId) {
        stData = {
          id: serahTerimaId,
          periodeId: String(hGet_(rows[i], h, 'periodeid', 1) || ''),
          tanggalSerah: toDateStr_(hGet_(rows[i], h, 'tanggalserah', 2)),
          penerobosEmail: String(hGet_(rows[i], h, 'penerobosemail', 3) || ''),
          total: Number(hGet_(rows[i], h, 'total', 4)) || 0,
          sumberTujuan: String(hGet_(rows[i], h, 'sumbertujuan', 5) || ''),
          catatan: String(hGet_(rows[i], h, 'catatan', 6) || ''),
          status: String(hGet_(rows[i], h, 'status', 7) || ''),
          dikonfirmasiBy: String(hGet_(rows[i], h, 'dikonfirmasiby', 8) || ''),
          dikonfirmasiAt: toDateStr_(hGet_(rows[i], h, 'dikonfirmasiat', 9))
        };
        break;
      }
    }
    if (!stData) return { success: false, message: 'Serah Terima tidak ditemukan' };

    // Ambil items KAS_PENEROBOS
    var kpSheet = ss.getSheetByName(CONFIG.SHEETS.KAS_PENEROBOS);
    var items = [];
    if (kpSheet) {
      var kpRows = kpSheet.getDataRange().getValues();
      var kpH = headerMap_(kpRows[0]);
      // Build jenis name map
      var jenisMap = {};
      var masterSheet = ss.getSheetByName(CONFIG.SHEETS.PEMASUKAN);
      if (masterSheet) {
        var mr = masterSheet.getDataRange().getValues();
        for (var mi = 1; mi < mr.length; mi++) jenisMap[String(mr[mi][0])] = String(mr[mi][1] || '');
      }
      for (var ki = 1; ki < kpRows.length; ki++) {
        if (String(hGet_(kpRows[ki], kpH, 'serahterimaid', 10)) === serahTerimaId) {
          var jenisId = String(hGet_(kpRows[ki], kpH, 'jenisid', 3) || '');
          items.push({
            jenisId: jenisId,
            jenisNama: jenisMap[jenisId] || jenisId,
            anggotaId: String(hGet_(kpRows[ki], kpH, 'anggotaid', 4) || ''),
            nominal: Number(hGet_(kpRows[ki], kpH, 'nominal', 5)) || 0,
            catatan: String(hGet_(kpRows[ki], kpH, 'catatan', 7) || '')
          });
        }
      }
    }

    // Info kelompok dari sheet Config jika ada
    var namaKelompok = 'Kas Kelompok';
    try {
      var cfgSheet = ss.getSheetByName('Config');
      if (cfgSheet) {
        var cfgRows = cfgSheet.getDataRange().getValues();
        for (var ci = 0; ci < cfgRows.length; ci++) {
          if (String(cfgRows[ci][0]).toLowerCase() === 'namakelompok') { namaKelompok = String(cfgRows[ci][1] || namaKelompok); break; }
        }
      }
    } catch(e) {}

    return { success: true, data: stData, items: items, namaKelompok: namaKelompok };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function getSerahTerimaForPDF_(ss, periodeId) {
  try {
    var stSheet = ss.getSheetByName(CONFIG.SHEETS.SERAH_TERIMA);
    if (!stSheet) return [];
    var rows = stSheet.getDataRange().getValues();
    var h = headerMap_(rows[0]);
    var kpSheet = ss.getSheetByName(CONFIG.SHEETS.KAS_PENEROBOS);
    var kpRows = kpSheet ? kpSheet.getDataRange().getValues() : [[]];
    var kpH = headerMap_(kpRows[0] || []);
    var jenisMap = {};
    var masterSheet = ss.getSheetByName(CONFIG.SHEETS.PEMASUKAN);
    if (masterSheet) {
      var mr = masterSheet.getDataRange().getValues();
      for (var mi = 1; mi < mr.length; mi++) jenisMap[String(mr[mi][0])] = String(mr[mi][1] || '');
    }
    var result = [];
    for (var i = 1; i < rows.length; i++) {
      if (String(hGet_(rows[i], h, 'periodeid', 1)) !== periodeId) continue;
      if (String(hGet_(rows[i], h, 'status', 7)) !== 'Dikonfirmasi') continue;
      var stId = String(hGet_(rows[i], h, 'id', 0));
      var items = [];
      for (var ki = 1; ki < kpRows.length; ki++) {
        if (String(hGet_(kpRows[ki], kpH, 'serahterimaid', 10)) !== stId) continue;
        var jId = String(hGet_(kpRows[ki], kpH, 'jenisid', 3) || '');
        items.push({ jenisNama: jenisMap[jId] || jId, nominal: Number(hGet_(kpRows[ki], kpH, 'nominal', 5)) || 0 });
      }
      result.push({
        tanggal: toDateStr_(hGet_(rows[i], h, 'tanggalserah', 2)),
        penerobosEmail: String(hGet_(rows[i], h, 'penerobosemail', 3) || ''),
        total: Number(hGet_(rows[i], h, 'total', 4)) || 0,
        sumberTujuan: String(hGet_(rows[i], h, 'sumbertujuan', 5) || ''),
        dikonfirmasiBy: String(hGet_(rows[i], h, 'dikonfirmasiby', 8) || ''),
        items: items
      });
    }
    return result;
  } catch(e) { return []; }
}

function migrasiTransaksiPenerobos() {
  try {
    var auth = requirePerm('master.manage');
    if (!auth.success) return { success: false, message: auth.message };
    var ss = getSS_();

    // Kumpulkan email semua user PENEROBOS
    var userList = getUserList_();
    var penerobosEmails = {};
    for (var u = 0; u < userList.length; u++) {
      if (userList[u].role === CONFIG.ROLES.PENEROBOS) penerobosEmails[userList[u].email] = true;
    }
    if (Object.keys(penerobosEmails).length === 0) return { success: true, count: 0, message: 'Tidak ada user PENEROBOS' };

    var sheetP = ss.getSheetByName(CONFIG.SHEETS.INPUT_PENERIMAAN);
    if (!sheetP || sheetP.getLastRow() < 2) return { success: true, count: 0 };

    var kpSheet = ss.getSheetByName(CONFIG.SHEETS.KAS_PENEROBOS);
    if (!kpSheet) {
      kpSheet = ss.insertSheet(CONFIG.SHEETS.KAS_PENEROBOS);
      kpSheet.appendRow(['ID', 'Periode ID', 'Tanggal', 'Jenis ID', 'Anggota ID', 'Nominal', 'Sumber Kas', 'Catatan', 'Penerobos Email', 'Status', 'Serah Terima ID', 'Created At']);
    }

    var rows = sheetP.getDataRange().getValues();
    var h = headerMap_(rows[0]);
    var toDelete = [];
    var count = 0;
    var now = toDateStr_(new Date());

    for (var i = rows.length - 1; i >= 1; i--) {
      var createdBy = String(hGet_(rows[i], h, 'createdby', 8) || '');
      if (!penerobosEmails[createdBy]) continue;
      var rowId = String(hGet_(rows[i], h, 'id', 0) || '');
      if (!rowId) continue;
      var periodeId2 = String(hGet_(rows[i], h, 'periodeid', 1) || '');
      var tgl2 = toDateStr_(hGet_(rows[i], h, 'tanggal', 4));
      var jenisId2 = String(hGet_(rows[i], h, 'jenisid', 2) || '');
      var anggotaId2 = String(hGet_(rows[i], h, 'anggotaid', 3) || '');
      var nominal2 = Number(hGet_(rows[i], h, 'nominal', 5)) || 0;
      var sumberKas2 = String(hGet_(rows[i], h, 'sumberkas', 6) || 'Tunai');
      var catatan2 = String(hGet_(rows[i], h, 'catatan', 7) || '');
      // Buat ID baru agar tidak bentrok
      var kpId = 'KP_MIG_' + rowId;
      kpSheet.appendRow([kpId, periodeId2, tgl2, jenisId2, anggotaId2, nominal2, sumberKas2, catatan2, createdBy, 'Aktif', '', now]);
      toDelete.push(i + 1);
      count++;
    }

    // Hapus dari INPUT_PENERIMAAN (dari bawah agar index tidak geser)
    for (var d = 0; d < toDelete.length; d++) {
      sheetP.deleteRow(toDelete[d]);
    }

    try { CacheService.getScriptCache().remove('dashboard_saldo'); CacheService.getScriptCache().remove('master_trx_data'); } catch(e) {}
    logActivity(auth.user.email, 'MIGRASI_KAS_PENEROBOS', 'Count: ' + count);
    return { success: true, count: count };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

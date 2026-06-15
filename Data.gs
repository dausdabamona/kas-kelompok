// ──────────────────────────────────────────────────────
// SPREADSHEET — buka sekali per eksekusi
// ──────────────────────────────────────────────────────
var _ss = null;
function getSS_() {
  if (!_ss) _ss = SpreadsheetApp.openById(getSpreadsheetId());
  return _ss;
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
      else if (h.indexOf('nama') !== -1 || h.indexOf('periode') !== -1) { if (colNama < 0) colNama = c; }
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

    return {
      success: true,
      user: auth.user,
      periode: periode,
      namaKelompok: saldoData.namaKelompok,
      kasTunai: saldoData.tunai,
      kasBank: saldoData.bank,
      totalKas: saldoData.tunai + saldoData.bank
    };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function calculateSaldo(periodeId, periode) {
  var ss = getSS_();
  // Mulai dari saldo awal periode (jika ada)
  var tunai = (periode && periode.saldoAwalTunai) ? Number(periode.saldoAwalTunai) : 0;
  var bank  = (periode && periode.saldoAwalBank)  ? Number(periode.saldoAwalBank)  : 0;

  var sheetP = ss.getSheetByName(CONFIG.SHEETS.INPUT_PENERIMAAN);
  if (sheetP && sheetP.getLastRow() > 1) {
    var dp = sheetP.getDataRange().getValues();
    for (var i = 1; i < dp.length; i++) {
      if (!periodeId || dp[i][1] === periodeId) {
        var nominal = Number(dp[i][5]) || 0;
        var sumber = dp[i][6];
        if (sumber === 'Tunai') tunai += nominal;
        else if (sumber === 'Bank') bank += nominal;
      }
    }
  }

  var sheetPK = ss.getSheetByName(CONFIG.SHEETS.INPUT_PENGELUARAN);
  if (sheetPK && sheetPK.getLastRow() > 1) {
    var dpk = sheetPK.getDataRange().getValues();
    for (var i = 1; i < dpk.length; i++) {
      if (!periodeId || dpk[i][1] === periodeId) {
        var nominal = Number(dpk[i][4]) || 0;
        var sumber = dpk[i][5];
        if (sumber === 'Tunai') tunai -= nominal;
        else if (sumber === 'Bank') bank -= nominal;
      }
    }
  }

  var sheetS = ss.getSheetByName(CONFIG.SHEETS.INPUT_SETORAN);
  if (sheetS && sheetS.getLastRow() > 1) {
    var ds = sheetS.getDataRange().getValues();
    for (var i = 1; i < ds.length; i++) {
      if (!periodeId || ds[i][1] === periodeId) {
        var nominal = Number(ds[i][3]) || 0;
        var arah = ds[i][4];
        if (arah === 'setor') { tunai -= nominal; bank += nominal; }
        else if (arah === 'tarik') { bank -= nominal; tunai += nominal; }
      }
    }
  }

  return { tunai: tunai, bank: bank };
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
    var tgl = data.tanggal ? new Date(data.tanggal) : now;

    if (data.tipe === 'masuk') {
      var sheet = ss.getSheetByName(CONFIG.SHEETS.INPUT_PENERIMAAN);
      if (!sheet) return { success: false, message: 'Sheet penerimaan tidak ditemukan' };
      sheet.appendRow([id, periode.id, data.jenisId, data.anggotaId || '', tgl, Number(data.nominal), data.sumberKas, data.catatan || '', auth.user.email, now]);
      logActivity(auth.user.email, 'PEMASUKAN', 'Nominal: ' + data.nominal);
    } else if (data.tipe === 'keluar') {
      var sheet = ss.getSheetByName(CONFIG.SHEETS.INPUT_PENGELUARAN);
      if (!sheet) return { success: false, message: 'Sheet pengeluaran tidak ditemukan' };
      sheet.appendRow([id, periode.id, data.jenisId, tgl, Number(data.nominal), data.sumberKas, data.catatan || '', auth.user.email, now]);
      logActivity(auth.user.email, 'PENGELUARAN', 'Nominal: ' + data.nominal);
    } else if (data.tipe === 'mutasi') {
      var sheet = ss.getSheetByName(CONFIG.SHEETS.INPUT_SETORAN);
      if (!sheet) {
        sheet = ss.insertSheet(CONFIG.SHEETS.INPUT_SETORAN);
        sheet.appendRow(['ID', 'PeriodeID', 'Tanggal', 'Nominal', 'Arah', 'CreatedBy', 'CreatedAt']);
      }
      sheet.appendRow([id, periode.id, tgl, Number(data.nominal), data.arah, auth.user.email, now]);
      logActivity(auth.user.email, 'MUTASI', 'Arah: ' + data.arah + ' Nominal: ' + data.nominal);
    }

    // Invalidate saldo cache setiap ada transaksi baru
    try { CacheService.getScriptCache().remove('dashboard_saldo'); } catch(e) {}
    return { success: true, id: id };
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
    var auth = checkAuth([CONFIG.ROLES.ADMIN, CONFIG.ROLES.BENDAHARA_1]);
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
    var auth = checkAuth([CONFIG.ROLES.ADMIN, CONFIG.ROLES.BENDAHARA_1]);
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
    var auth = checkAuth([CONFIG.ROLES.ADMIN]);
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
// BUKU IR
// ──────────────────────────────────────────────────────
function getBukuIRData() {
  try {
    var auth = checkAuth();
    if (!auth.success) return { success: false, message: auth.message };
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

    var dirincikanIds = {};
    if (sheetIR && sheetIR.getLastRow() > 1) {
      var irRows = sheetIR.getDataRange().getValues();
      for (var i = 1; i < irRows.length; i++) {
        if (irRows[i][1]) dirincikanIds[irRows[i][1]] = true;
      }
    }

    var belum = [], sudah = [];
    if (sheetP.getLastRow() > 1) {
      var pRows = sheetP.getDataRange().getValues();
      for (var i = 1; i < pRows.length; i++) {
        var row = pRows[i];
        if (!row[0]) continue;
        if (periodeId && row[1] !== periodeId) continue;
        if (bukuIRIds.indexOf(row[2]) === -1) continue;
        var item = {
          id: row[0], periodeId: row[1], jenisId: row[2], anggotaId: row[3],
          tanggal: row[4], nominal: Number(row[5]) || 0, sumberKas: row[6], catatan: row[7],
          anggota: anggotaMap[row[3]] || null
        };
        if (dirincikanIds[row[0]]) sudah.push(item);
        else belum.push(item);
      }
    }

    try { CacheService.getScriptCache().remove('master_trx_data'); } catch(e) {}
    return { success: true, data: { belumDirincikan: belum, sudahDirincikan: sudah } };
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
      sheet.appendRow(['ID', 'TransaksiID', 'PeriodeID', 'AnggotaID', 'Tanggal', 'IR', 'IR10', 'Cicilan', 'InfakDaerah', 'Index', 'Total', 'CreatedBy', 'CreatedAt']);
    }
    var id = generateID('IR');
    var now = new Date();
    sheet.appendRow([id, data.transaksiId, data.periodeId, data.anggotaId, data.tanggal,
      Number(data.ir) || 0, Number(data.ir10) || 0, Number(data.cicilan) || 0,
      Number(data.infakDaerah) || 0, Number(data.index) || 0, Number(data.total) || 0,
      auth.user.email, now]);
    try { CacheService.getScriptCache().remove('master_trx_data'); } catch(e) {}
    return { success: true, id: id };
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

    if (sheetP && sheetP.getLastRow() > 1) {
      var dp = sheetP.getDataRange().getValues();
      for (var i = 1; i < dp.length; i++) {
        if (!dp[i][0]) continue;
        if (periodeId && dp[i][1] !== periodeId) continue;
        var nominal = Number(dp[i][5]) || 0;
        var sumber = dp[i][6];
        pemasukan.push({ id: dp[i][0], jenis: masterPMap[dp[i][2]] || dp[i][2], tanggal: dp[i][4], nominal: nominal, sumber: sumber, catatan: dp[i][7] });
        if (sumber === 'Tunai') totalPTunai += nominal;
        else if (sumber === 'Bank') totalPBank += nominal;
      }
    }

    if (sheetPK && sheetPK.getLastRow() > 1) {
      var dpk = sheetPK.getDataRange().getValues();
      for (var i = 1; i < dpk.length; i++) {
        if (!dpk[i][0]) continue;
        if (periodeId && dpk[i][1] !== periodeId) continue;
        var nominal = Number(dpk[i][4]) || 0;
        var sumber = dpk[i][5];
        pengeluaran.push({ id: dpk[i][0], jenis: masterPKMap[dpk[i][2]] || dpk[i][2], tanggal: dpk[i][3], nominal: nominal, sumber: sumber, catatan: dpk[i][6] });
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
    var auth = checkAuth([CONFIG.ROLES.ADMIN]);
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
    var auth = checkAuth([CONFIG.ROLES.ADMIN]);
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
    var auth = checkAuth([CONFIG.ROLES.ADMIN]);
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
    var auth = checkAuth([CONFIG.ROLES.ADMIN]);
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
    var auth = checkAuth([CONFIG.ROLES.ADMIN]);
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
    var auth = checkAuth([CONFIG.ROLES.ADMIN]);
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
    if (!sheet) return { success: true, data: [] };
    var rows = sheet.getDataRange().getValues();
    var result = [];
    for (var i = 1; i < rows.length; i++) {
      if (rows[i][0]) result.push({ id: rows[i][0], nama: rows[i][1], tipe: rows[i][2], formula: rows[i][3], status: rows[i][4] });
    }
    try { CacheService.getScriptCache().remove('master_trx_data'); } catch(e) {}
    return { success: true, data: result };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function addPosSetoran(data) {
  try {
    var auth = checkAuth([CONFIG.ROLES.ADMIN]);
    if (!auth.success) return { success: false, message: auth.message };
    var ss = getSS_();
    var sheet = ss.getSheetByName(CONFIG.SHEETS.POS_SETORAN);
    if (!sheet) {
      sheet = ss.insertSheet(CONFIG.SHEETS.POS_SETORAN);
      sheet.appendRow(['ID', 'Nama', 'Tipe', 'Formula', 'Status', 'Target']);
    }
    var id = generateID('PST');
    sheet.appendRow([id, data.nama, data.tipe || 'manual', data.formula || '', data.status || 'Aktif', Number(data.target) || 0]);
    try { CacheService.getScriptCache().remove('master_trx_data'); } catch(e) {}
    return { success: true, id: id };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function updatePosSetoran(data) {
  try {
    var auth = checkAuth([CONFIG.ROLES.ADMIN]);
    if (!auth.success) return { success: false, message: auth.message };
    var ss = getSS_();
    var sheet = ss.getSheetByName(CONFIG.SHEETS.POS_SETORAN);
    if (!sheet) return { success: false, message: 'Sheet tidak ditemukan' };
    var rows = sheet.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      if (rows[i][0] === data.id) {
        sheet.getRange(i + 1, 2, 1, 5).setValues([[data.nama, data.tipe, data.formula || '', data.status, Number(data.target) || 0]]);
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
    var auth = checkAuth([CONFIG.ROLES.ADMIN]);
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
    var auth = checkAuth([CONFIG.ROLES.ADMIN]);
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
    var auth = checkAuth([CONFIG.ROLES.ADMIN]);
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
    var auth = checkAuth([CONFIG.ROLES.ADMIN]);
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
    var setoranMap = {};
    if (sheetSetoran && sheetSetoran.getLastRow() > 1) {
      var sRows = sheetSetoran.getDataRange().getValues();
      for (var i = 1; i < sRows.length; i++) {
        if (sRows[i][0]) setoranMap[sRows[i][1]] = {
          id: sRows[i][0], realisasi: Number(sRows[i][3]) || 0, catatan: sRows[i][4]
        };
      }
    }

    var result = [];
    for (var i = 1; i < posRows.length; i++) {
      if (!posRows[i][0] || posRows[i][4] !== 'Aktif') continue;
      var posId = posRows[i][0];
      var target = Number(posRows[i][5]) || 0;
      var realisasi = setoranMap[posId] ? setoranMap[posId].realisasi : 0;
      var sisa = target - realisasi;
      var pct = target > 0 ? Math.round((realisasi / target) * 100) : 0;
      result.push({
        id: posId, nama: posRows[i][1], tipe: posRows[i][2],
        target: target, realisasi: realisasi, sisa: sisa, persen: pct,
        status: pct >= 100 ? 'Lunas' : 'Belum Lunas',
        catatan: setoranMap[posId] ? setoranMap[posId].catatan : ''
      });
    }
    try { CacheService.getScriptCache().remove('master_trx_data'); } catch(e) {}
    return { success: true, data: result };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function submitRealisasiSetoran(data) {
  try {
    var auth = checkAuth([CONFIG.ROLES.ADMIN, CONFIG.ROLES.BENDAHARA_1, CONFIG.ROLES.BENDAHARA_2]);
    if (!auth.success) return { success: false, message: auth.message };
    var ss = getSS_();
    var sheet = ss.getSheetByName(CONFIG.SHEETS.SETORAN_DESA);
    if (!sheet) {
      sheet = ss.insertSheet(CONFIG.SHEETS.SETORAN_DESA);
      sheet.appendRow(['ID', 'PosID', 'Target', 'Realisasi', 'Catatan', 'UpdatedBy', 'UpdatedAt']);
    }
    var rows = sheet.getDataRange().getValues();
    var found = false;
    for (var i = 1; i < rows.length; i++) {
      if (rows[i][1] === data.posId) {
        sheet.getRange(i + 1, 3, 1, 5).setValues([[Number(data.target) || 0, Number(data.realisasi) || 0, data.catatan || '', auth.user.email, new Date()]]);
        found = true;
        break;
      }
    }
    if (!found) {
      var id = generateID('STR');
      sheet.appendRow([id, data.posId, Number(data.target) || 0, Number(data.realisasi) || 0, data.catatan || '', auth.user.email, new Date()]);
    }
    try {
      var cache2 = CacheService.getScriptCache();
      cache2.remove('master_trx_data');
      var periode2 = getPeriodeAktif();
      if (periode2) cache2.remove('laporan_setoran_' + periode2.id);
    } catch(e) {}
    return { success: true };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

// ──────────────────────────────────────────────────────
// PDF
// ──────────────────────────────────────────────────────
function generatePDF(periodeId) {
  try {
    var auth = checkAuth([CONFIG.ROLES.ADMIN, CONFIG.ROLES.BENDAHARA_1]);
    if (!auth.success) return { success: false, message: auth.message };
    var rekap = getRekapitulasiData();
    if (!rekap.success) return rekap;
    var html = buildPDFHTML(rekap);
    var blob = Utilities.newBlob(html, 'text/html', 'laporan.html');
    var folder = DriveApp.getRootFolder();
    var file = folder.createFile(blob);
    file.setName('Laporan Kas ' + (rekap.periode ? rekap.periode.nama : '') + '.html');
    try { CacheService.getScriptCache().remove('master_trx_data'); } catch(e) {}
    return { success: true, url: file.getUrl(), id: file.getId() };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function buildPDFHTML(data) {
  var html = '<html><head><meta charset="UTF-8"><style>body{font-family:Arial;font-size:12px;}table{border-collapse:collapse;width:100%;}td,th{border:1px solid #ccc;padding:4px 8px;}th{background:#eee;}</style></head><body>';
  html += '<h2>Laporan Kas Kelompok</h2>';
  if (data.periode) html += '<p>Periode: ' + data.periode.nama + '</p>';
  html += '<h3>Pemasukan</h3><table><tr><th>Jenis</th><th>Tanggal</th><th>Nominal</th><th>Sumber</th></tr>';
  (data.pemasukan || []).forEach(function(p) {
    html += '<tr><td>' + (p.jenis || '-') + '</td><td>' + fmtTanggal(p.tanggal) + '</td><td>' + fmtRp(p.nominal) + '</td><td>' + (p.sumber || '-') + '</td></tr>';
  });
  html += '</table><h3>Pengeluaran</h3><table><tr><th>Jenis</th><th>Tanggal</th><th>Nominal</th><th>Sumber</th></tr>';
  (data.pengeluaran || []).forEach(function(p) {
    html += '<tr><td>' + (p.jenis || '-') + '</td><td>' + fmtTanggal(p.tanggal) + '</td><td>' + fmtRp(p.nominal) + '</td><td>' + (p.sumber || '-') + '</td></tr>';
  });
  html += '</table>';
  html += '<p><strong>Saldo Tunai: ' + fmtRp(data.saldoTunai) + '</strong></p>';
  html += '<p><strong>Saldo Bank: ' + fmtRp(data.saldoBank) + '</strong></p>';
  html += '</body></html>';
  return html;
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
    var result = [];
    for (var i = 1; i < rows.length; i++) {
      if (!rows[i][0]) continue;
      if (periodeId && rows[i][1] !== periodeId) continue;
      result.push({
        id: rows[i][0], periodeId: rows[i][1], tanggal: rows[i][2],
        saldoAwal: Number(rows[i][3]) || 0, pemasukan: Number(rows[i][4]) || 0,
        pengeluaran: Number(rows[i][5]) || 0, saldoAkhirTeoritis: Number(rows[i][6]) || 0,
        saldoAkhirActual: Number(rows[i][7]) || 0, selisih: Number(rows[i][8]) || 0,
        status: rows[i][9], catatan: rows[i][10]
      });
    }
    try { CacheService.getScriptCache().remove('master_trx_data'); } catch(e) {}
    return { success: true, data: result };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function addBankTransaction(data) {
  try {
    var auth = checkAuth([CONFIG.ROLES.ADMIN, CONFIG.ROLES.BENDAHARA_1, CONFIG.ROLES.BENDAHARA_2]);
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
    sheet.appendRow([id, data.periodeId, data.tanggal,
      Number(data.saldoAwal) || 0, Number(data.pemasukan) || 0, Number(data.pengeluaran) || 0,
      saldoAkhirTeoritis, Number(data.saldoAkhirActual) || 0, selisih,
      selisih === 0 ? 'Balance' : 'Selisih', data.catatan || '', new Date()]);
    try { CacheService.getScriptCache().remove('master_trx_data'); } catch(e) {}
    return { success: true, id: id };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function updateBankDaily(data) {
  try {
    var auth = checkAuth([CONFIG.ROLES.ADMIN, CONFIG.ROLES.BENDAHARA_1]);
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
          selisih === 0 ? 'Balance' : 'Selisih', data.catatan || '', new Date()
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
    var result = [];
    for (var i = 1; i < rows.length; i++) {
      if (!rows[i][0]) continue;
      if (periodeId && rows[i][1] !== periodeId) continue;
      result.push({
        id: rows[i][0], periodeId: rows[i][1], tanggal: rows[i][2],
        keterangan: rows[i][3], nominal: Number(rows[i][4]) || 0,
        sumber: rows[i][5], status: rows[i][6],
        tanggalFound: rows[i][7], catatan: rows[i][8]
      });
    }
    try { CacheService.getScriptCache().remove('master_trx_data'); } catch(e) {}
    return { success: true, data: result };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function addManualBankIncome(data) {
  try {
    var auth = checkAuth([CONFIG.ROLES.ADMIN, CONFIG.ROLES.BENDAHARA_1]);
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
    var auth = checkAuth([CONFIG.ROLES.ADMIN, CONFIG.ROLES.BENDAHARA_1]);
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
    sheet.appendRow([id, periode.id, data.tanggal, data.keterangan, Number(data.nominal) || 0,
      data.sumber || 'Manual', 'Pending', '', data.catatan || '', new Date()]);
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
    var auth = checkAuth([CONFIG.ROLES.ADMIN, CONFIG.ROLES.BENDAHARA_1]);
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
    var auth = checkAuth([CONFIG.ROLES.ADMIN, CONFIG.ROLES.BENDAHARA_1]);
    if (!auth.success) return { success: false, message: auth.message };
    var ss = getSS_();
    var sheet = ss.getSheetByName(CONFIG.SHEETS.BANK_PENDING);
    if (!sheet) return { success: false, message: 'Sheet tidak ditemukan' };
    var rows = sheet.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      if (rows[i][0] === id) {
        sheet.getRange(i + 1, 7, 1, 4).setValues([[
          status,
          status === 'Found' ? new Date() : rows[i][7],
          catatan || rows[i][8],
          new Date()
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

function mapPosNamaToBukuIRCol(formula, nama) {
  // Kolom Detail Buku IR (0-indexed): [5]=IR, [6]=IR 1/10, [7]=Cicilan, [8]=InfakDaerah, [9]=Index
  if (formula) {
    var f = String(formula).toLowerCase().trim();
    if (f === 'ir' || f === 'col5' || f === '5') return 5;
    if (f.indexOf('1/10') !== -1 || f === 'ir10' || f === 'col6' || f === '6') return 6;
    if (f === 'cicilan' || f === 'col7' || f === '7') return 7;
    if (f.indexOf('infak') !== -1 || f === 'col8' || f === '8') return 8;
    if (f === 'index' || f === 'col9' || f === '9') return 9;
  }
  if (nama) {
    var n = String(nama).toLowerCase().trim();
    if (n === 'ir' && n.indexOf('1/10') === -1 && n.indexOf('10') === -1) return 5;
    if (n.indexOf('1/10') !== -1 || n === 'ir10' || n === 'ir 10') return 6;
    if (n.indexOf('cicilan') !== -1) return 7;
    if (n.indexOf('infak') !== -1) return 8;
    if (n.indexOf('index') !== -1) return 9;
    // fallback: if contains 'ir' without '1/10'
    if (n.indexOf('ir') !== -1 && n.indexOf('1/10') === -1) return 5;
  }
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

    // Hitung jumlah bulan sejak periode mulai
    var jumlahBulan = hitungJumlahBulan(periode.tanggalMulai);

    // Ambil anggota aktif
    var sheetAnggota = ss.getSheetByName(CONFIG.SHEETS.ANGGOTA);
    var anggotaAktif = [];
    var anggotaMap = {};
    if (sheetAnggota && sheetAnggota.getLastRow() > 1) {
      var angRows = sheetAnggota.getDataRange().getValues();
      for (var i = 1; i < angRows.length; i++) {
        if (!angRows[i][0]) continue;
        var statusAng = String(angRows[i][4] || '').trim().toLowerCase();
        if (statusAng !== 'nonaktif') {
          anggotaAktif.push({
            id: String(angRows[i][0]),
            nama: String(angRows[i][1] || ''),
            noTelp: String(angRows[i][2] || ''),
            ir: Number(angRows[i][5]) || 0,
            ir10: Number(angRows[i][6]) || 0,
            index: Number(angRows[i][7]) || 0,
            infakDaerah: Number(angRows[i][8]) || 0
          });
          anggotaMap[String(angRows[i][0])] = anggotaAktif[anggotaAktif.length - 1];
        }
      }
    }
    var jumlahJamaahAktif = anggotaAktif.length;

    // Ambil pos setoran aktif
    var sheetPos = ss.getSheetByName(CONFIG.SHEETS.POS_SETORAN);
    var posAktif = [];
    if (sheetPos && sheetPos.getLastRow() > 1) {
      var posRows = sheetPos.getDataRange().getValues();
      for (var i = 1; i < posRows.length; i++) {
        if (!posRows[i][0]) continue;
        if (String(posRows[i][4] || '').trim() === 'Aktif') {
          posAktif.push({
            id: String(posRows[i][0]),
            nama: String(posRows[i][1] || ''),
            tipe: String(posRows[i][2] || ''),
            formula: String(posRows[i][3] || ''),
            status: String(posRows[i][4] || '')
          });
        }
      }
    }

    // Ambil tarif dari Master Musyawaroh — map by nama
    var sheetMsy = ss.getSheetByName(CONFIG.SHEETS.MUSYAWARAH);
    var musyawarahMap = {};
    if (sheetMsy && sheetMsy.getLastRow() > 1) {
      var msyRows = sheetMsy.getDataRange().getValues();
      for (var i = 1; i < msyRows.length; i++) {
        if (msyRows[i][0]) {
          musyawarahMap[String(msyRows[i][1] || '').toLowerCase().trim()] = Number(msyRows[i][2]) || 0;
        }
      }
    }

    // Ambil Detail Buku IR untuk periode ini
    var sheetIR = ss.getSheetByName(CONFIG.SHEETS.BUKU_IR);
    // Kolom: [0]=ID, [1]=TransaksiID, [2]=PeriodeID, [3]=AnggotaID, [4]=Tanggal, [5]=IR, [6]=IR10, [7]=Cicilan, [8]=InfakDaerah, [9]=Index, [10]=Total
    var irData = [];
    if (sheetIR && sheetIR.getLastRow() > 1) {
      var irRows = sheetIR.getDataRange().getValues();
      for (var i = 1; i < irRows.length; i++) {
        if (!irRows[i][0]) continue;
        if (irRows[i][2] === periodeId) {
          irData.push(irRows[i]);
        }
      }
    }

    // Ambil Setoran Desa (sudahSetor) per pos
    var sheetSetoran = ss.getSheetByName(CONFIG.SHEETS.SETORAN_DESA);
    var setoranMap = {};
    if (sheetSetoran && sheetSetoran.getLastRow() > 1) {
      var sRows = sheetSetoran.getDataRange().getValues();
      for (var i = 1; i < sRows.length; i++) {
        if (sRows[i][0]) {
          setoranMap[String(sRows[i][1])] = Number(sRows[i][3]) || 0;
        }
      }
    }

    // Hitung per pos
    var totalTarget = 0, totalTerkumpul = 0, totalSudahSetor = 0;
    var dataPos = [];

    for (var p = 0; p < posAktif.length; p++) {
      var pos = posAktif[p];
      var colIdx = mapPosNamaToBukuIRCol(pos.formula, pos.nama);

      // Tarif dari musyawaroh — cari berdasarkan nama pos
      var tarif = 0;
      var posNamaLower = pos.nama.toLowerCase().trim();
      if (musyawarahMap[posNamaLower] !== undefined) {
        tarif = musyawarahMap[posNamaLower];
      } else {
        // Coba fuzzy match
        for (var mk in musyawarahMap) {
          if (posNamaLower.indexOf(mk) !== -1 || mk.indexOf(posNamaLower) !== -1) {
            tarif = musyawarahMap[mk];
            break;
          }
        }
      }

      var target = tarif * jumlahJamaahAktif * jumlahBulan;

      // Hitung terkumpul dari Detail Buku IR
      var terkumpul = 0;
      if (colIdx >= 0) {
        for (var r = 0; r < irData.length; r++) {
          terkumpul += Number(irData[r][colIdx]) || 0;
        }
      }

      var sudahSetor = setoranMap[pos.id] || 0;
      var sisa = terkumpul - sudahSetor;
      var kekurangan = target - terkumpul;
      var persen = target > 0 ? Math.round((terkumpul / target) * 100 * 10) / 10 : 0;

      var status, tindakLanjut;
      if (target === 0) {
        status = 'Belum Ada Target';
        tindakLanjut = 'Belum ada target yang ditetapkan dari musyawarah untuk pos ini.';
      } else if (terkumpul > target) {
        status = 'Lebih';
        tindakLanjut = 'Terkumpul melebihi target sebesar ' + fmtRp(terkumpul - target) + '. Kelebihan dapat digunakan untuk keperluan lain atau disimpan.';
      } else if (terkumpul >= target) {
        status = 'Tercapai';
        tindakLanjut = 'Target tercapai. Segera setor ke desa sebesar ' + fmtRp(terkumpul - sudahSetor) + ' jika belum dilakukan.';
      } else {
        status = 'Kurang';
        tindakLanjut = 'Masih kurang ' + fmtRp(kekurangan) + ' dari target. Lakukan penagihan kepada jamaah yang belum membayar.';
      }

      totalTarget += target;
      totalTerkumpul += terkumpul;
      totalSudahSetor += sudahSetor;

      dataPos.push({
        id: pos.id,
        nama: pos.nama,
        target: target,
        terkumpul: terkumpul,
        sudahSetor: sudahSetor,
        sisa: sisa,
        kekurangan: kekurangan,
        persen: persen,
        status: status,
        tindakLanjut: tindakLanjut
      });
    }

    var result = {
      periode: periode,
      summary: {
        totalTarget: totalTarget,
        totalTerkumpul: totalTerkumpul,
        totalSudahSetor: totalSudahSetor,
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

function getJamaahBelumBayar(posNama) {
  try {
    var auth = checkAuth();
    if (!auth.success) return { success: false, message: auth.message };

    var ss = getSS_();
    var periode = getPeriodeAktif();
    if (!periode) return { success: false, message: 'Tidak ada periode aktif' };
    var periodeId = periode.id;

    // Ambil pos setoran untuk cari formula
    var sheetPos = ss.getSheetByName(CONFIG.SHEETS.POS_SETORAN);
    var posFormula = '';
    if (sheetPos && sheetPos.getLastRow() > 1) {
      var posRows = sheetPos.getDataRange().getValues();
      for (var i = 1; i < posRows.length; i++) {
        if (String(posRows[i][1] || '') === posNama) {
          posFormula = String(posRows[i][3] || '');
          break;
        }
      }
    }
    var colIdx = mapPosNamaToBukuIRCol(posFormula, posNama);

    // Ambil semua anggota aktif
    var sheetAnggota = ss.getSheetByName(CONFIG.SHEETS.ANGGOTA);
    var anggotaMap = {};
    if (sheetAnggota && sheetAnggota.getLastRow() > 1) {
      var angRows = sheetAnggota.getDataRange().getValues();
      for (var i = 1; i < angRows.length; i++) {
        if (!angRows[i][0]) continue;
        var statusAng = String(angRows[i][4] || '').trim().toLowerCase();
        if (statusAng !== 'nonaktif') {
          var angId = String(angRows[i][0]);
          anggotaMap[angId] = {
            id: angId,
            nama: String(angRows[i][1] || ''),
            noTelp: String(angRows[i][2] || ''),
            ir: Number(angRows[i][5]) || 0,
            ir10: Number(angRows[i][6]) || 0,
            index: Number(angRows[i][7]) || 0,
            infakDaerah: Number(angRows[i][8]) || 0
          };
        }
      }
    }

    // Ambil Detail Buku IR untuk periode ini
    var sheetIR = ss.getSheetByName(CONFIG.SHEETS.BUKU_IR);
    var sudahBayarMap = {}; // anggotaId -> jumlah
    if (sheetIR && sheetIR.getLastRow() > 1 && colIdx >= 0) {
      var irRows = sheetIR.getDataRange().getValues();
      for (var i = 1; i < irRows.length; i++) {
        if (!irRows[i][0]) continue;
        if (irRows[i][2] !== periodeId) continue;
        var angId = String(irRows[i][3]);
        var jumlah = Number(irRows[i][colIdx]) || 0;
        if (jumlah > 0) {
          sudahBayarMap[angId] = (sudahBayarMap[angId] || 0) + jumlah;
        }
      }
    }

    var belumBayar = [], sudahBayar = [];
    for (var aid in anggotaMap) {
      var ang = anggotaMap[aid];
      // Tentukan target bayar per anggota berdasarkan kolom
      var targetBayar = 0;
      if (colIdx === 5) targetBayar = ang.ir;
      else if (colIdx === 6) targetBayar = ang.ir10;
      else if (colIdx === 8) targetBayar = ang.infakDaerah;
      else if (colIdx === 9) targetBayar = ang.index;

      if (sudahBayarMap[aid]) {
        sudahBayar.push({ id: aid, nama: ang.nama, noTelp: ang.noTelp, jumlahBayar: sudahBayarMap[aid] });
      } else {
        belumBayar.push({ id: aid, nama: ang.nama, noTelp: ang.noTelp, targetBayar: targetBayar });
      }
    }

    // Urutkan berdasarkan nama
    belumBayar.sort(function(a, b) { return a.nama.localeCompare(b.nama); });
    sudahBayar.sort(function(a, b) { return a.nama.localeCompare(b.nama); });

    return {
      success: true,
      posNama: posNama,
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

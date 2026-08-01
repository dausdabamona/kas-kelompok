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

// ══════════════════════════════════════════════════════
// FASE 1 — HELPER PENGENDALIAN INTERN
// ══════════════════════════════════════════════════════

// Baris transaksi dianggap dibatalkan (soft delete) bila kolom Status = 'Dibatalkan'.
// Baris lama tanpa kolom/nilai Status → dianggap AKTIF.
function barisDibatalkan_(row, h) {
  if (!h || h['status'] === undefined) return false;
  return String(row[h['status']] || '').toLowerCase().trim() === 'dibatalkan';
}

// K1: Kumpulan Transaksi ID (Input Penerimaan) yang DIBATALKAN.
// Sumber kebenaran tunggal untuk menyaring baris Detail Buku IR yang induknya
// sudah batal. Berlaku surut — tidak butuh kolom baru di Detail Buku IR.
function trxPenerimaanDibatalkan_() {
  var set = {};
  var sh = getSS_().getSheetByName(CONFIG.SHEETS.INPUT_PENERIMAAN);
  if (!sh || sh.getLastRow() < 2) return set;
  var rows = sh.getDataRange().getValues();
  var h = headerMap_(rows[0]);
  for (var i = 1; i < rows.length; i++) {
    if (barisDibatalkan_(rows[i], h)) set[String(hGet_(rows[i], h, 'id', 0))] = true;
  }
  return set;
}

// Penangguhan rincian: transaksi Buku IR yang boleh dirinci di periode berikutnya
// walau periode asalnya sudah ditutup. Uang TETAP di periode asal — yang berpindah
// hanya kewajiban setornya, karena baru diketahui saat rincian diisi.
function rincianDitangguhkan_(row, h) {
  if (!h || h['rincianditangguhkan'] === undefined) return false;
  return String(row[h['rincianditangguhkan']] || '').trim().toLowerCase() === 'ya';
}

// Baris Detail Buku IR dianggap batal bila transaksi induknya batal.
function rincianYatim_(irRow, irH, setBatal) {
  var trxId = String(hGet_(irRow, irH, 'transaksiid', 1) || '');
  return !!setBatal[trxId];
}

// Ambil periode berdasarkan ID dengan status APA PUN (OPEN/CLOSED).
function getPeriodeById_(periodeId) {
  try {
    if (!periodeId) return null;
    var ss = getSS_();
    var sheet = ss.getSheetByName(CONFIG.SHEETS.PERIOD);
    if (!sheet) return null;
    var rows = sheet.getDataRange().getValues();
    var h = headerMap_(rows[0]);
    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(periodeId)) {
        return {
          id: String(rows[i][0]),
          nama: String(hGet_(rows[i], h, 'nama', 1) || ''),
          tanggalMulai: toDateStr_(hGet_(rows[i], h, 'tglmulai', 2)),
          tanggalTutup: toDateStr_(hGet_(rows[i], h, 'tgltutup', 3)),
          status: String(hGet_(rows[i], h, 'status', 4) || ''),
          saldoAwalTunai: Number(hGet_(rows[i], h, 'saldoawaltunai', 5)) || 0,
          saldoAwalBank: Number(hGet_(rows[i], h, 'saldoawalbank', 6)) || 0,
          catatan: String(hGet_(rows[i], h, 'catatan', 7) || '')
        };
      }
    }
    return null;
  } catch(e) { return null; }
}

// Tolak bila periode baris (berdasarkan Periode ID di baris) BUKAN OPEN.
// Menutup T2: transaksi periode tertutup tidak dapat diubah/dihapus.
function assertPeriodeOpen_(periodeIdBaris) {
  var p = getPeriodeById_(periodeIdBaris);
  if (!p) return { ok: false, message: 'Periode transaksi tidak ditemukan.' };
  if (String(p.status) !== CONFIG.STATUS.OPEN) {
    return { ok: false, message: 'Transaksi ini milik periode yang sudah ditutup dan tidak dapat diubah.' };
  }
  return { ok: true, periode: p };
}

// Validasi nominal uang: bilangan bulat rupiah > 0, dalam batas wajar.
function validasiNominal_(n) {
  var v = Number(n);
  if (!isFinite(v)) return { ok: false, message: 'Nominal tidak valid.' };
  if (Math.floor(v) !== v) return { ok: false, message: 'Nominal harus bilangan bulat rupiah (tanpa desimal).' };
  if (v <= 0) return { ok: false, message: 'Nominal harus lebih besar dari 0.' };
  if (v > 100000000000) return { ok: false, message: 'Nominal melebihi batas wajar.' };
  return { ok: true, nilai: v };
}

// Validasi tanggal transaksi: tidak di masa depan & tidak sebelum awal periode.
function validasiTanggalPeriode_(rawTgl, periodeId) {
  var t = toDateStr_(rawTgl instanceof Date ? rawTgl : (rawTgl ? new Date(rawTgl) : null));
  if (!/^\d{4}-\d{2}-\d{2}$/.test(t)) t = toDateStr_(rawTgl); // fallback string apa adanya
  if (!/^\d{4}-\d{2}-\d{2}$/.test(t)) return { ok: false, message: 'Tanggal tidak valid.' };
  var today = toDateStr_(new Date());
  if (t > today) return { ok: false, message: 'Tanggal transaksi tidak boleh di masa depan.' };
  var p = getPeriodeById_(periodeId);
  if (p && /^\d{4}-\d{2}-\d{2}$/.test(p.tanggalMulai) && t < p.tanggalMulai) {
    return { ok: false, message: 'Tanggal (' + t + ') sebelum awal periode (' + p.tanggalMulai + ').' };
  }
  return { ok: true, tgl: t };
}

// FASE 4 (T11): isi No Bukti berseri pada baris TERAKHIR sheet (dipanggil
// tepat setelah appendRow, di dalam withLock_). prefix BKM/BKK, seq per periode,
// tanpa daur ulang (baris dibatalkan tetap memegang nomornya).
function _isiNoBukti_(sheet, prefix, periodeId) {
  try {
    var lastCol = sheet.getLastColumn();
    var header = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    var h = headerMap_(header);
    var cNB = h['nobukti']; if (cNB === undefined) return '';
    var cPid = h['periodeid'] !== undefined ? h['periodeid'] : 1;
    var maxSeq = 0;
    if (sheet.getLastRow() > 1) {
      var rows = sheet.getRange(1, 1, sheet.getLastRow(), lastCol).getValues();
      for (var i = 1; i < rows.length; i++) {
        if (String(rows[i][cPid]) !== String(periodeId)) continue;
        var m = String(rows[i][cNB] || '').match(/-(\d+)$/);
        if (m) { var n = parseInt(m[1], 10); if (n > maxSeq) maxSeq = n; }
      }
    }
    var pseq = _periodeSeqCode_(getSS_(), periodeId);
    var nb = prefix + '-' + pseq + '-' + ('0000' + (maxSeq + 1)).slice(-4);
    sheet.getRange(sheet.getLastRow(), cNB + 1).setValue(nb);
    return nb;
  } catch(e) { return ''; }
}

// FASE 5 (T12): set kolom approval pada baris TERAKHIR (setelah append).
function _isiApproval_(sheet, statusApproval, by, at) {
  try {
    var lastCol = sheet.getLastColumn();
    var h = headerMap_(sheet.getRange(1, 1, 1, lastCol).getValues()[0]);
    var rowNum = sheet.getLastRow();
    if (h['statusapproval'] !== undefined) sheet.getRange(rowNum, h['statusapproval'] + 1).setValue(statusApproval);
    if (by && h['disetujuiby'] !== undefined) sheet.getRange(rowNum, h['disetujuiby'] + 1).setValue(by);
    if (at && h['disetujuiat'] !== undefined) sheet.getRange(rowNum, h['disetujuiat'] + 1).setValue(at);
  } catch(e) {}
}

// Baris pengeluaran berstatus Draft (belum disetujui) — tidak dihitung ke saldo.
function barisDraft_(row, h) {
  if (!h || h['statusapproval'] === undefined) return false;
  return String(row[h['statusapproval']] || '').toLowerCase().trim() === 'draft';
}

// Log WAJIB untuk mutasi material: melempar error bila gagal menulis
// (berbeda dari logActivity yang silent). Menutup T3.
function logActivityWajib_(user, action, detail) {
  var ss = getSS_();
  var sheet = ss.getSheetByName(CONFIG.SHEETS.LOG);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEETS.LOG);
    sheet.appendRow(['Timestamp', 'User', 'Action', 'Detail']);
  }
  sheet.appendRow([new Date(), user, action, detail]);
}

// Periode ID sebuah patungan/terobosan (untuk kunci periode tagihan).
function getPatunganPeriodeId_(patunganId) {
  try {
    var ss = getSS_();
    var sheet = ss.getSheetByName(CONFIG.SHEETS.PATUNGAN);
    if (!sheet) return '';
    var rows = sheet.getDataRange().getValues();
    var h = headerMap_(rows[0]);
    for (var i = 1; i < rows.length; i++) {
      if (String(hGet_(rows[i], h, 'id', 0)) === String(patunganId)) return String(hGet_(rows[i], h, 'periodeid', 3) || '');
    }
    return '';
  } catch(e) { return ''; }
}

// FASE 4 (T10) — LAMPIRAN BUKTI
// Simpan 1 gambar base64 ke Drive + catat di sheet Lampiran.
function _simpanBukti1_(transaksiId, tipe, base64, email) {
  var folderId = PropertiesService.getScriptProperties().getProperty('FOLDER_BUKTI_ID');
  if (!folderId) throw new Error('Folder bukti belum dikonfigurasi (FOLDER_BUKTI_ID).');
  var m = String(base64).match(/^data:([^;]+);base64,(.*)$/);
  var mime = m ? m[1] : 'image/jpeg';
  var raw = m ? m[2] : base64;
  var bytes = Utilities.base64Decode(raw);
  var nama = 'bukti_' + transaksiId + '_' + new Date().getTime() + '.jpg';
  var file = DriveApp.getFolderById(folderId).createFile(Utilities.newBlob(bytes, mime, nama));
  var ss = getSS_();
  var sheet = ss.getSheetByName(CONFIG.SHEETS.LAMPIRAN);
  if (!sheet) { sheet = ss.insertSheet(CONFIG.SHEETS.LAMPIRAN); sheet.appendRow(['ID', 'Transaksi ID', 'Tipe', 'Nama File', 'Drive File ID', 'URL', 'Diunggah By', 'Diunggah At', 'Status']); }
  var id = generateID('LMP');
  sheet.appendRow([id, transaksiId, tipe || '', nama, file.getId(), file.getUrl(), email || '', toDateStr_(new Date()), 'Aktif']);
  return { id: id, url: file.getUrl() };
}

function _simpanBuktiList_(transaksiId, tipe, arr, email) {
  var n = 0;
  (arr || []).forEach(function(b) { if (b) { try { _simpanBukti1_(transaksiId, tipe, b, email); n++; } catch(e) {} } });
  return n;
}

// Jumlah lampiran Aktif untuk sebuah transaksi.
function countLampiran_(transaksiId) {
  var ss = getSS_();
  var sheet = ss.getSheetByName(CONFIG.SHEETS.LAMPIRAN);
  if (!sheet || sheet.getLastRow() < 2) return 0;
  var rows = sheet.getDataRange().getValues(); var h = headerMap_(rows[0]);
  var n = 0;
  for (var i = 1; i < rows.length; i++) {
    if (String(hGet_(rows[i], h, 'transaksiid', 1)) === String(transaksiId) && String(hGet_(rows[i], h, 'status', 8)) === 'Aktif') n++;
  }
  return n;
}

// Endpoint: unggah bukti untuk transaksi yang sudah ada (mis. retry offline).
function uploadBukti(data) {
  try {
    var auth = requirePermInput_('keluar');
    if (!auth.success) return auth;
    if (!data || !data.transaksiId || !data.base64) return { success: false, message: 'Data tidak lengkap.' };
    var r = _simpanBukti1_(String(data.transaksiId), String(data.tipe || ''), String(data.base64), auth.user.email);
    return { success: true, id: r.id, url: r.url };
  } catch(e) { return { success: false, message: e.message }; }
}

// Endpoint: daftar lampiran sebuah transaksi.
function getLampiran(transaksiId) {
  try {
    var auth = checkAuth();
    if (!auth.success) return auth;
    var ss = getSS_();
    var sheet = ss.getSheetByName(CONFIG.SHEETS.LAMPIRAN);
    if (!sheet || sheet.getLastRow() < 2) return { success: true, data: [] };
    var rows = sheet.getDataRange().getValues(); var h = headerMap_(rows[0]);
    var out = [];
    for (var i = 1; i < rows.length; i++) {
      if (String(hGet_(rows[i], h, 'transaksiid', 1)) !== String(transaksiId)) continue;
      if (String(hGet_(rows[i], h, 'status', 8)) !== 'Aktif') continue;
      out.push({ id: String(hGet_(rows[i], h, 'id', 0)), url: String(hGet_(rows[i], h, 'url', 5) || ''), nama: String(hGet_(rows[i], h, 'namafile', 3) || '') });
    }
    return { success: true, data: out };
  } catch(e) { return { success: false, message: e.message }; }
}

// Guard saldo negatif (T13): pastikan saldo sumber kas cukup untuk pengeluaran
// atau mutasi. override=true (khusus ADMIN + alasan) melewati guard.
function cekSaldoCukup_(periode, sumberKas, nominal) {
  var saldo = calculateSaldo(periode.id, periode);   // K3: error naik, tidak ditelan
  var s = (sumberKas === 'Bank') ? saldo.bank : saldo.tunai;
  if (nominal > s) {
    return { ok: false, message: 'Saldo ' + (sumberKas === 'Bank' ? 'Bank' : 'Tunai') + ' tidak mencukupi (saldo: Rp ' + Number(s).toLocaleString('id-ID') + ', diminta: Rp ' + Number(nominal).toLocaleString('id-ID') + ').', saldo: s };
  }
  return { ok: true, saldo: s };
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
  try {
    var auth = checkAuth();
    if (!auth.success) return { success: false, message: auth.message };
    var ss = getSS_();
    var sheet = ss.getSheetByName(CONFIG.SHEETS.PERIOD);
    if (!sheet || sheet.getLastRow() < 2) return { success: true, data: [] };
    var data = sheet.getDataRange().getValues();
    var h = headerMap_(data[0]);
    // Indeks saldo akhir per periode dari Saldo Tutup Buku (baris 'Tutup').
    var akhirMap = {};
    var shT = ss.getSheetByName(CONFIG.SHEETS.SALDO_TUTUP_BUKU);
    if (shT && shT.getLastRow() > 1) {
      var tr = shT.getDataRange().getValues(); var th = headerMap_(tr[0]);
      for (var t = 1; t < tr.length; t++) {
        if (String(hGet_(tr[t], th, 'status', 6)) !== 'Tutup') continue;
        akhirMap[String(hGet_(tr[t], th, 'periodeid', 1) || '')] = {
          tunai: Number(hGet_(tr[t], th, 'saldotunaiakhir', 3)) || 0,
          bank: Number(hGet_(tr[t], th, 'saldobankakhir', 4)) || 0
        };
      }
    }
    var result = [];
    for (var i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      var pid = String(hGet_(data[i], h, 'periode', 0) || data[i][0]);
      var status = String(hGet_(data[i], h, 'status', 4) || '');
      var awalTunai = Number(hGet_(data[i], h, 'saldoawaltunai', 5)) || 0;
      var awalBank = Number(hGet_(data[i], h, 'saldoawalbank', 6)) || 0;
      // Saldo akhir: CLOSED → dari arsip tutup buku; OPEN → hitung live.
      var akhirTunai = null, akhirBank = null;
      if (String(status).toUpperCase() === CONFIG.STATUS.CLOSED && akhirMap[pid]) {
        akhirTunai = akhirMap[pid].tunai; akhirBank = akhirMap[pid].bank;
      } else if (String(status).toUpperCase() === CONFIG.STATUS.OPEN) {
        try {
          var s = calculateSaldo(pid, { saldoAwalTunai: awalTunai, saldoAwalBank: awalBank });
          akhirTunai = s.tunai; akhirBank = s.bank;
        } catch(e) {}
      }
      result.push({
        id: pid,
        nama: String(hGet_(data[i], h, 'nama', 1) || ''),
        tanggalMulai: toDateStr_(hGet_(data[i], h, 'tglmulai', 2)),
        tanggalTutup: toDateStr_(hGet_(data[i], h, 'tgltutup', 3)),
        status: status,
        saldoAwalTunai: awalTunai,
        saldoAwalBank: awalBank,
        saldoAkhirTunai: akhirTunai,   // null bila tak tersedia
        saldoAkhirBank: akhirBank,
        saldoAkhirTotal: (akhirTunai === null && akhirBank === null) ? null : (Number(akhirTunai || 0) + Number(akhirBank || 0))
      });
    }
    return { success: true, data: result };
  } catch(e) {
    return { success: false, message: 'Gagal memuat periode: ' + e.message };
  }
}

// ──────────────────────────────────────────────────────
// DASHBOARD
// ──────────────────────────────────────────────────────
// T7: kas yang masih di tangan penerobos (status Aktif) + aging per penerobos.
function kasPenerobosAktif_(periodeId) {
  var ss = getSS_();
  var out = { total: 0, perPenerobos: [] };
  var sheet = ss.getSheetByName(CONFIG.SHEETS.KAS_PENEROBOS);
  if (!sheet || sheet.getLastRow() < 2) return out;
  var namaUser = {};
  try { getUserList_().forEach(function(u) { namaUser[u.email.toLowerCase()] = u.nama || u.email; }); } catch(e) {}
  var rows = sheet.getDataRange().getValues();
  var h = headerMap_(rows[0]);
  var today = new Date();
  var map = {};
  for (var i = 1; i < rows.length; i++) {
    if (String(hGet_(rows[i], h, 'status', 9)) !== 'Aktif') continue;
    if (periodeId && String(hGet_(rows[i], h, 'periodeid', 1)) !== periodeId) continue;
    var email = String(hGet_(rows[i], h, 'penerobosemail', 8) || '');
    var nominal = Number(hGet_(rows[i], h, 'nominal', 5)) || 0;
    var tglStr = toDateStr_(hGet_(rows[i], h, 'tanggal', 2));
    var umur = 0;
    if (/^\d{4}-\d{2}-\d{2}$/.test(tglStr)) { umur = Math.floor((today.getTime() - new Date(tglStr).getTime()) / 86400000); if (umur < 0) umur = 0; }
    out.total += nominal;
    var key = email.toLowerCase();
    if (!map[key]) map[key] = { email: email, nama: namaUser[key] || email, total: 0, tertuaHari: 0 };
    map[key].total += nominal;
    if (umur > map[key].tertuaHari) map[key].tertuaHari = umur;
  }
  Object.keys(map).forEach(function(k) { out.perPenerobos.push(map[k]); });
  out.perPenerobos.sort(function(a, b) { return b.tertuaHari - a.tertuaHari; });
  return out;
}

// L2c: ID transaksi sumber sebuah baris Kas Penerobos.
// Baris hasil migrasi berformat 'KP_MIG_<idSumber>'; awalan dibuang.
// Baris penerobos normal (TRX_*) TIDAK diturunkan dari Input Penerimaan → ''.
function _trxSumberIdPenerobos_(kpId) {
  var s = String(kpId || '');
  var pfx = 'KP_MIG_';
  if (s.indexOf(pfx) === 0) return s.substring(pfx.length);
  return '';
}

// Tanda tangan baris untuk mendeteksi duplikat internal (dua baris identik).
function _sigPenerobos_(row, h) {
  return [
    String(hGet_(row, h, 'periodeid', 1) || ''),
    toDateStr_(hGet_(row, h, 'tanggal', 2)),
    String(hGet_(row, h, 'jenisid', 3) || ''),
    String(hGet_(row, h, 'anggotaid', 4) || ''),
    String(Number(hGet_(row, h, 'nominal', 5)) || 0),
    String(hGet_(row, h, 'sumberkas', 6) || ''),
    String(hGet_(row, h, 'penerobosemail', 8) || '').toLowerCase()
  ].join('|');
}

// L2c: Deteksi kandidat duplikat Kas Penerobos + telusuri transaksi sumbernya.
// READ-ONLY. Mengembalikan analisis PER BARIS (semua status) — pemanggil menyaring.
function deteksiDuplikatPenerobos_() {
  var ss = getSS_();
  var sheet = ss.getSheetByName(CONFIG.SHEETS.KAS_PENEROBOS);
  if (!sheet || sheet.getLastRow() < 2) return [];
  // Indeks Input Penerimaan by id → { batal, ... }, + daftar baris aktif untuk pencocokan.
  var pen = {};
  var penAktif = []; // untuk deteksi "dicatat langsung"
  var shP = ss.getSheetByName(CONFIG.SHEETS.INPUT_PENERIMAAN);
  if (shP && shP.getLastRow() > 1) {
    var pr = shP.getDataRange().getValues();
    var ph = headerMap_(pr[0]);
    for (var p = 1; p < pr.length; p++) {
      var pidr = String(hGet_(pr[p], ph, 'id', 0) || '');
      if (!pidr) continue;
      var batalP = barisDibatalkan_(pr[p], ph);
      pen[pidr] = { batal: batalP };
      if (!batalP) {
        penAktif.push({
          id: pidr,
          tgl: toDateStr_(hGet_(pr[p], ph, 'tanggal', 4)),
          nominal: Number(hGet_(pr[p], ph, 'nominal', 5)) || 0,
          sumberKas: String(hGet_(pr[p], ph, 'sumberkas', 6) || ''),
          jenisId: String(hGet_(pr[p], ph, 'jenisid', 2) || '')
        });
      }
    }
  }
  // Cari baris Input Penerimaan yang "mungkin pasangan" catatan langsung (±30 hari).
  function cariPasangan_(tglStr, nominal, sumberKas, jenisId) {
    if (!/^\d{4}-\d{2}-\d{2}$/.test(tglStr)) return null;
    var t0 = new Date(tglStr).getTime();
    for (var q = 0; q < penAktif.length; q++) {
      var c = penAktif[q];
      if (c.nominal !== nominal) continue;
      if (sumberKas && c.sumberKas && c.sumberKas !== sumberKas) continue;
      if (jenisId && c.jenisId && c.jenisId !== jenisId) continue;
      if (!/^\d{4}-\d{2}-\d{2}$/.test(c.tgl)) continue;
      var beda = Math.abs(t0 - new Date(c.tgl).getTime()) / 86400000;
      if (beda <= 30) return { id: c.id, tanggal: c.tgl, nominal: c.nominal };
    }
    return null;
  }
  var namaUser = {}; try { getUserList_().forEach(function(u) { namaUser[String(u.email).toLowerCase()] = u.nama || u.email; }); } catch(e) {}
  var rows = sheet.getDataRange().getValues();
  var h = headerMap_(rows[0]);
  var today = new Date();
  // Kelompokkan tanda tangan baris Aktif untuk duplikat internal.
  var sigMap = {};
  for (var s1 = 1; s1 < rows.length; s1++) {
    if (String(hGet_(rows[s1], h, 'status', 9)) !== 'Aktif') continue;
    var sg = _sigPenerobos_(rows[s1], h);
    (sigMap[sg] = sigMap[sg] || []).push(String(hGet_(rows[s1], h, 'id', 0) || ''));
  }
  var out = [];
  for (var i = 1; i < rows.length; i++) {
    var id = String(hGet_(rows[i], h, 'id', 0) || '');
    if (!id) continue;
    var status = String(hGet_(rows[i], h, 'status', 9) || 'Aktif');
    var nominal = Number(hGet_(rows[i], h, 'nominal', 5)) || 0;
    var email = String(hGet_(rows[i], h, 'penerobosemail', 8) || '');
    var tglStr = toDateStr_(hGet_(rows[i], h, 'tanggal', 2));
    var umur = 0;
    if (/^\d{4}-\d{2}-\d{2}$/.test(tglStr)) { umur = Math.floor((today.getTime() - new Date(tglStr).getTime()) / 86400000); if (umur < 0) umur = 0; }
    var trxSumberId = _trxSumberIdPenerobos_(id);
    var sumberAda = false, sumberBatal = false;
    if (trxSumberId && pen[trxSumberId]) { sumberAda = true; sumberBatal = !!pen[trxSumberId].batal; }
    var sig = _sigPenerobos_(rows[i], h);
    var kembaran = (status === 'Aktif' && sigMap[sig] && sigMap[sig].length > 1)
      ? sigMap[sig].filter(function(x) { return x !== id; }) : [];
    var sumberKasRow = String(hGet_(rows[i], h, 'sumberkas', 6) || 'Tunai');
    var jenisIdRow = String(hGet_(rows[i], h, 'jenisid', 3) || '');
    var jenisDuplikat = 'BUKAN', risiko = 'NORMAL', kandidatPasangan = null;
    if (status === 'Aktif') {
      if (trxSumberId && sumberAda && !sumberBatal) {
        // 1) DUPLIKAT SILANG: transaksi sumber masih ada & aktif → uang dobel.
        jenisDuplikat = 'SILANG'; risiko = 'AMAN_DIBATALKAN';
      } else if (kembaran.length) {
        // 3) DUPLIKAT INTERNAL: dua baris Kas Penerobos identik.
        jenisDuplikat = 'INTERNAL'; risiko = 'AMAN_DIBATALKAN';
      } else {
        // 2) DICATAT LANGSUNG: cari baris Input Penerimaan mirip (±30 hari) — bendahara
        //    mungkin sudah mencatat langsung saat serah terima gagal. TANDAI, jangan simpulkan.
        kandidatPasangan = cariPasangan_(tglStr, nominal, sumberKasRow, jenisIdRow);
        if (kandidatPasangan) { jenisDuplikat = 'DICATAT LANGSUNG'; risiko = 'PERLU_DIKONFIRMASI'; }
        else if (trxSumberId) { jenisDuplikat = 'TIDAK ADA'; risiko = 'BAHAYA_UANG_HILANG'; }
        // baris penerobos normal tanpa jejak apa pun → tetap NORMAL (memang belum diserahkan).
      }
    }
    out.push({
      id: id,
      periodeId: String(hGet_(rows[i], h, 'periodeid', 1) || ''),
      tanggal: tglStr,
      jenisId: jenisIdRow,
      anggotaId: String(hGet_(rows[i], h, 'anggotaid', 4) || ''),
      nominal: nominal,
      sumberKas: sumberKasRow,
      penerobosEmail: email,
      penerobosNama: namaUser[email.toLowerCase()] || email,
      status: status,
      umurHari: umur,
      trxSumberId: trxSumberId,
      sumberAda: sumberAda,
      sumberBatal: sumberBatal,
      jenisDuplikat: jenisDuplikat,
      risiko: risiko,
      kembaranIds: kembaran,
      kandidatPasangan: kandidatPasangan
    });
  }
  return out;
}

// L2b: total baris Kas Penerobos berstatus Aktif yang periodenya sudah CLOSED.
// Uang masih di tangan penerobos padahal bukunya sudah ditutup — perlu diperingatkan.
function penerobosAktifPeriodeTertutup_() {
  var out = { total: 0, count: 0 };
  var ss = getSS_();
  var sheet = ss.getSheetByName(CONFIG.SHEETS.KAS_PENEROBOS);
  if (!sheet || sheet.getLastRow() < 2) return out;
  // Kumpulan periode yang berstatus OPEN.
  var openSet = {};
  var shP = ss.getSheetByName(CONFIG.SHEETS.PERIOD);
  if (shP && shP.getLastRow() > 1) {
    var pr = shP.getDataRange().getValues();
    var ph = headerMap_(pr[0]);
    for (var p = 1; p < pr.length; p++) {
      if (String(hGet_(pr[p], ph, 'status', 4)).toUpperCase() === CONFIG.STATUS.OPEN) openSet[String(pr[p][0])] = true;
    }
  }
  var rows = sheet.getDataRange().getValues();
  var h = headerMap_(rows[0]);
  for (var i = 1; i < rows.length; i++) {
    if (String(hGet_(rows[i], h, 'status', 9)) !== 'Aktif') continue;
    var pid = String(hGet_(rows[i], h, 'periodeid', 1) || '');
    if (openSet[pid]) continue; // periode masih OPEN → bukan "di luar buku"
    out.total += Number(hGet_(rows[i], h, 'nominal', 5)) || 0;
    out.count++;
  }
  return out;
}

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
      var tunai = 0, bank = 0, peringatan = '', sumberSaldo = 'kosong';
      if (periode) {
        var s = calculateSaldo(periode.id, periode);   // biarkan error naik — jangan ditelan
        tunai = s.tunai; bank = s.bank;
        sumberSaldo = 'periode';
      } else {
        // Tidak ada periode OPEN → JANGAN hitung ulang. Pakai arsip tutup buku.
        var akhir = getSaldoTutupBukuTerakhir_();
        if (akhir) {
          tunai = akhir.tunai; bank = akhir.bank;
          sumberSaldo = 'arsip';
          peringatan = 'Periode sudah ditutup (' + akhir.tanggalTutup + '). Angka di bawah adalah ' +
                       'saldo akhir tutup buku. Buka periode baru untuk mulai mencatat.';
        } else {
          peringatan = 'Belum ada periode aktif. Buka periode terlebih dahulu.';
        }
      }
      saldoData = { namaKelompok: namaKelompok, tunai: tunai, bank: bank, peringatan: peringatan, sumberSaldo: sumberSaldo };
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

    // T12: pengeluaran menunggu persetujuan (Draft).
    var draft = { count: 0, total: 0 };
    try { var dp2 = getDraftPengeluaran(); if (dp2 && dp2.success) { draft.count = (dp2.data || []).length; draft.total = dp2.total || 0; } } catch(e) {}

    // T7/L2b: pos ketiga — kas di tangan penerobos (belum masuk kas kelompok).
    // Mode 'periode' → kas penerobos periode itu, KOMPONEN Total Kas.
    // Mode 'arsip'/'kosong' → kas penerobos "saat ini" lintas periode, DI LUAR Total
    //   (karena Total berasal dari snapshot arsip beku — basis waktu berbeda).
    var sumberSaldo = saldoData.sumberSaldo || (periode ? 'periode' : 'kosong');
    var kp = { total: 0, perPenerobos: [] };
    try { kp = kasPenerobosAktif_(sumberSaldo === 'periode' ? periodeId : null); } catch(e) {}
    var penerobosDalamTotal = (sumberSaldo === 'periode');
    // L2b: baris Kas Penerobos di periode CLOSED yang masih Aktif (uang di luar buku).
    var penerobosTutup = { total: 0, count: 0 };
    try { penerobosTutup = penerobosAktifPeriodeTertutup_(); } catch(e) {}
    var agingHari = 7;
    try { var av = PropertiesService.getScriptProperties().getProperty('AGING_HARI'); if (av) agingHari = Number(av) || 7; } catch(e) {}
    var buildDate = '';
    try { buildDate = PropertiesService.getScriptProperties().getProperty('BUILD_DATE') || ''; } catch(e) {}

    return {
      success: true,
      user: auth.user,
      periode: periode,
      sumberSaldo: sumberSaldo,
      namaKelompok: saldoData.namaKelompok,
      kasTunai: saldoData.tunai,
      kasBank: saldoData.bank,
      kasKelompok: saldoData.tunai + saldoData.bank,
      kasPenerobos: kp.total,
      penerobosDetail: kp.perPenerobos,
      penerobosDalamTotal: penerobosDalamTotal,
      penerobosTutup: penerobosTutup,
      agingHari: agingHari,
      totalKas: saldoData.tunai + saldoData.bank + (penerobosDalamTotal ? kp.total : 0),
      belumDirincikanCount: belumCount,
      menungguApproval: draft,
      peringatan: saldoData.peringatan || '',
      buildDate: buildDate
    };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function calculateSaldo(periodeId, periode) {
  // K3: TANPA periode, saldo awal hilang dan filter transaksi mati → hasilnya salah.
  // Jangan pernah kembalikan angka. Pemanggil wajib menangani ketiadaan periode.
  if (!periodeId || !periode) {
    throw new Error('calculateSaldo: periode wajib. Gunakan getSaldoTutupBukuTerakhir_() bila tidak ada periode OPEN.');
  }
  var ss = getSS_();
  var tunai = Number(periode.saldoAwalTunai) || 0;
  var bank  = Number(periode.saldoAwalBank)  || 0;

  var sheetP = ss.getSheetByName(CONFIG.SHEETS.INPUT_PENERIMAAN);
  if (sheetP && sheetP.getLastRow() > 1) {
    var dp = sheetP.getDataRange().getValues();
    var dpH = headerMap_(dp[0]);
    for (var i = 1; i < dp.length; i++) {
      if (barisDibatalkan_(dp[i], dpH)) continue; // T3: abaikan yang dibatalkan
      if (String(hGet_(dp[i], dpH, 'periodeid', 1)) !== String(periodeId)) continue;  // ← tegas
      var nominal = Number(hGet_(dp[i], dpH, 'nominal', 5)) || 0;
      var sumber  = String(hGet_(dp[i], dpH, 'sumberkas', 6) || '');
      if (sumber === 'Tunai') tunai += nominal;
      else if (sumber === 'Bank') bank += nominal;
    }
  }

  var sheetPK = ss.getSheetByName(CONFIG.SHEETS.INPUT_PENGELUARAN);
  if (sheetPK && sheetPK.getLastRow() > 1) {
    var dpk = sheetPK.getDataRange().getValues();
    var dpkH = headerMap_(dpk[0]);
    for (var j = 1; j < dpk.length; j++) {
      if (barisDibatalkan_(dpk[j], dpkH)) continue; // T3
      if (barisDraft_(dpk[j], dpkH)) continue;       // T12: Draft belum masuk saldo
      if (String(hGet_(dpk[j], dpkH, 'periodeid', 1)) !== String(periodeId)) continue; // ← tegas
      var nomK = Number(hGet_(dpk[j], dpkH, 'nominal', 4)) || 0;
      var sumK = String(hGet_(dpk[j], dpkH, 'sumberkas', 5) || '');
      if (sumK === 'Tunai') tunai -= nomK;
      else if (sumK === 'Bank') bank -= nomK;
    }
  }

  var sheetS = ss.getSheetByName(CONFIG.SHEETS.INPUT_SETORAN);
  if (sheetS && sheetS.getLastRow() > 1) {
    var ds = sheetS.getDataRange().getValues();
    var dsH = headerMap_(ds[0]);
    for (var k = 1; k < ds.length; k++) {
      if (barisDibatalkan_(ds[k], dsH)) continue; // T3
      if (String(hGet_(ds[k], dsH, 'periodeid', 1)) !== String(periodeId)) continue;  // ← tegas
      var nomS = Number(hGet_(ds[k], dsH, 'nominal', 3)) || 0;
      var arah = String(hGet_(ds[k], dsH, 'arah', 4) || '');
      if (arah === 'setor')      { tunai -= nomS; bank += nomS; }
      else if (arah === 'tarik') { bank  -= nomS; tunai += nomS; }
    }
  }

  return { tunai: tunai, bank: bank };
}

// K3: Saldo "saat ini" yang aman dipanggil walau tidak ada periode OPEN.
// periode OPEN → calculateSaldo; tanpa periode → saldo arsip tutup buku (atau 0).
function saldoSaatIni_(periodeId, periode) {
  if (periodeId && periode) return calculateSaldo(periodeId, periode);
  var akhir = getSaldoTutupBukuTerakhir_();
  if (akhir) return { tunai: akhir.tunai, bank: akhir.bank };
  return { tunai: 0, bank: 0 };
}

// K3: Saldo akhir dari baris 'Tutup' TERAKHIR (periode CLOSED terbaru).
// Dipakai bila tidak ada periode OPEN. null bila belum pernah tutup buku.
function getSaldoTutupBukuTerakhir_() {
  var sheet = getSS_().getSheetByName(CONFIG.SHEETS.SALDO_TUTUP_BUKU);
  if (!sheet || sheet.getLastRow() < 2) return null;
  var rows = sheet.getDataRange().getValues();
  var h = headerMap_(rows[0]);
  for (var i = rows.length - 1; i >= 1; i--) {
    if (String(hGet_(rows[i], h, 'status', 6) || '').trim() !== 'Tutup') continue;
    return {
      periodeId: String(hGet_(rows[i], h, 'periodeid', 1) || ''),
      tanggalTutup: toDateStr_(hGet_(rows[i], h, 'tanggaltutup', 2)),
      tunai: Number(hGet_(rows[i], h, 'saldotunaiakhir', 3)) || 0,
      bank:  Number(hGet_(rows[i], h, 'saldobankakhir', 4)) || 0
    };
  }
  return null;
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
    return withLock_(function() {
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
    });
  } catch(e) {
    return { success: false, message: e.message };
  }
}

// FASE 2 — syarat wajib sebelum tutup buku (T5): tidak boleh ada uang
// "menggantung" atau data belum lengkap.
function cekSyaratTutupBuku_(periodeId) {
  var ss = getSS_();
  // Bank Pending berstatus Pending
  var shBP = ss.getSheetByName(CONFIG.SHEETS.BANK_PENDING);
  if (shBP && shBP.getLastRow() > 1) {
    var bp = shBP.getDataRange().getValues(); var hbp = headerMap_(bp[0]);
    for (var i = 1; i < bp.length; i++) {
      if (String(hGet_(bp[i], hbp, 'periodeid', 1)) === String(periodeId) && String(hGet_(bp[i], hbp, 'status', 6)) === 'Pending')
        return { ok: false, message: 'Masih ada transaksi Bank berstatus Pending. Selesaikan rekonsiliasi bank dulu.' };
    }
  }
  // Kas Penerobos berstatus Aktif (uang masih di tangan penerobos)
  var shKP = ss.getSheetByName(CONFIG.SHEETS.KAS_PENEROBOS);
  if (shKP && shKP.getLastRow() > 1) {
    var kp = shKP.getDataRange().getValues(); var hkp = headerMap_(kp[0]);
    for (var i = 1; i < kp.length; i++) {
      if (String(hGet_(kp[i], hkp, 'periodeid', 1)) === String(periodeId) && String(hGet_(kp[i], hkp, 'status', 9)) === 'Aktif')
        return { ok: false, message: 'Masih ada Kas Penerobos yang belum diserahterimakan. Selesaikan serah terima dulu.' };
    }
  }
  // Serah Terima berstatus Menunggu
  var shST = ss.getSheetByName(CONFIG.SHEETS.SERAH_TERIMA);
  if (shST && shST.getLastRow() > 1) {
    var st = shST.getDataRange().getValues(); var hst = headerMap_(st[0]);
    for (var i = 1; i < st.length; i++) {
      if (String(hGet_(st[i], hst, 'periodeid', 1)) === String(periodeId) && String(hGet_(st[i], hst, 'status', 7)) === 'Menunggu')
        return { ok: false, message: 'Masih ada Serah Terima berstatus Menunggu. Konfirmasi dulu.' };
    }
  }
  // Buku IR belum dirincikan (yang sudah DITANGGUHKAN tidak memblokir).
  try {
    var bir = getBukuIRData();
    var belum = (bir && bir.success && bir.data && bir.data.belumDirincikan) ? bir.data.belumDirincikan : [];
    var blokir = belum.filter(function(t) { return !t.ditangguhkan; });
    if (blokir.length > 0)
      return { ok: false, bisaTangguhkan: true, jumlahBelumDirinci: blokir.length,
        message: 'Masih ada ' + blokir.length + ' transaksi Buku IR yang belum dirincikan.' };
  } catch(e) {}
  return { ok: true };
}

// Tandai transaksi Buku IR periode ini yang BELUM dirinci sebagai "ditangguhkan",
// agar tutup buku tidak terhalang dan rinciannya bisa diisi di periode berikutnya.
// Uang tidak dipindah — hanya penanda pada baris transaksi. Mengembalikan jumlahnya.
function _tangguhkanRincianBukuIR_(periodeId, email) {
  var ss = getSS_();
  var sheet = ss.getSheetByName(CONFIG.SHEETS.INPUT_PENERIMAAN);
  if (!sheet || sheet.getLastRow() < 2) return 0;
  ensureColumns_(sheet, ['Rincian Ditangguhkan', 'Ditangguhkan At']);
  // Jenis berkategori Buku IR.
  var bukuIRIds = {};
  var shM = ss.getSheetByName(CONFIG.SHEETS.PEMASUKAN);
  if (shM && shM.getLastRow() > 1) {
    var mr = shM.getDataRange().getValues(); var mh = headerMap_(mr[0]);
    for (var m = 1; m < mr.length; m++) {
      var kd = String(hGet_(mr[m], mh, 'kode', 0) || '');
      if (kd && String(hGet_(mr[m], mh, 'kategori', 2) || '').toLowerCase().indexOf('buku ir') !== -1) bukuIRIds[kd] = true;
    }
  }
  // Transaksi yang sudah punya rincian.
  var sudah = {};
  var shIR = ss.getSheetByName(CONFIG.SHEETS.BUKU_IR);
  if (shIR && shIR.getLastRow() > 1) {
    var ir = shIR.getDataRange().getValues(); var ih = headerMap_(ir[0]);
    for (var k = 1; k < ir.length; k++) {
      var t = String(hGet_(ir[k], ih, 'transaksiid', 1) || '');
      if (t) sudah[t] = true;
    }
  }
  var rows = sheet.getDataRange().getValues();
  var h = headerMap_(rows[0]);
  var colTgh = h['rincianditangguhkan'], colAt = h['ditangguhkanat'];
  if (colTgh === undefined) return 0;
  var n = 0, daftar = [];
  var nowS = toDateStr_(new Date());
  for (var i = 1; i < rows.length; i++) {
    var id = String(hGet_(rows[i], h, 'id', 0) || '');
    if (!id || barisDibatalkan_(rows[i], h)) continue;
    if (String(hGet_(rows[i], h, 'periodeid', 1)) !== String(periodeId)) continue;
    if (!bukuIRIds[String(hGet_(rows[i], h, 'jenisid', 2) || '')]) continue;
    if (sudah[id]) continue;                       // sudah dirinci
    if (rincianDitangguhkan_(rows[i], h)) continue; // sudah ditandai
    sheet.getRange(i + 1, colTgh + 1).setValue('Ya');
    if (colAt !== undefined) sheet.getRange(i + 1, colAt + 1).setValue(nowS);
    daftar.push(id + '(' + (Number(hGet_(rows[i], h, 'nominal', 5)) || 0) + ')');
    n++;
  }
  if (n > 0) {
    logActivityWajib_(email, 'PRIVILEGED_TANGGUHKAN_RINCIAN_IR',
      'Periode: ' + periodeId + ' | ' + n + ' transaksi ditangguhkan ke periode berikutnya | ' + daftar.join(', '));
    try { CacheService.getScriptCache().remove('buku_ir_data'); } catch(e) {}
  }
  return n;
}

// Buat baris penyesuaian "Selisih Kas" agar saldo sistem = aktual (T5.7).
// Tanggal = TANGGAL TUTUP BUKU (bukan hari ini), agar tetap di dalam periode.
// selisih > 0 (aktual > sistem) → pemasukan; < 0 → pengeluaran.
function buatPenyesuaianSelisih_(periode, sumberKas, selisih, email, tglTutup) {
  if (!selisih) return;
  var ss = getSS_();
  var tgl = tglTutup || toDateStr_(new Date());
  var nominal = Math.abs(selisih);
  var catatan = 'Penyesuaian selisih kas tutup buku (' + sumberKas + ')';
  var masuk = (selisih > 0);
  var sheet = ss.getSheetByName(masuk ? CONFIG.SHEETS.INPUT_PENERIMAAN : CONFIG.SHEETS.INPUT_PENGELUARAN);
  if (!sheet) return;
  // Tulis berbasis nama kolom — tahan terhadap perubahan urutan kolom.
  var lastCol = sheet.getLastColumn();
  var h = headerMap_(sheet.getRange(1, 1, 1, lastCol).getValues()[0]);
  var baris = new Array(lastCol);
  function set(key, val) { if (h[key] !== undefined) baris[h[key]] = val; }
  set('id', generateID('ADJ'));
  set('periodeid', periode.id);
  set('jenisid', 'SELISIH_KAS');
  set('tanggal', tgl);
  set('nominal', nominal);
  set('sumberkas', sumberKas);
  set('catatan', catatan);
  set('createdby', email);
  set('createdat', tgl);
  set('status', 'Aktif');
  if (!masuk) set('statusapproval', 'Disetujui');  // penyesuaian tutup buku tidak lewat maker-checker
  for (var i = 0; i < lastCol; i++) if (baris[i] === undefined) baris[i] = '';
  sheet.appendRow(baris);
  _isiNoBukti_(sheet, masuk ? 'BKM' : 'BKK', periode.id);
}

// Arsipkan laporan periode ke Drive (T14). Butuh Script Property FOLDER_ARSIP_ID.
// Mengembalikan {fileId, url, hash} atau null bila folder belum dikonfigurasi.
function arsipkanLaporan_(periodeId, periodeNama) {
  try {
    var folderId = PropertiesService.getScriptProperties().getProperty('FOLDER_ARSIP_ID');
    if (!folderId) return null;
    var pdf = generatePDF(periodeId);
    if (!pdf || !pdf.success) return null;
    var html = pdf.html || '';
    var blob = Utilities.newBlob(html, 'text/html', 'laporan.html').getAs('application/pdf')
      .setName('Laporan_' + String(periodeNama).replace(/[^A-Za-z0-9]+/g, '_') + '.pdf');
    var folder = DriveApp.getFolderById(folderId);
    var file = folder.createFile(blob);
    var hashBytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, html, Utilities.Charset.UTF_8);
    return { fileId: file.getId(), url: file.getUrl(), hash: Utilities.base64Encode(hashBytes) };
  } catch(e) { return null; }
}

// URL arsip laporan untuk sebuah periode (dari Saldo Tutup Buku).
function _arsipUrlPeriode_(periodeId) {
  var ss = getSS_();
  var sheet = ss.getSheetByName(CONFIG.SHEETS.SALDO_TUTUP_BUKU);
  if (!sheet || sheet.getLastRow() < 2) return '';
  var rows = sheet.getDataRange().getValues();
  var h = headerMap_(rows[0]);
  for (var i = rows.length - 1; i >= 1; i--) {
    if (String(hGet_(rows[i], h, 'periodeid', 1)) === String(periodeId) && String(hGet_(rows[i], h, 'status', 6)) === 'Tutup') {
      return h['arsipurl'] !== undefined ? String(rows[i][h['arsipurl']] || '') : '';
    }
  }
  return '';
}

function tutupBuku(data) {
  try {
    var auth = requirePerm('periode.manage');
    if (!auth.success) return { success: false, message: auth.message };
    return withLock_(function() {
    var periode = getPeriodeAktif();
    if (!periode) return { success: false, message: 'Tidak ada periode aktif yang bisa ditutup' };

    var ss = getSS_();
    var now = new Date();

    // Opsi: tangguhkan rincian Buku IR yang belum diisi ke periode berikutnya.
    // Uang TETAP di periode ini; hanya rincian (dan kewajiban setor yang timbul
    // darinya) yang dikerjakan di periode berikutnya.
    var jmlTangguh = 0;
    if (data.tangguhkanRincian) {
      jmlTangguh = _tangguhkanRincianBukuIR_(periode.id, auth.user.email);
    }

    // FASE 2: syarat wajib sebelum tutup buku (T5).
    var syarat = cekSyaratTutupBuku_(periode.id);
    if (!syarat.ok) return { success: false, message: syarat.message, bisaTangguhkan: !!syarat.bisaTangguhkan, jumlahBelumDirinci: syarat.jumlahBelumDirinci || 0 };

    // Saldo AKTUAL (hasil cash count) dari pengurus.
    var aktualTunai = Number(data.saldoTunaiAktual) || 0;
    var aktualBank  = Number(data.saldoBankAktual)  || 0;
    if (aktualTunai < 0 || aktualBank < 0) return { success: false, message: 'Saldo aktual tidak boleh negatif.' };

    // Saldo SISTEM dihitung di SERVER (jangan percaya angka klien).
    var sis = calculateSaldo(periode.id, periode);
    var selisihTunai = aktualTunai - sis.tunai;
    var selisihBank  = aktualBank  - sis.bank;
    var selisihTotal = selisihTunai + selisihBank;
    var alasanSelisih = String(data.alasanSelisih || '').trim();
    if (selisihTotal !== 0 && !alasanSelisih) {
      return { success: false, adaSelisih: true, selisihTunai: selisihTunai, selisihBank: selisihBank,
        message: 'Ada selisih kas — Tunai Rp ' + selisihTunai.toLocaleString('id-ID') + ', Bank Rp ' + selisihBank.toLocaleString('id-ID') + '. Wajib isi alasan selisih untuk melanjutkan.' };
    }

    // Buat baris penyesuaian agar saldo sistem = aktual setelah tutup buku (T5.7).
    // Tanggal penyesuaian = tanggal tutup buku (tetap di dalam periode).
    var tglTutup = toDateStr_(now);
    buatPenyesuaianSelisih_(periode, 'Tunai', selisihTunai, auth.user.email, tglTutup);
    buatPenyesuaianSelisih_(periode, 'Bank',  selisihBank,  auth.user.email, tglTutup);

    // Arsip laporan (opsional; aktif bila Script Property FOLDER_ARSIP_ID diset) — T14.
    var arsip = arsipkanLaporan_(periode.id, periode.nama) || {};

    // Catat ke Saldo Tutup Buku (aktual + sistem + selisih + arsip).
    var sheetSaldo = ss.getSheetByName(CONFIG.SHEETS.SALDO_TUTUP_BUKU);
    if (!sheetSaldo) {
      sheetSaldo = ss.insertSheet(CONFIG.SHEETS.SALDO_TUTUP_BUKU);
      sheetSaldo.appendRow(['ID', 'Periode ID', 'Tanggal Tutup', 'Saldo Tunai Akhir', 'Saldo Bank Akhir', 'Total Kas', 'Status', 'Catatan', 'Created By', 'Created At', 'Saldo Tunai Sistem', 'Saldo Bank Sistem', 'Selisih Tunai', 'Selisih Bank', 'Selisih Total', 'Alasan Selisih', 'Arsip File ID', 'Arsip URL', 'Arsip Hash']);
    }
    var sldId = generateID('SLD');
    var saldoTunai = aktualTunai, saldoBank = aktualBank; // saldo awal periode baru = aktual
    sheetSaldo.appendRow([sldId, periode.id, toDateStr_(now), aktualTunai, aktualBank, aktualTunai + aktualBank, 'Tutup', data.catatan || '', auth.user.email, toDateStr_(now),
      sis.tunai, sis.bank, selisihTunai, selisihBank, selisihTotal, alasanSelisih, arsip.fileId || '', arsip.url || '', arsip.hash || '']);

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

    // Otomatis buka periode baru agar SELALU ada periode aktif.
    // Saldo awal periode baru = saldo akhir aktual periode yang ditutup.
    var newPeriode = null;
    if (sheetPeriod) {
      var bulanID = ['Januari','Februari','Maret','April','Mei','Juni','Juli','Agustus','September','Oktober','November','Desember'];
      var namaBaru = (data.namaPeriodeBaru && String(data.namaPeriodeBaru).trim())
        || ('Periode ' + bulanID[now.getMonth()] + ' ' + now.getFullYear());
      var newId = generateID('PER');
      // Urutan kolom sama seperti bukaPeriode:
      // [Periode, Nama, Tgl Mulai, Tgl Tutup, Status, Saldo Awal Tunai, Saldo Awal Bank, Catatan]
      sheetPeriod.appendRow([newId, namaBaru, toDateStr_(now), '', CONFIG.STATUS.OPEN,
        saldoTunai, saldoBank, 'Lanjutan dari ' + periode.nama]);
      newPeriode = { id: newId, nama: namaBaru, saldoAwalTunai: saldoTunai, saldoAwalBank: saldoBank };
    }

    try {
      var c = CacheService.getScriptCache();
      c.remove('dashboard_saldo');
      c.remove('master_trx_data');
    } catch(e) {}

    logActivityWajib_(auth.user.email, 'TUTUP_BUKU', 'Tutup: ' + periode.nama + ' | Aktual T/B: ' + aktualTunai + '/' + aktualBank + ' | Sistem T/B: ' + sis.tunai + '/' + sis.bank + ' | Selisih: ' + selisihTotal + (alasanSelisih ? ' | Alasan: ' + alasanSelisih : '') + (newPeriode ? ' | Buka: ' + newPeriode.nama : ''));
    return { success: true, id: sldId, selisihTunai: selisihTunai, selisihBank: selisihBank, newPeriode: newPeriode, arsip: arsip.url || '' };
    });
  } catch(e) {
    return { success: false, message: e.message };
  }
}

// Saldo akhir (aktual) periode yang paling terakhir ditutup — untuk rollforward.
function saldoAkhirTerakhir_() {
  var ss = getSS_();
  var sheet = ss.getSheetByName(CONFIG.SHEETS.SALDO_TUTUP_BUKU);
  if (!sheet || sheet.getLastRow() < 2) return null;
  var rows = sheet.getDataRange().getValues();
  var h = headerMap_(rows[0]);
  for (var i = rows.length - 1; i >= 1; i--) {
    if (String(hGet_(rows[i], h, 'status', 6)) === 'Tutup') {
      return {
        periodeId: String(hGet_(rows[i], h, 'periodeid', 1) || ''),
        tunai: Number(hGet_(rows[i], h, 'saldotunaiakhir', 3)) || 0,
        bank: Number(hGet_(rows[i], h, 'saldobankakhir', 4)) || 0
      };
    }
  }
  return null;
}

// Info untuk form Buka Periode: saldo awal yang WAJIB (rollforward dari tutup buku
// terakhir). adaPrev=false → periode pertama, saldo awal bebas diisi.
function getInfoBukaPeriode() {
  var auth = requirePerm('periode.manage');
  if (!auth.success) return { success: false, message: auth.message };
  try {
    var prev = saldoAkhirTerakhir_();
    // Cek apakah masih ada periode OPEN (tak boleh buka dua-duanya).
    var adaOpen = false;
    var ss = getSS_();
    var sp = ss.getSheetByName(CONFIG.SHEETS.PERIOD);
    if (sp && sp.getLastRow() > 1) {
      var pr = sp.getDataRange().getValues(); var ph = headerMap_(pr[0]);
      for (var i = 1; i < pr.length; i++) {
        if (String(hGet_(pr[i], ph, 'status', 4)).trim() === CONFIG.STATUS.OPEN) { adaOpen = true; break; }
      }
    }
    // Tanggal mulai saran = sehari setelah tanggal tutup periode lalu (biar berkesinambungan).
    var tglMulaiSaran = '';
    var tanggalTutupTerakhir = '';
    try {
      var akhir = getSaldoTutupBukuTerakhir_();
      if (akhir && /^\d{4}-\d{2}-\d{2}$/.test(akhir.tanggalTutup)) {
        tanggalTutupTerakhir = akhir.tanggalTutup;
        var d = new Date(akhir.tanggalTutup);
        d.setDate(d.getDate() + 1);
        tglMulaiSaran = toDateStr_(d);
      }
    } catch(e) {}
    return {
      success: true,
      adaOpen: adaOpen,
      adaPrev: !!prev,
      rollforwardTunai: prev ? prev.tunai : 0,
      rollforwardBank: prev ? prev.bank : 0,
      tanggalTutupTerakhir: tanggalTutupTerakhir,
      tglMulaiSaran: tglMulaiSaran
    };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

// Buka kembali periode yang sudah CLOSED (mis. salah tutup buku). PRIVILEGED — ADMIN.
// Periode OPEN saat ini otomatis ditutup (hanya boleh 1 periode OPEN). Wajib alasan +
// ketik ulang nama periode untuk konfirmasi. Arsip tutup buku periode itu ditandai
// 'Dibuka Kembali' agar tak lagi dianggap sebagai penutupan.
function bukaKembaliPeriode(periodeId, alasan, konfirmasiNama) {
  var auth = requirePerm('periode.manage');
  if (!auth.success) return { success: false, message: auth.message };
  if (!periodeId) return { success: false, message: 'ID periode wajib.' };
  var alasanBersih = String(alasan || '').trim();
  if (alasanBersih.length < 10) return { success: false, message: 'Alasan wajib diisi (minimal 10 karakter).' };
  return withLock_(function() {
    var ss = getSS_();
    var sheet = ss.getSheetByName(CONFIG.SHEETS.PERIOD);
    if (!sheet || sheet.getLastRow() < 2) return { success: false, message: 'Master Period tidak ditemukan.' };
    var rows = sheet.getDataRange().getValues();
    var h = headerMap_(rows[0]);
    var colStatus  = (h['status']   !== undefined ? h['status']   : 4) + 1;
    var colTutup   = (h['tgltutup'] !== undefined ? h['tgltutup'] : 3) + 1;
    var colCatatan = (h['catatan']  !== undefined ? h['catatan']  : 7) + 1;
    var targetRow = -1, targetNama = '', targetStatus = '';
    var openRows = [];
    for (var i = 1; i < rows.length; i++) {
      var pid = String(hGet_(rows[i], h, 'periode', 0) || rows[i][0]);
      var st = String(hGet_(rows[i], h, 'status', 4) || '').toUpperCase();
      if (pid === String(periodeId)) { targetRow = i + 1; targetNama = String(hGet_(rows[i], h, 'nama', 1) || ''); targetStatus = st; }
      else if (st === CONFIG.STATUS.OPEN) openRows.push({ row: i + 1, nama: String(hGet_(rows[i], h, 'nama', 1) || '') });
    }
    if (targetRow < 0) return { success: false, message: 'Periode tidak ditemukan.' };
    if (targetStatus !== CONFIG.STATUS.CLOSED) return { success: false, message: 'Hanya periode CLOSED yang bisa dibuka kembali (status saat ini: ' + targetStatus + ').' };
    // Konfirmasi: ketik ulang nama periode.
    if (String(konfirmasiNama || '').trim() !== targetNama) {
      return { success: false, butuhKonfirmasiNama: true, nama: targetNama,
        message: 'Buka kembali periode "' + targetNama + '"? Periode aktif saat ini akan otomatis ditutup. Ketik ulang nama periode untuk konfirmasi.' };
    }
    var now = new Date();
    // Tutup periode OPEN saat ini (otomatis).
    var ditutup = [];
    for (var o = 0; o < openRows.length; o++) {
      sheet.getRange(openRows[o].row, colStatus).setValue(CONFIG.STATUS.CLOSED);
      sheet.getRange(openRows[o].row, colTutup).setValue(toDateStr_(now));
      var catLama = String(sheet.getRange(openRows[o].row, colCatatan).getValue() || '');
      sheet.getRange(openRows[o].row, colCatatan).setValue((catLama ? catLama + ' | ' : '') + 'Ditutup otomatis: periode ' + targetNama + ' dibuka kembali');
      ditutup.push(openRows[o].nama);
    }
    // Buka kembali target: status OPEN, kosongkan Tgl Tutup.
    sheet.getRange(targetRow, colStatus).setValue(CONFIG.STATUS.OPEN);
    sheet.getRange(targetRow, colTutup).setValue('');
    // Tandai arsip tutup buku periode target agar tak lagi dianggap 'Tutup'.
    var arsipDitandai = 0;
    var shT = ss.getSheetByName(CONFIG.SHEETS.SALDO_TUTUP_BUKU);
    if (shT && shT.getLastRow() > 1) {
      var tr = shT.getDataRange().getValues(); var th = headerMap_(tr[0]);
      var colTS = (th['status'] !== undefined ? th['status'] : 6) + 1;
      for (var t = 1; t < tr.length; t++) {
        if (String(hGet_(tr[t], th, 'periodeid', 1)) === String(periodeId) && String(hGet_(tr[t], th, 'status', 6)) === 'Tutup') {
          shT.getRange(t + 1, colTS).setValue('Dibuka Kembali'); arsipDitandai++;
        }
      }
    }
    logActivityWajib_(auth.user.email, 'PRIVILEGED_BUKA_KEMBALI_PERIODE',
      'Periode: ' + targetNama + ' | ALASAN: ' + alasanBersih + ' | Ditutup otomatis: ' + (ditutup.join(', ') || '-') + ' | Arsip ditandai: ' + arsipDitandai);
    try { var c = CacheService.getScriptCache(); c.remove('dashboard_saldo'); c.remove('buku_ir_data'); c.remove('master_trx_data'); } catch(e) {}
    return { success: true, dibukaKembali: targetNama, ditutup: ditutup };
  });
}

function bukaPeriode(data) {
  try {
    var auth = requirePerm('periode.manage');
    if (!auth.success) return { success: false, message: auth.message };
    return withLock_(function() {
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

    // FASE 2: rollforward terkunci (T1). Saldo awal = saldo akhir periode CLOSED
    // terakhir. Secara default kita LANGSUNG memakai nilai rollforward (tidak
    // membandingkan input agar tidak gagal karena beda pembulatan/format). Hanya
    // bila pengurus SENGAJA menyatakan "Ada selisih" (override + alasan) input
    // dipakai & selisihnya dicatat sebagai penyesuaian eksplisit.
    var prev = saldoAkhirTerakhir_();
    var awalTunai, awalBank, adjust = null;
    if (prev) {
      var mauOverride = !!data.overrideRollforward;
      if (mauOverride) {
        if (!String(data.alasanRollforward || '').trim()) {
          return { success: false, message: 'Untuk saldo awal berbeda dari tutup buku, alasan selisih wajib diisi.' };
        }
        var claimT = Number(data.saldoAwalTunai), claimB = Number(data.saldoAwalBank);
        awalTunai = isFinite(claimT) ? claimT : prev.tunai;
        awalBank  = isFinite(claimB) ? claimB : prev.bank;
        var difT = Math.round(awalTunai) - Math.round(prev.tunai);
        var difB = Math.round(awalBank) - Math.round(prev.bank);
        if (difT !== 0 || difB !== 0) {
          adjust = { tunai: difT, bank: difB, alasan: String(data.alasanRollforward).trim() };
        }
      } else {
        // Default: kunci ke rollforward. Input diabaikan → selalu bisa disimpan.
        awalTunai = prev.tunai; awalBank = prev.bank;
      }
    } else {
      awalTunai = Number(data.saldoAwalTunai) || 0;
      awalBank  = Number(data.saldoAwalBank)  || 0;
    }

    var id = generateID('PER');
    sheetPeriod.appendRow([
      id,
      data.nama,
      toDateStr_(new Date(data.tglMulai)),
      '',
      CONFIG.STATUS.OPEN,
      awalTunai,
      awalBank,
      data.catatan || ''
    ]);

    // Penyesuaian rollforward (bila override) dicatat sebagai transaksi eksplisit.
    if (adjust) {
      var tglMulaiStr = toDateStr_(new Date(data.tglMulai));
      buatPenyesuaianSelisih_({ id: id }, 'Tunai', adjust.tunai, auth.user.email, tglMulaiStr);
      buatPenyesuaianSelisih_({ id: id }, 'Bank',  adjust.bank,  auth.user.email, tglMulaiStr);
      logActivityWajib_(auth.user.email, 'ROLLFORWARD_OVERRIDE', 'Periode: ' + data.nama + ' | Selisih T/B: ' + adjust.tunai + '/' + adjust.bank + ' | Alasan: ' + adjust.alasan);
    }

    try { CacheService.getScriptCache().remove('dashboard_saldo'); } catch(e) {}
    logActivity(auth.user.email, 'BUKA_PERIODE', 'Periode: ' + data.nama);
    return { success: true, id: id };
    });
  } catch(e) {
    return { success: false, message: e.message };
  }
}

// ──────────────────────────────────────────────────────
// TRANSAKSI
// ──────────────────────────────────────────────────────
function submitTransaksi(data) {
  try {
    // FASE 5.1: enforcement per-arah (masuk vs keluar/mutasi) dgn fallback ke 'trx.input'.
    var _tipeInput = (data && data.tipe === 'masuk') ? 'masuk' : 'keluar';
    var auth = requirePermInput_(_tipeInput);
    if (!auth.success) return { success: false, message: auth.message };
    return withLock_(function() {
    var ss = getSS_();
    var periode = getPeriodeAktif();
    if (!periode) return { success: false, message: 'Tidak ada periode aktif' };

    var id = generateID('TRX');
    var now = new Date();

    // FASE 1: validasi nominal, tanggal, sumber kas / arah (server-side, T4/T8/A3).
    var nomCek = validasiNominal_(data.nominal);
    if (!nomCek.ok) return { success: false, message: nomCek.message };
    var nominal = nomCek.nilai;
    if (data.tipe === 'masuk' || data.tipe === 'keluar') {
      if (data.sumberKas !== 'Tunai' && data.sumberKas !== 'Bank') return { success: false, message: 'Sumber Kas harus Tunai atau Bank.' };
    } else if (data.tipe === 'mutasi') {
      if (data.arah !== 'setor' && data.arah !== 'tarik') return { success: false, message: 'Arah mutasi harus setor atau tarik.' };
    }
    var tglCek = validasiTanggalPeriode_(data.tanggal ? data.tanggal : now, periode.id);
    if (!tglCek.ok) return { success: false, message: tglCek.message };
    var tgl = tglCek.tgl;

    // T10: pengeluaran di atas ambang wajib ada bukti (lampiran).
    // SEMENTARA DINONAKTIFKAN: hanya diwajibkan bila Script Property WAJIB_BUKTI = 'true'.
    // Default (properti kosong/'false') → bukti bersifat OPSIONAL (tak menolak simpan).
    if (data.tipe === 'keluar') {
      var wajibBukti = false;
      try { wajibBukti = String(PropertiesService.getScriptProperties().getProperty('WAJIB_BUKTI') || '').toLowerCase() === 'true'; } catch(e) {}
      if (wajibBukti) {
        var ambang = 500000;
        try { var av = PropertiesService.getScriptProperties().getProperty('AMBANG_BUKTI'); if (av) ambang = Number(av) || 500000; } catch(e) {}
        var jmlBukti = (data.buktiList && data.buktiList.length) ? data.buktiList.length : 0;
        if (nominal > ambang && jmlBukti === 0) {
          return { success: false, butuhBukti: true, message: 'Pengeluaran di atas Rp ' + ambang.toLocaleString('id-ID') + ' wajib melampirkan minimal 1 bukti (foto nota).' };
        }
      }
    }

    // Guard saldo negatif (T13) untuk pengeluaran & mutasi. Override khusus ADMIN + alasan wajib.
    var override = !!data.override && auth.user.role === CONFIG.ROLES.ADMIN;
    if (override && !String(data.alasanOverride || '').trim()) return { success: false, message: 'Override saldo wajib disertai alasan.' };
    if (!override) {
      if (data.tipe === 'keluar') {
        var ck = cekSaldoCukup_(periode, data.sumberKas, nominal);
        if (!ck.ok) return { success: false, message: ck.message, saldoKurang: true };
      } else if (data.tipe === 'mutasi') {
        var ckm = cekSaldoCukup_(periode, (data.arah === 'setor') ? 'Tunai' : 'Bank', nominal);
        if (!ckm.ok) return { success: false, message: ckm.message, saldoKurang: true };
      }
    } else {
      logActivityWajib_(auth.user.email, 'OVERRIDE_SALDO', 'Alasan: ' + data.alasanOverride + ' | tipe: ' + data.tipe + ' | nominal: ' + nominal);
    }

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
        kpSheet.appendRow([id, periode.id, tgl, data.jenisId, data.anggotaId || '', nominal, data.sumberKas || 'Tunai', data.catatan || '', auth.user.email, 'Aktif', '', toDateStr_(now)]);
        logActivity(auth.user.email, 'KAS_PENEROBOS_VIA_TRX', 'Nominal: ' + nominal);
        try { CacheService.getScriptCache().remove('dashboard_saldo'); } catch(e) {}
        return { success: true, id: id, viaPenerobos: true };
      }
      var sheet = ss.getSheetByName(CONFIG.SHEETS.INPUT_PENERIMAAN);
      if (!sheet) return { success: false, message: 'Sheet penerimaan tidak ditemukan' };
      sheet.appendRow([id, periode.id, data.jenisId, data.anggotaId || '', tgl, nominal, data.sumberKas, data.catatan || '', auth.user.email, toDateStr_(now), 'Aktif']);
      _isiNoBukti_(sheet, 'BKM', periode.id);
      logActivity(auth.user.email, 'PEMASUKAN', 'Nominal: ' + nominal);
    } else if (data.tipe === 'keluar') {
      var sheet = ss.getSheetByName(CONFIG.SHEETS.INPUT_PENGELUARAN);
      if (!sheet) return { success: false, message: 'Sheet pengeluaran tidak ditemukan' };
      sheet.appendRow([id, periode.id, data.jenisId, tgl, nominal, data.sumberKas, data.catatan || '', auth.user.email, toDateStr_(now), 'Aktif']);
      _isiNoBukti_(sheet, 'BKK', periode.id);
      if (data.buktiList && data.buktiList.length) { try { _simpanBuktiList_(id, 'keluar', data.buktiList, auth.user.email); } catch(e) {} }
      // T12 maker-checker: pengeluaran di atas ambang jadi Draft (belum masuk saldo).
      var ambangApp = 1000000;
      try { var aa = PropertiesService.getScriptProperties().getProperty('AMBANG_APPROVAL'); if (aa) ambangApp = Number(aa) || 1000000; } catch(e) {}
      var perluApproval = nominal > ambangApp;
      _isiApproval_(sheet, perluApproval ? 'Draft' : 'Disetujui', perluApproval ? '' : auth.user.email, perluApproval ? '' : toDateStr_(now));
      logActivity(auth.user.email, 'PENGELUARAN', 'Nominal: ' + nominal + (perluApproval ? ' [DRAFT menunggu persetujuan]' : ''));
      try { CacheService.getScriptCache().remove('dashboard_saldo'); } catch(e) {}
      return { success: true, id: id, perluApproval: perluApproval };
    } else if (data.tipe === 'mutasi') {
      var sheet = ss.getSheetByName(CONFIG.SHEETS.INPUT_SETORAN);
      if (!sheet) {
        sheet = ss.insertSheet(CONFIG.SHEETS.INPUT_SETORAN);
        sheet.appendRow(['ID', 'Periode ID', 'Tanggal', 'Nominal', 'Arah', 'Created By', 'Created At', 'Status', 'Dibatalkan By', 'Dibatalkan At', 'Alasan Batal']);
      }
      sheet.appendRow([id, periode.id, tgl, nominal, data.arah, auth.user.email, toDateStr_(now), 'Aktif']);
      logActivity(auth.user.email, 'MUTASI', 'Arah: ' + data.arah + ' Nominal: ' + nominal);
    }

    // Invalidate saldo cache setiap ada transaksi baru
    try { CacheService.getScriptCache().remove('dashboard_saldo'); } catch(e) {}
    return { success: true, id: id };
    });
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function updateTransaksi(data) {
  try {
    var cap = (data.sumberKas === 'Tunai') ? 'trx.edit.tunai' : 'trx.edit.bank';
    var auth = requirePerm(cap);
    if (!auth.success) return { success: false, message: auth.message };
    return withLock_(function() {
    if (data.sumberKas !== 'Tunai' && data.sumberKas !== 'Bank') return { success: false, message: 'Sumber Kas harus Tunai atau Bank.' };
    var nomCek = validasiNominal_(data.nominal);
    if (!nomCek.ok) return { success: false, message: nomCek.message };
    var nominal = nomCek.nilai;
    var masuk = (data.tipe === 'masuk');
    if (!masuk && data.tipe !== 'keluar') return { success: false, message: 'Tipe transaksi tidak valid' };

    var ss = getSS_();
    var sheet = ss.getSheetByName(masuk ? CONFIG.SHEETS.INPUT_PENERIMAAN : CONFIG.SHEETS.INPUT_PENGELUARAN);
    if (!sheet) return { success: false, message: 'Sheet tidak ditemukan' };
    var rows = sheet.getDataRange().getValues();
    var h = headerMap_(rows[0]);
    var iTgl = masuk ? 4 : 3, iNom = masuk ? 5 : 4, iSmb = masuk ? 6 : 5, iCat = masuk ? 7 : 6;
    for (var i = 1; i < rows.length; i++) {
      if (String(hGet_(rows[i], h, 'id', 0)) !== String(data.id)) continue;
      if (barisDibatalkan_(rows[i], h)) return { success: false, message: 'Transaksi sudah dibatalkan, tidak dapat diedit.' };
      var pidBaris = String(hGet_(rows[i], h, 'periodeid', 1) || '');
      var ap = assertPeriodeOpen_(pidBaris); // T2: kunci periode tertutup
      if (!ap.ok) return { success: false, message: ap.message };
      var tglCek = validasiTanggalPeriode_(data.tanggal, pidBaris);
      if (!tglCek.ok) return { success: false, message: tglCek.message };

      var lama = { jenis: String(hGet_(rows[i], h, 'jenisid', 2) || ''), tanggal: toDateStr_(hGet_(rows[i], h, 'tanggal', iTgl)), nominal: Number(hGet_(rows[i], h, 'nominal', iNom)) || 0, sumberKas: String(hGet_(rows[i], h, 'sumberkas', iSmb) || '') };
      var rowNum = i + 1;
      sheet.getRange(rowNum, (h['jenisid'] !== undefined ? h['jenisid'] : 2) + 1).setValue(data.jenisId);
      if (masuk) sheet.getRange(rowNum, (h['anggotaid'] !== undefined ? h['anggotaid'] : 3) + 1).setValue(data.anggotaId || '');
      sheet.getRange(rowNum, (h['tanggal'] !== undefined ? h['tanggal'] : iTgl) + 1).setValue(tglCek.tgl);
      sheet.getRange(rowNum, (h['nominal'] !== undefined ? h['nominal'] : iNom) + 1).setValue(nominal);
      sheet.getRange(rowNum, (h['sumberkas'] !== undefined ? h['sumberkas'] : iSmb) + 1).setValue(data.sumberKas);
      sheet.getRange(rowNum, (h['catatan'] !== undefined ? h['catatan'] : iCat) + 1).setValue(data.catatan || '');
      try { CacheService.getScriptCache().remove('dashboard_saldo'); } catch(e) {}
      logActivityWajib_(auth.user.email, 'EDIT_' + (masuk ? 'PEMASUKAN' : 'PENGELUARAN'),
        'ID: ' + data.id + ' | LAMA: ' + JSON.stringify(lama) + ' | BARU: ' + JSON.stringify({ jenis: data.jenisId, tanggal: tglCek.tgl, nominal: nominal, sumberKas: data.sumberKas }));
      return { success: true };
    }
    return { success: false, message: 'Transaksi tidak ditemukan' };
    });
  } catch(e) {
    return { success: false, message: e.message };
  }
}

// SOFT DELETE (T3): tidak menghapus baris; menandai Status=Dibatalkan +
// mewajibkan alasan + mencatat snapshot nilai lama di log. Menolak baris
// dari periode yang sudah ditutup (T2).
function deleteTransaksi(data) {
  try {
    var cap = (data.sumberKas === 'Tunai') ? 'trx.edit.tunai' : 'trx.edit.bank';
    var auth = requirePerm(cap);
    if (!auth.success) return { success: false, message: auth.message };
    var alasan = String(data.alasan || '').trim();
    if (!alasan) return { success: false, message: 'Alasan pembatalan wajib diisi.' };
    return withLock_(function() {
    var masuk = (data.tipe === 'masuk');
    var ss = getSS_();
    var sheet = ss.getSheetByName(masuk ? CONFIG.SHEETS.INPUT_PENERIMAAN : CONFIG.SHEETS.INPUT_PENGELUARAN);
    if (!sheet) return { success: false, message: 'Sheet tidak ditemukan' };
    var rows = sheet.getDataRange().getValues();
    var h = headerMap_(rows[0]);
    if (h['status'] === undefined) return { success: false, message: 'Kolom Status belum ada. Jalankan migrasiPengendalian dulu.' };
    var iTgl = masuk ? 4 : 3, iNom = masuk ? 5 : 4, iSmb = masuk ? 6 : 5;
    for (var i = 1; i < rows.length; i++) {
      if (String(hGet_(rows[i], h, 'id', 0)) !== String(data.id)) continue;
      if (barisDibatalkan_(rows[i], h)) return { success: false, message: 'Transaksi sudah dibatalkan.' };
      var pidBaris = String(hGet_(rows[i], h, 'periodeid', 1) || '');
      var ap = assertPeriodeOpen_(pidBaris); // T2
      if (!ap.ok) return { success: false, message: ap.message };

      var snap = { jenis: String(hGet_(rows[i], h, 'jenisid', 2) || ''), tanggal: toDateStr_(hGet_(rows[i], h, 'tanggal', iTgl)), nominal: Number(hGet_(rows[i], h, 'nominal', iNom)) || 0, sumberKas: String(hGet_(rows[i], h, 'sumberkas', iSmb) || '') };
      var rowNum = i + 1;
      var nowS = toDateStr_(new Date());
      sheet.getRange(rowNum, h['status'] + 1).setValue('Dibatalkan');
      if (h['dibatalkanby'] !== undefined) sheet.getRange(rowNum, h['dibatalkanby'] + 1).setValue(auth.user.email);
      if (h['dibatalkanat'] !== undefined) sheet.getRange(rowNum, h['dibatalkanat'] + 1).setValue(nowS);
      if (h['alasanbatal'] !== undefined) sheet.getRange(rowNum, h['alasanbatal'] + 1).setValue(alasan);
      // K2: buang juga cache Buku IR & master agar pembatalan penerimaan langsung terlihat.
      try { var cd = CacheService.getScriptCache(); cd.remove('dashboard_saldo'); cd.remove('buku_ir_data'); cd.remove('master_trx_data'); } catch(e) {}
      logActivityWajib_(auth.user.email, 'BATAL_' + (masuk ? 'PEMASUKAN' : 'PENGELUARAN'),
        'ID: ' + data.id + ' | ALASAN: ' + alasan + ' | NILAI: ' + JSON.stringify(snap));
      return { success: true };
    }
    return { success: false, message: 'Transaksi tidak ditemukan' };
    });
  } catch(e) {
    return { success: false, message: e.message };
  }
}

// FASE 5 (T12) — maker-checker: setujui pengeluaran Draft. Penyetuju ≠ pembuat.
function setujuiPengeluaran(id) {
  try {
    var auth = requirePerm('trx.approve');
    if (!auth.success) return { success: false, message: auth.message };
    return withLock_(function() {
    var ss = getSS_();
    var sheet = ss.getSheetByName(CONFIG.SHEETS.INPUT_PENGELUARAN);
    if (!sheet) return { success: false, message: 'Sheet tidak ditemukan' };
    var rows = sheet.getDataRange().getValues();
    var h = headerMap_(rows[0]);
    if (h['statusapproval'] === undefined) return { success: false, message: 'Kolom approval belum ada. Jalankan migrasiPengendalian.' };
    for (var i = 1; i < rows.length; i++) {
      if (String(hGet_(rows[i], h, 'id', 0)) !== String(id)) continue;
      if (barisDibatalkan_(rows[i], h)) return { success: false, message: 'Transaksi sudah dibatalkan.' };
      var ap = assertPeriodeOpen_(String(hGet_(rows[i], h, 'periodeid', 1) || ''));
      if (!ap.ok) return { success: false, message: ap.message };
      var pembuat = String(hGet_(rows[i], h, 'createdby', 7) || '').toLowerCase();
      if (pembuat === String(auth.user.email).toLowerCase()) return { success: false, message: 'Anda tidak boleh menyetujui pengeluaran yang Anda buat sendiri.' };
      if (String(rows[i][h['statusapproval']] || '').toLowerCase() !== 'draft') return { success: false, message: 'Pengeluaran ini tidak berstatus Draft.' };
      var rowNum = i + 1, now = toDateStr_(new Date());
      sheet.getRange(rowNum, h['statusapproval'] + 1).setValue('Disetujui');
      if (h['disetujuiby'] !== undefined) sheet.getRange(rowNum, h['disetujuiby'] + 1).setValue(auth.user.email);
      if (h['disetujuiat'] !== undefined) sheet.getRange(rowNum, h['disetujuiat'] + 1).setValue(now);
      try { CacheService.getScriptCache().remove('dashboard_saldo'); } catch(e) {}
      logActivityWajib_(auth.user.email, 'SETUJUI_PENGELUARAN', 'ID: ' + id + ' | Pembuat: ' + pembuat);
      return { success: true };
    }
    return { success: false, message: 'Transaksi tidak ditemukan' };
    });
  } catch(e) { return { success: false, message: e.message }; }
}

function getDraftPengeluaran() {
  try {
    var auth = checkAuth(); if (!auth.success) return auth;
    var ss = getSS_(); var periode = getPeriodeAktif();
    var sheet = ss.getSheetByName(CONFIG.SHEETS.INPUT_PENGELUARAN);
    if (!sheet || sheet.getLastRow() < 2) return { success: true, data: [], total: 0 };
    var namaKeluar = {};
    var mk = ss.getSheetByName(CONFIG.SHEETS.PENGELUARAN);
    if (mk) { var mr = mk.getDataRange().getValues(); for (var j = 1; j < mr.length; j++) { if (mr[j][0]) namaKeluar[String(mr[j][0])] = String(mr[j][1] || ''); } }
    var rows = sheet.getDataRange().getValues(); var h = headerMap_(rows[0]);
    var out = [], total = 0;
    for (var i = 1; i < rows.length; i++) {
      if (barisDibatalkan_(rows[i], h) || !barisDraft_(rows[i], h)) continue;
      if (periode && String(hGet_(rows[i], h, 'periodeid', 1)) !== periode.id) continue;
      var nom = Number(hGet_(rows[i], h, 'nominal', 4)) || 0; total += nom;
      out.push({ id: String(hGet_(rows[i], h, 'id', 0)), jenis: namaKeluar[String(hGet_(rows[i], h, 'jenisid', 2))] || '', tanggal: toDateStr_(hGet_(rows[i], h, 'tanggal', 3)), nominal: nom, sumberKas: String(hGet_(rows[i], h, 'sumberkas', 5) || ''), catatan: String(hGet_(rows[i], h, 'catatan', 6) || ''), pembuat: String(hGet_(rows[i], h, 'createdby', 7) || '') });
    }
    return { success: true, data: out, total: total };
  } catch(e) { return { success: false, message: e.message }; }
}

// ──────────────────────────────────────────────────────
// IMPOR CSV — banyak transaksi bank sekaligus (mis. rekening koran)
// payload = { jenisMasukId, jenisKeluarId, sumberKas, rows:[{tanggal,tipe,nominal,keterangan}] }
// Semua baris sumber kas sama (default 'Bank'); tiap baris memakai jenis
// default sesuai tipe. Dibungkus withLock_ agar atomik.
// ──────────────────────────────────────────────────────
function importTransaksiCSV(payload) {
  try {
    // FASE 5.1: enforcement per-arah dilakukan setelah tahu isi baris (masuk/keluar).
    var auth = checkAuth();
    if (!auth.success) return { success: false, message: auth.message };
    payload = payload || {};
    var rows = payload.rows || [];
    if (!rows.length) return { success: false, message: 'Tidak ada baris untuk diimpor.' };
    if (rows.length > 500) return { success: false, message: 'Maksimum 500 baris per impor.' };

    var ss = getSS_();
    var periode = getPeriodeAktif();
    if (!periode) return { success: false, message: 'Tidak ada periode aktif.' };

    var sumberKas = (String(payload.sumberKas || 'Bank') === 'Tunai') ? 'Tunai' : 'Bank';
    var jenisMasukId = String(payload.jenisMasukId || '');
    var jenisKeluarId = String(payload.jenisKeluarId || '');

    // Validasi FK jenis (bila dipakai)
    function jenisValid_(sheetName, id) {
      if (!id) return false;
      var sh = ss.getSheetByName(sheetName);
      if (!sh) return false;
      var r = sh.getDataRange().getValues();
      for (var i = 1; i < r.length; i++) { if (String(r[i][0]) === id) return true; }
      return false;
    }
    var adaMasuk = rows.some(function(x) { return String(x.tipe).toLowerCase() === 'masuk'; });
    var adaKeluar = rows.some(function(x) { return String(x.tipe).toLowerCase() === 'keluar'; });
    // FASE 5.1: cek izin per-arah sesuai isi baris CSV.
    if (adaMasuk && !userCanInput_(auth.user.role, 'masuk')) return { success: false, message: 'Akses ditolak: tidak berwenang input pemasukan.' };
    if (adaKeluar && !userCanInput_(auth.user.role, 'keluar')) return { success: false, message: 'Akses ditolak: tidak berwenang input pengeluaran.' };
    if (adaMasuk && !jenisValid_(CONFIG.SHEETS.PEMASUKAN, jenisMasukId)) return { success: false, message: 'Pilih Jenis Pemasukan yang valid (ada baris "masuk").' };
    if (adaKeluar && !jenisValid_(CONFIG.SHEETS.PENGELUARAN, jenisKeluarId)) return { success: false, message: 'Pilih Jenis Pengeluaran yang valid (ada baris "keluar").' };

    return withLock_(function() {
      var shIn = ss.getSheetByName(CONFIG.SHEETS.INPUT_PENERIMAAN);
      var shOut = ss.getSheetByName(CONFIG.SHEETS.INPUT_PENGELUARAN);
      var now = new Date();
      var berhasil = 0; var gagal = [];
      for (var i = 0; i < rows.length; i++) {
        var row = rows[i] || {};
        var tipe = String(row.tipe || '').toLowerCase().trim();
        var nominal = Number(String(row.nominal).replace(/[^0-9]/g, '')) || 0;
        var tgl = toDateStr_(row.tanggal ? new Date(row.tanggal) : now);
        var ket = String(row.keterangan || '');
        if (tipe !== 'masuk' && tipe !== 'keluar') { gagal.push({ baris: i + 1, alasan: 'Tipe harus masuk/keluar' }); continue; }
        var nc = validasiNominal_(nominal);
        if (!nc.ok) { gagal.push({ baris: i + 1, alasan: nc.message }); continue; }
        var tc = validasiTanggalPeriode_(row.tanggal ? row.tanggal : now, periode.id);
        if (!tc.ok) { gagal.push({ baris: i + 1, alasan: tc.message }); continue; }
        tgl = tc.tgl;
        var id = generateID('TRX');
        if (tipe === 'masuk') {
          if (!shIn) { gagal.push({ baris: i + 1, alasan: 'Sheet penerimaan tidak ada' }); continue; }
          shIn.appendRow([id, periode.id, jenisMasukId, '', tgl, nominal, sumberKas, ket, auth.user.email, toDateStr_(now), 'Aktif']);
          _isiNoBukti_(shIn, 'BKM', periode.id);
        } else {
          if (!shOut) { gagal.push({ baris: i + 1, alasan: 'Sheet pengeluaran tidak ada' }); continue; }
          shOut.appendRow([id, periode.id, jenisKeluarId, tgl, nominal, sumberKas, ket, auth.user.email, toDateStr_(now), 'Aktif']);
          _isiNoBukti_(shOut, 'BKK', periode.id);
        }
        berhasil++;
      }
      try { var c = CacheService.getScriptCache(); c.remove('dashboard_saldo'); c.remove('master_trx_data'); } catch(e) {}
      logActivity(auth.user.email, 'IMPOR_CSV', 'Berhasil: ' + berhasil + ', Gagal: ' + gagal.length);
      return { success: true, berhasil: berhasil, gagal: gagal, total: rows.length };
    });
  } catch(e) {
    return { success: false, message: e.message };
  }
}

// ──────────────────────────────────────────────────────
// RIWAYAT TRANSAKSI SAYA — transaksi yang diinput user login
// Mengembalikan gabungan pemasukan/pengeluaran/mutasi/kas penerobos
// milik user (berdasarkan Created By = email), urut terbaru dulu.
// ──────────────────────────────────────────────────────
function getRiwayatTransaksiSaya() {
  try {
    var auth = checkAuth();
    if (!auth.success) return { success: false, message: auth.message };
    var email = auth.user.email;
    var ss = getSS_();

    // Peta nama jenis pemasukan & pengeluaran (id -> nama)
    var namaMasuk = {}, namaKeluar = {};
    var sheetP = ss.getSheetByName(CONFIG.SHEETS.PEMASUKAN);
    if (sheetP) {
      var rp = sheetP.getDataRange().getValues();
      for (var i = 1; i < rp.length; i++) { if (rp[i][0]) namaMasuk[String(rp[i][0])] = String(rp[i][1] || ''); }
    }
    var sheetPK = ss.getSheetByName(CONFIG.SHEETS.PENGELUARAN);
    if (sheetPK) {
      var rpk = sheetPK.getDataRange().getValues();
      for (var i = 1; i < rpk.length; i++) { if (rpk[i][0]) namaKeluar[String(rpk[i][0])] = String(rpk[i][1] || ''); }
    }

    var list = [];

    // Pemasukan (Input Penerimaan)
    var shIn = ss.getSheetByName(CONFIG.SHEETS.INPUT_PENERIMAAN);
    if (shIn) {
      var rin = shIn.getDataRange().getValues();
      var hin = headerMap_(rin[0]);
      for (var i = 1; i < rin.length; i++) {
        if (!hGet_(rin[i], hin, 'id', 0)) continue;
        if (!emailSama_(hGet_(rin[i], hin, 'createdby', 8), email)) continue;
        if (barisDibatalkan_(rin[i], hin)) continue; // T3
        var jid = String(hGet_(rin[i], hin, 'jenisid', 2) || '');
        list.push({
          id: String(hGet_(rin[i], hin, 'id', 0)),
          tipe: 'masuk',
          jenis: namaMasuk[jid] || jid || 'Pemasukan',
          jenisId: jid,
          noBukti: hin['nobukti'] !== undefined ? String(rin[i][hin['nobukti']] || '') : '',
          anggotaId: String(hGet_(rin[i], hin, 'anggotaid', 3) || ''),
          nominal: Number(hGet_(rin[i], hin, 'nominal', 5)) || 0,
          sumberKas: String(hGet_(rin[i], hin, 'sumberkas', 6) || ''),
          tanggal: toDateStr_(hGet_(rin[i], hin, 'tanggal', 4)),
          catatan: String(hGet_(rin[i], hin, 'catatan', 7) || ''),
          createdAt: toDateStr_(hGet_(rin[i], hin, 'createdat', 9)),
          seq: i
        });
      }
    }

    // Pengeluaran (Input Pengeluaran)
    var shOut = ss.getSheetByName(CONFIG.SHEETS.INPUT_PENGELUARAN);
    if (shOut) {
      var rout = shOut.getDataRange().getValues();
      var hout = headerMap_(rout[0]);
      for (var i = 1; i < rout.length; i++) {
        if (!hGet_(rout[i], hout, 'id', 0)) continue;
        if (!emailSama_(hGet_(rout[i], hout, 'createdby', 7), email)) continue;
        if (barisDibatalkan_(rout[i], hout)) continue; // T3
        var jid2 = String(hGet_(rout[i], hout, 'jenisid', 2) || '');
        list.push({
          id: String(hGet_(rout[i], hout, 'id', 0)),
          tipe: 'keluar',
          jenis: namaKeluar[jid2] || jid2 || 'Pengeluaran',
          jenisId: jid2,
          noBukti: hout['nobukti'] !== undefined ? String(rout[i][hout['nobukti']] || '') : '',
          anggotaId: '',
          nominal: Number(hGet_(rout[i], hout, 'nominal', 4)) || 0,
          sumberKas: String(hGet_(rout[i], hout, 'sumberkas', 5) || ''),
          tanggal: toDateStr_(hGet_(rout[i], hout, 'tanggal', 3)),
          catatan: String(hGet_(rout[i], hout, 'catatan', 6) || ''),
          createdAt: toDateStr_(hGet_(rout[i], hout, 'createdat', 8)),
          seq: i
        });
      }
    }

    // Mutasi (Input Setoran Bank)
    var shMut = ss.getSheetByName(CONFIG.SHEETS.INPUT_SETORAN);
    if (shMut) {
      var rmut = shMut.getDataRange().getValues();
      var hmut = headerMap_(rmut[0]);
      for (var i = 1; i < rmut.length; i++) {
        if (!hGet_(rmut[i], hmut, 'id', 0)) continue;
        if (!emailSama_(hGet_(rmut[i], hmut, 'createdby', 5), email)) continue;
        if (barisDibatalkan_(rmut[i], hmut)) continue; // T3
        var arah = String(hGet_(rmut[i], hmut, 'arah', 4) || '');
        list.push({
          id: String(hGet_(rmut[i], hmut, 'id', 0)),
          tipe: 'mutasi',
          jenis: arah === 'tarik' ? 'Tarik (Bank → Tunai)' : 'Setor (Tunai → Bank)',
          nominal: Number(hGet_(rmut[i], hmut, 'nominal', 3)) || 0,
          sumberKas: '',
          tanggal: toDateStr_(hGet_(rmut[i], hmut, 'tanggal', 2)),
          catatan: '',
          createdAt: toDateStr_(hGet_(rmut[i], hmut, 'createdat', 6)),
          seq: i
        });
      }
    }

    // Kas Penerobos (transaksi masuk via penerobos)
    var shKP = ss.getSheetByName(CONFIG.SHEETS.KAS_PENEROBOS);
    if (shKP) {
      var rkp = shKP.getDataRange().getValues();
      var hkp = headerMap_(rkp[0]);
      for (var i = 1; i < rkp.length; i++) {
        if (!hGet_(rkp[i], hkp, 'id', 0)) continue;
        if (!emailSama_(hGet_(rkp[i], hkp, 'penerobosemail', 8), email)) continue;
        var jid3 = String(hGet_(rkp[i], hkp, 'jenisid', 3) || '');
        list.push({
          id: String(hGet_(rkp[i], hkp, 'id', 0)),
          tipe: 'masuk',
          jenis: (namaMasuk[jid3] || jid3 || 'Pemasukan') + ' (Penerobos)',
          nominal: Number(hGet_(rkp[i], hkp, 'nominal', 5)) || 0,
          sumberKas: String(hGet_(rkp[i], hkp, 'sumberkas', 6) || ''),
          tanggal: toDateStr_(hGet_(rkp[i], hkp, 'tanggal', 2)),
          catatan: String(hGet_(rkp[i], hkp, 'catatan', 7) || ''),
          createdAt: toDateStr_(hGet_(rkp[i], hkp, 'createdat', 11)),
          seq: i
        });
      }
    }

    // Urut terbaru dulu: tanggal desc, lalu createdAt desc, lalu seq desc
    list.sort(function(a, b) {
      var ka = a.createdAt || a.tanggal || '';
      var kb = b.createdAt || b.tanggal || '';
      if (ka !== kb) return ka < kb ? 1 : -1;
      var ta = a.tanggal || '', tb = b.tanggal || '';
      if (ta !== tb) return ta < tb ? 1 : -1;
      return (b.seq || 0) - (a.seq || 0);
    });
    list.forEach(function(it) { delete it.seq; });

    var periode = getPeriodeAktif();
    var periodeOpen = !!(periode && String(periode.status) === CONFIG.STATUS.OPEN);
    return { success: true, data: list, total: list.length, periodeOpen: periodeOpen };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

// ──────────────────────────────────────────────────────
// ADMIN CONSOLE (desktop) — agregasi data nyata, KHUSUS ADMIN
// Dipakai halaman ?view=admin. Menggabungkan saldo, periode, arus kas
// 6 bulan, buku besar terbaru, komposisi pemasukan, peran, keamanan,
// perlu-tindakan, dan activity log — dari sheet yang sudah ada.
// ──────────────────────────────────────────────────────
function getAdminConsole() {
  try {
    var auth = checkAuth();
    if (!auth.success) return { success: false, message: auth.message };
    if (auth.user.role !== CONFIG.ROLES.ADMIN) return { success: false, message: 'Khusus Admin.' };

    var ss = getSS_();
    var periode = getPeriodeAktif();
    var periodeId = periode ? periode.id : null;

    var namaKelompok = 'Kas Kelompok';
    try {
      var sheetK = ss.getSheetByName(CONFIG.SHEETS.KELOMPOK);
      if (sheetK && sheetK.getLastRow() > 1) namaKelompok = sheetK.getRange(2, 2).getValue() || namaKelompok;
    } catch(e) {}

    var saldo = saldoSaatIni_(periodeId, periode);   // K3: aman tanpa periode OPEN

    // Peta nama jenis & anggota & user
    var namaMasuk = {}, katMasuk = {}, namaKeluar = {}, namaAnggota = {}, namaUser = {};
    var sheetP = ss.getSheetByName(CONFIG.SHEETS.PEMASUKAN);
    if (sheetP) { var rp = sheetP.getDataRange().getValues(); for (var i = 1; i < rp.length; i++) { if (rp[i][0]) { namaMasuk[String(rp[i][0])] = String(rp[i][1] || ''); katMasuk[String(rp[i][0])] = String(rp[i][2] || 'Umum'); } } }
    var sheetPK = ss.getSheetByName(CONFIG.SHEETS.PENGELUARAN);
    if (sheetPK) { var rpk = sheetPK.getDataRange().getValues(); for (var i = 1; i < rpk.length; i++) { if (rpk[i][0]) namaKeluar[String(rpk[i][0])] = String(rpk[i][1] || ''); } }
    var sheetA = ss.getSheetByName(CONFIG.SHEETS.ANGGOTA);
    if (sheetA) { var ra = sheetA.getDataRange().getValues(); for (var i = 1; i < ra.length; i++) { if (ra[i][0]) namaAnggota[String(ra[i][0])] = String(ra[i][1] || ''); } }
    var users = getUserList_();
    users.forEach(function(u) { namaUser[u.email.toLowerCase()] = u.nama || u.email; });

    // Baca transaksi masuk & keluar
    var trx = []; // {tgl, ym, jenis, kat, nama, sumber, nominal, arah, oleh, periodeId}
    var shIn = ss.getSheetByName(CONFIG.SHEETS.INPUT_PENERIMAAN);
    if (shIn) {
      var rin = shIn.getDataRange().getValues(); var hin = headerMap_(rin[0]);
      for (var i = 1; i < rin.length; i++) {
        if (!hGet_(rin[i], hin, 'id', 0)) continue;
        if (barisDibatalkan_(rin[i], hin)) continue; // T3
        var jid = String(hGet_(rin[i], hin, 'jenisid', 2) || '');
        var tgl = toDateStr_(hGet_(rin[i], hin, 'tanggal', 4));
        var aid = String(hGet_(rin[i], hin, 'anggotaid', 3) || '');
        var email = String(hGet_(rin[i], hin, 'createdby', 8) || '').toLowerCase();
        trx.push({ tgl: tgl, ym: tgl.substring(0, 7), jenis: namaMasuk[jid] || 'Pemasukan', kat: katMasuk[jid] || 'Umum',
          nama: namaAnggota[aid] || '', sumber: String(hGet_(rin[i], hin, 'sumberkas', 6) || 'Tunai'),
          nominal: Number(hGet_(rin[i], hin, 'nominal', 5)) || 0, arah: 'in', oleh: namaUser[email] || email,
          periodeId: String(hGet_(rin[i], hin, 'periodeid', 1) || '') });
      }
    }
    var shOut = ss.getSheetByName(CONFIG.SHEETS.INPUT_PENGELUARAN);
    if (shOut) {
      var rout = shOut.getDataRange().getValues(); var hout = headerMap_(rout[0]);
      for (var i = 1; i < rout.length; i++) {
        if (!hGet_(rout[i], hout, 'id', 0)) continue;
        if (barisDibatalkan_(rout[i], hout)) continue; // T3
        var jid2 = String(hGet_(rout[i], hout, 'jenisid', 2) || '');
        var tgl2 = toDateStr_(hGet_(rout[i], hout, 'tanggal', 3));
        var email2 = String(hGet_(rout[i], hout, 'createdby', 7) || '').toLowerCase();
        trx.push({ tgl: tgl2, ym: tgl2.substring(0, 7), jenis: namaKeluar[jid2] || 'Pengeluaran', kat: 'Umum',
          nama: String(hGet_(rout[i], hout, 'catatan', 6) || ''), sumber: String(hGet_(rout[i], hout, 'sumberkas', 5) || 'Tunai'),
          nominal: Number(hGet_(rout[i], hout, 'nominal', 4)) || 0, arah: 'out', oleh: namaUser[email2] || email2,
          periodeId: String(hGet_(rout[i], hout, 'periodeid', 1) || '') });
      }
    }

    // Arus kas 6 bulan terakhir (dari semua transaksi, per bulan)
    var byMonth = {};
    trx.forEach(function(t) { if (!t.ym) return; if (!byMonth[t.ym]) byMonth[t.ym] = { masuk: 0, keluar: 0 }; byMonth[t.ym][t.arah === 'in' ? 'masuk' : 'keluar'] += t.nominal; });
    var months = Object.keys(byMonth).sort();
    var arus6 = months.slice(-6).map(function(m) { return { ym: m, masuk: byMonth[m].masuk, keluar: byMonth[m].keluar }; });

    // Masuk/keluar bulan berjalan
    var nowYm = toDateStr_(new Date()).substring(0, 7);
    var masukBulan = (byMonth[nowYm] || {}).masuk || 0;
    var keluarBulan = (byMonth[nowYm] || {}).keluar || 0;

    // Buku besar: transaksi periode berjalan terbaru (maks 12)
    var bukuBesar = trx.filter(function(t) { return !periodeId || t.periodeId === periodeId; })
      .sort(function(a, b) { return a.tgl < b.tgl ? 1 : (a.tgl > b.tgl ? -1 : 0); })
      .slice(0, 12);
    var trxCountPeriode = trx.filter(function(t) { return !periodeId || t.periodeId === periodeId; }).length;

    // Komposisi pemasukan (per jenis, periode berjalan) → persentase
    var komp = {}, totalMasukP = 0;
    trx.forEach(function(t) { if (t.arah !== 'in') return; if (periodeId && t.periodeId !== periodeId) return; komp[t.jenis] = (komp[t.jenis] || 0) + t.nominal; totalMasukP += t.nominal; });
    var komposisi = Object.keys(komp).map(function(k) { return { nama: k, nilai: komp[k], pct: totalMasukP ? Math.round(komp[k] / totalMasukP * 100) : 0 }; })
      .sort(function(a, b) { return b.nilai - a.nilai; }).slice(0, 6);

    // Peran & user
    var roleCount = {}; var aktif = 0, nonaktif = 0;
    users.forEach(function(u) {
      var act = String(u.status || 'Aktif').toLowerCase() !== 'nonaktif';
      if (act) aktif++; else nonaktif++;
      roleCount[u.role] = (roleCount[u.role] || 0) + 1;
    });

    // Keamanan: perangkat aktif & sesi aktif
    var perangkatAktif = 0, sesiAktif = 0;
    var shDev = ss.getSheetByName(CONFIG.SHEETS.PERANGKAT);
    if (shDev) { var rd = shDev.getDataRange().getValues(); var hd = headerMap_(rd[0]); for (var i = 1; i < rd.length; i++) { if (hGet_(rd[i], hd, 'id', 0) && String(hGet_(rd[i], hd, 'status', 6)) === 'Aktif') perangkatAktif++; } }
    var shSes = ss.getSheetByName(CONFIG.SHEETS.SESI);
    var nowMs = new Date().getTime();
    if (shSes) { var rs = shSes.getDataRange().getValues(); var hs = headerMap_(rs[0]); for (var i = 1; i < rs.length; i++) { if (String(hGet_(rs[i], hs, 'status', 5)) !== 'Aktif') continue; var ev = hGet_(rs[i], hs, 'kadaluarsa', 4); var em = (ev instanceof Date) ? ev.getTime() : new Date(ev).getTime(); if (nowMs <= em) sesiAktif++; } }

    // Perlu tindakan
    var belumDirinci = 0;
    try { var bc = CacheService.getScriptCache().get('buku_ir_data'); if (bc) { var bd = JSON.parse(bc); belumDirinci = (bd.data && bd.data.belumDirincikan) ? bd.data.belumDirincikan.length : 0; } } catch(e) {}
    var serahMenunggu = 0;
    var shST = ss.getSheetByName(CONFIG.SHEETS.SERAH_TERIMA);
    if (shST) { var rst = shST.getDataRange().getValues(); var hst = headerMap_(rst[0]); for (var i = 1; i < rst.length; i++) { if (hGet_(rst[i], hst, 'id', 0) && String(hGet_(rst[i], hst, 'status', 7)) === 'Menunggu') serahMenunggu++; } }

    // Activity log terbaru (maks 8)
    var logs = [];
    var shLog = ss.getSheetByName(CONFIG.SHEETS.LOG);
    if (shLog) {
      var rl = shLog.getDataRange().getValues();
      for (var i = rl.length - 1; i >= 1 && logs.length < 8; i--) {
        if (!rl[i][0]) continue;
        var em3 = String(rl[i][1] || '').toLowerCase();
        logs.push({ who: namaUser[em3] || rl[i][1] || 'Sistem', action: String(rl[i][2] || ''), detail: String(rl[i][3] || ''), time: toDateStr_(rl[i][0]) });
      }
    }

    var inisial = (auth.user.nama || 'A').split(' ').map(function(w) { return w.charAt(0); }).join('').substring(0, 2).toUpperCase();

    return {
      success: true,
      namaKelompok: namaKelompok,
      user: { nama: auth.user.nama, role: auth.user.role, inisial: inisial },
      periode: periode ? { nama: periode.nama, mulai: periode.tanggalMulai || '', status: periode.status, saldoAwalTunai: periode.saldoAwalTunai || 0, saldoAwalBank: periode.saldoAwalBank || 0, transaksi: trxCountPeriode } : null,
      kasTunai: saldo.tunai, kasBank: saldo.bank, totalKas: saldo.tunai + saldo.bank,
      masukBulan: masukBulan, keluarBulan: keluarBulan,
      arus6: arus6, bukuBesar: bukuBesar, komposisi: komposisi,
      roles: roleCount, userAktif: aktif, userNonaktif: nonaktif,
      perangkatAktif: perangkatAktif, sesiAktif: sesiAktif,
      belumDirinci: belumDirinci, serahMenunggu: serahMenunggu,
      logs: logs
    };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

// Activity Log terbaru (KHUSUS ADMIN) untuk halaman konsol.
function getActivityLog(limit) {
  try {
    var auth = checkAuth();
    if (!auth.success) return { success: false, message: auth.message };
    if (auth.user.role !== CONFIG.ROLES.ADMIN) return { success: false, message: 'Khusus Admin.' };
    var max = Math.min(Number(limit) || 100, 300);
    var ss = getSS_();
    // Peta email → nama
    var namaUser = {};
    getUserList_().forEach(function(u) { namaUser[u.email.toLowerCase()] = u.nama || u.email; });
    var sheet = ss.getSheetByName(CONFIG.SHEETS.LOG);
    if (!sheet) return { success: true, data: [] };
    var rows = sheet.getDataRange().getValues();
    var out = [];
    for (var i = rows.length - 1; i >= 1 && out.length < max; i--) {
      if (!rows[i][0]) continue;
      var em = String(rows[i][1] || '').toLowerCase();
      out.push({
        waktu: toDateStr_(rows[i][0]),
        user: namaUser[em] || rows[i][1] || 'Sistem',
        email: String(rows[i][1] || ''),
        action: String(rows[i][2] || ''),
        detail: String(rows[i][3] || '')
      });
    }
    return { success: true, data: out };
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
    // Nominal tiap grade: bilangan bulat >= 0, minimal satu > 0.
    var adaNominal = false;
    var gradeKeys = Object.keys(data.gradeConfig);
    for (var gi = 0; gi < gradeKeys.length; gi++) {
      var gv = Number(data.gradeConfig[gradeKeys[gi]]);
      if (!isFinite(gv) || gv < 0 || Math.floor(gv) !== gv) return { success: false, message: 'Nominal grade harus bilangan bulat ≥ 0.' };
      if (gv > 0) adaNominal = true;
    }
    if (!adaNominal) return { success: false, message: 'Minimal satu grade memiliki nominal > 0.' };

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
    return withLock_(function() {
    var ss = getSS_();
    var sheet = ss.getSheetByName(CONFIG.SHEETS.TAGIHAN_PATUNGAN);
    if (!sheet) return { success: false, message: 'Sheet penerobosan tidak ditemukan' };
    var rows = sheet.getDataRange().getValues();
    var h = headerMap_(rows[0]);
    var today = toDateStr_(new Date());
    for (var i = 1; i < rows.length; i++) {
      if (String(hGet_(rows[i], h, 'id', 0)) === tagihanId) {
        var apPat = assertPeriodeOpen_(getPatunganPeriodeId_(String(hGet_(rows[i], h, 'patunganid', 1) || ''))); // T2
        if (!apPat.ok) return { success: false, message: apPat.message };
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
    });
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function batalBayarTagihanPatungan(tagihanId) {
  try {
    var auth = requirePerm('terobosan.batal');
    if (!auth.success) return { success: false, message: auth.message };
    return withLock_(function() {
    var ss = getSS_();
    var sheet = ss.getSheetByName(CONFIG.SHEETS.TAGIHAN_PATUNGAN);
    if (!sheet) return { success: false, message: 'Sheet penerobosan tidak ditemukan' };
    var rows = sheet.getDataRange().getValues();
    var h = headerMap_(rows[0]);
    for (var i = 1; i < rows.length; i++) {
      if (String(hGet_(rows[i], h, 'id', 0)) === tagihanId) {
        var apPat2 = assertPeriodeOpen_(getPatunganPeriodeId_(String(hGet_(rows[i], h, 'patunganid', 1) || ''))); // T2
        if (!apPat2.ok) return { success: false, message: apPat2.message };
        var colStatus = (h['statusbayar'] !== undefined ? h['statusbayar'] : 6) + 1;
        var colTgl = (h['tanggalbayar'] !== undefined ? h['tanggalbayar'] : 7) + 1;
        sheet.getRange(i + 1, colStatus).setValue('Belum');
        sheet.getRange(i + 1, colTgl).setValue('');
        return { success: true };
      }
    }
    return { success: false, message: 'Penerobosan tidak ditemukan' };
    });
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
      var setBatalBIR = trxPenerimaanDibatalkan_();      // ← K1: jangan muat rincian yatim
      var irRows = sheetIR.getDataRange().getValues();
      var irH = headerMap_(irRows[0]);
      for (var i = 1; i < irRows.length; i++) {
        var trxId = String(hGet_(irRows[i], irH, 'transaksiid', 1) || '');
        if (!trxId) continue;
        if (rincianYatim_(irRows[i], irH, setBatalBIR)) continue;   // ← K1
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
        if (barisDibatalkan_(row, pHdr)) continue;   // ← K2: penerimaan batal jangan dihitung "belum dirincikan"
        if (bukuIRIds.indexOf(row[c_jid]) === -1) continue;
        var trxIdStr = String(row[c_id]);
        var ditangguhkan = rincianDitangguhkan_(row, pHdr);
        var periodeLain = (periodeId && row[c_pid] !== periodeId);
        // Transaksi periode lain hanya ikut bila DITANGGUHKAN dan belum dirinci —
        // supaya bisa diselesaikan di periode berjalan.
        if (periodeLain && !(ditangguhkan && !rincianMap[trxIdStr])) continue;
        var item = {
          id: trxIdStr, periodeId: String(row[c_pid]), jenisId: String(row[c_jid]),
          anggotaId: String(row[c_aid] || ''),
          tanggal: toDateStr_(row[c_tgl]),
          nominal: Number(row[c_nom]) || 0,
          sumberKas: String(row[c_kas] || ''),
          catatan: String(row[c_cat] || ''),
          anggota: anggotaMap[String(row[c_aid])] || null,
          ditangguhkan: ditangguhkan,
          dariPeriodeLain: !!periodeLain
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

// K2: bersihkan cache lama setelah deploy patch. Jalankan sekali dari editor.
function bersihkanCache() {
  var c = CacheService.getScriptCache();
  c.remove('buku_ir_data');
  c.remove('dashboard_saldo');
  c.remove('master_trx_data');
  return { success: true };
}

function submitRincianIR(data) {
  try {
    // FASE 1: endpoint mutasi wajib requirePerm (sebelumnya hanya checkAuth).
    var auth = requirePerm('bukuIR.input');
    if (!auth.success) return { success: false, message: auth.message };
    return withLock_(function() {
    var ss = getSS_();
    var sheet = ss.getSheetByName(CONFIG.SHEETS.BUKU_IR);
    if (!sheet) {
      sheet = ss.insertSheet(CONFIG.SHEETS.BUKU_IR);
      // Kolom Total dihapus — nilai derived, dihitung saat read (IR+IR10+Cicilan+InfakDaerah+Index)
      sheet.appendRow(['ID', 'TransaksiID', 'PeriodeID', 'AnggotaID', 'Tanggal', 'IR', 'IR10', 'Cicilan', 'InfakDaerah', 'Index', 'CreatedBy', 'CreatedAt']);
    }
    // Baca transaksi penerimaan yang dirujuk. Nominal/periode/anggota/tanggal
    // diambil dari SHEET (bukan klien) — menutup T6 & A1.
    var sheetP = ss.getSheetByName(CONFIG.SHEETS.INPUT_PENERIMAAN);
    if (!sheetP) return { success: false, message: 'Sheet penerimaan tidak ditemukan.' };
    var pRows = sheetP.getDataRange().getValues();
    var pH = headerMap_(pRows[0]);
    var trxRow = null;
    for (var vi = 1; vi < pRows.length; vi++) {
      if (String(hGet_(pRows[vi], pH, 'id', 0)) === String(data.transaksiId)) { trxRow = pRows[vi]; break; }
    }
    if (!trxRow) return { success: false, message: 'Transaksi tidak ditemukan.' };
    if (barisDibatalkan_(trxRow, pH)) return { success: false, message: 'Transaksi sudah dibatalkan.' };
    var trxNominal   = Number(hGet_(trxRow, pH, 'nominal', 5)) || 0;
    var trxPeriodeId = String(hGet_(trxRow, pH, 'periodeid', 1) || '');
    var trxAnggotaId = String(hGet_(trxRow, pH, 'anggotaid', 3) || '');
    var trxTanggal   = toDateStr_(hGet_(trxRow, pH, 'tanggal', 4));

    // Kunci periode tertutup (T2) — KECUALI transaksi yang rinciannya DITANGGUHKAN.
    // Rincian tidak mengubah uang periode lama (kas tetap di sana); yang dicatat di
    // periode berjalan hanyalah komposisinya, sehingga kewajiban setornya jatuh di
    // periode tempat rincian diisi (tempat ia akan disetor).
    var ditangguhkan = rincianDitangguhkan_(trxRow, pH);
    var periodeRincian = trxPeriodeId;
    if (ditangguhkan) {
      var pAktif = getPeriodeAktif();
      if (!pAktif) return { success: false, message: 'Tidak ada periode aktif untuk mencatat rincian tangguhan.' };
      periodeRincian = pAktif.id;
    } else {
      var ap = assertPeriodeOpen_(trxPeriodeId);
      if (!ap.ok) return { success: false, message: ap.message };
    }

    var ir = Number(data.ir) || 0;
    var ir10 = Number(data.ir10) || 0;
    var cicilan = Number(data.cicilan) || 0;
    var infakDaerah = Number(data.infakDaerah) || 0;
    var index = Number(data.index) || 0;
    if (ir < 0 || ir10 < 0 || cicilan < 0 || infakDaerah < 0 || index < 0) {
      return { success: false, message: 'Komponen rincian tidak boleh negatif.' };
    }
    // Validasi balance di SERVER (T6): total komponen = nominal transaksi.
    var totalKomp = ir + ir10 + cicilan + infakDaerah + index;
    if (totalKomp !== trxNominal) {
      return { success: false, message: 'Rincian tidak seimbang: total Rp ' + totalKomp.toLocaleString('id-ID') + ' ≠ nominal transaksi Rp ' + trxNominal.toLocaleString('id-ID') + ' (selisih Rp ' + Math.abs(totalKomp - trxNominal).toLocaleString('id-ID') + ').' };
    }

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
      // Anggota/tanggal dari transaksi (server), bukan klien (A1). PeriodeID memakai
      // periodeRincian: sama dengan transaksi, atau periode berjalan bila ditangguhkan.
      sheet.appendRow([id, data.transaksiId, periodeRincian, trxAnggotaId, trxTanggal,
        ir, ir10, cicilan, infakDaerah, index, auth.user.email, toDateStr_(new Date())]);
      if (ditangguhkan) {
        logActivityWajib_(auth.user.email, 'RINCIAN_TANGGUHAN_DIISI',
          'Transaksi: ' + data.transaksiId + ' | periode asal: ' + trxPeriodeId + ' | dicatat di periode: ' + periodeRincian);
      }
    }
    try { var c = CacheService.getScriptCache(); c.remove('master_trx_data'); c.remove('buku_ir_data'); c.remove('dashboard_saldo'); } catch(e) {}
    return { success: true, id: id, updated: existingRow > 0 };
    });
  } catch(e) {
    return { success: false, message: e.message };
  }
}

// ──────────────────────────────────────────────────────
// REKAPITULASI
// ──────────────────────────────────────────────────────
function getRekapitulasiData(periodeIdParam) {
  try {
    var auth = checkAuth();
    if (!auth.success) return { success: false, message: auth.message };
    var ss = getSS_();
    // Bila periodeIdParam diberikan, laporkan periode itu (mis. periode CLOSED untuk
    // laporan historis); jika tidak, pakai periode aktif.
    var periode = periodeIdParam ? getPeriodeById_(periodeIdParam) : getPeriodeAktif();
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
        if (barisDibatalkan_(dp[i], dpH)) continue; // T3
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
        if (barisDibatalkan_(dpk[i], dpkH)) continue; // T3
        if (barisDraft_(dpk[i], dpkH)) continue;       // T12: Draft belum final
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

    var saldo = saldoSaatIni_(periodeId, periode);   // K3: aman tanpa periode OPEN

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
      var setBatal = trxPenerimaanDibatalkan_();          // ← K1
      var irRows = sheetIR.getDataRange().getValues();
      irColMap = headerMap_(irRows[0]);
      var ir_perid = irColMap['periodeid'] !== undefined ? irColMap['periodeid'] : 2;
      for (var j = 1; j < irRows.length; j++) {
        if (!irRows[j][0]) continue;
        if (rincianYatim_(irRows[j], irColMap, setBatal)) continue;   // ← K1
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
        if (barisDibatalkan_(pmRows[k], pmH)) continue; // T3
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
      var danaTersedia = null, danaRef = '';   // dana penutup kewajiban (earmark), utk tipe 'tetap'
      if (sumberTipe === 'bukuir') {
        // Per kolom — sesuai pilihan dropdown di tabel Pos Setoran (Sumber Ref)
        var colIdx = mapPosNamaToBukuIRCol(sumberRef, sumberRef, irColMap);
        if (colIdx >= 0) {
          for (var r = 0; r < irData.length; r++) target += Number(irData[r][colIdx]) || 0;
        }
        sumberKet = 'Buku IR: ' + sumberRef;
      } else if (sumberTipe === 'tetap') {
        // Kewajiban TETAP per bulan (keputusan musyawaroh) — besarnya tidak bergantung
        // pemasukan. Target = nilai kolom Target x jumlah bulan dalam periode.
        var perBulan = Number(posRows[i][posH['target'] !== undefined ? posH['target'] : 5]) || 0;
        var nBulan = 1;
        try { if (periode && periode.tanggalMulai) nBulan = Math.max(1, hitungJumlahBulan(periode.tanggalMulai)); } catch(e) {}
        target = perBulan * nBulan;
        sumberKet = 'Tetap ' + perBulan.toLocaleString('id-ID') + '/bulan' + (nBulan > 1 ? ' × ' + nBulan + ' bulan' : '');
        // Sumber Ref (opsional) = kolom Buku IR yang jadi dana utama penutup kewajiban ini
        // (mis. Index untuk Jatah Desa). Hanya informasi — tidak mengubah besar kewajiban.
        if (sumberRef) {
          var cIdx = mapPosNamaToBukuIRCol(sumberRef, sumberRef, irColMap);
          if (cIdx >= 0) {
            danaTersedia = 0;
            for (var d = 0; d < irData.length; d++) danaTersedia += Number(irData[d][cIdx]) || 0;
            danaRef = sumberRef;
          }
        }
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
        isAuto: sumberTipe !== 'manual',
        // Earmark: dana yang tersedia untuk menutup kewajiban ini + kurang/lebihnya.
        danaRef: danaRef,
        danaTersedia: danaTersedia,
        danaKurang: (danaTersedia === null) ? null : Math.max(0, target - danaTersedia),
        danaLebih: (danaTersedia === null) ? null : Math.max(0, danaTersedia - target)
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
    return withLock_(function() {
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
    if (realisasi < 0 || Math.floor(realisasi) !== realisasi) return { success: false, message: 'Realisasi harus bilangan bulat ≥ 0.' };
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
    });
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

  // T9: KOREKSI, bukan overwrite. Bila baris lama ada & nilainya BERUBAH,
  // batalkan baris lama (soft delete, alasan otomatis) lalu buat baris baru.
  // Tanggal baris lama TIDAK PERNAH diubah.
  if (existingId) {
    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][c_id]) !== String(existingId)) continue;
      if (barisDibatalkan_(rows[i], h)) break; // sudah dibatalkan → buat baru
      var samaNominal = (Number(rows[i][c_nom]) || 0) === (Number(info.nominal) || 0);
      var samaJenis = String(rows[i][c_jid]) === String(info.jenisId);
      var samaKas = String(rows[i][c_kas]) === String(info.sumberKas);
      if (samaNominal && samaJenis && samaKas) {
        // Tidak ada perubahan material → cukup perbarui catatan.
        sheet.getRange(i + 1, c_cat + 1).setValue(info.catatan);
        return existingId;
      }
      // Ada perubahan → batalkan baris lama (jejak tetap ada).
      if (h['status'] !== undefined) {
        sheet.getRange(i + 1, h['status'] + 1).setValue('Dibatalkan');
        if (h['dibatalkanby'] !== undefined) sheet.getRange(i + 1, h['dibatalkanby'] + 1).setValue(info.email || '');
        if (h['dibatalkanat'] !== undefined) sheet.getRange(i + 1, h['dibatalkanat'] + 1).setValue(toDateStr_(new Date()));
        if (h['alasanbatal'] !== undefined) sheet.getRange(i + 1, h['alasanbatal'] + 1).setValue('Koreksi realisasi setoran');
      }
      try { logActivityWajib_(info.email || 'sistem', 'KOREKSI_SETORAN', 'Batal ' + existingId + ' → nominal baru ' + info.nominal); } catch(e) {}
      break;
    }
  }
  // Buat baris baru (Status Aktif).
  var newId = generateID('PNK');
  sheet.appendRow([newId, info.periodeId, info.jenisId, toDateStr_(new Date()),
    info.nominal, info.sumberKas, info.catatan, info.email, toDateStr_(new Date()), 'Aktif']);
  _isiNoBukti_(sheet, 'BKK', info.periodeId);
  return newId;
}

// ──────────────────────────────────────────────────────
// PDF
// ──────────────────────────────────────────────────────
function generatePDF(periodeId) {
  try {
    var auth = checkAuth();
    if (!auth.success) return { success: false, message: auth.message };

    // T14: laporan periode yang SUDAH DITUTUP diutamakan dari ARSIP (snapshot saat
    // tutup buku). Bila arsip TIDAK ada (mis. FOLDER_ARSIP_ID belum diset waktu tutup
    // buku), laporan tetap dibuat dari DATA HISTORIS periode itu (mode 'historis').
    var historis = false;
    if (periodeId) {
      var p = getPeriodeById_(periodeId);
      if (p && String(p.status) === CONFIG.STATUS.CLOSED) {
        var arsipUrl = _arsipUrlPeriode_(periodeId);
        if (arsipUrl) return { success: true, dariArsip: true, arsipUrl: arsipUrl, periodeNama: p.nama };
        historis = true; // tak ada arsip → susun dari data historis periode ini
      }
    }

    // Laporkan periode yang diminta (periodeId) atau periode aktif bila null.
    var rekap = getRekapitulasiData(periodeId || null);
    if (!rekap.success) return rekap;

    // Lampiran seksi tambahan hanya untuk laporan periode AKTIF (data hidup).
    // Untuk laporan historis, seksi tsb dilewati agar tidak mencampur data periode lain.
    if (!historis) {
      // Seksi D ambil dari POS SETORAN (tempat realisasi benar-benar dicatat).
      // Sebelumnya memakai getLaporanSetoran yang mencocokkan realisasi lewat kode
      // jenis pemasukan, padahal sheet Setoran Desa menyimpan PosID → tak pernah
      // cocok sehingga "Sudah Setor" selalu Rp 0 meski setoran sudah dilakukan.
      try {
        var rs = getRekapSetoran();
        if (rs && rs.success) {
          var sdBaris = (rs.data || []).map(function(p) {
            return { nama: p.nama, kewajiban: p.target, sudahSetor: p.realisasi, sisa: p.sisa, status: p.status };
          });
          var sT = 0, sS = 0;
          sdBaris.forEach(function(b) { sT += Number(b.kewajiban) || 0; sS += Number(b.sudahSetor) || 0; });
          rekap.setoran = { success: true, data: sdBaris,
            summary: { totalTarget: sT, totalSudahSetor: sS, totalSisa: sT - sS } };
        }
      } catch(e) {}
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
    }

    var html = buildPDFHTML(rekap);
    // Kirim HTML string ke client — browser akan render langsung via window.open
    return { success: true, html: html };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function buildPDFHTML(data) {
  var p = data.periode || {};
  var isFinal = (p.status === 'Tutup' || String(p.status).toUpperCase() === 'CLOSED');
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
    return withLock_(function() {
    var ss = getSS_();
    var sheet = ss.getSheetByName(CONFIG.SHEETS.BANK_DAILY);
    if (!sheet) {
      sheet = ss.insertSheet(CONFIG.SHEETS.BANK_DAILY);
      sheet.appendRow(['ID', 'PeriodeID', 'Tanggal', 'Saldo Awal', 'Pemasukan', 'Pengeluaran', 'Saldo Akhir Teoritis', 'Saldo Akhir Actual', 'Selisih', 'Status', 'Catatan', 'Last Updated']);
    }
    var tglCekBnk = validasiTanggalPeriode_(data.tanggal, data.periodeId);
    if (!tglCekBnk.ok) return { success: false, message: tglCekBnk.message };
    if ([data.saldoAwal, data.pemasukan, data.pengeluaran, data.saldoAkhirActual].some(function(v) { return Number(v) < 0; })) {
      return { success: false, message: 'Nilai saldo/mutasi bank tidak boleh negatif.' };
    }
    var id = generateID('BNK');
    var saldoAkhirTeoritis = (Number(data.saldoAwal) || 0) + (Number(data.pemasukan) || 0) - (Number(data.pengeluaran) || 0);
    var selisih = (Number(data.saldoAkhirActual) || 0) - saldoAkhirTeoritis;
    sheet.appendRow([id, data.periodeId, tglCekBnk.tgl,
      Number(data.saldoAwal) || 0, Number(data.pemasukan) || 0, Number(data.pengeluaran) || 0,
      saldoAkhirTeoritis, Number(data.saldoAkhirActual) || 0, selisih,
      selisih === 0 ? 'Balance' : 'Selisih', data.catatan || '', toDateStr_(new Date())]);
    try { CacheService.getScriptCache().remove('master_trx_data'); } catch(e) {}
    return { success: true, id: id };
    });
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function updateBankDaily(data) {
  try {
    var auth = requirePerm('bank.manage');
    if (!auth.success) return { success: false, message: auth.message };
    return withLock_(function() {
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
    });
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
    var nomCekPnd = validasiNominal_(data.nominal);
    if (!nomCekPnd.ok) return { success: false, message: nomCekPnd.message };
    var tglCekPnd = validasiTanggalPeriode_(data.tanggal ? data.tanggal : new Date(), periode.id);
    if (!tglCekPnd.ok) return { success: false, message: tglCekPnd.message };
    var id = generateID('PND');
    sheet.appendRow([id, periode.id, tglCekPnd.tgl, data.keterangan, nomCekPnd.nilai,
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
    var hpnd = headerMap_(rows[0]);
    for (var i = 1; i < rows.length; i++) {
      if (rows[i][0] === id) {
        var apPnd = assertPeriodeOpen_(String(hGet_(rows[i], hpnd, 'periodeid', 1) || '')); // T2
        if (!apPnd.ok) return { success: false, message: apPnd.message };
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

// Bagi hasil: hitung hak Kelompok / Desa / Daerah dari pemasukan periode.
// Tiap jenis pemasukan punya %Kelompok, %Desa, %Daerah (Master Pemasukan).
// periodeId opsional → default periode aktif. Baris dibatalkan diabaikan.
function getBagiHasil(periodeId) {
  try {
    var auth = checkAuth();
    if (!auth.success) return { success: false, message: auth.message };
    var ss = getSS_();
    var periode = periodeId ? getPeriodeById_(periodeId) : getPeriodeAktif();
    if (!periode) return { success: false, message: 'Tidak ada periode untuk dihitung.' };
    var pid = periode.id;

    // Master Pemasukan → jenis + persentase + kategori.
    var jenis = {}; // kode → {nama, kategori, pctK, pctD, pctDr}
    var shM = ss.getSheetByName(CONFIG.SHEETS.PEMASUKAN);
    if (shM && shM.getLastRow() > 1) {
      var mr = shM.getDataRange().getValues(); var mh = headerMap_(mr[0]);
      for (var i = 1; i < mr.length; i++) {
        var kode = String(hGet_(mr[i], mh, 'kode', 0) || '');
        if (!kode) continue;
        jenis[kode] = {
          nama: String(hGet_(mr[i], mh, 'namapemasukan', 1) || hGet_(mr[i], mh, 'nama', 1) || kode),
          kategori: String(hGet_(mr[i], mh, 'kategori', 2) || ''),
          pctK: Number(hGet_(mr[i], mh, '%kelompok', 3)) || 0,
          pctD: Number(hGet_(mr[i], mh, '%desa', 4)) || 0,
          pctDr: Number(hGet_(mr[i], mh, '%daerah', 5)) || 0
        };
      }
    }
    function isBukuIR_(j) { return j && String(j.kategori).toLowerCase().indexOf('buku ir') !== -1; }

    // Total pemasukan per jenis untuk periode ini (skip dibatalkan).
    var masukPerJenis = {};
    var shP = ss.getSheetByName(CONFIG.SHEETS.INPUT_PENERIMAAN);
    if (shP && shP.getLastRow() > 1) {
      var pr = shP.getDataRange().getValues(); var ph = headerMap_(pr[0]);
      for (var r = 1; r < pr.length; r++) {
        if (!hGet_(pr[r], ph, 'id', 0)) continue;
        if (barisDibatalkan_(pr[r], ph)) continue;
        if (String(hGet_(pr[r], ph, 'periodeid', 1)) !== String(pid)) continue;
        var jid = String(hGet_(pr[r], ph, 'jenisid', 2) || '');
        masukPerJenis[jid] = (masukPerJenis[jid] || 0) + (Number(hGet_(pr[r], ph, 'nominal', 5)) || 0);
      }
    }

    // Komponen Buku IR (dari Detail Buku IR) + pemilik tiap komponen.
    // Aturan mengikuti Setoran Desa: IR, 1/10 IR, Cicilan, Infak Daerah = Desa; Index = Kelompok.
    var IR_KOMP = [
      { key: 'ir', kolom: 5, nama: 'IR', hak: 'desa' },
      { key: 'ir10', kolom: 6, nama: '1/10 IR', hak: 'desa' },
      { key: 'cicilan', kolom: 7, nama: 'Cicilan', hak: 'desa' },
      { key: 'infakdaerah', kolom: 8, nama: 'Infak Daerah', hak: 'desa' },
      { key: 'index', kolom: 9, nama: 'Index', hak: 'kelompok' }
    ];
    var irSum = { ir: 0, ir10: 0, cicilan: 0, infakdaerah: 0, index: 0 };
    var adaBukuIR = false;
    var shIR = ss.getSheetByName(CONFIG.SHEETS.BUKU_IR);
    if (shIR && shIR.getLastRow() > 1) {
      var setBatal = trxPenerimaanDibatalkan_();
      var ir = shIR.getDataRange().getValues(); var ih = headerMap_(ir[0]);
      for (var k = 1; k < ir.length; k++) {
        if (!ir[k][0]) continue;
        if (rincianYatim_(ir[k], ih, setBatal)) continue;
        if (String(hGet_(ir[k], ih, 'periodeid', 2)) !== String(pid)) continue;
        adaBukuIR = true;
        IR_KOMP.forEach(function(c) { irSum[c.key] += Number(hGet_(ir[k], ih, c.key, c.kolom)) || 0; });
      }
    }

    var totKel = 0, totDesa = 0, totDaerah = 0, totMasuk = 0;
    var rincian = [];
    // 1) Jenis non-Buku IR → pakai persentase master.
    Object.keys(masukPerJenis).forEach(function(jid) {
      var m = masukPerJenis[jid] || 0;
      if (m <= 0) return;
      var j = jenis[jid] || { nama: jid, kategori: '', pctK: 100, pctD: 0, pctDr: 0 };
      if (isBukuIR_(j)) return; // Buku IR ditangani per-komponen di bawah
      var kel = Math.round(m * j.pctK / 100);
      var desa = Math.round(m * j.pctD / 100);
      var daerah = Math.round(m * j.pctDr / 100);
      totMasuk += m; totKel += kel; totDesa += desa; totDaerah += daerah;
      rincian.push({ jenis: j.nama, masuk: m, pctK: j.pctK, pctD: j.pctD, pctDr: j.pctDr, kelompok: kel, desa: desa, daerah: daerah });
    });
    // 2) Komponen Buku IR → satu baris per komponen sesuai pemiliknya.
    if (adaBukuIR) {
      IR_KOMP.forEach(function(c) {
        var m = irSum[c.key] || 0;
        if (m <= 0) return;
        var kel = c.hak === 'kelompok' ? m : 0;
        var desa = c.hak === 'desa' ? m : 0;
        var daerah = c.hak === 'daerah' ? m : 0;
        totMasuk += m; totKel += kel; totDesa += desa; totDaerah += daerah;
        rincian.push({
          jenis: 'Buku IR · ' + c.nama, masuk: m,
          pctK: kel ? 100 : 0, pctD: desa ? 100 : 0, pctDr: daerah ? 100 : 0,
          kelompok: kel, desa: desa, daerah: daerah
        });
      });
    }
    rincian.sort(function(a, b) { return b.masuk - a.masuk; });

    return {
      success: true,
      periode: { id: pid, nama: periode.nama, status: periode.status },
      totalMasuk: totMasuk,
      kelompok: totKel, desa: totDesa, daerah: totDaerah,
      rincian: rincian
    };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

// Rincian IR periode berjalan: daftar per jamaah + komponen (IR, 1/10 IR,
// Cicilan, Infak Daerah, Index) + total per baris & total kolom.
function getRincianIR(periodeId) {
  try {
    var auth = checkAuth();
    if (!auth.success) return { success: false, message: auth.message };
    var ss = getSS_();
    var periode = periodeId ? getPeriodeById_(periodeId) : getPeriodeAktif();
    if (!periode) return { success: false, message: 'Tidak ada periode untuk ditampilkan.' };
    var pid = periode.id;
    // Nama jamaah.
    var angMap = {};
    var shA = ss.getSheetByName(CONFIG.SHEETS.ANGGOTA);
    if (shA && shA.getLastRow() > 1) { var ar = shA.getDataRange().getValues(); for (var a = 1; a < ar.length; a++) if (ar[a][0]) angMap[String(ar[a][0])] = String(ar[a][1] || ''); }
    var setBatal = trxPenerimaanDibatalkan_();
    var rows = [], tot = { ir: 0, ir10: 0, cicilan: 0, infakdaerah: 0, index: 0, total: 0 };
    var shIR = ss.getSheetByName(CONFIG.SHEETS.BUKU_IR);
    if (shIR && shIR.getLastRow() > 1) {
      var ir = shIR.getDataRange().getValues(); var ih = headerMap_(ir[0]);
      for (var k = 1; k < ir.length; k++) {
        if (!ir[k][0]) continue;
        if (rincianYatim_(ir[k], ih, setBatal)) continue;
        if (String(hGet_(ir[k], ih, 'periodeid', 2)) !== String(pid)) continue;
        var c = {
          ir: Number(hGet_(ir[k], ih, 'ir', 5)) || 0,
          ir10: Number(hGet_(ir[k], ih, 'ir10', 6)) || 0,
          cicilan: Number(hGet_(ir[k], ih, 'cicilan', 7)) || 0,
          infakdaerah: Number(hGet_(ir[k], ih, 'infakdaerah', 8)) || 0,
          index: Number(hGet_(ir[k], ih, 'index', 9)) || 0
        };
        var total = c.ir + c.ir10 + c.cicilan + c.infakdaerah + c.index;
        if (total <= 0) continue;
        var aid = String(hGet_(ir[k], ih, 'anggotaid', 3) || '');
        rows.push({ tanggal: toDateStr_(hGet_(ir[k], ih, 'tanggal', 4)), anggota: angMap[aid] || aid || '-',
          ir: c.ir, ir10: c.ir10, cicilan: c.cicilan, infakdaerah: c.infakdaerah, index: c.index, total: total });
        tot.ir += c.ir; tot.ir10 += c.ir10; tot.cicilan += c.cicilan; tot.infakdaerah += c.infakdaerah; tot.index += c.index; tot.total += total;
      }
    }
    rows.sort(function(x, y) { return String(x.tanggal).localeCompare(String(y.tanggal)); });
    return { success: true, periode: { id: pid, nama: periode.nama, status: periode.status }, rows: rows, total: tot, jumlah: rows.length };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

// Kepemilikan saldo: dari Total Kas saat ini, berapa MILIK Kelompok vs masih
// TITIPAN/kewajiban ke Desa & Daerah (hak yang belum disetor).
//   Milik Desa   = hak Desa periode − yang sudah disetor ke desa
//   Milik Daerah = hak Daerah periode − yang sudah disetor ke daerah (0 bila tak dilacak)
//   Milik Kelompok = Total Kas − Milik Desa − Milik Daerah
function getKepemilikanSaldo() {
  try {
    var auth = checkAuth();
    if (!auth.success) return { success: false, message: auth.message };
    var periode = getPeriodeAktif();
    var pid = periode ? periode.id : null;
    var saldo = saldoSaatIni_(pid, periode); // {tunai, bank}
    var totalKas = (Number(saldo.tunai) || 0) + (Number(saldo.bank) || 0);

    // Kewajiban yang BELUM disetor, dihitung PER POS dari realisasi nyata
    // (Pos Setoran). Dihitung per pos agar kelebihan setor di satu pos tidak
    // menutupi kekurangan pos lain.
    var sisaDesa = 0, sisaDaerah = 0, kewajiban = 0, sudahSetor = 0;
    var rincianPos = [];   // rincian kewajiban per pos (utk penjelasan titipan)
    if (periode) {
      try {
        var rs = getRekapSetoran();
        if (rs && rs.success) {
          (rs.data || []).forEach(function(p) {
            var t = Number(p.target) || 0, rl = Number(p.realisasi) || 0;
            kewajiban += t; sudahSetor += rl;
            var sisa = Math.max(0, t - rl);
            var nm = String(p.nama || '').toLowerCase();
            var rf = String(p.sumberRef || '').toLowerCase();
            var keDaerah = (rf === 'infakdaerah' || nm.indexOf('daerah') !== -1);
            if (keDaerah) sisaDaerah += sisa; else sisaDesa += sisa;
            rincianPos.push({ nama: p.nama, tujuan: keDaerah ? 'Daerah' : 'Desa',
              kewajiban: t, disetor: rl, sisa: sisa, lebih: Math.max(0, rl - t), dasar: p.sumberKet || '' });
          });
        }
      } catch(e) {}
    }

    // Buku IR yang BELUM DIRINCI = kewajiban yang belum bisa dihitung per komponen.
    // Selama belum dirinci, uangnya belum boleh dianggap milik kelompok.
    var belumDirinci = 0;
    if (periode) {
      try {
        var ssb = getSS_();
        var kodeBukuIR = {};
        var shMb = ssb.getSheetByName(CONFIG.SHEETS.PEMASUKAN);
        if (shMb && shMb.getLastRow() > 1) {
          var mrb = shMb.getDataRange().getValues(); var mhb = headerMap_(mrb[0]);
          for (var m = 1; m < mrb.length; m++) {
            var kd = String(hGet_(mrb[m], mhb, 'kode', 0) || '');
            if (kd && String(hGet_(mrb[m], mhb, 'kategori', 2) || '').toLowerCase().indexOf('buku ir') !== -1) kodeBukuIR[kd] = true;
          }
        }
        var masukBukuIR = 0;
        var shPb = ssb.getSheetByName(CONFIG.SHEETS.INPUT_PENERIMAAN);
        if (shPb && shPb.getLastRow() > 1) {
          var prb = shPb.getDataRange().getValues(); var phb = headerMap_(prb[0]);
          for (var q = 1; q < prb.length; q++) {
            if (!hGet_(prb[q], phb, 'id', 0)) continue;
            if (barisDibatalkan_(prb[q], phb)) continue;
            if (String(hGet_(prb[q], phb, 'periodeid', 1)) !== String(pid)) continue;
            if (kodeBukuIR[String(hGet_(prb[q], phb, 'jenisid', 2) || '')]) masukBukuIR += Number(hGet_(prb[q], phb, 'nominal', 5)) || 0;
          }
        }
        var dirinci = 0;
        try { var ri = getRincianIR(pid); if (ri && ri.success && ri.total) dirinci = Number(ri.total.total) || 0; } catch(e) {}
        belumDirinci = Math.max(0, masukBukuIR - dirinci);
      } catch(e) {}
    }

    var milikDesa = sisaDesa + belumDirinci;
    var milikDaerah = sisaDaerah;
    var milikKelompok = totalKas - milikDesa - milikDaerah;

    // Rekonsiliasi Total Kas: saldo awal + pemasukan − pengeluaran.
    var saldoAwal = 0, totalMasuk = 0, totalKeluar = 0;
    if (periode) {
      saldoAwal = (Number(periode.saldoAwalTunai) || 0) + (Number(periode.saldoAwalBank) || 0);
      try {
        var ssr = getSS_();
        var shPm = ssr.getSheetByName(CONFIG.SHEETS.INPUT_PENERIMAAN);
        if (shPm && shPm.getLastRow() > 1) {
          var pm = shPm.getDataRange().getValues(); var pmh = headerMap_(pm[0]);
          for (var x = 1; x < pm.length; x++) {
            if (!hGet_(pm[x], pmh, 'id', 0) || barisDibatalkan_(pm[x], pmh)) continue;
            if (String(hGet_(pm[x], pmh, 'periodeid', 1)) !== String(pid)) continue;
            totalMasuk += Number(hGet_(pm[x], pmh, 'nominal', 5)) || 0;
          }
        }
        var shPk = ssr.getSheetByName(CONFIG.SHEETS.INPUT_PENGELUARAN);
        if (shPk && shPk.getLastRow() > 1) {
          var pk = shPk.getDataRange().getValues(); var pkh = headerMap_(pk[0]);
          for (var y = 1; y < pk.length; y++) {
            if (!hGet_(pk[y], pkh, 'id', 0) || barisDibatalkan_(pk[y], pkh) || barisDraft_(pk[y], pkh)) continue;
            if (String(hGet_(pk[y], pkh, 'periodeid', 1)) !== String(pid)) continue;
            totalKeluar += Number(hGet_(pk[y], pkh, 'nominal', 4)) || 0;
          }
        }
      } catch(e) {}
    }

    // Sumber hak kelompok dari pemasukan periode ini (konteks, bukan penjumlahan
    // langsung ke Milik Kelompok — karena sebagian sudah terpakai untuk pengeluaran).
    var sumberKelompok = [], hakKelompokPeriode = 0;
    if (periode) {
      try {
        var bh2 = getBagiHasil(pid);
        if (bh2 && bh2.success) {
          hakKelompokPeriode = Number(bh2.kelompok) || 0;
          (bh2.rincian || []).forEach(function(x) {
            if ((Number(x.kelompok) || 0) > 0) sumberKelompok.push({ jenis: x.jenis, nilai: x.kelompok });
          });
          sumberKelompok.sort(function(a, b) { return b.nilai - a.nilai; });
        }
      } catch(e) {}
    }

    return {
      success: true,
      periode: periode ? { id: pid, nama: periode.nama, status: periode.status } : null,
      totalKas: totalKas,
      milikKelompok: milikKelompok,
      milikDesa: milikDesa,
      milikDaerah: milikDaerah,
      kewajiban: kewajiban,
      sudahSetor: sudahSetor,
      sisaKewajiban: sisaDesa + sisaDaerah,
      sisaDesa: sisaDesa,
      sisaDaerah: sisaDaerah,
      belumDirinci: belumDirinci,
      kelompokNegatif: milikKelompok < 0,
      // Rincian untuk halaman penjelasan
      saldoAwal: saldoAwal, totalMasuk: totalMasuk, totalKeluar: totalKeluar,
      rincianPos: rincianPos,
      sumberKelompok: sumberKelompok,
      hakKelompokPeriode: hakKelompokPeriode
    };
  } catch(e) {
    return { success: false, message: e.message };
  }
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
        if (barisDibatalkan_(pRows[i], pH)) continue; // T3
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
        if (barisDibatalkan_(pRows[i], pH)) continue; // T3
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
      var setBatalIR = trxPenerimaanDibatalkan_();          // ← K1
      var irRows = sheetIR.getDataRange().getValues();
      var irH = headerMap_(irRows[0]);
      for (var i = 1; i < irRows.length; i++) {
        if (!irRows[i][0]) continue;
        if (rincianYatim_(irRows[i], irH, setBatalIR)) continue;   // ← K1
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
        if (barisDibatalkan_(pRows[i], pH)) continue; // T3
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
    // FASE 1: endpoint mutasi wajib requirePerm (sebelumnya hanya checkAuth).
    var auth = requirePerm('pembelaan.manage');
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
    // FASE 1: endpoint mutasi wajib requirePerm (sebelumnya hanya checkAuth).
    var auth = requirePerm('pembelaan.manage');
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
        if (barisDibatalkan_(pRows[i], pH)) continue; // T3
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
    return withLock_(function() {
    var ss = getSS_();
    var periode = getPeriodeAktif();
    if (!periode) return { success: false, message: 'Tidak ada periode aktif' };

    var nomCek = validasiNominal_(data.nominal);
    if (!nomCek.ok) return { success: false, message: nomCek.message };
    var tglCek = validasiTanggalPeriode_(data.tanggal ? data.tanggal : new Date(), periode.id);
    if (!tglCek.ok) return { success: false, message: tglCek.message };

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
    var tgl = tglCek.tgl;
    sheet.appendRow([id, periode.id, tgl, data.jenisId, data.anggotaId || '', nomCek.nilai, data.sumberKas || 'Tunai', data.catatan || '', auth.user.email, 'Aktif', '', toDateStr_(new Date())]);
    logActivity(auth.user.email, 'KAS_PENEROBOS', 'Nominal: ' + nomCek.nilai);
    return { success: true, id: id };
    });
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
      if (!isAdmin && !emailSama_(email, auth.user.email)) continue;
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
    var diag = _diagnosaSerahTerima_(periode);
    return { success: true, data: result, periode: periode, dapatSerahTerima: diag.bisa, alasanSerahTerima: diag.alasan };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

// L2d: jelaskan kenapa serah terima mungkin tak bisa diproses — supaya UI tidak
// sekadar menyembunyikan tombol tanpa alasan.
function _diagnosaSerahTerima_(periode) {
  if (!periode || String(periode.status).toUpperCase() !== CONFIG.STATUS.OPEN) {
    return { bisa: false, alasan: 'Periode sedang tidak terbuka. Buka periode baru agar serah terima bisa diproses dan dikonfirmasi bendahara.' };
  }
  // Harus ada bendahara/admin aktif yang bisa mengonfirmasi (role cocok Sumber Tujuan).
  var adaTunai = false, adaBank = false;
  try {
    getUserList_().forEach(function(u) {
      if (String(u.status || 'Aktif').toLowerCase() === 'nonaktif') return;
      if (u.role === CONFIG.ROLES.ADMIN) { adaTunai = true; adaBank = true; }
      if (u.role === CONFIG.ROLES.BENDAHARA_1) adaTunai = true;
      if (u.role === CONFIG.ROLES.BENDAHARA_2) adaBank = true;
    });
  } catch(e) {}
  if (!adaTunai && !adaBank) {
    return { bisa: false, alasan: 'Belum ada Bendahara/Admin aktif yang bisa mengonfirmasi serah terima. Tambahkan pengurus dengan role Bendahara 1 (Tunai) / Bendahara 2 (Bank).' };
  }
  var catatan = '';
  if (!adaTunai) catatan = 'Catatan: belum ada penerima Tunai (Bendahara 1).';
  else if (!adaBank) catatan = 'Catatan: belum ada penerima Bank (Bendahara 2).';
  return { bisa: true, alasan: catatan };
}

// K-cleanup: batalkan baris Kas Penerobos yang nyangkut agar tidak mengunci tutup buku.
// Hanya baris berstatus Aktif yang belum diserahterimakan. Wajib alasan.
// L2c: halaman Kelola Kas Penerobos (Admin) — READ-ONLY.
// Menampilkan SELURUH baris (semua periode & status) + penilaian risiko duplikat.
function getKelolaKasPenerobos() {
  var auth = requirePerm('penerobos.kelola');
  if (!auth.success) return { success: false, message: auth.message };
  try {
    var ss = getSS_();
    var analisis = deteksiDuplikatPenerobos_();
    // Peta nama jenis & anggota & status periode.
    var jenisMap = {}, angMap = {}, periodeStatus = {};
    var shM = ss.getSheetByName(CONFIG.SHEETS.PEMASUKAN);
    if (shM) { var mr = shM.getDataRange().getValues(); for (var i = 1; i < mr.length; i++) if (mr[i][0]) jenisMap[String(mr[i][0])] = String(mr[i][1] || ''); }
    var shA = ss.getSheetByName(CONFIG.SHEETS.ANGGOTA);
    if (shA) { var ar = shA.getDataRange().getValues(); for (var j = 1; j < ar.length; j++) if (ar[j][0]) angMap[String(ar[j][0])] = String(ar[j][1] || ''); }
    var shP = ss.getSheetByName(CONFIG.SHEETS.PERIOD);
    if (shP) { var pr = shP.getDataRange().getValues(); var ph = headerMap_(pr[0]); for (var k = 1; k < pr.length; k++) if (pr[k][0]) periodeStatus[String(pr[k][0])] = String(hGet_(pr[k], ph, 'status', 4) || '').toUpperCase(); }
    var agingHari = 7;
    try { var av = PropertiesService.getScriptProperties().getProperty('AGING_HARI'); if (av) agingHari = Number(av) || 7; } catch(e) {}
    var data = analisis.map(function(a) {
      a.jenis = jenisMap[a.jenisId] || a.jenisId;
      a.anggota = angMap[a.anggotaId] || '';
      a.periodeStatus = periodeStatus[a.periodeId] || '';
      a.periodeTertutup = (a.periodeStatus === CONFIG.STATUS.CLOSED);
      return a;
    });
    data.sort(function(x, y) { return String(y.tanggal).localeCompare(String(x.tanggal)); });
    return { success: true, data: data, agingHari: agingHari };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

// L2c: Batalkan baris Kas Penerobos (duplikat / salah catat / uang batal). SOFT DELETE.
// Baris TIDAK dihapus. Wajib alasan (≥15 char); risiko BAHAYA/PERLU_DIKONFIRMASI
// wajib ketik ulang nominal. Boleh di periode CLOSED (Kas Penerobos tak pernah masuk
// calculateSaldo) — tapi ditandai PRIVILEGED untuk direviu ketua kelompok.
function batalkanKasPenerobos(id, alasan, konfirmasiNominal) {
  var auth = requirePerm('penerobos.kelola');
  if (!auth.success) return { success: false, message: auth.message };
  if (!id) return { success: false, message: 'ID wajib.' };
  var alasanBersih = String(alasan || '').trim();
  if (alasanBersih.length < 15) return { success: false, message: 'Alasan pembatalan wajib diisi (minimal 15 karakter).' };
  return withLock_(function() {
    var ss = getSS_();
    var sheet = ss.getSheetByName(CONFIG.SHEETS.KAS_PENEROBOS);
    if (!sheet || sheet.getLastRow() < 2) return { success: false, message: 'Data Kas Penerobos tidak ditemukan.' };
    // Pastikan kolom pembatalan tersedia (idempoten, tidak menghapus data lama).
    ensureColumns_(sheet, ['Dibatalkan By', 'Dibatalkan At', 'Alasan Batal']);
    var rows = sheet.getDataRange().getValues();
    var h = headerMap_(rows[0]);
    for (var i = 1; i < rows.length; i++) {
      if (String(hGet_(rows[i], h, 'id', 0)) !== String(id)) continue;
      var status = String(hGet_(rows[i], h, 'status', 9) || 'Aktif');
      if (status === 'Diserahkan') return { success: false, message: 'Sudah diserahterimakan — tidak bisa dibatalkan dari sini.' };
      if (status !== 'Aktif') return { success: false, message: 'Baris ini sudah berstatus ' + status + '.' };
      var pidBaris = String(hGet_(rows[i], h, 'periodeid', 1) || '');
      // Periode CLOSED BOLEH (kas penerobos tak masuk calculateSaldo) → tandai PRIVILEGED.
      var periodeTertutup = false;
      var apKP = assertPeriodeOpen_(pidBaris);
      if (!apKP.ok) periodeTertutup = true;
      // Penilaian risiko dari deteksi.
      var risiko = 'PERLU_DITELITI';
      try { var an = deteksiDuplikatPenerobos_(); for (var z = 0; z < an.length; z++) if (an[z].id === String(id)) { risiko = an[z].risiko; break; } } catch(e) {}
      var nominal = Number(hGet_(rows[i], h, 'nominal', 5)) || 0;
      if (risiko === 'BAHAYA_UANG_HILANG' || risiko === 'PERLU_DIKONFIRMASI') {
        if (Number(konfirmasiNominal) !== nominal) {
          return { success: false, butuhKonfirmasiNominal: true, nominal: nominal, risiko: risiko,
            message: (risiko === 'BAHAYA_UANG_HILANG'
              ? 'Tidak ada jejak transaksi sumber. Membatalkannya berarti uang Rp ' + nominal.toLocaleString('id-ID') + ' hilang total dari sistem.'
              : 'Baris ini kemungkinan sudah dicatat langsung oleh bendahara. Cocokkan dulu dengan pasangannya sebelum membatalkan (uang Rp ' + nominal.toLocaleString('id-ID') + ').') +
              ' Ketik ulang nominal untuk konfirmasi.' };
        }
      }
      var r = i + 1;
      var snapshot = JSON.stringify({
        id: id, periodeId: pidBaris, nominal: nominal,
        jenisId: String(hGet_(rows[i], h, 'jenisid', 3) || ''), anggotaId: String(hGet_(rows[i], h, 'anggotaid', 4) || ''),
        sumberKas: String(hGet_(rows[i], h, 'sumberkas', 6) || ''), penerobosEmail: String(hGet_(rows[i], h, 'penerobosemail', 8) || '')
      });
      sheet.getRange(r, h['status'] + 1).setValue('Dibatalkan');
      if (h['dibatalkanby'] !== undefined) sheet.getRange(r, h['dibatalkanby'] + 1).setValue(auth.user.email);
      if (h['dibatalkanat'] !== undefined) sheet.getRange(r, h['dibatalkanat'] + 1).setValue(toDateStr_(new Date()));
      if (h['alasanbatal'] !== undefined) sheet.getRange(r, h['alasanbatal'] + 1).setValue(alasanBersih);
      // Aksi di periode tertutup ditandai PRIVILEGED agar muncul di Aktivitas Istimewa.
      var aksiLog = periodeTertutup ? 'PRIVILEGED_BATAL_KAS_PENEROBOS_TUTUP' : 'BATAL_KAS_PENEROBOS';
      logActivityWajib_(auth.user.email, aksiLog, 'ID: ' + id + ' | ALASAN: ' + alasanBersih + ' | RISIKO: ' + risiko + ' | NILAI: ' + snapshot);
      try { CacheService.getScriptCache().remove('dashboard_saldo'); } catch(e) {}
      return { success: true, periodeTertutup: periodeTertutup };
    }
    return { success: false, message: 'Kas Penerobos tidak ditemukan.' };
  });
}

// L2c: Aktivitas Istimewa — entri Activity Log yang bersifat PRIVILEGED (untuk direviu
// ketua kelompok). READ-ONLY. Butuh penerobos.kelola atau user.manage.
function getAktivitasIstimewa(limit) {
  var auth = getCurrentUser();
  if (!auth) return { success: false, message: 'Belum login' };
  if (!userCan_(auth.role, 'penerobos.kelola') && !userCan_(auth.role, 'user.manage')) {
    return { success: false, message: 'Akses ditolak' };
  }
  try {
    var lim = Number(limit) || 200;
    var ss = getSS_();
    var sheet = ss.getSheetByName(CONFIG.SHEETS.LOG);
    if (!sheet || sheet.getLastRow() < 2) return { success: true, data: [] };
    var rows = sheet.getDataRange().getValues();
    var h = headerMap_(rows[0]);
    var out = [];
    for (var i = rows.length - 1; i >= 1 && out.length < lim; i--) {
      var aksi = String(hGet_(rows[i], h, 'action', 2) || '');
      if (aksi.indexOf('PRIVILEGED') !== 0) continue;
      out.push({
        waktu: toDateStr_(hGet_(rows[i], h, 'timestamp', 0)),
        email: String(hGet_(rows[i], h, 'user', 1) || ''),
        aksi: aksi,
        detail: String(hGet_(rows[i], h, 'detail', 3) || '')
      });
    }
    return { success: true, data: out };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

// L2c/L0: Pemeriksaan integritas ringkas — saat ini fokus duplikat Kas Penerobos.
function cekIntegritas() {
  var auth = requirePerm('penerobos.kelola');
  if (!auth.success) return { success: false, message: auth.message };
  try {
    var analisis = deteksiDuplikatPenerobos_();
    var kandidat = analisis.filter(function(a) { return a.status === 'Aktif' && a.jenisDuplikat !== 'BUKAN'; });
    var bahaya = kandidat.filter(function(a) { return a.risiko === 'BAHAYA_UANG_HILANG'; });
    var konfirmasi = kandidat.filter(function(a) { return a.risiko === 'PERLU_DIKONFIRMASI'; });
    var aman = kandidat.filter(function(a) { return a.risiko === 'AMAN_DIBATALKAN'; });
    var kategori = [{
      kode: 'DUPLIKAT_PENEROBOS',
      judul: 'Kandidat catatan ganda Kas Penerobos',
      jumlah: kandidat.length,
      ringkas: kandidat.length
        ? (aman.length + ' aman dibatalkan, ' + konfirmasi.length + ' perlu dikonfirmasi, ' + bahaya.length + ' BAHAYA (uang bisa hilang) — tinjau di Kelola Kas Penerobos.')
        : 'Tidak ada kandidat catatan ganda.',
      detail: kandidat
    }];
    return { success: true, kategori: kategori };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function buatSerahTerima(data) {
  try {
    var auth = requirePerm('penerobos.input');
    if (!auth.success) return { success: false, message: auth.message };
    return withLock_(function() {
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
        if (!isAdmin && !emailSama_(email, auth.user.email)) return { success: false, message: 'Item bukan milik Anda: ' + targetId };
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
    });
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
      if (isPenerobos && !emailSama_(email, auth.user.email)) continue;
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
    return withLock_(function() {
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
          catatan:  String(hGet_(kpRows[ki], kpH, 'catatan', 7) || ''),
          tanggal:  toDateStr_(hGet_(kpRows[ki], kpH, 'tanggal', 2)) // T8: tanggal ASLI terima dari jamaah
        });
      }
    }
    if (items.length === 0) return { success: false, message: 'Tidak ada item kas penerobos untuk serah terima ini' };

    // T8 cut-off: tanggal asli penerimaan wajib berada dalam rentang periode aktif.
    var periodeAktif = getPeriodeAktif();
    for (var vj = 0; vj < items.length; vj++) {
      var vc = validasiTanggalPeriode_(items[vj].tanggal, periodeAktif ? periodeAktif.id : stData.periodeId);
      if (!vc.ok) return { success: false, message: 'Tanggal terima item (' + (items[vj].tanggal || '-') + ') di luar rentang periode aktif — ' + vc.message + ' Lakukan koreksi lewat prosedur, jangan geser tanggal.' };
    }

    // Buat INPUT_PENERIMAAN per item — Tanggal = tanggal ASLI terima; Created At = tanggal konfirmasi.
    var sheetP = ss.getSheetByName(CONFIG.SHEETS.INPUT_PENERIMAAN);
    if (!sheetP) return { success: false, message: 'Sheet penerimaan tidak ditemukan' };
    var now = toDateStr_(new Date());
    var penIds = [];
    for (var ji = 0; ji < items.length; ji++) {
      var penId = generateID('TRX');
      var catatanEntry = 'Serah Terima dari: ' + stData.penerobosEmail + (items[ji].catatan ? ' | ' + items[ji].catatan : '');
      sheetP.appendRow([penId, stData.periodeId, items[ji].jenisId, items[ji].anggotaId,
        items[ji].tanggal || now, items[ji].nominal, stData.sumberTujuan, catatanEntry, auth.user.email, now, 'Aktif']);
      _isiNoBukti_(sheetP, 'BKM', stData.periodeId);
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
    });
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
      if (userList[u].role === CONFIG.ROLES.PENEROBOS) penerobosEmails[String(userList[u].email || '').toLowerCase().trim()] = true;
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
      if (!penerobosEmails[createdBy.toLowerCase().trim()]) continue;
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

// ══════════════════════════════════════════════════════
// INPUT REKAP SETORAN — banyak setoran jamaah sekaligus
// Desain: docs/DESAIN-INPUT-REKAP.md (dikunci 24-10-2026)
//
// Satu baris rekap = satu Penerimaan (jenis berkategori Buku IR) + satu
// baris Detail Buku IR. Hasilnya sama persis dengan input lewat dua form
// (submitTransaksi + submitRincianIR), karena keduanya memakai helper
// bersama _barisPenerimaan_ & _barisRincianIR_ (K10).
//
// Alur: mode 'periksa' (tanpa menulis) → pratinjau → mode 'simpan'
// (pemeriksaan diulang di server, lalu tulis semua atau tidak sama sekali).
// ══════════════════════════════════════════════════════

// Pemetaan judul kolom buku → komponen rincian Buku IR (K1).
// Sumber kebenaran tunggal pemetaan; kunci = komponenRincianIR_().kunci.
function kolomRekapSetoran_() {
  return [
    { judul: 'ir',        kunci: 'ir',          label: 'IR' },
    { judul: 'cilepus',   kunci: 'cicilan',     label: 'Cicilan' },
    { judul: 'daerah',    kunci: 'infakDaerah', label: 'Infak Daerah' },
    { judul: 'klp',       kunci: 'index',       label: 'Index' },
    { judul: 'tunggakan', kunci: 'sewaKos',     label: 'Sewa Kos' }
  ];
}

var REKAP_MAKS_BARIS_ = 200;   // K13

// ──────────────────────────────────────────────────────
// ENDPOINT
// payload = {
//   mode: 'periksa' | 'simpan',
//   tanggal: 'YYYY-MM-DD',
//   jenisId: '<kode jenis Buku IR>',
//   kodeRekap: 'REKAP-2026-10-23',
//   teks: '<tabel yang ditempel dari spreadsheet/buku>',
//   pilihanAnggota: { '<no baris>': '<anggotaId>' }   // pilihan manual (K5)
// }
// ──────────────────────────────────────────────────────
function importRekapSetoran(payload) {
  try {
    var auth = authRekapSetoran_();
    if (!auth.success) return auth;
    payload = payload || {};
    var mode = (payload.mode === 'simpan') ? 'simpan' : 'periksa';

    if (mode === 'periksa') return periksaRekapSetoran_(payload);

    return withLock_(function() {
      // Pemeriksaan diulang DI DALAM lock: hasil periksa dari klien tidak
      // dipercaya, dan cek kode rekap (K8) harus atomik dengan penulisan.
      var hasil = periksaRekapSetoran_(payload);
      if (!hasil.success) return hasil;
      if (!hasil.bisaSimpan) {
        return { success: false, message: 'Masih ada data yang salah. Periksa ulang rekap sebelum menyimpan.', hasil: hasil };
      }
      return simpanRekapSetoran_(hasil, auth.user.email);
    });
  } catch(e) {
    return { success: false, message: e.message };
  }
}

// K11: butuh izin input pemasukan + rincian Buku IR. PENEROBOS tidak boleh,
// karena pemasukan penerobos masuk Kas Penerobos, bukan kas utama.
function authRekapSetoran_() {
  var auth = requirePermInput_('masuk');
  if (!auth.success) return { success: false, message: auth.message };
  if (auth.user.role === CONFIG.ROLES.PENEROBOS) return { success: false, message: 'Akses ditolak: Input Rekap tidak tersedia untuk peran Penerobos.' };
  if (!userCan_(auth.user.role, 'bukuIR.input')) return { success: false, message: 'Akses ditolak: tidak berwenang input rincian Buku IR.' };
  return auth;
}

// ──────────────────────────────────────────────────────
// PERIKSA — tidak menulis apa pun
// ──────────────────────────────────────────────────────
function periksaRekapSetoran_(payload) {
  var galat = [];        // galat tingkat rekap (membuat Simpan mati)
  var peringatanRekap = [];

  var periode = getPeriodeAktif();
  if (!periode) return { success: false, message: 'Tidak ada periode aktif.' };

  // Tanggal (K7)
  var tgl = '';
  if (!payload.tanggal) galat.push('Tanggal rekap belum diisi.');
  else {
    var tc = validasiTanggalPeriode_(payload.tanggal, periode.id);
    if (!tc.ok) galat.push(tc.message); else tgl = tc.tgl;
  }

  // Jenis pemasukan berkategori Buku IR (K6)
  var daftarJenis = jenisBukuIRAktif_();
  var jenisId = String(payload.jenisId || '');
  if (!daftarJenis.length) galat.push('Belum ada Jenis Pemasukan berkategori "Buku IR" di Master Pemasukan.');
  else if (!jenisId) galat.push('Jenis Pemasukan belum dipilih.');
  else if (!daftarJenis.some(function(j) { return j.id === jenisId; })) galat.push('Jenis Pemasukan harus berkategori "Buku IR".');

  // Kode rekap (K8)
  var kodeRekap = String(payload.kodeRekap || '').trim();
  if (!kodeRekap) galat.push('Kode Rekap belum diisi.');
  else if (kodeRekapSudahDipakai_(kodeRekap)) galat.push('Kode Rekap "' + kodeRekap + '" sudah pernah disimpan. Rekap ini kemungkinan sudah diinput.');

  // Baca tabel tempelan
  var tabel = uraiTeksRekap_(String(payload.teks || ''));
  if (tabel.galat) galat.push(tabel.galat);
  if (tabel.baris.length > REKAP_MAKS_BARIS_) galat.push('Maksimal ' + REKAP_MAKS_BARIS_ + ' baris per rekap (ditempel ' + tabel.baris.length + ').');

  // Kolom komponen di sheet Detail Buku IR
  var irRows = sheetValues_(CONFIG.SHEETS.BUKU_IR);
  var kompAda = irRows.length ? komponenRincianAda_(headerMap_(irRows[0])) : [];
  var kunciAda = {};
  kompAda.forEach(function(k) { kunciAda[k.kunci] = true; });
  if (!irRows.length) galat.push('Sheet "' + CONFIG.SHEETS.BUKU_IR + '" belum ada.');

  // Anggota
  var anggota = daftarAnggotaRekap_();
  var pilihan = payload.pilihanAnggota || {};

  // Hitung kemunculan nama untuk peringatan nama dobel (K4)
  var hitungNama = {};
  tabel.baris.forEach(function(b) { var k = normNama_(b.nama); if (k) hitungNama[k] = (hitungNama[k] || 0) + 1; });

  var kolom = kolomRekapSetoran_();
  var baris = [];
  var ringkas = {
    jumlahBaris: 0, siap: 0, peringatan: 0, ditahan: 0, salah: 0,
    perKomponen: {}, totalTunai: 0, totalBank: 0, totalDitahan: 0, totalDisimpan: 0, totalSemua: 0,
    grandTotalBuku: tabel.grandTotal, cocokGrandTotal: null
  };
  kolom.forEach(function(c) { ringkas.perKomponen[c.kunci] = 0; });

  tabel.baris.forEach(function(r) {
    var pesan = [], salah = false, warn = false, tahan = false;

    // Komponen + angka
    var komponen = {}, jumlahKomp = 0;
    kolom.forEach(function(c) {
      var v = r.angka[c.kunci];
      if (v === null) { salah = true; pesan.push('Angka ' + c.label + ' tidak terbaca.'); v = 0; }
      komponen[c.kunci] = v;
      jumlahKomp += v;
      if (v > 0 && !kunciAda[c.kunci]) { salah = true; pesan.push('Kolom ' + c.label + ' belum ada di sheet Detail Buku IR (jalankan migrasi).'); }
    });
    var total = r.angka.total;
    if (total === null) { salah = true; pesan.push('Total tidak terbaca.'); total = 0; }
    else {
      var nc = validasiNominal_(total);
      if (!nc.ok) { salah = true; pesan.push('Total: ' + nc.message); }
      if (jumlahKomp !== total) {
        salah = true;
        pesan.push('Jumlah komponen Rp ' + jumlahKomp.toLocaleString('id-ID') + ' ≠ Total Rp ' + total.toLocaleString('id-ID') + '.');
      }
    }

    // Sumber kas (K2, K3)
    var sk = sumberKasDariKet_(r.ket);
    if (sk.galat) { salah = true; pesan.push(sk.galat); }
    if (sk.tahan) { tahan = true; pesan.push('Cara bayar perlu dikonfirmasi — baris ditahan, tidak disimpan.'); }

    // Anggota (K5)
    var cocok = cocokkanAnggota_(r.nama, anggota);
    var anggotaId = '', anggotaNama = '', caraCocok = '';
    var pilih = String(pilihan[String(r.no)] || '');
    if (pilih) {
      var a = anggota.filter(function(x) { return x.id === pilih; })[0];
      if (a) { anggotaId = a.id; anggotaNama = a.nama; caraCocok = 'manual'; }
      else { salah = true; pesan.push('Anggota pilihan tidak ditemukan.'); }
    } else if (cocok.tepat) {
      anggotaId = cocok.tepat.id; anggotaNama = cocok.tepat.nama; caraCocok = cocok.cara;
    } else if (!tahan) {
      salah = true;
      pesan.push(cocok.kandidat.length > 1
        ? 'Nama cocok dengan ' + cocok.kandidat.length + ' anggota — pilih salah satu.'
        : 'Nama tidak ditemukan di daftar anggota — pilih manual.');
    }

    if (!r.nama) { salah = true; pesan.push('Nama kosong.'); }
    if (hitungNama[normNama_(r.nama)] > 1) { warn = true; pesan.push('Nama ini muncul lebih dari sekali di rekap.'); }

    var status = salah ? 'salah' : (tahan ? 'ditahan' : (warn ? 'peringatan' : 'siap'));
    baris.push({
      no: r.no, namaBuku: r.nama, anggotaId: anggotaId, anggotaNama: anggotaNama, caraCocok: caraCocok,
      kandidat: (anggotaId ? [] : cocok.kandidat.map(function(k) { return { id: k.id, nama: k.nama }; })),
      komponen: komponen, total: total, ket: r.ket, sumberKas: sk.sumberKas || '', catatan: r.catatan,
      status: status, pesan: pesan
    });

    ringkas.jumlahBaris++;
    ringkas[status]++;
    ringkas.totalSemua += total;
    if (status === 'ditahan') { ringkas.totalDitahan += total; return; }
    if (status === 'salah') return;
    ringkas.totalDisimpan += total;
    if (sk.sumberKas === 'Bank') ringkas.totalBank += total; else ringkas.totalTunai += total;
    kolom.forEach(function(c) { ringkas.perKomponen[c.kunci] += komponen[c.kunci]; });
  });

  if (!tabel.baris.length && !tabel.galat) galat.push('Tidak ada baris jamaah yang terbaca dari teks yang ditempel.');
  if (tabel.grandTotal !== null) {
    ringkas.cocokGrandTotal = (tabel.grandTotal === ringkas.totalSemua);
    if (!ringkas.cocokGrandTotal) {
      peringatanRekap.push('Jumlah semua baris Rp ' + ringkas.totalSemua.toLocaleString('id-ID') + ' berbeda dengan GRAND TOTAL buku Rp ' + tabel.grandTotal.toLocaleString('id-ID') + '.');
    }
  }
  var adaYangDisimpan = (ringkas.siap + ringkas.peringatan) > 0;
  if (tabel.baris.length && !adaYangDisimpan && ringkas.salah === 0) galat.push('Semua baris ditahan — tidak ada yang bisa disimpan.');

  return {
    success: true,
    bisaSimpan: galat.length === 0 && ringkas.salah === 0 && adaYangDisimpan,
    galat: galat, peringatan: peringatanRekap,
    periode: { id: periode.id, nama: periode.nama },
    tanggal: tgl, jenisId: jenisId, kodeRekap: kodeRekap,
    kolom: kolom, baris: baris, ringkasan: ringkas,
    daftarJenis: daftarJenis,
    daftarAnggota: anggota.map(function(a) { return { id: a.id, nama: a.nama }; })
  };
}

// ──────────────────────────────────────────────────────
// SIMPAN — dipanggil di dalam withLock_ setelah periksa lolos
// ──────────────────────────────────────────────────────
function simpanRekapSetoran_(hasil, email) {
  var ss = getSS_();
  var periodeId = hasil.periode.id;
  var now = toDateStr_(new Date());
  var simpan = hasil.baris.filter(function(b) { return b.status === 'siap' || b.status === 'peringatan'; });

  // ── 1. Input Penerimaan ──
  var shP = ss.getSheetByName(CONFIG.SHEETS.INPUT_PENERIMAAN);
  if (!shP) return { success: false, message: 'Sheet penerimaan tidak ditemukan.' };
  ensureColumns_(shP, ['Kode Rekap']);
  var lastColP = shP.getLastColumn();
  var dataP = shP.getRange(1, 1, Math.max(1, shP.getLastRow()), lastColP).getValues();
  var hP = headerMap_(dataP[0]);
  var cNB = hP['nobukti'], cKode = hP['koderekap'];
  var seq = (cNB !== undefined) ? _maxSeqNoBukti_(dataP, hP, periodeId) : 0;
  var pseq = (cNB !== undefined) ? _periodeSeqCode_(ss, periodeId) : 0;

  var barisP = [], trxIds = [], noBukti = [];
  simpan.forEach(function(b) {
    var id = generateID('TRX');
    trxIds.push(id);
    var catatan = 'Rekap ' + hasil.kodeRekap + ' #' + b.no + ' (' + b.namaBuku + ')' + (b.catatan ? ' — ' + b.catatan : '');
    var row = _barisPenerimaan_({
      id: id, periodeId: periodeId, jenisId: hasil.jenisId, anggotaId: b.anggotaId,
      tanggal: hasil.tanggal, nominal: b.total, sumberKas: b.sumberKas, catatan: catatan,
      email: email, createdAt: now
    });
    while (row.length < lastColP) row.push('');
    if (cNB !== undefined) {
      seq++;
      var nb = 'BKM-' + pseq + '-' + ('0000' + seq).slice(-4);
      row[cNB] = nb; noBukti.push(nb);
    }
    row[cKode] = hasil.kodeRekap;
    barisP.push(row);
  });
  shP.getRange(shP.getLastRow() + 1, 1, barisP.length, lastColP).setValues(barisP);

  // ── 2. Detail Buku IR ──
  var shIR = ss.getSheetByName(CONFIG.SHEETS.BUKU_IR);
  var lastColIR = shIR.getLastColumn();
  var irHeader = shIR.getRange(1, 1, 1, lastColIR).getValues()[0];
  var irH = headerMap_(irHeader);
  var komp = komponenRincianAda_(irH);
  var barisIR = simpan.map(function(b, i) {
    var row = _barisRincianIR_(irHeader.length, irH, komp, {
      id: generateID('IR'), transaksiId: trxIds[i], periodeId: periodeId, anggotaId: b.anggotaId,
      tanggal: hasil.tanggal, nilai: b.komponen, email: email, createdAt: now
    });
    if (irH['rapelbulan'] !== undefined) row[irH['rapelbulan']] = 1;   // K12
    return row;
  });
  shIR.getRange(shIR.getLastRow() + 1, 1, barisIR.length, lastColIR).setValues(barisIR);

  // ── 3. Jejak & cache ──
  logActivityWajib_(email, 'IMPOR_REKAP', 'Kode: ' + hasil.kodeRekap + ' | tanggal: ' + hasil.tanggal +
    ' | baris: ' + simpan.length + ' | total: ' + hasil.ringkasan.totalDisimpan +
    ' | ditahan: ' + hasil.ringkasan.ditahan + ' (' + hasil.ringkasan.totalDitahan + ')');
  try {
    var c = CacheService.getScriptCache();
    c.removeAll(['dashboard_saldo', 'master_trx_data', 'buku_ir_data', 'laporan_setoran_' + periodeId]);
  } catch(e) {}

  return {
    success: true,
    tersimpan: simpan.length,
    totalTersimpan: hasil.ringkasan.totalDisimpan,
    noBuktiAwal: noBukti[0] || '', noBuktiAkhir: noBukti[noBukti.length - 1] || '',
    ditahan: hasil.baris.filter(function(b) { return b.status === 'ditahan'; })
      .map(function(b) { return { no: b.no, nama: b.namaBuku, total: b.total, ket: b.ket }; }),
    kodeRekap: hasil.kodeRekap
  };
}

// ──────────────────────────────────────────────────────
// HELPER
// ──────────────────────────────────────────────────────

// Urai teks tempelan (tab dari spreadsheet, atau koma/titik koma).
// Kolom dibaca berdasarkan JUDUL. Baris GRAND TOTAL dibaca sebagai pembanding.
function uraiTeksRekap_(teks) {
  var hasil = { baris: [], grandTotal: null, galat: '' };
  var lines = String(teks || '').replace(/\r/g, '').split('\n');
  var pemisah = teks.indexOf('\t') !== -1 ? '\t' : (teks.indexOf(';') !== -1 && teks.indexOf(',') === -1 ? ';' : ',');
  var rows = lines.map(function(l) { return pecahBarisCSV_(l, pemisah); });

  // Cari baris judul: memuat 'nama' dan 'total'
  var hIdx = -1, map = {};
  for (var i = 0; i < rows.length; i++) {
    var m = {};
    rows[i].forEach(function(cel, j) { var k = String(cel).toLowerCase().replace(/[^a-z0-9]/g, ''); if (k && m[k] === undefined) m[k] = j; });
    if (m['nama'] !== undefined && m['total'] !== undefined) { hIdx = i; map = m; break; }
  }
  if (hIdx < 0) { hasil.galat = 'Baris judul tidak ditemukan. Pastikan ikut menyalin baris judul yang memuat kolom "Nama" dan "Total".'; return hasil; }

  var kolom = kolomRekapSetoran_();
  var nomor = 0;
  for (var r = hIdx + 1; r < rows.length; r++) {
    var row = rows[r];
    var ambil = function(key) { return map[key] !== undefined ? String(row[map[key]] || '').trim() : ''; };
    var nama = ambil('nama');
    var totalRaw = ambil('total');
    var semuaKosong = row.every(function(c) { return !String(c).trim(); });
    if (semuaKosong) continue;
    var namaN = nama.toLowerCase().replace(/[^a-z]/g, '');
    if (namaN === 'grandtotal' || namaN === 'total' || namaN === 'jumlah') {
      hasil.grandTotal = angkaRekap_(totalRaw);
      continue;
    }
    if (!nama && !totalRaw) continue;
    nomor++;
    var noBuku = parseInt(ambil('no'), 10);
    var angka = { total: angkaRekap_(totalRaw) };
    kolom.forEach(function(c) { angka[c.kunci] = angkaRekap_(ambil(c.judul)); });
    hasil.baris.push({ no: isFinite(noBuku) ? noBuku : nomor, nama: nama, angka: angka, ket: ambil('ket'), catatan: ambil('catatan') });
  }
  return hasil;
}

// Pecah satu baris CSV dengan dukungan tanda kutip ("a, b").
function pecahBarisCSV_(line, sep) {
  if (sep === '\t') return line.split('\t');
  var out = [], cur = '', q = false;
  for (var i = 0; i < line.length; i++) {
    var ch = line.charAt(i);
    if (q) {
      if (ch === '"' && line.charAt(i + 1) === '"') { cur += '"'; i++; }
      else if (ch === '"') q = false;
      else cur += ch;
    } else if (ch === '"') q = true;
    else if (ch === sep) { out.push(cur); cur = ''; }
    else cur += ch;
  }
  out.push(cur);
  return out;
}

// "1.000.000" / "Rp 1.000.000" / "1000000" / "" / "-" → angka; tak terbaca → null.
function angkaRekap_(s) {
  var t = String(s == null ? '' : s).trim().replace(/^rp\.?/i, '').replace(/[\s.,]/g, '');
  if (t === '' || t === '-') return 0;
  if (!/^\d+$/.test(t)) return null;
  return parseInt(t, 10);
}

// K2/K3: Ket → Sumber Kas.
function sumberKasDariKet_(ket) {
  var k = String(ket || '').trim().toUpperCase();
  if (k.indexOf('?') !== -1) return { tahan: true };
  if (k === '' || k === 'CASH' || k === 'TUNAI' || k === 'C') return { sumberKas: 'Tunai' };
  if (k === 'TF' || k === 'TRANSFER' || k === 'BANK' || k === 'T') return { sumberKas: 'Bank' };
  return { galat: 'Ket "' + ket + '" tidak dikenal (isi CASH, TF, kosong, atau ? bila perlu konfirmasi).' };
}

function normNama_(s) {
  return String(s || '').toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim();
}

// K5: (1) sama persis; (2) satu-satunya anggota yang namanya memuat nama buku
// sebagai kata utuh (nama panggilan → nama lengkap). Lebih dari satu → pilih manual.
function cocokkanAnggota_(namaBuku, anggota) {
  var n = normNama_(namaBuku);
  if (!n) return { tepat: null, kandidat: [] };
  var persis = anggota.filter(function(a) { return a.norm === n; });
  if (persis.length === 1) return { tepat: persis[0], cara: 'persis', kandidat: persis };
  if (persis.length > 1) return { tepat: null, kandidat: persis };
  var kata = ' ' + n + ' ';
  var memuat = anggota.filter(function(a) { return (' ' + a.norm + ' ').indexOf(kata) !== -1; });
  if (memuat.length === 1) return { tepat: memuat[0], cara: 'panggilan', kandidat: memuat };
  return { tepat: null, kandidat: memuat };
}

// Anggota yang tidak berstatus nonaktif.
function daftarAnggotaRekap_() {
  var rows = sheetValues_(CONFIG.SHEETS.ANGGOTA);
  var out = [];
  for (var i = 1; i < rows.length; i++) {
    if (!rows[i][0]) continue;
    if (String(rows[i][4] || '').toLowerCase().trim() === 'nonaktif') continue;
    out.push({ id: String(rows[i][0]), nama: String(rows[i][1] || ''), norm: normNama_(rows[i][1]) });
  }
  out.sort(function(a, b) { return a.nama.localeCompare(b.nama); });
  return out;
}

// Jenis pemasukan aktif berkategori Buku IR (aturan yang sama dengan
// _tangguhkanRincianBukuIR_: kategori memuat 'buku ir').
function jenisBukuIRAktif_() {
  var rows = sheetValues_(CONFIG.SHEETS.PEMASUKAN);
  if (rows.length < 2) return [];
  var h = headerMap_(rows[0]);
  var out = [];
  for (var i = 1; i < rows.length; i++) {
    var kd = String(hGet_(rows[i], h, 'kode', 0) || '');
    if (!kd) continue;
    if (String(hGet_(rows[i], h, 'kategori', 2) || '').toLowerCase().indexOf('buku ir') === -1) continue;
    if (String(hGet_(rows[i], h, 'status', 3) || '').toLowerCase().trim() === 'nonaktif') continue;
    out.push({ id: kd, nama: String(rows[i][1] || '') });
  }
  return out;
}

// K8: kode rekap sudah dipakai oleh penerimaan yang masih aktif?
function kodeRekapSudahDipakai_(kode) {
  var rows = sheetValues_(CONFIG.SHEETS.INPUT_PENERIMAAN);
  if (rows.length < 2) return false;
  var h = headerMap_(rows[0]);
  if (h['koderekap'] === undefined) return false;
  var k = String(kode).trim().toLowerCase();
  for (var i = 1; i < rows.length; i++) {
    if (barisDibatalkan_(rows[i], h)) continue;
    if (String(rows[i][h['koderekap']] || '').trim().toLowerCase() === k) return true;
  }
  return false;
}

// ──────────────────────────────────────────────────────
// UJI DARI EDITOR APPS SCRIPT (G2) — mode PERIKSA saja, tidak menulis.
// Masuk denylist apiCall, jadi tidak bisa dipanggil dari aplikasi.
// Ganti JENIS_ID dengan kode jenis Buku IR di Master Pemasukan, lalu Run.
// Hasil di View → Logs (Ctrl+Enter).
// ──────────────────────────────────────────────────────
function ujiRekapSetoran() {
  var JENIS_ID = '';   // ← isi, mis. 'PM-001'. Kosong = pakai jenis Buku IR pertama.
  var teks = [
    'No\tTanggal\tNama\tIR\tCilepus\tDaerah\tKLP\tTunggakan\tTotal\tKet\tCatatan',
    '1\t23/10/2026\tHeru\t400.000\t540.000\t30.000\t30.000\t\t1.000.000\tCASH\tDicoret di buku, ikut grand total',
    '2\t23/10/2026\tHeru\t400.000\t540.000\t30.000\t30.000\t\t1.000.000\tCASH',
    '3\t23/10/2026\tHusni\t300.000\t\t\t\t\t300.000\tCASH',
    '12\t23/10/2026\tAli\t1.000.000\t200.000\t5.000\t100.000\t1.000.000\t2.305.000\tTF\tTunggakan 500.000 + 500.000',
    '27\t23/10/2026\tRoni\t3.000.000\t\t\t\t\t3.000.000\t?\tCara bayar perlu konfirmasi',
    '\t\tGRAND TOTAL\t\t\t\t\t\t7.605.000'
  ].join('\n');

  __REQ_USER_ = { email: Session.getEffectiveUser().getEmail(), role: CONFIG.ROLES.ADMIN, nama: 'Uji Editor' };
  try {
    var jenis = JENIS_ID || ((jenisBukuIRAktif_()[0] || {}).id || '');
    var r = importRekapSetoran({ mode: 'periksa', tanggal: toDateStr_(new Date()), jenisId: jenis, kodeRekap: 'UJI-' + new Date().getTime(), teks: teks });
    Logger.log('bisaSimpan: ' + r.bisaSimpan);
    Logger.log('galat: ' + JSON.stringify(r.galat));
    Logger.log('peringatan: ' + JSON.stringify(r.peringatan));
    Logger.log('ringkasan: ' + JSON.stringify(r.ringkasan));
    (r.baris || []).forEach(function(b) {
      Logger.log('#' + b.no + ' ' + b.namaBuku + ' → ' + (b.anggotaNama || '-') + ' [' + b.status + '] ' +
        JSON.stringify(b.komponen) + ' total=' + b.total + ' ' + b.sumberKas + ' | ' + b.pesan.join(' '));
    });
    return r;
  } finally {
    __REQ_USER_ = null;
  }
}

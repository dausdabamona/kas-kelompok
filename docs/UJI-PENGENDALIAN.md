# Uji Pengendalian Intern — Kas Kelompok

Checklist uji manual per fase. Jalankan di **deployment staging** setelah
`clasp push` + **Deploy versi baru**. Sebelum uji Fase 1, jalankan sekali dari editor:
**Run → `migrasiPengendalian`** (menambah kolom `Status`/`Dibatalkan By`/`Dibatalkan At`/`Alasan Batal`).

---

## FASE 1 — Kunci Integritas Dasar

### Status implementasi
**Sudah:**
- Helper: `getPeriodeById_`, `assertPeriodeOpen_`, `validasiNominal_`, `validasiTanggalPeriode_`, `barisDibatalkan_`, `logActivityWajib_`, `cekSaldoCukup_` (`Data.gs`).
- Skema + `migrasiPengendalian()` (kolom soft-delete di `Input Penerimaan/Pengeluaran/Setoran Bank`).
- **Validasi nominal/tanggal/sumber kas** di: `submitTransaksi`, `submitKasPenerobos`, `importTransaksiCSV`, `updateTransaksi`.
- **Kunci periode tertutup** (`assertPeriodeOpen_` per baris) di: `updateTransaksi`, `deleteTransaksi`, `submitRincianIR`.
- **Validasi rincian IR balance di server** + ambil nominal/periode/anggota/tanggal dari sheet (T6/A1).
- **Soft delete** `deleteTransaksi` (Status=Dibatalkan + alasan wajib + log snapshot `logActivityWajib_`). Frontend `hapusTrx` meminta alasan.
- **Filter `Dibatalkan`** di pembaca: `calculateSaldo`, `getRekapitulasiData` (→ PDF), `getAdminConsole`, `getRiwayatTransaksiSaya`.
- **Guard saldo negatif** di `submitTransaksi` (pengeluaran & mutasi) + override ADMIN (`data.override` + `data.alasanOverride`, dicatat `OVERRIDE_SALDO`).

**Sisa audit Fase 1 — SUDAH DILENGKAPI:**
- ✅ `assertPeriodeOpen_` di `updatePendingStatus`, `bayar`/`batalBayarTagihanPatungan` (via `getPatunganPeriodeId_`). `submitRealisasiSetoran` operasi pada periode aktif + validasi realisasi ≥ 0.
- ✅ Validasi nominal/tanggal di `addPendingTransaction`, `addBankTransaction` (tanggal + non-negatif), `addPatungan` (nominal grade ≥ 0, min. satu > 0). `addManualBankIncome` aman (delegasi `submitTransaksi`).
- ✅ Filter `Dibatalkan` di `getRekapSetoran`, `getLaporanSetoran`, `getBukuIRBelumSerah`, `getKasPenerobos`, `getTagihanPenerobos`.

**Residual (1, frekuensi rendah — untuk Fase 2/kecil):**
- Bila sebuah **penerimaan yang punya rincian Buku IR dibatalkan**, baris rincian di `Detail Buku IR` (sheet tanpa soft-delete) masih terhitung di rekap Setoran Desa. Solusi: reader Buku IR menyaring rincian yang `transaksiId`-nya menunjuk penerimaan berstatus `Dibatalkan`. Dicatat untuk ditangani.

### Checklist uji manual

| # | Langkah | Hasil diharapkan |
|---|---------|------------------|
| **1** | Input Pemasukan nominal **−50.000** (via form / `apiCall` submitTransaksi tipe masuk) | **Ditolak**: "Nominal harus lebih besar dari 0." Tidak ada baris baru. |
| **2** | Input transaksi nominal **0** atau **1000.5** | Ditolak (harus bilangan bulat > 0). |
| **3** | Input transaksi bertanggal **besok** | Ditolak: "Tanggal transaksi tidak boleh di masa depan." |
| **4** | Input transaksi bertanggal **sebelum Tgl Mulai periode** | Ditolak: "Tanggal (…) sebelum awal periode (…)." |
| **5** | Input **Pengeluaran Tunai** melebihi saldo tunai | Ditolak: "Saldo Tunai tidak mencukupi (saldo: Rp X, diminta: Rp Y)." |
| **6** | Ulangi #5 sebagai **ADMIN** dengan `override:true` tanpa alasan | Ditolak: "Override saldo wajib disertai alasan." Dengan alasan → berhasil + tercatat `OVERRIDE_SALDO` di Activity Log. |
| **7** | Rincian Buku IR yang **tidak balance** (mis. IR+…+Index ≠ nominal transaksi) | Ditolak: "Rincian tidak seimbang: total Rp … ≠ nominal Rp … (selisih …)." |
| **8** | Rincian Buku IR **balance** (jumlah = nominal) | Berhasil disimpan. |
| **9** | **Hapus** transaksi via Riwayat/Rekap tanpa isi alasan (Batal di prompt) | Tidak ada aksi. Isi alasan → transaksi **tidak hilang**, tampil hilang dari daftar aktif; di sheet Status=`Dibatalkan`, `Alasan Batal` terisi, dan saldo berkurang sesuai. |
| **10** | Cek **Activity Log** untuk pembatalan #9 | Ada entri `BATAL_…` memuat ALASAN + snapshot nilai (jenis/tanggal/nominal/sumber). |
| **11** | **Tutup buku** periode, lalu coba **edit/hapus** transaksi periode itu (pakai ID lama) | Ditolak: "Transaksi ini milik periode yang sudah ditutup dan tidak dapat diubah." |
| **12** | Setelah membatalkan sebuah transaksi, buka **Dashboard**, **Rekapitulasi**, **Riwayat**, **PDF** | Semua konsisten: transaksi dibatalkan **tidak** ikut dihitung/ditampilkan. Saldo di keempat tempat sama. |
| **13** | Impor CSV berisi 1 baris tanggal masa depan & 1 baris nominal negatif | Kedua baris masuk daftar **gagal** dengan alasan; baris valid tetap terimpor. |

> **Regresi wajib dicek:** jalankan alur normal (input masuk/keluar/mutasi, rincian IR,
> serah terima) sebelum & sesudah membatalkan satu transaksi — pastikan angka tetap benar.

---

## FASE 2 — Integritas Periode & Saldo (T1, T5, T14)

### Prasyarat
- `migrasiPengendalian` **dijalankan ulang** (menambah kolom selisih/arsip di `Saldo Tutup Buku` + jenis `Selisih Kas` di master).
- (Opsional, untuk arsip) Script Property **`FOLDER_ARSIP_ID`** = ID folder Drive tempat menyimpan PDF; setujui izin **Drive** saat run pertama. Tanpa ini, tutup buku tetap jalan tetapi tanpa arsip.

### Implementasi
- **T5 tutup buku bermakna:** `cekSyaratTutupBuku_` (tolak bila ada Bank Pending, Buku IR belum dirinci, Kas Penerobos Aktif, Serah Terima Menunggu). `tutupBuku` menghitung **saldo sistem** (server), menyimpan sistem+aktual+**selisih** (Tunai/Bank/Total), **menolak** bila selisih ≠ 0 tanpa alasan, dan membuat **baris penyesuaian `Selisih Kas`** agar sistem = aktual.
- **T1 rollforward:** `bukaPeriode` mengunci saldo awal = saldo akhir periode CLOSED terakhir (`saldoAkhirTerakhir_`); beda → wajib `overrideRollforward`+alasan, selisih jadi penyesuaian eksplisit. Auto-open dari `tutupBuku` sudah rollforward = aktual.
- **T14 snapshot:** `arsipkanLaporan_` simpan PDF ke Drive + `Arsip File ID/URL/Hash`. `generatePDF(periodeId)` untuk periode CLOSED **selalu** kembalikan `arsipUrl` (bukan regen). Tombol "Laporan (arsip)" di menu Periode.

### Checklist uji
| # | Langkah | Hasil |
|---|---------|-------|
| 1 | Tutup buku saat masih ada **Serah Terima Menunggu** | Ditolak: "Masih ada Serah Terima berstatus Menunggu…". |
| 2 | Tutup buku saat masih ada **Kas Penerobos Aktif** | Ditolak. |
| 3 | Tutup buku saat masih ada **Buku IR belum dirinci** | Ditolak (jumlahnya disebut). |
| 4 | Tutup buku dengan saldo aktual **beda** dari sistem, batal saat diminta alasan | Tidak jadi tutup. |
| 5 | Idem #4, isi alasan | Berhasil. Cek `Saldo Tutup Buku`: kolom Sistem, Selisih, Alasan terisi; ada baris `Selisih Kas` di penerimaan/pengeluaran; saldo periode baru = aktual. |
| 6 | Tutup buku **tanpa selisih** (aktual = sistem) | Berhasil tanpa minta alasan; tak ada baris penyesuaian. |
| 7 | Setelah tutup, **Buka Periode Baru manual** dengan saldo awal ≠ saldo akhir sebelumnya, tanpa override | Ditolak: "Saldo awal harus = saldo akhir periode sebelumnya…". |
| 8 | (Jika `FOLDER_ARSIP_ID` diset) Menu **Periode** → periode CLOSED → **Laporan (arsip)** | Membuka berkas PDF arsip di Drive. |
| 9 | Buka `generatePDF(periodeIdCLOSED)` saat arsip belum diset | Balas: "Laporan periode tertutup tidak diarsipkan…". |

**Residual Fase 1 (rincian Buku IR dari penerimaan dibatalkan)** — belum ditangani; dianjurkan sebelum produksi.

---

## FASE 3 — Kas Penerobos & Cut-off (T7, T8, T9)

### Implementasi
- **T7:** `kasPenerobosAktif_` (total + aging per penerobos). `getDashboardData` menambah pos **Kas di Tangan Penerobos** (`kasPenerobos`, `penerobosDetail`, `agingHari`); `totalKas` = Tunai+Bank+Penerobos. Frontend menampilkan kartu terpisah + peringatan bila umur > `AGING_HARI` (Script Property, default 7). *(Seksi PDF penerobos: menyusul.)*
- **T8:** `konfirmasiSerahTerima` mencatat `Input Penerimaan` dengan **tanggal ASLI** (`Tanggal` di Kas Penerobos), bukan tanggal konfirmasi; menolak bila tanggal asli di luar rentang periode aktif.
- **T9:** `upsertSetoranPengeluaran_` tidak lagi menimpa — bila realisasi berubah, **batalkan baris lama** (soft delete, alasan "Koreksi realisasi setoran") + **buat baris baru**; tanggal lama tak diubah.

### Checklist uji
| # | Langkah | Hasil |
|---|---------|-------|
| 1 | Penerobos input pemasukan (jadi Kas Penerobos Aktif) → buka Dashboard (non-penerobos) | Muncul kartu **Kas di Tangan Penerobos** dengan nominal + umur; masuk hitungan **Total Kas**. |
| 2 | Biarkan > `AGING_HARI` (atau set `AGING_HARI`=0) | Baris penerobos diberi tanda **⚠ N hari** merah. |
| 3 | Serah terima → konfirmasi | Baris `Input Penerimaan` bertanggal **= tanggal asli terima** (bukan hari konfirmasi); `Created At` = hari konfirmasi. |
| 4 | Item kas penerobos bertanggal di luar periode aktif → konfirmasi | Ditolak dengan pesan cut-off. |
| 5 | Ubah **realisasi setoran** yang sudah tercatat | Baris pengeluaran lama jadi **Dibatalkan** (alasan "Koreksi realisasi setoran"), muncul **baris baru**; saldo benar; tanggal lama tak berubah. |

---

## FASE 4 — Bukti & Nomor Bukti (T10, T11)

### Prasyarat
- **Run → `migrasiPengendalian`** (kolom `No Bukti`, sheet `Lampiran`, backfill No Bukti data lama).
- **Lampiran bukti** butuh Script Property **`FOLDER_BUKTI_ID`** (folder Drive) + setujui izin **Drive**. Ambang wajib-bukti: **`AMBANG_BUKTI`** (default Rp 500.000).

### Implementasi
- **T11 nomor bukti:** kolom `No Bukti` di penerimaan (`BKM-<periode>-0001`) & pengeluaran (`BKK-…`), berseri per periode via `_isiNoBukti_` (dalam `withLock_`); baris dibatalkan tetap memegang nomor (tak didaur ulang). Backfill data lama urut tanggal+ID. Tampil di Riwayat.
- **T10 lampiran:** sheet `Lampiran`; `submitTransaksi` pengeluaran menerima `buktiList` (foto dikompres klien ~800px/JPEG 0.6), disimpan ke Drive. Pengeluaran > `AMBANG_BUKTI` **wajib** ≥1 bukti (ditolak `butuhBukti`). Endpoint `uploadBukti`/`getLampiran`.

### Checklist uji
| # | Langkah | Hasil |
|---|---------|-------|
| 1 | Input pemasukan & pengeluaran | Dapat **No Bukti** (BKM-…/BKK-…) berseri; tampil di Riwayat. |
| 2 | Batalkan satu transaksi lalu input baru | Nomor tak didaur ulang (yang dibatalkan tetap memegang nomornya). |
| 3 | Input **pengeluaran > AMBANG_BUKTI** tanpa foto | Ditolak: "…wajib melampirkan minimal 1 bukti". |
| 4 | Idem #3 dengan foto (set `FOLDER_BUKTI_ID`) | Berhasil; berkas muncul di folder Drive; tercatat di sheet `Lampiran`. |
| 5 | Pengeluaran kecil (≤ ambang) tanpa foto | Tetap boleh disimpan. |

**Sisa (didokumentasikan):** daftar "Pengeluaran tanpa bukti" & antrian offline "bukti tertunda" belum dibuat; seksi Kas Penerobos & bukti di PDF menyusul.

---

## FASE 5 — Maker-Checker, Log Hak Akses & Anti-XSS (T12, keamanan)

### Prasyarat
- **Run → `migrasiPengendalian`** (kolom `Status Approval`, `Disetujui By`, `Disetujui At` di `Input Pengeluaran`).
- Ambang approval: Script Property **`AMBANG_APPROVAL`** (default Rp 1.000.000).
- Kapabilitas **`trx.approve`** (default: ADMIN + Bendahara 1) — cek di Pengaturan → Hak Akses.

### Implementasi
- **T12 maker-checker:** pengeluaran > `AMBANG_APPROVAL` disimpan berstatus **Draft** (`_isiApproval_`), **belum masuk saldo** (`calculateSaldo` & `getRekapitulasiData` melewati `barisDraft_`). Disetujui lewat `setujuiPengeluaran(id)` (butuh `trx.approve`) → status **Disetujui**, baru dihitung. **Penyetuju ≠ pembuat** (dicegah server). Dashboard menampilkan tombol **Persetujuan Pengeluaran** (jumlah + total) bila ada Draft; halaman `pageApproval`/`renderApproval` untuk menyetujui. `getDashboardData.menungguApproval` dan `getDraftPengeluaran`.
- **Log hak akses:** `savePermMatrix` mencatat **selisih (diff)** matriks lama→baru sebagai `PRIVILEGED_HAK_AKSES` via `logActivityWajib_` (gagal log = gagal simpan).
- **Anti-XSS:** semua teks bebas dari pengguna (jenis, catatan, alasan, nama jamaah/pos) di-`escapeHtml`/`_esc` sebelum masuk `innerHTML`.

### Checklist uji
| # | Langkah | Hasil |
|---|---------|-------|
| 1 | Input **pengeluaran > AMBANG_APPROVAL** | Tersimpan **Draft**; pesan "menunggu persetujuan"; **saldo belum berubah**; tak muncul di rekap saldo. |
| 2 | Login sebagai **pembuat** Draft → buka Persetujuan | Tak bisa menyetujui transaksi sendiri (ditolak server: penyetuju ≠ pembuat). |
| 3 | Login sebagai **penyetuju lain** (punya `trx.approve`) → Setujui | Status jadi **Disetujui**; **saldo bertambah** sesuai nominal; hilang dari daftar Draft. |
| 4 | Pengeluaran **≤ AMBANG_APPROVAL** | Langsung **Disetujui** (tanpa antre). |
| 5 | Ubah matriks Hak Akses (Pengaturan) lalu cek `Log Activity` | Ada baris `PRIVILEGED_HAK_AKSES` berisi diff peran per-kapabilitas. |
| 6 | Isi catatan/jenis transaksi dengan teks `<img src=x onerror=alert(1)>` | Tampil sebagai teks apa adanya (tidak dieksekusi). |

---

## FASE 5.1 — Perketat Hak Akses Default & Pisah Input Masuk/Keluar (T12)

### Prasyarat
- **Run → `migrasiPengendalian`**: menyisipkan baris `trx.input.masuk` & `trx.input.keluar` ke sheet **Hak Akses**, **meniru nilai `trx.input` lama** (perilaku instalasi lama tetap sama; idempoten). Instalasi baru memakai default ketat.

### Implementasi
- **Pisah kapabilitas:** `trx.input` (lama) dipertahankan; enforcement kini per-arah:
  - `trx.input.masuk` — catat pemasukan. Default: ADMIN, Bendahara 1, Bendahara 2, Penulis, Penerobos.
  - `trx.input.keluar` — catat pengeluaran/mutasi. Default: ADMIN, Bendahara 1, Bendahara 2.
- **Enforcement server** (`requirePermInput_`/`userCanInput_`, fallback ke `trx.input` bila kolom baru belum ada): `submitTransaksi` (per `data.tipe`), `importTransaksiCSV` (per baris masuk/keluar), `uploadBukti` (keluar).
- **Frontend** (`App.canInput(tipe)`, fallback sama): tab **Masuk** muncul bila `trx.input.masuk`; tab **Keluar**/**Mutasi** bila `trx.input.keluar`; menu **Input CSV** bila salah satu.
- **Default lain diperketat:** `bukuIR.input` (tanpa Penerobos), `pembelaan.manage` (ADMIN+Bend.1+Bend.2), `terobosan.bayar` (tanpa Penulis), `grade.edit` (ADMIN+Bend.1). Hanya mempengaruhi **default**; matriks tersimpan tak berubah otomatis.

### Checklist uji
| # | Langkah | Hasil |
|---|---------|-------|
| 1 | (Instalasi lama) Run `migrasiPengendalian` → cek sheet Hak Akses | Muncul baris `trx.input.masuk` & `trx.input.keluar` = salinan nilai `trx.input`; akses semua user tak berubah. |
| 2 | Set peran (mis. Penulis) `trx.input.keluar`=OFF, `trx.input.masuk`=ON → login sbg Penulis → Input Transaksi | Hanya tab **Masuk** tampil; tab Keluar/Mutasi hilang. |
| 3 | Paksa panggil `submitTransaksi` tipe `keluar` sbg peran tanpa `trx.input.keluar` | Ditolak server: "Akses ditolak". |
| 4 | Import CSV berisi baris `keluar` sbg peran tanpa `trx.input.keluar` | Ditolak: "tidak berwenang input pengeluaran". |
| 5 | Peran dengan `trx.input.masuk`=ON → import CSV berisi hanya baris `masuk` | Berhasil. |

---

## PATCH K — Regresi Soft Delete & Saldo Tanpa Periode (K1, K2, K3)

### Prasyarat
- Deploy kode terbaru, lalu **Run → `bersihkanCache`** sekali dari editor (buang cache `buku_ir_data`/`dashboard_saldo`/`master_trx_data` lama).

### Implementasi
- **K1 (uang riil):** `trxPenerimaanDibatalkan_()` + `rincianYatim_()` — rincian Buku IR sah hanya bila transaksi penerimaan induknya masih Aktif. Difilter di `getRekapSetoran` & `getBukuIRBelumSerah`. Berlaku surut, tanpa kolom baru di `Detail Buku IR`.
- **K2 (deadlock tutup buku):** `getBukuIRData` kini melewati penerimaan `Dibatalkan` (`barisDibatalkan_`) → transaksi batal tak lagi dihitung "belum dirincikan".
- **K3 (saldo salah):** `calculateSaldo` **melempar error** bila dipanggil tanpa periode (tidak lagi mengembalikan angka ngawur). `getSaldoTutupBukuTerakhir_`/`saldoSaatIni_` dipakai saat tak ada periode OPEN. `getDashboardData` menampilkan **peringatan** + saldo arsip tutup buku; `cekSaldoCukup_` tak menelan error.
- **Rapikan:** `buatPenyesuaianSelisih_` menulis berbasis nama kolom + tanggal = tanggal tutup buku + pengeluaran penyesuaian berstatus **Disetujui** (agar dihitung `calculateSaldo`). Endpoint `batalkanKasPenerobos(id, alasan)` (cap `penerobos.input`/ADMIN) untuk melepas baris Kas Penerobos yang nyangkut.

### Checklist uji
| # | Langkah | Hasil yang diharapkan |
|---|---------|------------------------|
| 1 | Catat penerimaan Buku IR Rp 500.000 → rincikan → **batalkan** transaksinya | Kas berkurang 500.000 **dan** Rekap Setoran Desa **ikut berkurang** sesuai rincian. |
| 2 | Setelah #1, buka menu Buku IR | Transaksi dibatalkan **tidak** muncul di "Belum Dirincikan". |
| 3 | Setelah #1, coba Tutup Buku | **Tidak** diblokir pesan "masih ada N belum dirincikan". |
| 4 | Set semua periode CLOSED, buka dashboard | Saldo = saldo akhir tutup buku + kotak peringatan kuning. **Bukan** angka ngawur, **bukan** 0. |
| 5 | Tutup buku dengan selisih (mis. aktual bank kurang Rp 10.000), isi alasan | Muncul baris BKK "Selisih Kas" Rp 10.000 status **Disetujui**; saldo periode baru = aktual. |
| 6 | Cek `Master Pengeluaran` baris SELISIH_KAS | Ada, kode persis `SELISIH_KAS`. |
| 7 | Kas Penerobos Aktif yang nyangkut → tombol **Batalkan** + alasan | Status jadi Dibatalkan; tak lagi mengunci tutup buku; hilang dari Total Kas. |
| 8 | Muat ulang dashboard | Total Kas = Tunai + Bank + Kas Penerobos, ketiganya berlabel jelas. |

---

## FASE 6 L1–L3 — Penutupan Celah Sisa (gate kritis)

Inti K1–K3 sudah terpasang pada patch sebelumnya; Fase 6 menutup 4 celah sisa:
- `bersihkanCache` dijadikan **editor-only** (masuk `API_DENYLIST_`) — bukan lagi endpoint terbuka.
- `batalkanKasPenerobos` menambah `assertPeriodeOpen_` pada periode baris (T2).
- `getBukuIRData` tidak lagi memuat `rincianMap` milik penerimaan yang dibatalkan (K1 tuntas di semua pembaca).
- Dashboard menyembunyikan tombol **+ Input Transaksi** bila tidak ada periode aktif (K3).

### Audit pembaca `Detail Buku IR` (hasil verifikasi)
| Fungsi | Sumber | Status filter |
|--------|--------|---------------|
| `getBukuIRData` | Input Penerimaan + rincianMap (BUKU_IR) | `barisDibatalkan_` (K2) + `rincianYatim_` (K1) ✓ |
| `getRekapSetoran` | BUKU_IR per kolom | `rincianYatim_` ✓ |
| `getBukuIRBelumSerah` | BUKU_IR | `rincianYatim_` ✓ |
| `getLaporanSetoran` | Input Penerimaan (%Desa/%Daerah), **tidak** baca BUKU_IR | `barisDibatalkan_` ✓ |
| `submitRincianIR` | tulis BUKU_IR (bukan pembaca laporan) | — |

Tidak ada pembaca keempat yang terlewat.

### Checklist gate
| # | Langkah | Hasil |
|---|---------|-------|
| 1 (gerbang) | Penerimaan Buku IR Rp 500.000 → rincikan → **batalkan** | Kas −500.000 **dan** Rekap Setoran Desa ikut berkurang. |
| 2 | Buka menu Buku IR | Transaksi batal tak muncul di "Belum Dirincikan". |
| 3 | Tutup Buku setelah #1 | Tak diblokir "belum dirincikan". |
| 4 | Semua periode CLOSED → dashboard | Saldo = arsip tutup buku + peringatan kuning; tombol Input Transaksi hilang. |
| 5 | Kas Penerobos → `batalkanKasPenerobos` (periode OPEN) | Hilang dari pos penerobos; tutup buku tak terblokir. Periode CLOSED → ditolak. |

---

## L2b & L2c — Tampilan Kas Penerobos + Kelola Duplikat (Admin)

### Prasyarat
- **Run → `migrasiPengendalian`** (menambah kolom `Dibatalkan By/At`, `Alasan Batal`, `Koreksi Ref` di sheet `Kas Penerobos`).
- Set Script Property **`BUILD_DATE`** (mis. tanggal deploy) agar penanda versi muncul di footer — memudahkan deteksi HTML tertinggal saat deploy.
- Kapabilitas baru **`penerobos.kelola`** (default hanya ADMIN) — cek di Pengaturan → Hak Akses.

### Implementasi
- **L2b:** `getDashboardData` mengirim `sumberSaldo` ('periode'/'arsip'/'kosong'), `penerobosDalamTotal`, `penerobosTutup` (baris Aktif di periode CLOSED), `buildDate`. Kartu **Kas di Tangan Penerobos** tampil di dua mode: mode periode = komponen Total Kas; mode arsip = **di luar** Total (label tegas + garis merah). Peringatan merah bila ada baris di periode tutup. Aging merah bila umur > `AGING_HARI`.
- **L2c:** `deteksiDuplikatPenerobos_` (READ-ONLY) menilai tiap baris: SILANG (sumber masih aktif → 🟢 aman), INTERNAL (kembar → 🟢 aman), sumber hilang/batal → 🔴 bahaya. `getKelolaKasPenerobos` (cap `penerobos.kelola`) menampilkan semua baris + penilaian + status transaksi sumber. `batalkanKasPenerobos(id, alasan, konfirmasiNominal)` soft-delete (alasan ≥10 char; 🔴 wajib ketik ulang nominal; periode CLOSED ditolak → arahkan ke koreksi; `Diserahkan` ditolak). `koreksiKasPenerobosPeriodeTertutup` menandai `Dikoreksi` + log PRIVILEGED + peringatan arsip. `cekIntegritas` menyertakan kategori duplikat penerobos.

### Checklist uji
| # | Langkah | Hasil |
|---|---------|-------|
| 1 | Deploy ulang + set `BUILD_DATE`, buka dashboard | Kartu Kas Penerobos muncul; footer menampilkan `build <tgl>`; **tidak ada** `undefined`. |
| 2 | Dashboard tanpa periode OPEN | Total Kas = arsip (Tunai+Bank) saja. Kartu penerobos tampil **terpisah**, berlabel "tidak termasuk Total Kas". |
| 3 | Dashboard dengan periode OPEN | Total Kas = Tunai + Bank + Penerobos periode itu. |
| 4 | Buka Kelola Kas Penerobos | Semua baris tampil; kolom Sumber = ADA/DIBATALKAN/TIDAK DITEMUKAN; baris `KP_MIG_*` diberi badge 🔴/🟡. |
| 5 | Batalkan tanpa alasan (atau <10 char) | Ditolak. |
| 6 | Batalkan baris 🔴 tanpa ketik ulang nominal | Ditolak (`butuhKonfirmasiNominal`). |
| 7 | Batalkan baris di periode CLOSED | Ditolak; diarahkan ke menu Koreksi. |
| 8 | Batalkan satu baris (periode OPEN, skenario uji) | Status → Dibatalkan; hilang dari kartu dashboard & Total; **baris tetap ada di sheet**; Activity Log memuat alasan + snapshot + RISIKO. |
| 9 | `cekIntegritas()` | Kategori `DUPLIKAT_PENEROBOS` muncul dengan penilaian risiko per baris. |
| 10 | Coba serah terima baris Dibatalkan | Ditolak (tidak bisa dipilih). |

---

## L2d — Akar Macetnya Serah Terima (normalisasi email + diagnostik)

### Temuan (kenapa alur macet)
1. **Email case-sensitive (utama).** `getKasPenerobos`, `buatSerahTerima`, `getSerahTerimaList` membandingkan `email !== auth.user.email` **mentah**. Sheet menyimpan `Brianfpratama@gmail.com`, login `brianfpratama@gmail.com` → baris Brian **tak pernah terlihat** → tak bisa membuat serah terima. Kelas bug sama di `getRiwayatTransaksiSaya` (4 loop) dan key `migrasiTransaksiPenerobos`.
2. **Ketergantungan periode.** `konfirmasiSerahTerima` & `getSerahTerimaList` butuh periode OPEN. Setelah tutup buku tanpa periode baru, serah terima tak bisa dikonfirmasi → baris terjebak. (`cekSyaratTutupBuku_` sudah memblokir tutup buku saat kas penerobos Aktif — dikonfirmasi jalan.)
3. **Role penerima.** Konfirmasi butuh `serahterima.konfirmasi` + role cocok Sumber Tujuan (B1→Tunai, B2→Bank). Bila tak ada pengurus dengan role itu, tak ada yang bisa mengonfirmasi.

### Perbaikan
- Helper `emailSama_(a,b)` (`toLowerCase().trim()` kedua sisi) menggantikan semua perbandingan identitas email di `Data.gs` (riwayat ×4, getKasPenerobos, buatSerahTerima, getSerahTerimaList) + key migrasi dinormalisasi.
- `getKasPenerobos` mengirim `dapatSerahTerima` + `alasanSerahTerima` via `_diagnosaSerahTerima_` (cek periode OPEN + ada bendahara/admin aktif). UI menampilkan alasan bila serah terima tak bisa dibuat — bukan tombol hilang diam-diam.

### Checklist uji
| # | Langkah | Hasil |
|---|---------|-------|
| 1 | Login penerobos yang casing emailnya beda dari sheet | Kas penerobosnya **tetap terlihat**. |
| 2 | Buka Kas Penerobos saat tak ada periode OPEN | Banner kuning: "Periode sedang tidak terbuka…". Tombol serah terima tidak muncul, **dengan penjelasan**. |
| 3 | Tak ada Bendahara/Admin aktif | Banner: "Belum ada Bendahara/Admin aktif…". |
| 4 | Tutup buku saat masih ada kas penerobos Aktif | Diblokir (`cekSyaratTutupBuku_`). |

---

## L2c (revisi) — Deteksi "dicatat langsung", periode tutup, Aktivitas Istimewa

Melengkapi L2b/L2c sebelumnya dengan penyempurnaan sesuai insiden serah terima:
- **Kategori baru DICATAT LANGSUNG:** bila tak ada ID sumber cocok TAPI ada baris Input Penerimaan (tanggal ±30 hari, nominal + sumber kas + jenis sama), `deteksiDuplikatPenerobos_` menandai `PERLU_DIKONFIRMASI` (🟡) + menampilkan **kandidat pasangannya** (ID + tanggal + nominal). Tidak menyimpulkan — hanya menandai.
- **Risiko:** sumber ADA & aktif → 🟢 AMAN; ada pasangan → 🟡 PERLU_DIKONFIRMASI; tak ada jejak → 🔴 BAHAYA. 🟡 & 🔴 sama-sama wajib ketik ulang nominal.
- **Periode CLOSED boleh dibatalkan** (kas penerobos tak masuk `calculateSaldo`) — dicatat `PRIVILEGED_BATAL_KAS_PENEROBOS_TUTUP` + peringatan UI. (Menu koreksi-tutup lama dihapus.)
- **Alasan min 15 karakter** + alasan siap-pakai di UI.
- **Aktivitas Istimewa:** `getAktivitasIstimewa` menampilkan entri log berawalan `PRIVILEGED` untuk direviu ketua.

### Checklist uji
| # | Langkah | Hasil |
|---|---------|-------|
| 1 | Baris `KP_MIG_*` yang sumbernya hilang, tapi ada penerimaan mirip ±30 hari | Badge 🟡 + baris "Mungkin sudah dicatat langsung: <id> …". |
| 2 | Batalkan baris 🟡/🔴 tanpa ketik ulang nominal | Ditolak (`butuhKonfirmasiNominal`). |
| 3 | Batalkan dengan alasan < 15 karakter | Ditolak. |
| 4 | Batalkan baris di periode CLOSED | Boleh (setelah konfirmasi); tercatat `PRIVILEGED_BATAL_KAS_PENEROBOS_TUTUP`. |
| 5 | Buka Aktivitas Istimewa | Aksi PRIVILEGED muncul (waktu, user, detail). |
| 6 | `cekIntegritas()` | Ringkas memuat jumlah aman / perlu dikonfirmasi / bahaya. |

---

## Perbaikan — Tidak bisa buka periode baru (rollforward form)

**Akar masalah:** `bukaPeriode` mengunci saldo awal = saldo akhir tutup buku sebelumnya
dan menolak bila berbeda, tetapi form Buka Periode menampilkan field saldo **kosong (0)**
tanpa prefill rollforward & tanpa opsi "Ada selisih". Karena 0 ≠ saldo tutup buku, setiap
percakapan ditolak dan periode tak pernah bisa dibuat. Ditambah `showToast` (tak
terdefinisi) melempar error di success handler sehingga dashboard tak dimuat ulang.

**Perbaikan:**
- Endpoint baru `getInfoBukaPeriode` (cap `periode.manage`): mengirim `adaOpen`, `adaPrev`,
  dan saldo rollforward. Form otomatis **prefill** saldo awal = saldo tutup buku terakhir
  (read-only) + checkbox **"Ada selisih"** untuk override dengan alasan wajib.
- Bila masih ada periode OPEN, form menampilkan pesan jelas (tutup dulu).
- `showToast` → `toast` (4 lokasi) — memperbaiki bug success handler.

### Checklist uji
| # | Langkah | Hasil |
|---|---------|-------|
| 1 | Setelah tutup buku, buka menu Buka Periode | Saldo awal terisi otomatis = saldo akhir tutup buku (read-only). |
| 2 | Isi nama + tanggal, klik Buka Periode (tanpa ubah saldo) | Periode baru terbuka; dashboard tampil periode aktif. |
| 3 | Centang "Ada selisih", ubah saldo, isi alasan | Periode terbuka; selisih tercatat sebagai penyesuaian. |
| 4 | Centang "Ada selisih" tanpa alasan | Ditolak (alasan wajib). |
| 5 | Ada periode OPEN → buka menu | Pesan "tutup periode aktif dulu", form tidak muncul. |

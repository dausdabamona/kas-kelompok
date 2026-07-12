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

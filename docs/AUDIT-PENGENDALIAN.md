# Audit Pengendalian Intern — Kas Kelompok (FASE 0)

> Verifikasi temuan pengendalian intern atas pencatatan kas. **Tidak ada kode diubah
> di fase ini.** Acuan: `docs/KEAMANAN.md`. Semua nomor baris merujuk kode saat audit.

Legenda status: **VALID** (masih ada) · **DIPERBAIKI** (sudah tertutup) ·
**SEBAGIAN** (tertutup parsial) · **SALAH** (klaim keliru).

---

## 1. Tabel Verifikasi Temuan

| # | Status | Fungsi · Baris | Catatan verifikasi |
|---|--------|----------------|--------------------|
| **T1** | **VALID** | `bukaPeriode()` `Data.gs:355` | Saldo awal = `data.saldoAwalTunai/Bank` (baris 383-384), input bebas. Tidak membaca `Saldo Tutup Buku` periode CLOSED terakhir, tidak ada `cekRollforward_`. *Catatan:* perubahan terbaru pada `tutupBuku()` (auto-open periode) sudah melakukan rollforward untuk jalur tutup-buku, **tetapi** endpoint `bukaPeriode()` sendiri tetap menerima saldo awal sembarang (dipakai untuk periode pertama / buka manual). Rantai belum terkunci. |
| **T2** | **VALID** | `updateTransaksi()` `Data.gs:483`, `deleteTransaksi()` `Data.gs:552` | Baris dicari `hGet_(rows[i],h,'id',0) === data.id` **lintas seluruh sheet** (baris 502, 527, 569) tanpa cek `Periode ID` baris = periode aktif. `if (periode.status !== CONFIG.STATUS.OPEN)` (baris 491, 560) adalah **dead code** — `getPeriodeAktif()` (`Data.gs:54`) hanya mengembalikan objek saat status OPEN, selain itu `null`. Selama ADA periode OPEN, transaksi milik periode CLOSED (ID-nya diketahui) tetap bisa diedit/dihapus. Sama untuk `submitRincianIR()` mode update (`Data.gs:1600`, cari `transaksiId` lintas-periode). |
| **T3** | **VALID** | `deleteTransaksi()` `Data.gs:570`, `logActivity()` `Code.gs:170` | Hapus keras `sheet.deleteRow(i+1)` (baris 570). `logActivity(..., 'HAPUS_...', 'ID: ' + data.id)` (baris 572) — tanpa nominal/jenis/tanggal → transaksi terhapus tak bisa direkonstruksi. `logActivity` dibungkus `try{}catch(e){}` diam (`Code.gs`), boleh gagal tanpa jejak. |
| **T4** | **VALID** | `submitTransaksi()` `Data.gs:~440-470`, `submitKasPenerobos()`, `submitRincianIR()`, `submitRealisasiSetoran()` | Tidak ada `nominal > 0` di server: nilai ditulis via `Number(data.nominal)` (baris 450, 457, 462, 470) tanpa guard. Hanya `importTransaksiCSV` (`Data.gs:632`) & `submitRincianIR` per-baris (`:1277`) yang punya `<= 0`. Pemasukan **negatif** lolos → efektif pengeluaran yang melewati `trx.edit.*` dan tak tampak sebagai pengeluaran di laporan. `trx.input` default = semua role (lihat T12). |
| **T5** | **VALID** | `tutupBuku()` `Data.gs:302` | Menyimpan hanya `saldoTunai/saldoBank` **aktual dari klien** (baris 312-313, 323). Tidak memanggil `calculateSaldo()` untuk saldo sistem, tidak menghitung/menyimpan selisih, tidak memblokir bila selisih ≠ 0. Skema `SALDO_TUTUP_BUKU` (`Code.gs:276`) tak punya kolom Selisih (padahal `Bank Daily` punya). |
| **T6** | **VALID** | `submitRincianIR()` `Data.gs:1564` | Membaca `ir, ir10, cicilan, infakDaerah, index` (baris 1587-1591) lalu menulis, **tanpa** membandingkan jumlahnya dengan `nominal` baris `Input Penerimaan` milik `transaksiId`. Validasi "balance" hanya di klien (`cekTotalRincian`). Selain itu `periodeId`, `anggotaId`, `tanggal` diambil dari **klien** (baris 1614), bukan dibaca ulang dari transaksi. |
| **T7** | **VALID** | `calculateSaldo()` `Data.gs:172` | Menjumlah hanya `Input Penerimaan`, `Input Pengeluaran`, `Input Setoran Bank` (mutasi). Sheet `KAS_PENEROBOS` status `Aktif` **tidak pernah** ikut. Uang di tangan penerobos tak muncul di saldo, tanpa aging/batas setor. |
| **T8** | **VALID** | `konfirmasiSerahTerima()` `Data.gs:3835-3841`, `submitTransaksi()` | `var now = toDateStr_(new Date())` lalu `Input Penerimaan` dibuat dengan `now` (tgl konfirmasi), **bukan** kolom `Tanggal` asli di `Kas Penerobos` (item hanya membaca jenis/anggota/nominal/catatan, baris 3823-3826). `submitTransaksi` memakai `data.tanggal` bebas (`Data.gs:404`) tanpa validasi masuk rentang periode. |
| **T9** | **VALID** | `upsertSetoranPengeluaran_()` `Data.gs:2266` | Bila `existingId` ditemukan → `setValue` menimpa jenis, nominal, sumber kas, catatan, **dan tanggal di-reset ke `toDateStr_(new Date())`** (baris 2289). Pengeluaran posted dimutasi diam-diam; tak ada entri koreksi. |
| **T10** | **VALID** | seluruh kode | Tidak ada sheet `Lampiran`, tidak ada `<input type=file>`, tidak ada penyimpanan Drive untuk bukti. |
| **T11** | **VALID** | `generateID()` `Code.gs:166` | ID = `prefix + '_' + timestamp + '_' + random`. Tidak ada nomor bukti berseri (BKM/BKK) per periode; kelengkapan berseri tak bisa diuji. |
| **T12** | **VALID** | `getDefaultPermMatrix_()` `Code.gs:82`, `userCan_()` `Auth.gs` | `trx.input` default `allRoles` (`Code.gs:97`). `userCan_` `return true` untuk ADMIN tanpa cek (`Auth.gs:84`). Tak ada maker-checker. Tak ada pemisahan input masuk vs keluar. `Activity Log` di spreadsheet yang sama & bisa diedit manual. |
| **T13** | **VALID** | `submitTransaksi()`, mutasi | Tak ada guard "saldo tak cukup". Pengeluaran/mutasi tunai→bank bisa membuat saldo negatif (terbukti empiris: dashboard sempat Kas Bank −405.309). |
| **T14** | **VALID** | `generatePDF()` `Data.gs:2304` | Selalu `getRekapitulasiData()` dari data hidup (baris 2308). Tak ada snapshot/arsip; karena data lama masih bisa diubah (T2), laporan bisa berubah kemudian. |
| **T15** | **SEBAGIAN** | `Javascript.html` | Helper `escapeHtml()`/`_esc()` **sudah dibuat** (V5) dan dipakai **±72** kali, tapi ada **±140** sink `innerHTML`. Jadi tertutup ~separuh; sink data user yang belum di-escape masih ada. Bukan lagi "hanya 3 pemanggilan". |

---

## 2. Temuan Tambahan (di luar daftar)

| # | Tingkat | Lokasi | Temuan |
|---|---------|--------|--------|
| **A1** | TINGGI | `submitRincianIR()` `Data.gs:1614` | `periodeId`, `anggotaId`, `tanggal` rincian Buku IR diambil dari **input klien**, bukan diturunkan dari baris transaksi `Input Penerimaan` yang dirujuk. Klien bisa mengarahkan rincian ke periode/anggota lain. (Terkait T6.) |
| **A2** | SEDANG | `importTransaksiCSV()` `Data.gs` | Impor massal memakai `row.tanggal` apa adanya — tak ada validasi rentang periode, dan otomatis melewati maker-checker (Fase 5). Bisa menyuntik banyak transaksi bertanggal sembarang. |
| **A3** | SEDANG | `submitTransaksi()` | `sumberKas` tidak divalidasi ke himpunan {`Tunai`,`Bank`}. Nilai sembarang tersimpan → merusak agregasi saldo per sumber. |
| **A4** | SEDANG | `submitPemeriksaanSaldo()` vs `tutupBuku()` | Dua jalur pencatatan saldo akhir ke sheet yang sama (`SALDO_TUTUP_BUKU`) dengan status `Pemeriksaan`/`Tutup`. Berpotensi ambigu saat rollforward (Fase 2 harus memilih baris `Tutup` yang benar). |
| **A5** | RENDAH | `konfirmasiSerahTerima()` | Perlu dipastikan status `KAS_PENEROBOS` di-set `Diserahkan` saat konfirmasi (agar T7 aging akurat & tak dobel hitung). Perlu diverifikasi saat Fase 3. |

---

## 3. Rencana Perubahan Skema (semua via `ensureColumns_`, idempoten, hanya menambah)

| Sheet | Kolom Baru | Tipe | Alasan / Fase |
|-------|-----------|------|---------------|
| `Input Penerimaan` | `Status` (Aktif/Dibatalkan) · `Dibatalkan By` · `Dibatalkan At` · `Alasan Batal` | teks/tanggal | Soft delete + jejak (T3) — **F1** |
| `Input Penerimaan` | `No Bukti` (BKM-{periode}-000n) | teks | Nomor bukti berseri (T11) — **F4** |
| `Input Pengeluaran` | `Status` · `Dibatalkan By/At` · `Alasan Batal` | teks/tanggal | Soft delete (T3) — **F1** |
| `Input Pengeluaran` | `Status Approval` (Draft/Disetujui) · `Disetujui By` · `Disetujui At` | teks/tanggal | Maker-checker (T12) — **F5** |
| `Input Pengeluaran` | `No Bukti` (BKK-…) | teks | T11 — **F4** |
| `Input Setoran Bank` | `Status` · `Dibatalkan By/At` · `Alasan Batal` | teks/tanggal | Soft delete (T3) — **F1** |
| `Saldo Tutup Buku` | `Saldo Tunai Sistem` · `Saldo Bank Sistem` · `Selisih Tunai` · `Selisih Bank` · `Selisih Total` · `Alasan Selisih` | angka/teks | Tutup buku bermakna (T5) — **F2** |
| `Saldo Tutup Buku` | `Arsip File ID` · `Arsip URL` · `Arsip Hash` | teks | Snapshot laporan (T14) — **F2** |
| `Kas Penerobos` | *(punya `Tanggal` & `Status`)* — tambah `Umur Hari` (derived, opsional) | — | Aging (T7) — **F3** (dihitung saat baca, kolom tak wajib) |
| **Sheet baru** `Lampiran` | `ID` · `Transaksi ID` · `Tipe` · `Nama File` · `Drive File ID` · `URL` · `Diunggah By` · `Diunggah At` · `Status` | — | Bukti (T10) — **F4** |
| **Master Pemasukan/Pengeluaran** | pastikan ada jenis khusus `Selisih Kas` & `Penyesuaian Saldo Awal` | baris master | Entri koreksi eksplisit (T5/T1) — **F2** |

**Script Properties baru:** `FOLDER_ARSIP_ID`, `FOLDER_BUKTI_ID`, `AGING_HARI` (default 7),
`AMBANG_BUKTI` (default 500000), `AMBANG_APPROVAL` (default 1000000).

---

## 4. Analisis Dampak Regresi

**Penambahan kolom** (append di kanan via `ensureColumns_`) aman untuk semua pembaca
karena akses lewat `headerMap_`/`hGet_` (berbasis nama, bukan indeks tetap). Yang **berisiko**
adalah **soft delete** (F1): setiap pembaca ketiga sheet transaksi harus **memfilter
`Status = Dibatalkan`**. Fungsi yang wajib diaudit & difilter:

- Saldo & rekap: `calculateSaldo` `:172`, `getRekapitulasiData` `:1628`, `getDashboardData` `:117`, `getAdminConsole`.
- Buku IR & setoran: `getBukuIRData`, `getRekapSetoran`, `getLaporanSetoran`, `getBukuIRBelumSerah`.
- Riwayat & PDF: `getRiwayatTransaksiSaya`, `generatePDF`/`buildPDFHTML`, `getRekapitulasiData`.
- Penerobos & serah terima: `getKasPenerobos`, `getTagihanPenerobos`, `getSerahTerimaList/Detail`.
- Pembelaan/patungan bila membaca pengeluaran terkait.

**Risiko tertinggi F1:** ada pembaca yang terlewat filter → saldo/laporan salah hitung
(baris dibatalkan ikut terjumlah). Mitigasi: buat helper tunggal `bacaTransaksiAktif_(sheet)`
yang dipakai semua pembaca, sehingga filter terpusat di satu tempat.

**Maker-checker (F5):** `calculateSaldo` harus mengecualikan pengeluaran `Draft` — pembaca
yang sama seperti di atas perlu tahu status approval.

**Rollforward & tutup buku (F1→F2):** perubahan `tutupBuku`/`bukaPeriode` memengaruhi
auto-open periode yang baru dibuat; uji end-to-end tutup→buka.

---

## 5. Urutan Fase & Estimasi Risiko

| Fase | Isi | Risiko regresi | Alasan |
|------|-----|----------------|--------|
| **1** | Kunci periode (T2), validasi nominal/tanggal (T4/T8), validasi rincian IR server (T6), soft delete + log wajib (T3), guard saldo negatif (T13) | **Tinggi** | Menyentuh jalur baca saldo di banyak fungsi (filter Dibatalkan). Butuh helper terpusat + uji negatif menyeluruh. |
| **2** | Rollforward terkunci (T1), tutup buku + selisih + blokir (T5), snapshot arsip (T14) | **Sedang** | Terisolasi di `bukaPeriode`/`tutupBuku`/`generatePDF`; skema `Saldo Tutup Buku` bertambah. |
| **3** | Kas penerobos sebagai pos + aging (T7), cut-off tanggal (T8), koreksi setoran bukan overwrite (T9) | **Sedang** | Mengubah `calculateSaldo` (tambah pos) & `konfirmasiSerahTerima`; dashboard/PDF ikut. |
| **4** | Nomor bukti berseri (T11), lampiran bukti (T10) | **Sedang** | Penomoran data lama (migrasi) + unggah Drive; konkurensi nomor di `withLock_`. |
| **5** | Perketat hak akses (T12), batasi ADMIN + log PRIVILEGED, maker-checker pengeluaran, tutup XSS sisa (T15) | **Tinggi** | Mengubah matriks default (perlu persetujuan) & `calculateSaldo` (kecualikan Draft); menyentuh UI luas. |

**Catatan urutan:** F1 fondasi (soft delete + kunci periode) harus lebih dulu karena
F2/F3/F5 bergantung padanya (mis. koreksi setoran F3 memakai soft delete F1; maker-checker
F5 memakai status approval yang seskema dengan status batal F1).

---

## Kesimpulan Fase 0
14 dari 15 temuan **VALID** (masih ada); **T15 SEBAGIAN** (escaping sudah separuh).
Ditambah 5 temuan baru (A1–A5). Rencana skema & regresi siap.

> **CATATAN:** Bagian di atas adalah snapshot audit awal (Fase 0). Fase 1–5 sudah
> dikerjakan. Status akhir tiap temuan ada di bagian 6 & 7 di bawah.

---

## 6. Status Akhir Temuan (setelah Fase 1–5.1)

Legenda: **DITUTUP** (perbaikan terpasang & teruji) · **DITUTUP\*** (tertutup dengan sisa
kecil terdokumentasi).

| # | Status akhir | Fase | Ringkas perbaikan |
|---|--------------|------|-------------------|
| **T1** | **DITUTUP** | F2 | `bukaPeriode` mengunci rollforward = `saldoAkhirTerakhir_`; beda wajib `overrideRollforward`+alasan → baris penyesuaian. |
| **T2** | **DITUTUP** | F1 | `assertPeriodeOpen_(periodeId baris)` di `updateTransaksi`/`deleteTransaksi`/`submitRincianIR`; tak lagi lintas-periode. |
| **T3** | **DITUTUP** | F1 | Soft delete (`Status`, `Dibatalkan By/At`, `Alasan Batal`) + `logActivityWajib_` dengan snapshot; hapus keras dihilangkan. |
| **T4** | **DITUTUP** | F1 | `validasiNominal_` (>0, integer, ≤100e9) server-side di semua jalur input. |
| **T5** | **DITUTUP** | F2 | `tutupBuku` hitung saldo sistem (`calculateSaldo`), simpan sistem/aktual/selisih, blokir selisih≠0 tanpa alasan, baris penyesuaian `Selisih Kas`. |
| **T6** | **DITUTUP** | F1 | `submitRincianIR` baca nominal/periode/anggota/tanggal dari transaksi induk; validasi jumlah komponen === nominal. |
| **T7** | **DITUTUP** | F3 | `kasPenerobosAktif_` jadi pos ketiga `Total Kas` + aging (`AGING_HARI`). |
| **T8** | **DITUTUP** | F3 | Cut-off tanggal (rentang periode) di `submitTransaksi`/`konfirmasiSerahTerima`; pakai tanggal asli terima. |
| **T9** | **DITUTUP** | F3 | `upsertSetoranPengeluaran_` soft-cancel + baris baru (bukan overwrite); tanggal lama tak di-reset. |
| **T10** | **DITUTUP\*** | F4 | Sheet `Lampiran` + unggah Drive; pengeluaran > `AMBANG_BUKTI` wajib bukti. *Sisa:* daftar "tanpa bukti" & antrian offline. |
| **T11** | **DITUTUP** | F4 | No Bukti berseri BKM/BKK per periode via `_isiNoBukti_` (dalam lock) + backfill data lama. |
| **T12** | **DITUTUP** | F5/5.1 | Maker-checker pengeluaran (Draft→Disetujui, penyetuju≠pembuat); `trx.input` dipecah masuk/keluar; default diperketat; log `PRIVILEGED_HAK_AKSES`. |
| **T13** | **DITUTUP** | F1 | `cekSaldoCukup_` guard saldo negatif; override khusus ADMIN + alasan. |
| **T14** | **DITUTUP\*** | F2 | `arsipkanLaporan_` snapshot PDF ke Drive saat tutup buku; `generatePDF` periode CLOSED baca arsip. *Sisa:* seksi penerobos/bukti di PDF. |
| **T15** | **DITUTUP** | F5 | Sink `innerHTML` data user dibungkus `escapeHtml`/`_esc`; temuan V5 ditutup. |
| **A1** | **DITUTUP** | F1 | Rincian IR diturunkan dari transaksi induk, bukan input klien. |
| **A2** | **DITUTUP** | F1 | `importTransaksiCSV` validasi nominal/tanggal per baris + enforcement per-arah (5.1). |
| **A3** | **DITUTUP** | F1 | `sumberKas` divalidasi {Tunai,Bank}. |
| **A4** | **DITUTUP** | F2 | `getSaldoTutupBukuTerakhir_` memilih baris `Tutup` terakhir; rollforward jelas. |
| **A5** | **DITUTUP** | F3 | Status `KAS_PENEROBOS` di-set saat serah terima; aging akurat. |

---

## 7. Temuan Regresi Pasca-Perbaikan (K1–K3)

Muncul **akibat** perbaikan Fase 1 (soft delete) — pembatalan `Input Penerimaan` tidak
merambat ke data turunannya. Diperbaiki pada patch lanjutan setelah Fase 5.1.

| # | Tingkat | Akar masalah | Perbaikan |
|---|---------|--------------|-----------|
| **K1** | TINGGI (uang riil) | `Detail Buku IR` tak punya Status; rincian dari penerimaan yang dibatalkan tetap terhitung di Setoran Desa. | `trxPenerimaanDibatalkan_()` + `rincianYatim_()` sebagai sumber kebenaran (induk batal → rincian batal). Difilter di `getRekapSetoran` & `getBukuIRBelumSerah`. Berlaku surut, tanpa kolom baru. |
| **K2** | TINGGI (deadlock) | `getBukuIRData` tak memfilter penerimaan Dibatalkan → transaksi batal dianggap "belum dirincikan" → `cekSyaratTutupBuku_` mengunci tutup buku selamanya. | Tambah `barisDibatalkan_` di loop `getBukuIRData`; `bersihkanCache()` untuk buang cache lama. |
| **K3** | TINGGI (saldo salah) | `calculateSaldo(null)` mengembalikan angka ngawur; `getDashboardData`/`cekSaldoCukup_` menelan error diam-diam. | `calculateSaldo` melempar error bila tanpa periode; `getSaldoTutupBukuTerakhir_`/`saldoSaatIni_` untuk kondisi tanpa periode OPEN; dashboard menampilkan peringatan. Rapikan: `buatPenyesuaianSelisih_` berbasis header + tanggal tutup + status `Disetujui`; endpoint `batalkanKasPenerobos`. |

**Status:** K1–K3 **DITUTUP**. Semua fase pengendalian intern (Fase 0–5.1 + patch K) selesai.

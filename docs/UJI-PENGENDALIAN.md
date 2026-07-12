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

**Sisa audit Fase 1 (perlu dilengkapi sebelum tandai selesai):**
- `assertPeriodeOpen_` di `bayarTagihanPatungan`/`batalBayarTagihanPatungan` (via periode patungan), `updatePendingStatus`, `submitRealisasiSetoran`.
- Validasi nominal di `addBankTransaction`, `addPendingTransaction`, `addPatungan` (dan `submitPemeriksaanSaldo`/`tutupBuku` = angka saldo, dibahas Fase 2). `addManualBankIncome` sudah aman (delegasi ke `submitTransaksi`).
- Filter `Dibatalkan` di pembaca sekunder: `getRekapSetoran`, `getLaporanSetoran`, `getBukuIRBelumSerah`, `getKasPenerobos`, `getTagihanPenerobos`. (Tambahkan `if (barisDibatalkan_(row, h)) continue;` — pola sama.)
- Buku IR yang penerimaannya dibatalkan: rincian terkait perlu ikut diabaikan di perhitungan setoran (linkage `transaksiId`).

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

## FASE 2–5
*(diisi saat fase terkait dikerjakan.)*

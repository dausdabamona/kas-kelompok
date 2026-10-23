# Desain Fitur — Input Rekap Setoran

Status: **DRAFT — menunggu dikunci**
Tanggal: 23 Oktober 2026
Repo: `dausdabamona/kas-kelompok` (acuan commit `d033428`, 24-09-2026)
Menggantikan: rencana MCP Kas Kelompok (`mcp-kas-kelompok-desain.md`) — **ditunda**, input dilakukan lewat aplikasi.

---

## 1. Masalah

Setiap pengajian, bendahara menerima setoran dari puluhan jamaah sekaligus dan mencatatnya di buku tulis (contoh: rekap 23-10-2026, 27 baris, Rp21.402.000).

Di aplikasi saat ini, satu jamaah butuh **dua form**:

1. **Penerimaan**: nominal total, anggota, tanggal, sumber kas.
2. **Rincian Buku IR**: memecah total ke komponen (IR, Cicilan, Infak Daerah, Index, dst.).

Untuk 27 jamaah berarti 54 kali isi form. Ini lambat dan rawan salah ketik. Fitur impor CSV yang ada hanya untuk rekening koran (tanpa anggota dan tanpa rincian), jadi tidak bisa dipakai.

## 2. Tujuan

Bendahara **menempel rekap dari spreadsheet/buku** satu kali, aplikasi **memeriksa semuanya**, menampilkan **pratinjau**, lalu menyimpan seluruh baris dengan satu tombol. Hasil akhirnya **sama persis** dengan jika setiap jamaah diinput manual lewat dua form tadi.

## 3. Keputusan yang dikunci

| No | Keputusan | Pilihan |
|----|-----------|---------|
| K1 | Pemetaan kolom buku → komponen aplikasi | **IR → IR**, **Cilepus → Cicilan**, **Daerah → Infak Daerah**, **KLP → Index**, **Tunggakan → Sewa Kos** |
| K2 | Kolom Ket → Sumber Kas | **CASH → Tunai**, **TF → Bank**, **kosong → Tunai** |
| K3 | Ket bertanda **`?`** (perlu konfirmasi) | Baris **ditahan**: tampil di pratinjau, **tidak disimpan**, dilaporkan agar diinput setelah dikonfirmasi |
| K4 | Nama jamaah sama muncul lebih dari sekali | **Boleh**, tetap disimpan semua, tetapi ditandai kuning di pratinjau |
| K5 | Pencocokan nama ke master Anggota | Tidak peka huruf besar/kecil dan spasi. Nama yang tidak cocok **wajib dipilih manual** dari daftar anggota di pratinjau. |
| K6 | Jenis pemasukan | Dipilih sekali per rekap, hanya dari jenis berkategori **"Buku IR"** di Master Pemasukan |
| K7 | Satu rekap = satu tanggal | Tanggal diisi sekali, wajib berada dalam **periode aktif** |
| K8 | Cegah input dobel | Setiap rekap diberi **Kode Rekap** (mis. `REKAP-2026-10-23`). Kode yang sama ditolak pada simpan kedua. |
| K9 | Semua atau tidak sama sekali | Semua baris divalidasi dulu. Bila ada satu baris salah (selain baris ditahan K3), **tidak ada yang disimpan**. |
| K10 | Tidak menulis ulang aturan | Logika tulis penerimaan dan rincian IR **dipindah ke helper bersama**. Form lama dan fitur baru memakai helper yang sama, sehingga perilaku form lama tidak berubah. |
| K11 | Siapa yang boleh | Pengguna yang punya izin **input pemasukan** dan **`bukuIR.input`**. Peran PENEROBOS **tidak** mendapat menu ini. |
| K12 | Rapelan | Semua baris disimpan dengan rapel **1 bulan**. Rapel lebih dari 1 diedit lewat form rincian lama bila perlu. |
| K13 | Batas | Maksimal **200 baris** per rekap |

> **Catatan pemetaan K1:** Sewa Kos dicatat sebagai **milik kelompok**. Jika tunggakan Ali (Rp1.000.000 = 500.000 + 500.000) sebenarnya tunggakan **IR**, uang itu milik **desa** dan seharusnya masuk IR dengan rapel 3 bulan. Mohon dipastikan sebelum dikunci.

## 4. Format yang ditempel

Bendahara menyalin tabel dari Google Sheet/Excel (blok sel termasuk baris judul) lalu menempelnya ke kotak teks. Pemisah tab (hasil salin dari spreadsheet) maupun koma diterima.

Kolom yang dibaca berdasarkan **nama judul**, bukan urutan:

| Judul kolom | Wajib | Keterangan |
|-------------|-------|------------|
| `Nama` | Ya | Dicocokkan ke master Anggota |
| `IR` | – | Angka, boleh bertitik ribuan (`1.000.000`) |
| `Cilepus` | – | |
| `Daerah` | – | |
| `KLP` | – | |
| `Tunggakan` | – | |
| `Total` | Ya | Dipakai untuk memeriksa penjumlahan |
| `Ket` | – | CASH / TF / kosong / `?` |
| `Catatan` | – | Disalin ke catatan transaksi |

Kolom lain (No, Tanggal) diabaikan. Baris berjudul **GRAND TOTAL** dibaca sebagai pembanding, bukan sebagai jamaah.

## 5. Alur layar

1. **Tempel**: isi Tanggal, pilih Jenis Pemasukan (Buku IR), isi Kode Rekap (otomatis dari tanggal, bisa diubah), lalu tempel tabel. Tombol **Periksa**.
2. **Pratinjau**: tabel satu baris per jamaah, berisi nama anggota hasil cocok, komponen, total, dan sumber kas.
   - Hijau: siap disimpan.
   - Kuning: peringatan (nama dobel), tetap disimpan.
   - Abu-abu: ditahan (`?`).
   - Merah: salah. Total tidak cocok, nama belum dipilih, angka tidak terbaca, atau nominal 0.
   - Ringkasan di bawah: jumlah baris, total per komponen, total Tunai vs Bank, dan perbandingan dengan GRAND TOTAL dari buku.
   - Tombol **Simpan** aktif hanya bila tidak ada baris merah.
3. **Hasil**: jumlah tersimpan, nomor BKM pertama–terakhir, daftar baris ditahan untuk ditindaklanjuti.

## 6. Sisi server (Data.gs)

### 6.1 Fungsi baru

`importRekapSetoran(payload)`, dengan `payload.mode = 'periksa' | 'simpan'`

```
payload = {
  mode: 'periksa' | 'simpan',
  tanggal: '2026-10-23',
  jenisId: '<kode jenis Buku IR>',
  kodeRekap: 'REKAP-2026-10-23',
  grandTotalBuku: 21402000,        // opsional, untuk pembanding
  baris: [
    { no: 1, nama: 'Heru', anggotaId: '<opsional, hasil pilih manual>',
      ir: 400000, cicilan: 540000, infakDaerah: 30000, index: 30000, sewaKos: 0,
      total: 1000000, ket: 'CASH', catatan: '...' },
    ...
  ]
}
```

- **Periksa**: tidak menulis apa pun. Mengembalikan status setiap baris dan ringkasan.
- **Simpan**: menjalankan pemeriksaan yang sama **lagi di server** (hasil periksa dari klien tidak dipercaya), lalu menulis di dalam `withLock_`.

### 6.2 Yang ditulis per baris (baris ditahan dilewati)

1. **Input Penerimaan**: ID `TRX…`, periode aktif, jenisId, anggotaId, tanggal, total, sumber kas, catatan `Rekap <kodeRekap> #<no> — <catatan>`, pembuat, status Aktif, dan nomor **BKM** berurutan.
2. **Detail Buku IR**: komponen sesuai K1, RapelBulan 1, terkait ke TransaksiID di atas.
3. Satu baris **Activity Log**: `IMPOR_REKAP`, berisi kode rekap, jumlah baris, dan total.

Kode Rekap disimpan di kolom baru **`Kode Rekap`** pada Input Penerimaan (ditambahkan otomatis lewat `ensureColumns_`) untuk pemeriksaan K8.

### 6.3 Helper bersama (K10)

| Helper | Dipakai oleh |
|--------|--------------|
| `_tulisPenerimaan_(ss, periode, data, email)` | `submitTransaksi` (jalur masuk non-penerobos), `importRekapSetoran` |
| `_tulisRincianIR_(ss, trxId, periodeId, anggotaId, tanggal, nilai, rapel, email)` | `submitRincianIR` (jalur baris baru), `importRekapSetoran` |

Validasi yang sudah ada dipakai ulang tanpa diubah: `validasiNominal_`, `validasiTanggalPeriode_`, `komponenRincianAda_`, `_isiNoBukti_`, `requirePermInput_`, `requirePerm`.

### 6.4 Batas waktu GAS

Batas eksekusi Apps Script 6 menit. Penulisan dikumpulkan lalu ditulis dengan `setValues` sekaligus per sheet, bukan `appendRow` per baris. Untuk 200 baris, waktunya jauh di bawah batas.

## 7. Sisi tampilan

- Menu baru **"Input Rekap"** di bawah menu transaksi, hanya tampil bagi peran yang berhak (K11).
- Kode di `Javascript.html` dan `Stylesheet.html`, mengikuti gaya form yang ada (tema terang, ramah HP).
- Semua teks berbahasa Indonesia.

## 8. Yang tidak termasuk

- Membatalkan satu rekap sekaligus. Pembatalan tetap per transaksi lewat fitur batal yang sudah ada.
- Impor pengeluaran.
- Membaca foto buku tulis (OCR).
- MCP. Rencana MCP ditunda; bisa dibuka lagi nanti dengan memakai `importRekapSetoran` yang sama.

## 9. Uji terima (data 23-10-2026)

| Uji | Hasil yang diharapkan |
|-----|-----------------------|
| Tempel 27 baris + GRAND TOTAL | 26 baris hijau/kuning, 1 baris abu-abu (Roni, bila Ket diisi `?`) |
| Heru dua kali | Kuning, keduanya tersimpan |
| Total pratinjau vs buku | Rp21.402.000 sama dengan GRAND TOTAL (Rp18.402.000 tersimpan + Rp3.000.000 ditahan) |
| Komponen Ali | IR 1.000.000, Cicilan 200.000, Infak Daerah 5.000, Index 100.000, Sewa Kos 1.000.000 |
| Ubah satu Total jadi salah | Baris merah, tombol Simpan mati |
| Simpan dua kali dengan kode rekap sama | Simpan kedua ditolak |
| Form penerimaan + rincian lama | Tetap berjalan seperti sebelumnya |
| Saldo, Rekap Setoran Desa, Pantau IR, Kepemilikan Saldo | Ikut berubah sesuai data baru, tanpa error |

## 10. Tahapan dan gerbang

| Tahap | Pekerjaan | Gerbang |
|-------|-----------|---------|
| G1 | Dokumen ini dikunci | Balasan **"KUNCI"** |
| G2 | Helper bersama + `importRekapSetoran` (server) | Uji fungsi di editor Apps Script pada **salinan spreadsheet** → "LANJUT" |
| G3 | Layar Input Rekap | Coba alur lengkap di deployment uji → "LANJUT" |
| G4 | Input data 23-10-2026 di aplikasi asli | Angka cocok dengan buku → selesai |

## 11. Pertanyaan terbuka

1. **Tunggakan Ali**: tunggakan Sewa Kos, atau tunggakan IR? (lihat catatan K1)
2. **Roni**: setelah dikonfirmasi CASH atau TF, diinput lewat form biasa atau lewat rekap susulan?
3. **Nama di buku vs master Anggota**: apakah semua 27 nama sudah terdaftar? Nama ganda seperti "Ikhsan-Era" (dua orang atau satu keluarga?) dan "Tajudin S." perlu dicocokkan sekali; pilihan manual tidak diingat aplikasi (YAGNI).

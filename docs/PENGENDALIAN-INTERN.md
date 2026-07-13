# Pengendalian Intern Kas Kelompok — Panduan Pengurus

> Ditulis untuk **pengurus kelompok**, bukan programmer. Berisi ringkasan pengaman
> yang berlaku di aplikasi + prosedur rutin yang perlu dijalankan.

## Pengaman yang sudah berjalan di aplikasi

- **Tidak ada penghapusan permanen.** Transaksi yang salah **dibatalkan** (ditandai),
  bukan dihapus. Jejaknya tetap tersimpan untuk pemeriksaan.
- **Periode yang sudah ditutup terkunci.** Angka periode lama tidak bisa diubah diam-diam.
- **Pengeluaran besar butuh persetujuan** orang kedua (bukan si pembuat).
- **Nominal harus masuk akal** (lebih dari nol, bukan tanggal masa depan).
- **Saldo tidak boleh minus** tanpa persetujuan khusus + alasan.
- **Bukti wajib** untuk pengeluaran di atas ambang tertentu.
- **Uang di tangan penerobos** ditampilkan terpisah dan diberi peringatan bila sudah lama
  belum diserahkan.

## Prosedur bulanan pengurus

1. **Hitung uang fisik (cash count).** Cocokkan uang tunai & saldo bank nyata dengan angka
   di aplikasi. Selisih dicatat dengan alasan saat tutup buku.
2. **Rekonsiliasi bank.** Pastikan mutasi bank di aplikasi sama dengan rekening koran.
3. **Tinjau aktivitas istimewa.** Buka daftar aktivitas istimewa (override saldo, perubahan
   hak akses, koreksi periode tutup) — pastikan semuanya memang atas sepengetahuan pengurus.
4. **Serah terima kas penerobos.** Uang yang masih di tangan penerobos segera
   diserahterimakan; jangan dibiarkan menggantung sampai tutup buku.
5. **Tutup buku.** Setelah semua beres, lakukan tutup buku. Periode baru terbuka otomatis
   dengan saldo awal = saldo akhir periode lalu.

## ⚠️ Sebelum membatalkan baris Kas Penerobos mana pun

Cocokkan dulu dengan:

1. Catatan/kuitansi penerobos yang bersangkutan.
2. Uang fisik yang benar-benar ada di tangannya.
3. Hasil **Pemeriksaan Integritas** di aplikasi.

**Membatalkan baris yang transaksi sumbernya sudah tidak ada BUKAN koreksi hitung ganda —
itu menghapus uang dari sistem.** Kalau uangnya memang ada di tangan penerobos, yang benar
adalah **melakukan serah terima**, bukan membatalkan.

Aplikasi memberi tanda:
- 🟢 **Aman dibatalkan** — uang memang tercatat dua kali.
- 🔴 **Bahaya: uang hilang** — sumbernya tidak ada; membatalkan = uang lenyap dari catatan.
  Aplikasi akan meminta Anda mengetik ulang nominal sebagai konfirmasi.
- 🟡 **Perlu diteliti** — cek manual dulu.

## Membatalkan baris di periode yang sudah ditutup

Baris Kas Penerobos **boleh** dibatalkan walau periodenya sudah ditutup — karena kas
penerobos memang **tidak pernah** ikut dihitung dalam saldo periode, jadi membatalkannya
tidak mengubah angka periode yang sudah dilaporkan. Namun aksi ini dicatat sebagai
**Aktivitas Istimewa** dan diberi peringatan di layar, supaya bisa ditinjau ketua kelompok
di menu **Aktivitas Istimewa**.

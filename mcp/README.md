# MCP Kas Kelompok — baca saja

Menghubungkan Claude ke aplikasi Kas Kelompok. Claude menerima angka yang
**sudah dihitung aplikasi**, bukan isi spreadsheet mentah.

Bedanya penting. Membaca sheet langsung akan salah, karena:
- baris berstatus `Dibatalkan` dan pengeluaran `Draft` ikut terjumlah;
- hak kelompok / desa / daerah tidak terpisah;
- Buku IR yang belum dirinci terhitung sebagai hak kelompok, padahal bukan.

## Tidak bisa menulis

Tidak ada satu pun alat di sini yang mengubah data. Pembatasnya tiga lapis:

1. **Daftar-putih di server** (`MCP_ALLOW_` pada `Auth.gs`). Endpoint baru tidak
   otomatis terbuka — harus didaftarkan satu per satu.
2. **Token terpisah** (`MCP_TOKEN`), bukan session token pengguna. Dicabut kapan
   saja tanpa mengganggu siapa pun yang sedang login.
3. **Izin peran tetap berlaku.** Permintaan dijalankan atas nama satu user di
   Master User (`MCP_USER_EMAIL`), sehingga `requirePerm` di tiap endpoint tetap
   menyaring.

Selama `MCP_TOKEN` belum diisi, endpoint-nya mati total.

## Pemasangan

### 1. Aktifkan di sisi Apps Script

Buka editor Apps Script, pilih fungsi **`aktifkanApiMCP`**, klik Run. Buka
**Execution log** — token 48 karakter tercetak di sana. Salin.

Lalu **Project Settings → Script Properties**, tambahkan:

| Properti | Isi |
|---|---|
| `MCP_USER_EMAIL` | email salah satu user di Master User |

Saran: buat user khusus, misalnya `mcp@kelompok.local`, dengan peran yang
izinnya secukupnya. Jangan pakai akun admin pribadi — kalau token bocor,
yang terbuka hanya sebatas peran itu.

Mengganti token: jalankan `aktifkanApiMCP` lagi (yang lama langsung mati).
Mematikan: jalankan `matikanApiMCP`.

### 2. Pasang dependensi

```bash
cd mcp
npm install
```

### 3. Daftarkan ke Claude Code

```bash
claude mcp add kas-kelompok \
  --env KAS_EXEC_URL=https://script.google.com/macros/s/XXXX/exec \
  --env KAS_TOKEN=token-dari-langkah-1 \
  -- node /path/lengkap/ke/kas-kelompok/mcp/server.js
```

Repo ini juga menyertakan `.mcp.json` yang membaca `KAS_EXEC_URL` dan
`KAS_TOKEN` dari environment — berguna bila beberapa orang memakai repo yang
sama dengan token masing-masing.

### 4. Daftarkan ke Claude Desktop

Tambahkan ke `claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "kas-kelompok": {
      "command": "node",
      "args": ["D:\\Drive\\5. NGODING\\kas-kelompok\\mcp\\server.js"],
      "env": {
        "KAS_EXEC_URL": "https://script.google.com/macros/s/XXXX/exec",
        "KAS_TOKEN": "token-dari-langkah-1"
      }
    }
  }
}
```

**Jangan commit token ke git.** Simpan hanya di konfigurasi lokal.

## Alat yang tersedia

| Alat | Isinya |
|---|---|
| `kas_ringkasan` | Saldo tunai & bank, total kas, kas di tangan penerobos |
| `kas_kepemilikan_saldo` | Hak kelompok, titipan desa/daerah/jamaah, belum dirinci |
| `kas_pengeluaran_rutin` | Perkiraan vs realisasi, sisa pokok cicilan |
| `kas_pantau_ir` | Setoran IR tiap jamaah beberapa bulan terakhir |
| `kas_rekap_setoran` | Kewajiban setor per pos: target, realisasi, sisa |
| `kas_rincian_ir` | Rincian Buku IR per jamaah menurut komponennya |
| `kas_buku_ir` | Yang sudah & belum dirincikan |
| `kas_transaksi` | Seluruh transaksi periode berjalan |
| `kas_rekapitulasi` | Pemasukan/pengeluaran per jenis, mutasi, saldo |
| `kas_periode` | Daftar periode & statusnya |
| `kas_jamaah` | Daftar jamaah & acuan setoran bulanannya |

## Kalau gagal

**"API tidak aktif"** — `MCP_TOKEN` belum diisi atau kurang dari 32 karakter.
Jalankan `aktifkanApiMCP`.

**"Token tidak sah"** — `KAS_TOKEN` tidak sama dengan Script Property-nya.

**"Balasan bukan JSON"** — `KAS_EXEC_URL` salah, atau deployment tidak dapat
diakses. Pastikan URL berakhiran `/exec`, bukan `/dev`.

**"Endpoint tidak diizinkan"** — alat memanggil fungsi yang belum terdaftar di
`MCP_ALLOW_`. Ini memang pengaman, bukan kerusakan.

**"Akses ditolak"** — peran `MCP_USER_EMAIL` tidak punya izin untuk endpoint itu.

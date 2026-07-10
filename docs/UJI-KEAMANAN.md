# Uji Keamanan — Kas Kelompok (Fase 5)

Acuan: `docs/KEAMANAN.md`. Dokumen ini adalah **rencana uji + bukti kode** untuk Fase 1–4.
Kolom **Status**:
- ✅ **terverifikasi-kode** — kontrol ada & bisa dibaca di kode (referensi disertakan).
- 🧪 **perlu uji staging** — WAJIB dijalankan pemilik di deployment staging (butuh runtime GAS/MailApp; tidak bisa diverifikasi dari repo).

> Jalankan `clasp push` ke **deployment staging** (bukan produksi), lalu uji lewat URL web app staging.

---

## A. Prasyarat staging (lakukan dulu)

1. `clasp push` kode terbaru ke project staging.
2. Dari editor Apps Script: **Run → `migrasiKeamanan`** (buat sheet `Perangkat`/`Sesi` + kolom PIN + pepper). Setujui izin.
3. Pastikan ada ≥1 user **ADMIN** aktif di `Master User` dengan email yang bisa menerima email (untuk uji OTP).
4. Deploy: **Execute as: Me** + **Who has access: Anyone**. Salin URL `/exec`.
5. `appsscript.json` memuat `"timeZone": "Asia/Jayapura"` (lihat §D).

---

## B. Checklist verifikasi

### 1. Endpoint mutasi TANPA session token valid → ditolak
- **Status:** ✅ terverifikasi-kode + 🧪 uji staging
- **Bukti:** dispatcher `apiCall` set `__REQ_USER_ = resolveUser_(token)`; token invalid → `null` (`Auth.gs` `apiCall`, `resolveUser_`→`verifySession_`). `requirePerm`/`checkAuth` → `getCurrentUser()` mengembalikan `null` → `{success:false, message:'Belum login'}`.
- **Uji:** dari Console browser di halaman staging:
  `google.script.run.withSuccessHandler(console.log).apiCall('token-palsu','submitTransaksi',[{tipe:'masuk',jenisId:'x',nominal:1000}])`
  → **Harus** balas `{success:false, message:'Belum login'}` dan **tidak** menulis baris.

### 2. Role rendah tidak menembus capability terlarang
- **Status:** ✅ + 🧪
- **Bukti:** tiap endpoint mutasi memanggil `requirePerm('<cap>')`; matriks default `getDefaultPermMatrix_` (`Code.gs`). Contoh: `deleteUser`→`user.manage` (hanya ADMIN), `tutupBuku`→`periode.manage` (ADMIN).
- **Uji:** login sebagai role **PENULIS**, lalu panggil endpoint terlarang via `apiCall(<tokenPenulis>,'deleteUser',['x@y.com'])` → **Harus** `{success:false, message:'Akses ditolak'}`.

### 3. `loginWithEmail` lama tidak memberi identitas tanpa OTP/PIN
- **Status:** ✅ + 🧪
- **Bukti:** `loginWithEmail` (Fase 1) hanya sukses bila email == `Session.getActiveUser().getEmail()`; pada deploy anonim Session kosong → ditolak. Frontend tak lagi memanggilnya (Fase 2c). `getCurrentUser` mengutamakan `__REQ_USER_`.
- **Uji:** buka URL staging tanpa login Google (incognito) → panggil `apiCall('', 'loginWithEmail', ['admin@contoh.com'])` → **Harus** gagal (tidak menerbitkan token/sesi).
- **Catatan:** tidy-up tersisa (belum dikerjakan): hapus total fallback Session di `getCurrentUser` + denylist `loginWithEmail` di dispatcher (lihat §E).

### 4. OTP kadaluarsa / salah → ditolak; >5 gagal → lockout
- **Status:** ✅ + 🧪
- **Bukti:** `verifyOtp_` cek expiry (`data.e`) & hash; salah → `recordAuthFail_`. `isAuthLocked_` blokir setelah `SEC_MAX_FAIL_=5` selama `SEC_LOCK_TTL_S_=900`s. Throttle kirim `_otpThrottleOk_` (cooldown 60s, maks 5/jam).
- **Uji:** minta OTP → tunggu >5 menit → verifikasi → **Harus** "OTP kadaluarsa". Lalu masukkan OTP salah **6×** → percobaan ke-6 **Harus** "Terlalu banyak percobaan. Coba lagi dalam 15 menit."

### 5. Perangkat dicabut → sesi terkait invalid
- **Status:** ✅ + 🧪
- **Bukti:** `revokeDevice` set `Status=Dicabut` + `invalidateSessionsByDevice_` (semua sesi device → `Dicabut`); `verifySession_` menolak status ≠ `Aktif`.
- **Uji:** login PIN di perangkat A (dapat sesi) → di Admin **Kelola Perangkat** → **Cabut** perangkat A → aksi berikutnya di perangkat A **Harus** memicu "Sesi berakhir" (login ulang).

### 6. XSS tertutup: nama jamaah `<img onerror=...>` tidak tereksekusi
- **Status:** ✅ + 🧪
- **Bukti:** `App.escapeHtml` diterapkan pada data dinamis (Fase 1, 53+ titik); server `esc_` tersedia.
- **Uji:** tambah jamaah bernama `<img src=x onerror=alert(1)>` → buka daftar Jamaah → **Harus** tampil sebagai teks, **tidak** ada alert.

### 7. Dua submit bersamaan tidak merusak saldo (LockService)
- **Status:** ✅ + 🧪
- **Bukti:** `withLock_` (`Code.gs`) membungkus 15 fungsi tulis (Fase 3).
- **Uji:** dari Console jalankan 2 `apiCall(...,'submitTransaksi',[...])` hampir bersamaan → keduanya tercatat benar; saldo akhir = penjumlahan tepat (tidak ada baris hilang / saldo dobel). Salah satu boleh balas "Sistem sedang sibuk" bila >15s (lalu ulang).

### 8. `appsscript.json` + rahasia di Script Properties
- **Status:** 🧪 (manual)
- **Uji:** `appsscript.json` `timeZone` = `Asia/Jayapura`; **Project Settings → Script Properties** memuat `SECURITY_PEPPER` (dibuat otomatis oleh `getPepper_`/`migrasiKeamanan`). Spreadsheet ID: lihat §E (masih hardcoded — tidy-up).

### 9. Anti-lockout admin (min. 1 ADMIN aktif)
- **Status:** ✅
- **Bukti:** `countActiveAdmins_` mencegah `deleteUser`/`updateUser`/`setUserStatus` menyisakan 0 ADMIN aktif; `getPermMatrix_` memaksa ADMIN selalu penuh (anti-lockout).

---

## C. Alur happy-path end-to-end (uji staging)
1. Buka URL staging → layar **Email**.
2. Masukkan email admin → perangkat baru → **OTP dikirim** → cek email → masukkan OTP → **set PIN**.
3. Masuk dashboard. **Logout**.
4. Login lagi → email sama → langsung diminta **PIN** (perangkat terpercaya) → masuk.
5. Input 1 transaksi → cek tercatat di sheet + saldo dashboard berubah.

---

## D. Tindakan manual pemilik (WAJIB sebelum produksi)
- [ ] `clasp push` + jalankan **`migrasiKeamanan`** sekali.
- [ ] Setujui izin **MailApp** (kirim email) & **Spreadsheet** saat run pertama.
- [ ] Tambah/lengkapi `appsscript.json` dengan `"timeZone": "Asia/Jayapura"` sebelum push.
- [ ] Uji **kirim OTP** benar-benar masuk email (cek kuota MailApp: ±100/hari akun consumer).
- [ ] Deploy **Execute as: Me + Anyone**, uji seluruh checklist §B di staging.
- [ ] Pastikan minimal 1 ADMIN sudah **bootstrap perangkat + PIN** sebelum menutup akses lama.
- [ ] (Opsional) Buat **trigger harian** memanggil `cleanupSessions_`.

---

## E. Tidy-up
1. ✅ **SELESAI** — `getCurrentUser` kini hanya `__REQ_USER_` (fallback Session dihapus); `loginWithEmail`/`getGoogleEmail`/`getCurrentUser` masuk `API_DENYLIST_` → **V2 tertutup 100%**.
2. ✅ **SELESAI** — `getSpreadsheetId()` baca `SPREADSHEET_ID` dari Script Properties (fallback + auto-populate) → **V6 ditutup**.
3. ✅ **SELESAI** — `setupSheets`/`repairSheetHeaders`/`migratePosSetoran`/`migrasiKeamanan` dijaga `_setupAccessOk_()` (bootstrap-safe: izinkan bila belum ada ADMIN; selain itu hanya ADMIN terverifikasi via editor) → **V4 ditutup**. Pemanggilan anonim via `google.script.run` ditolak.
4. ⏳ **Terbuka (opsional)** — fallback non-`Proxy` untuk `App.srv()` bila ada perangkat WebView sangat lawas (Proxy didukung Android 5+; risiko rendah).

> Dampak tidy-up #1: setelah ini, aplikasi **wajib** punya sesi token valid untuk semua endpoint. Pastikan minimal 1 ADMIN sudah bootstrap perangkat+PIN (via OTP) sebelum push ke produksi.

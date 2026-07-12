# Audit & Rencana Keamanan — Kas Kelompok

> **Status:** FASE 0 (audit & desain). Belum ada kode aplikasi yang diubah.
> Dokumen ini adalah dasar untuk Fase 1–4. Perubahan kode menunggu instruksi "LANJUT".

Konteks deploy (menentukan seluruh model ancaman):

- Web app di-deploy **"Execute as: Me (owner)"** + **"Access: Anyone (even anonymous)"**.
- Akibatnya:
  1. `Session.getActiveUser().getEmail()` sering **kosong** → tidak boleh jadi sumber identitas.
  2. URL web app **publik** → setiap fungsi `google.script.run` = endpoint internet terbuka.
  3. `PropertiesService.getUserProperties()` mengembalikan properti **pemilik skrip** untuk
     SEMUA pengunjung anonim → **tidak bisa** membedakan pengguna → haram dipakai untuk sesi.

---

## a. Temuan Kerentanan Saat Ini

| # | Tingkat | Lokasi | Temuan | Dampak |
|---|---------|--------|--------|--------|
| V1 | **KRITIS** | `Auth.gs:123 loginWithEmail()` | Login **tanpa kredensial** — cukup kirim email terdaftar, langsung dapat sesi penuh sesuai role. Email mudah ditebak/diketahui. | Siapa saja bisa masuk sebagai ADMIN/Bendahara dan memutasi kas. |
| V2 | **KRITIS** | `Auth.gs:1 getCurrentUser()` | Mengambil identitas dari `UserProperties('userEmail')` lebih dulu, baru `Session`. Pada deploy "execute as owner + anonymous", `UserProperties` = milik **pemilik** untuk semua pengunjung → sesi bocor/tercampur antar pengguna anonim. | Satu login menular ke semua pengunjung; mustahil membedakan pengguna. |
| V3 | **TINGGI** | `submitTransaksi` (`Data.gs:394`), `submitRincianIR` (`1278`), `submitKesanggupanPembelaan` (`3112`), `updateStatusPembelaan` (`3147`) | Endpoint **mutasi** hanya dijaga `checkAuth()` (asal login), **tanpa** `requirePerm(cap)`. | User role terendah (mis. PENULIS/PENEROBOS) bisa menulis data di luar wewenangnya. |
| V4 | **TINGGI** | `setupSheets` (`Code.gs:316`), `repairSheetHeaders` (`499`), `migratePosSetoran` (`439`) | Fungsi setup/migrasi skema **global tanpa guard apa pun** → bisa dipanggil via `google.script.run` oleh siapa saja. | Perusakan/penulisan ulang header sheet oleh anonim. |
| V5 | **DITUTUP (Fase 5)** | `Javascript.html` | ~~Data user dirender via `innerHTML` tanpa escaping.~~ Semua teks bebas dari pengguna (nama jamaah/pos, catatan, keterangan, jenis, alasan) kini dibungkus `escapeHtml`/`_esc` sebelum masuk `innerHTML`. | Tertutup: XSS tersimpan tidak lagi tereksekusi (dirender sebagai teks). |
| V6 | **SEDANG** | `Code.gs:184 getSpreadsheetId()` | Spreadsheet ID **hardcoded** dan fungsi bisa dipanggil dari klien. | Kebocoran ID + rahasia tidak terpusat. Harus pindah ke Script Properties. |
| V7 | **SEDANG** | `Auth.gs:123` | Tidak ada **rate-limit / lockout** pada login. | Brute-force PIN/OTP tanpa hambatan. |
| V8 | RENDAH | `getGoogleEmail`, `logActivity`, `generateID`, `fmtRp` dll. | Helper global ikut ter-ekspos sebagai endpoint. | Permukaan serang lebih luas; sebaiknya diakhiri `_` atau dijaga. |

---

## b. Inventaris Fungsi Server (dipanggil via `google.script.run`)

Konvensi kode: helper diakhiri `_` (mis. `getUserList_`) = **tidak** ditujukan untuk klien.
Fungsi tanpa `_` di bawah ini = permukaan endpoint publik. **M** = Mutasi (tulis), **R** = Baca.

### Auth & User (`Auth.gs`)
| Fungsi | M/R | Guard sekarang | Capability yang seharusnya |
|--------|-----|----------------|-----------------------------|
| `loginWithEmail` | M(sesi) | — (tanpa kredensial) | **DIGANTI** alur OTP/PIN (lihat §c) |
| `logoutUser` | M(sesi) | `getCurrentUser` | token valid |
| `getCurrentUser` | R | UserProperties/Session | **DIGANTI** `resolveUser_(token)` |
| `getPermMatrix` | R | `checkAuth([ADMIN])` | `hakakses.view` (ADMIN) |
| `savePermMatrix` | M | `checkAuth([ADMIN])` | `hakakses.manage` (ADMIN) |
| `getUserListAdmin` | R | `requirePerm('user.manage')` | `user.manage` ✓ |
| `addUser` `updateUser` `setUserStatus` `deleteUser` | M | `requirePerm('user.manage')` | `user.manage` ✓ |

### Setup / Sistem (`Code.gs`)
| Fungsi | M/R | Guard | Rencana |
|--------|-----|-------|---------|
| `doGet` | entry | — | render Index; tidak menerima token |
| `setupSheets` `repairSheetHeaders` `migratePosSetoran` | M(skema) | **tanpa guard (V4)** | jadikan hanya-editor: guard `requirePerm('sistem.setup')` + jangan diekspos ke klien |
| `getSpreadsheetId` | R(rahasia) | — (V6) | pindah ke Script Properties, jadikan helper `_` |
| `include` `fmtRp` `fmtTanggal` `generateID` `logActivity` | helper | — | akhiri `_` / tidak dipanggil klien |

### Transaksi & Keuangan (`Data.gs`) — mutasi
| Fungsi | Guard sekarang | Capability yang seharusnya |
|--------|----------------|-----------------------------|
| `submitTransaksi` | **checkAuth (V3)** | `trx.create` (baru) — cabang penerobos → `penerobos.input` |
| `updateTransaksi` `deleteTransaksi` | `trx.edit.tunai`/`trx.edit.bank` | ✓ pertahankan |
| `submitPemeriksaanSaldo` | `saldo.input` | ✓ |
| `tutupBuku` `bukaPeriode` | `periode.manage` | ✓ |
| `submitRincianIR` | **checkAuth (V3)** | `bukuir.rincian` (baru) |
| `addAnggota` `updateAnggota` `deleteAnggota` | `jamaah.add/edit/delete` | ✓ |
| `addPatungan` `bayarTagihanPatungan` `batalBayarTagihanPatungan` `updateAnggotaGrade` | `terobosan.*` / `grade.edit` | ✓ |
| `addPos*` `updatePos*` `deletePos*` `add/update/deleteMusyawaroh` `addPosSetoran…` | `master.manage` | ✓ |
| `submitRealisasiSetoran` | `setoran.realisasi` | ✓ |
| `addBankTransaction` | `bank.input` | ✓ |
| `updateBankDaily` `addManualBankIncome` `addPendingTransaction` `submitRekonsiliasiBank` `updatePendingStatus` | `bank.manage` | ✓ |
| `submitKesanggupanPembelaan` | **checkAuth (V3)** | `pembelaan.input` (baru) |
| `updateStatusPembelaan` | **checkAuth (V3)** | `pembelaan.manage` (baru) |
| `submitKasPenerobos` `buatSerahTerima` | `penerobos.input` | ✓ |
| `konfirmasiSerahTerima` | `serahterima.konfirmasi` | ✓ |
| `migrasiTransaksiPenerobos` | `master.manage` | ✓ |

### Data baca-saja (`Data.gs`) — cukup `checkAuth()` (token valid), tanpa cap tulis
`getPeriodeAktif`, `getAllPeriode`, `getDashboardData`, `getPemeriksaanSaldo`,
`getRiwayatTransaksiSaya`, `getTransaksiMasterData`, `getMasterPemasukan`,
`getMasterPengeluaran`, `getAnggota`, `getPatunganList`, `getTagihanPatungan`,
`getAnggotaWithGrade`, `getBukuIRData`, `getRekapitulasiData`, `getPosSetoranAll`,
`getMusyawaroh`, `getRekapSetoran`, `generatePDF`, `getBankDaily`, `getBankPending`,
`getRekonsiliasiData`, `findPendingTransactions`, `getLaporanSetoran`,
`getTagihanPenerobos`, `getBukuIRBelumSerah`, `getPembelaanData`, `getJamaahBelumBayar`,
`getKasPenerobos`, `getSerahTerimaList`, `getSerahTerimaDetail`.

> Catatan: beberapa "read" memuat data sensitif (saldo, laporan). Setelah token wajib,
> semua ini otomatis tertutup untuk anonim. Pengetatan cap-per-baca bisa menyusul.

**Ringkasan permukaan:** ~100 fungsi global tanpa `_`; **64** titik pemanggilan
`google.script.run` di `Javascript.html`; **123** sink `innerHTML`.

---

## c. Desain Target

### Skema sheet baru / kolom baru (via fungsi migrasi idempoten — tidak menghapus data)

**Sheet `Perangkat`** (satu baris per perangkat terpercaya):
`ID | Email | Device ID (hash) | Nama Perangkat | Terakhir Dipakai | Status | Created At`

**Sheet `Sesi`** (satu baris per sesi aktif):
`Token (hash) | Email | Device ID (hash) | Dibuat | Kadaluarsa | IP/UA (opsional) | Status`

**Sheet `OTP`** (kode bootstrap sementara, TTL pendek):
`Email | OTP (hash) | Kadaluarsa | Percobaan | Status`

**Master User** — tambah kolom:
`... | PIN Hash | Salt | Perlu Ganti PIN | Gagal Login | Terkunci Sampai`

Semua rahasia (device token, OTP, PIN) **hanya** disimpan sebagai
`hashSecret_(nilai)` = Base64( SHA-256( nilai + salt-per-user + PEPPER ) ).
Plaintext hanya hidup di **email** (OTP) dan **perangkat** (device token di localStorage, PIN di kepala user).
`PEPPER` dan `Spreadsheet ID` → **Script Properties** (bukan hardcode).

### Alur

1. **Bootstrap perangkat baru (OTP):**
   `requestOtp_(email)` → verifikasi email ada & aktif di Master User → buat OTP 6 digit,
   simpan hash + TTL 10 menit di sheet `OTP` → kirim via `MailApp.sendEmail`.
   `verifyOtp_(email, otp)` → cocokkan hash + cek TTL + hitung percobaan (lockout) →
   jika benar → `issueDeviceToken_(email)`.
2. **Terbitkan token perangkat:** `issueDeviceToken_(email)` → token acak
   (`Utilities.getUuid()` + entropi) → simpan **hash**-nya di sheet `Perangkat` → kembalikan
   **plaintext** ke klien (disimpan di `localStorage`). Sekaligus user set **PIN 6 digit**
   (disimpan hash di Master User).
3. **Login harian (perangkat terpercaya):** klien kirim `deviceToken` + `PIN` →
   `verifyDevice_(deviceToken)` cocokkan hash → `verifyPin_(email, pin)` → bila cocok →
   `issueSession_(email, deviceId)`. Efektif 2FA: **punya** (perangkat) + **tahu** (PIN).
4. **Session token:** `issueSession_()` → token acak, simpan hash + kadaluarsa (mis. 12 jam)
   di sheet `Sesi`, kembalikan plaintext ke klien.
5. **Tiap panggilan server** menyertakan session token → `verifySession_(token)` →
   kembalikan `{email, role}` atau null (juga perpanjang "Terakhir Dipakai").

### Helper baru & tanda tangan

```javascript
// — kriptografi —
function hashSecret_(plain, salt)            // -> String (Base64 SHA-256 + salt + PEPPER)
function randomToken_()                       // -> String token acak kuat
function getPepper_()                         // -> String dari Script Properties

// — OTP / bootstrap —
function requestOtp_(email)                   // -> { success }  (kirim email)
function verifyOtp_(email, otp)               // -> { success, deviceToken?, needPin? }
function issueDeviceToken_(email)             // -> { deviceToken }  (simpan hash)
function setPin_(email, pin)                  // -> { success }

// — login harian —
function verifyDevice_(deviceToken)           // -> { ok, email? }
function verifyPin_(email, pin)               // -> { ok, locked? }
function issueSession_(email, deviceId)       // -> { sessionToken, expiresAt }
function verifySession_(token)                // -> { email, role } | null

// — endpoint klien (pengganti login lama) —
function apiRequestOtp(email)                 // wrapper publik -> requestOtp_
function apiVerifyOtp(email, otp)             // -> deviceToken + minta set PIN
function apiSetPin(deviceToken, pin)          // set PIN pertama kali
function apiLogin(deviceToken, pin)           // -> sessionToken (login harian)
function apiLogout(token)                      // hapus sesi
```

### Rendering aman (V5)
Sediakan `esc_()` (server, sudah ada di `Data.gs:2328`) & `App._esc()` (klien, sudah ada) —
**wajib** membungkus semua data user sebelum masuk `innerHTML`. Fase 4 menyisir 123 sink.

---

## d. Dampak Refactor (sesi jadi token-based)

Karena identitas tak lagi dari `UserProperties`, `getCurrentUser()` harus tahu **token**
milik pemanggil. Dua pola dipertimbangkan:

- **Pola per-fungsi (banyak sentuhan):** tambahkan `token` sebagai argumen pertama di
  ~100 endpoint + ubah `checkAuth(token,…)`/`requirePerm(token,cap)` + 64 call-site frontend.
  Risiko regresi tinggi, banyak diff.

- **Pola dispatcher (DIREKOMENDASIKAN — paling sedikit menyentuh server):**
  1 endpoint tunggal `apiCall(token, fnName, argsArray)` yang: `verifySession_(token)` →
  set variabel request-scoped `__REQ_USER` → panggil `this[fnName].apply(null, argsArray)`.
  `getCurrentUser()` cukup diubah **isinya** menjadi `return __REQ_USER || null;`
  (dari `resolveUser_`), sehingga **`checkAuth()` & `requirePerm()` tak berubah tanda tangan**
  dan ~35 fungsi mutasi + ~30 fungsi baca **tidak disentuh sama sekali**.
  Frontend: satu wrapper `App.callServer(fnName, args) -> Promise` menggantikan pemanggilan
  langsung; 64 call-site diarahkan ke wrapper (mekanis, seragam).

**Estimasi terdampak:**
- Server: **inti** hanya `getCurrentUser` (diubah total) + `apiCall` (baru) + fungsi Auth
  baru (§c). ~35 endpoint mutasi & ~30 baca **tetap** (hanya lewat dispatcher).
  Whitelist `fnName` wajib agar dispatcher tak jadi celah RCE fungsi internal.
- Frontend: 64 titik `google.script.run` → `App.callServer`; layar login dirombak
  (OTP → PIN); tambah simpan `deviceToken`/`sessionToken` di `localStorage`.

> **Keputusan yang butuh konfirmasi Anda (jangan ditebak):**
> 1. Panjang PIN tetap **6 digit**? TTL sesi harian berapa jam (usul: 12 jam)?
> 2. OTP dikirim via `MailApp` dari akun pemilik — kuota harian MailApp terbatas
>    (±100/hari consumer, 1500/hari Workspace). Cukup?
> 3. Apakah PIN diset **user sendiri** saat bootstrap, atau **Admin** yang menetapkan awal?
> 4. Perlukah 1 perangkat = 1 user, atau boleh banyak user berbagi 1 perangkat (kiosk)?

---

## e. Rencana Fase 1–4 + Urutan + Risiko Regresi

| Fase | Isi | Output | Risiko regresi |
|------|-----|--------|----------------|
| **1 — Fondasi kripto & skema** | Script Properties (PEPPER, SPREADSHEET_ID); `hashSecret_`, `randomToken_`; migrasi idempoten buat sheet `Perangkat`/`Sesi`/`OTP` + kolom PIN di Master User; guard `sistem.setup` untuk setup/migrasi (tutup V4). | Helper + skema siap; belum mengubah alur login. | **Rendah** — hanya menambah; tidak menyentuh endpoint lama. |
| **2 — Alur OTP → device → PIN → sesi** | `requestOtp_/verifyOtp_/issueDeviceToken_/setPin_/verifyDevice_/verifyPin_/issueSession_/verifySession_` + endpoint `api*`; layar login baru (OTP+PIN) di frontend; simpan token di localStorage. **Login lama `loginWithEmail` dinonaktifkan** (tutup V1). | Bisa login via OTP+PIN paralel dengan sistem lama (feature-flag). | **Sedang** — perubahan login; uji end-to-end wajib. |
| **3 — Dispatcher & wajib token** | `apiCall` + whitelist; `getCurrentUser` → `__REQ_USER`; frontend `App.callServer` gantikan 64 call-site; tutup V2. Tambah cap baru & pasang `requirePerm` di 4 endpoint V3. | Semua panggilan lewat token; anonim tertutup total. | **Tinggi** — menyentuh semua call-site; perlu regresi menyeluruh + rollback plan. |
| **4 — Hardening** | Sisir 123 sink `innerHTML` → `esc_`/`_esc` (tutup V5); rate-limit/lockout (tutup V7); pindah rahasia sisa ke Script Properties (tutup V6); audit log login. | Aplikasi diperkeras. | **Rendah–Sedang** — sebagian besar defensif. |

**Urutan wajib:** 1 → 2 → 3 → 4. Fase 3 hanya boleh jalan setelah Fase 2 terbukti
(ada minimal 1 ADMIN yang sudah bootstrap perangkat+PIN, agar tidak lockout).

**Mitigasi anti-lockout:** sebelum Fase 3 mewajibkan token, pastikan jalur darurat
(fungsi editor `setupSheets`-style) untuk mereset perangkat/PintuADMIN dari Apps Script editor.

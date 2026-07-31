@echo off
setlocal enabledelayedexpansion
chcp 65001 >nul
cd /d "%~dp0"

REM ============================================================
REM  deploy.bat — clasp push + buat versi deploy baru sekaligus
REM  (setara: Deploy > Manage deployments > Edit > New version)
REM  URL /exec tetap sama, hanya kodenya diperbarui ke versi baru.
REM ============================================================

REM  >>> ISI SEKALI: ID deployment web app Anda (lihat langkah di bawah) <<<
set "DEPLOYMENT_ID=AKfycbzAqaIYpcoVWJhJaZ01VpDKJ-FIcf6doR_2tefxO3r1_1BFINu_tBVj94mgdqFOC_96"
set "BRANCH=claude/vigilant-goldberg-5aiati"

echo.
echo === [1/3] git pull (tarik kode terbaru dari GitHub) ===
call git pull --ff-only origin %BRANCH%
if errorlevel 1 (
  echo.
  echo ############################################################
  echo  GAGAL menarik dari GitHub ^(mungkin ada perubahan lokal yang
  echo  belum di-commit, atau riwayat menyimpang^).
  echo  Deploy DIBATALKAN agar tidak mengunggah kode lama ke Apps Script.
  echo.
  echo  Perbaikan umum ^(hati-hati, membuang perubahan lokal^):
  echo     git stash            ^(simpan sementara perubahan lokal^)
  echo     git pull origin %BRANCH%
  echo  atau jika yakin mau ikut GitHub sepenuhnya:
  echo     git fetch origin ^&^& git reset --hard origin/%BRANCH%
  echo ############################################################
  goto :end
)

echo.
echo === [2/3] clasp push (unggah kode) ===
call clasp push --force
if errorlevel 1 goto :error

if "%DEPLOYMENT_ID%"=="" (
  echo.
  echo -----------------------------------------------------------
  echo  DEPLOYMENT_ID belum diisi.
  echo  Daftar deployment Anda:
  echo -----------------------------------------------------------
  call clasp deployments
  echo.
  echo  Salin ID deployment WEB APP ^(baris yang ada kata "exec" / deskripsi web app^),
  echo  buka deploy.bat dengan Notepad, isikan ke baris:
  echo        set "DEPLOYMENT_ID=..."
  echo  lalu jalankan deploy.bat lagi.
  echo.
  goto :end
)

echo.
echo === [3/3] clasp deploy (buat versi baru pada deployment yang sama) ===
call clasp deploy --deploymentId %DEPLOYMENT_ID% --description "auto-deploy %date% %time%"
if errorlevel 1 goto :error

echo.
echo ============================================================
echo  SELESAI. Versi baru sudah LIVE di URL /exec yang sama.
echo ============================================================
goto :end

:error
echo.
echo ############################################################
echo  GAGAL. Periksa pesan error di atas.
echo  Bila error "User has not enabled the Apps Script API":
echo    buka https://script.google.com/home/usersettings
echo    aktifkan "Google Apps Script API", tunggu 1-2 menit, ulangi.
echo ############################################################

:end
echo.
pause
endlocal

@echo off
rem Klik dua kali untuk mengaktifkan add-in AI Assistant.
rem Pilihan aplikasi diingat: lain kali langsung terbuka setelah hitung mundur 3 detik (tekan G untuk ganti).
rem Add-in dimuat dari GitHub Pages (tanpa server lokal). Buka Word/Excel/PowerPoint lewat skrip ini;
rem kalau dibuka dengan cara biasa, tombol AI Assistant tidak muncul. Jendela ini menutup sendiri kalau lancar.
rem Bisa juga langsung dengan nama aplikasinya:  "Aktifkan AI Assistant.cmd" excel
title Aktifkan AI Assistant
cd /d "%~dp0"

where node >nul 2>nul
if errorlevel 1 (
  echo.
  echo   Node.js belum terpasang. Pasang dulu dari https://nodejs.org lalu jalankan lagi.
  echo.
  pause
  exit /b 1
)
if not exist "node_modules\office-addin-dev-settings" (
  echo.
  echo   Menyiapkan add-in untuk pertama kali, mohon tunggu...
  call npm install
)

set "REMEMBER=%LOCALAPPDATA%\Office AI Assistant\last-apps.txt"
set "APPS=%*"
if defined APPS goto run

set "LAST="
if exist "%REMEMBER%" set /p LAST=<"%REMEMBER%"
if not defined LAST goto menu
echo.
echo   Membuka AI Assistant di: %LAST%
echo   Tekan G kalau ingin memilih aplikasi lain.
choice /c LG /n /t 3 /d L >nul
if errorlevel 2 goto menu
set "APPS=%LAST%"
goto run

:menu
echo.
echo   Buka add-in AI Assistant di:
echo.
echo     1. Word
echo     2. Excel
echo     3. PowerPoint
echo     4. Ketiganya
echo.
choice /c 1234 /n /m "  Pilih 1-4: "
rem Hasil choice disimpan dulu: perintah set di antaranya membuat "if errorlevel" tidak terbaca.
set "PICK=%errorlevel%"
set "APPS=word"
if "%PICK%"=="2" set "APPS=excel"
if "%PICK%"=="3" set "APPS=powerpoint"
if "%PICK%"=="4" set "APPS=word excel powerpoint"
if not exist "%LOCALAPPDATA%\Office AI Assistant" mkdir "%LOCALAPPDATA%\Office AI Assistant"
>"%REMEMBER%" echo %APPS%

:run
echo.
node tools\start.js %APPS% --background --hosted
set "CODE=%errorlevel%"
if not "%CODE%"=="0" (
  echo.
  pause
  exit /b %CODE%
)
echo.
echo   Selesai. Jendela ini menutup sendiri...
timeout /t 4 >nul

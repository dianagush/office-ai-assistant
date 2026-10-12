@echo off
rem Klik dua kali agar tombol AI Assistant SELALU ada di Word, Excel, dan PowerPoint,
rem walaupun dibuka dengan cara biasa. Cukup dijalankan sekali.
rem Skrip ini mengubah pengaturan keamanan Office dan membuat satu folder berbagi hanya-baca;
rem Windows akan meminta persetujuan administrator. Lihat dulu tanpa mengubah apa pun:  "Pasang Permanen.cmd" -Periksa
rem Membatalkan semuanya:  "Pasang Permanen.cmd" -Lepas
title Pasang AI Assistant permanen
cd /d "%~dp0"
echo.
powershell -NoProfile -ExecutionPolicy Bypass -File "tools\pasang-permanen.ps1" %*
echo.
pause

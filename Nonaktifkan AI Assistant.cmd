@echo off
rem Klik dua kali untuk menonaktifkan add-in AI Assistant:
rem server add-in dimatikan dan add-in dilepas dari Office.
title Nonaktifkan AI Assistant
cd /d "%~dp0"
echo.
node tools\stop.js
echo.
pause

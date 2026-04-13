@echo off
cd /d "%~dp0"
start "" powershell -WindowStyle Hidden -NoProfile -ExecutionPolicy Bypass -STA -File "v8\zip_multiplas_pastas_pro_v8.ps1"

@echo off
cd /d "%~dp0"
start "" powershell -WindowStyle Hidden -NoProfile -ExecutionPolicy Bypass -STA -File "zip_multiplas_pastas_pro_v6_7.ps1"
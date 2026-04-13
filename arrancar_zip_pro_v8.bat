@echo off
setlocal

set "PS1=%~dp0v8\zip_multiplas_pastas_pro_v8.ps1"

powershell -STA -ExecutionPolicy Bypass -File "%PS1%"

endlocal
pause

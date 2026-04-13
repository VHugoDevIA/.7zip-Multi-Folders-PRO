@echo off
setlocal

set "SCRIPT_DIR=%~dp0"
set "PS1=%SCRIPT_DIR%v8\zip_multiplas_pastas_pro_v8.ps1"

powershell -STA -ExecutionPolicy Bypass -File "%PS1%"

endlocal
pause

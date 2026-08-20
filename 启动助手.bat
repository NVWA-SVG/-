@echo off
cd /d "%~dp0"
set "SELF_TEST="
if /I "%~1"=="--self-test" set "SELF_TEST=-SelfTest"
powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "$env:RENAMER_APP_DIR='%~dp0'; $code=[System.IO.File]::ReadAllText('%~dp0app.ps1',[System.Text.Encoding]::UTF8); & ([ScriptBlock]::Create($code)) %SELF_TEST%" 1>nul
if errorlevel 1 (
  echo.
  echo Launch failed. Please send the error above to the developer.
  pause
)

@echo off
setlocal EnableExtensions
cd /d "%~dp0.."
echo Performance Manager - workbook sync refusal evidence (#138)
echo Save your work and close all Excel windows before continuing.
echo.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Verify-SyncRefusals.ps1" %*
set "CODE=%ERRORLEVEL%"
echo.
pause
exit /b %CODE%

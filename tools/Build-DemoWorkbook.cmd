@echo off
setlocal EnableExtensions
cd /d "%~dp0.."
echo Performance Manager - build and smoke-test the demo workbook (#47)
echo Save your work and close all Excel windows before continuing.
echo.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Build-DemoWorkbook.ps1" -AllowVbomAccess %*
set "CODE=%ERRORLEVEL%"
echo.
pause
exit /b %CODE%

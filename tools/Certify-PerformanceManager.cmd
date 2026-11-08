@echo off
setlocal EnableExtensions
cd /d "%~dp0.."
echo Performance Manager - Excel certification
echo Save your work and close all Excel windows before continuing.
echo.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Certify-PerformanceManager.ps1" %*
set "CODE=%ERRORLEVEL%"
echo.
pause
exit /b %CODE%

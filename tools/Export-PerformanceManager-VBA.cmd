@echo off
setlocal
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Manage-PerformanceManagerVBA.ps1" -Mode Export -AllowVbomAccess %*
set "CODE=%ERRORLEVEL%"
pause
exit /b %CODE%

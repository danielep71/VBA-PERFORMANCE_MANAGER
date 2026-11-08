@echo off
call "%~dp0Certify-PerformanceManager.cmd" -TimerRollbackRedCheck
exit /b %ERRORLEVEL%

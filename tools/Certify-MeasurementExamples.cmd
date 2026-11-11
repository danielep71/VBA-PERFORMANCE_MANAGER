@echo off
call "%~dp0Certify-PerformanceManager.cmd" -MeasurementExamples %*
exit /b %ERRORLEVEL%

@echo off
cd /d "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0push_dashboard_csv.ps1"
set "EC=%ERRORLEVEL%"
echo.
if not "%EC%"=="0" echo [Exit code %EC%]
pause
exit /b %EC%

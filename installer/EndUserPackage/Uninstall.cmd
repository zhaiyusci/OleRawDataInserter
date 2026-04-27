@echo off
setlocal
cd /d "%~dp0"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File ".\uninstall.ps1"
set "EXITCODE=%ERRORLEVEL%"
echo.
if not "%EXITCODE%"=="0" (
  echo Uninstall failed. Please send install.log to the developer.
) else (
  echo Uninstall completed. Restart Microsoft Word.
)
echo.
pause
exit /b %EXITCODE%

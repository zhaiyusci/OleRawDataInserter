@echo off
setlocal
cd /d "%~dp0"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File ".\install.ps1"
set "EXITCODE=%ERRORLEVEL%"
echo.
if not "%EXITCODE%"=="0" (
  echo Installation failed. Please send install.log to the developer.
) else (
  echo Installation completed. Restart Microsoft Word to use the add-in.
)
echo.
pause
exit /b %EXITCODE%

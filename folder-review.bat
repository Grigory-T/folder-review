@echo off
setlocal EnableExtensions DisableDelayedExpansion
cd /d "%~dp0" || exit /b 30
set "POWERSHELL_EXE=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
if not exist "%POWERSHELL_EXE%" (
    echo ERROR: Windows PowerShell was not found.
    exit /b 30
)
if not exist "%~dp0support\launcher.ps1" (
    echo ERROR: Extract the complete folder-review ZIP before starting.
    exit /b 30
)
"%POWERSHELL_EXE%" -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%~dp0support\launcher.ps1" -NoPause
set "REVIEW_EXIT=%ERRORLEVEL%"
echo.
if "%REVIEW_EXIT%"=="0" echo COMPLETE - report created beside folder-review.bat.
if "%REVIEW_EXIT%"=="2" echo PARTIAL - report created; review scan_status.
if not "%REVIEW_EXIT%"=="0" if not "%REVIEW_EXIT%"=="2" echo FAILED - read the error above.
if /i "%~1"=="--no-pause" exit /b %REVIEW_EXIT%
echo Press any key to close this window.
pause >nul
exit /b %REVIEW_EXIT%

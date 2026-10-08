@echo off
setlocal

cd /d "%~dp0" || exit /b 1

set "UV_EXE="
if exist "%~dp0uv.exe" set "UV_EXE=%~dp0uv.exe"
if not defined UV_EXE for /f "delims=" %%I in ('where uv.exe 2^>nul') do if not defined UV_EXE set "UV_EXE=%%I"
if not defined UV_EXE if exist "%USERPROFILE%\.local\bin\uv.exe" set "UV_EXE=%USERPROFILE%\.local\bin\uv.exe"
if not defined UV_EXE for /d %%D in ("%APPDATA%\Python\Python*") do if exist "%%~fD\Scripts\uv.exe" if not defined UV_EXE set "UV_EXE=%%~fD\Scripts\uv.exe"

if not defined UV_EXE (
    echo ERROR: uv.exe was not found.
    exit /b 1
)

echo Using uv: %UV_EXE%

if not exist ".venv\Scripts\python.exe" (
    echo Creating .venv...
    "%UV_EXE%" venv ".venv" --python 3.13
    if errorlevel 1 goto :failed
)

echo Checking dependencies...
"%UV_EXE%" sync --frozen --no-dev --quiet
if errorlevel 1 goto :failed

echo Running folder review...
".venv\Scripts\python.exe" "folder_review.py" %*
set "RUN_EXIT=%ERRORLEVEL%"
if "%RUN_EXIT%"=="2" goto :partial
if not "%RUN_EXIT%"=="0" (
    echo ERROR: folder review failed with exit code %RUN_EXIT%.
    exit /b %RUN_EXIT%
)

echo Completed successfully.
exit /b 0

:partial
echo WARNING: folder review completed with partial results.
echo Review the scan_status sheet in the generated workbook.
exit /b 2

:failed
echo ERROR: folder review failed with exit code %ERRORLEVEL%.
exit /b %ERRORLEVEL%

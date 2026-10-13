@echo off
setlocal

cd /d "%~dp0" || exit /b 1

if not exist "%~dp0.venv\Scripts\python.exe" (
    echo ERROR: The local Python environment is missing. Run SETUP.bat once first.
    echo SETUP.bat prepares dependencies; RUN.bat never downloads or installs anything.
    exit /b 1
)

echo Running folder review...
".venv\Scripts\python.exe" -B "folder_review.py" %*
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

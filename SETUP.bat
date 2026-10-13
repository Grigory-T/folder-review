@echo off
setlocal
cd /d "%~dp0" || exit /b 1
set "UV_EXE="
if exist "%~dp0uv.exe" set "UV_EXE=%~dp0uv.exe"
if not defined UV_EXE for /f "delims=" %%I in ('where uv.exe 2^>nul') do if not defined UV_EXE set "UV_EXE=%%I"
if not defined UV_EXE if exist "%USERPROFILE%\.local\bin\uv.exe" set "UV_EXE=%USERPROFILE%\.local\bin\uv.exe"
if not defined UV_EXE (
    echo ERROR: Install uv before preparing the Python version.
    exit /b 1
)
echo Preparing the local environment. This setup step may access the Internet.
"%UV_EXE%" sync --frozen --no-dev --python 3.13
exit /b %ERRORLEVEL%

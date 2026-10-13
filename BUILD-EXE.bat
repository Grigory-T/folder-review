@echo off
setlocal
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0windows_build\build.ps1" %*
exit /b %ERRORLEVEL%

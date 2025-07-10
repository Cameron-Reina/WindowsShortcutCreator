@echo off
set SCRIPT_DIR=%~dp0bin
set SCRIPT_PATH=%SCRIPT_DIR%\WindowsShortcutCreator.ps1

:: Run PowerShell as admin and close this window immediately
start "" powershell -ExecutionPolicy Bypass -File "%SCRIPT_PATH%"

exit

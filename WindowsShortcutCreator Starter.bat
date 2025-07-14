@echo off
title WindowsShortcutCreator Launcher
echo ======================================================
echo        WindowsShortcutCreator - Version 3.0
echo              Made by: Cameron Reina
echo ======================================================
echo.

:: Check if running as administrator
net session >nul 2>&1
if %errorlevel% == 0 (
    echo [+] Running with Administrator privileges
    echo     Public desktop access: Available
    goto :RunScript
) else (
    echo [!] Running with standard user privileges
    echo     Public desktop access: Limited
    echo.
    echo [*] Launching PowerShell script...
    echo [*] The script will ask if you want to elevate privileges
    echo.
    goto :RunScript
)

:RunScript
set SCRIPT_DIR=%~dp0bin
set SCRIPT_PATH=%SCRIPT_DIR%\WindowsShortcutCreator.ps1

echo [*] Checking script path...
if not exist "%SCRIPT_PATH%" (
    echo [!] ERROR: Script not found at: %SCRIPT_PATH%
    echo.
    pause
    exit /b 1
)

echo [+] Script found: %SCRIPT_PATH%
echo.

echo [*] Launching PowerShell script...
echo [*] Note: This will launch PowerShell with the WindowsShortcutCreator
echo.

:: Use -NoExit to keep PowerShell window open after script completion
:: The script itself handles the restart loop and exit conditions
powershell.exe -ExecutionPolicy Bypass -NoExit -Command "& '%SCRIPT_PATH%'"

:: Check the exit code to determine if this was an intentional restart
if %errorlevel% == 0 (
    :: Exit code 0 means intentional exit (likely restarting with admin privileges)
    exit /b 0
) else (
    :: Non-zero exit code means unexpected closure
    echo.
    echo [!] PowerShell window closed unexpectedly.
    pause
)

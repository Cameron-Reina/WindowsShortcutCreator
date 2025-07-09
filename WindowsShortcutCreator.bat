@echo off
setlocal EnableDelayedExpansion

:: Elevate with UAC if not running as admin
net session >nul 2>&1
if %errorlevel% NEQ 0 (
    powershell -Command "Start-Process cmd -WorkingDirectory '%~dp0' -ArgumentList '/k \"\"%~f0\"\"'" -Verb RunAs"
    exit /b
)

@title WindowsShortcutCreator - Version 1.14
@echo WindowsShortcutCreator - Version 1.14
@echo Made by: Cameron Reina
@echo ------------------------------------------------------

:: Set working directory to script location
cd /d "%~dp0"

:: Define folders and key file paths
set "PathsDir=%CD%\paths"
set "TempDir=%CD%\temp"
set "LogDir=%CD%\logs"
set "InputFile=%PathsDir%\paths.txt"
set "Cscript=%SystemRoot%\System32\cscript.exe"

:: Create necessary folders if missing
for %%F in ("%PathsDir%" "%TempDir%" "%LogDir%") do (
    if not exist %%F (
        mkdir "%%F"
    )
)

:: Create a placeholder paths.txt if missing
if not exist "%InputFile%" (
    echo # Add full paths to executable files, one per line. > "%InputFile%"
    echo # Lines beginning with # are comments and will be ignored. >> "%InputFile%"
)

:: Get timestamp for log filename
for /f %%a in ('powershell -NoProfile -Command "Get-Date -Format yyyy-MM-dd_HH-mm-ss"') do set "Stamp=%%a"
set "LogFile=%LogDir%\shortcut_log_%Stamp%.txt"

:: Create temporary VBScript to handle shortcut creation
set "VBSFile=%TempDir%\create_shortcut.vbs"
> "%VBSFile%" echo Set oWS = CreateObject("WScript.Shell")
>> "%VBSFile%" echo Set oLink = oWS.CreateShortcut(WScript.Arguments(0))
>> "%VBSFile%" echo oLink.TargetPath = WScript.Arguments(1)
>> "%VBSFile%" echo oLink.Save

:: Write log header
>> "%LogFile%" echo Shortcut creation log — %DATE% %TIME%
>> "%LogFile%" echo ------------------------------------------------------

echo Creating shortcuts...
set /a createdCount=0

:: Read each line from paths.txt and skip comments
for /f "usebackq tokens=*" %%A in ("%InputFile%") do (
    echo %%A | findstr /b /c:"#">nul
    if errorlevel 1 (
        set "rawPath=%%A"
        set "rawPath=!rawPath:"=!"
        set "baseName=%%~nA"
        set "shortcutFile=%PUBLIC%\Desktop\!baseName!.lnk"

        call "%Cscript%" //nologo "%VBSFile%" "!shortcutFile!" "!rawPath!"

        if exist "!shortcutFile!" (
            echo Created: !baseName!.lnk
            >> "%LogFile%" echo SUCCESS: !baseName!.lnk → !rawPath!
            set /a createdCount+=1
        ) else (
            echo FAILED: !rawPath!
            >> "%LogFile%" echo FAILED: !rawPath!
        )
    )
)

:: Delete temporary VBScript
del "%VBSFile%" >nul 2>&1

:: Delete the entire temp folder
rd /s /q "%TempDir%"


:: Wait for user input before exiting
echo.
echo Press ENTER to exit...
pause >nul
exit

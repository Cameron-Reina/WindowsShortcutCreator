# WindowsShortcutCreator.ps1
# Version 2.0 - Full PowerShell Rewrite
# Made by: Cameron Reina

# --- Set Window Title and Banner ---
$Host.UI.RawUI.WindowTitle = "WindowsShortcutCreator - Version 2.0"
Write-Host "WindowsShortcutCreator - Version 2.0"
Write-Host "Made by: Cameron Reina"
Write-Host "------------------------------------------------------"

# --- UAC Elevation ---
$currUser = [Security.Principal.WindowsIdentity]::GetCurrent()
$principal = New-Object Security.Principal.WindowsPrincipal($currUser)
if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host "This script needs to run as administrator."
    Write-Host "A UAC prompt will appear."
    Start-Process -FilePath "powershell.exe" -ArgumentList @("-ExecutionPolicy", "Bypass", "-File", "$($MyInvocation.MyCommand.Definition)") -Verb RunAs
    exit
}

# --- Setup Environment ---
$baseDir   = Split-Path -Parent $PSCommandPath
$pathsDir  = "$baseDir\paths"
$logDir    = "$baseDir\logs"
$inputFile = "$pathsDir\paths.txt"
$prefFile  = "$pathsDir\shortcut_destination.txt"
$timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
$logFile   = "$logDir\shortcut_log_$timestamp.txt"

# --- Create folders ---
foreach ($folder in @($pathsDir, $logDir)) {
    if (!(Test-Path $folder)) { New-Item -ItemType Directory -Path $folder | Out-Null }
}

# --- Logger helper ---
Add-Content $logFile "`n=== WindowsShortcutCreator Log - $(Get-Date) ==="
Add-Content $logFile "Script directory: $baseDir"
Add-Content $logFile "------------------------------------------------------`n"

# --- Get shortcut targets ---
Write-Host "How do you want to select shortcut targets?"
Write-Host "[1] Use an existing .txt file"
Write-Host "[2] Select files via GUI"
$targetChoice = Read-Host "Enter choice [1/2] (default: 2)"
if (-not $targetChoice) { $targetChoice = "2" }

if ($targetChoice -eq "1") {
    $inputFilePrompt = Read-Host "Enter path to .txt file with shortcut targets (default: $inputFile)"
    if ($inputFilePrompt) { $inputFile = $inputFilePrompt }
    if (!(Test-Path $inputFile)) {
        Write-Host "File not found: $inputFile. Script aborted."
        Read-Host "Press ENTER to exit..."
        exit
    }
    Add-Content $logFile "paths.txt loaded from user-specified file: $inputFile"
} else {
    Add-Type -AssemblyName System.Windows.Forms
    $fileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $fileDialog.Title = "Select files to create shortcuts for"
    $fileDialog.InitialDirectory = [Environment]::GetFolderPath("Desktop")
    $fileDialog.Filter = "All files (*.*)|*.*"
    $fileDialog.Multiselect = $true
    if ($fileDialog.ShowDialog() -eq "OK") {
        $fileDialog.FileNames | Set-Content -Encoding ASCII $inputFile
        Add-Content $logFile "paths.txt created via GUI picker"
    } else {
        Write-Host "`nNo files selected. Script aborted."
        Read-Host "Press ENTER to exit..."
        exit
    }
}

$paths = Get-Content $inputFile | Where-Object { $_ -and ($_ -notmatch "^#") }

# --- Destination setup ---
$prevDest = if (Test-Path $prefFile) { Get-Content $prefFile } else { "" }
$defaultDest = [Environment]::GetFolderPath("CommonDesktopDirectory")

Write-Host "`nChoose shortcut destination:"
Write-Host "[1] Use previously used location: $prevDest"
Write-Host "[2] Use default location: $defaultDest"
Write-Host "[3] Select folder via GUI"

$choice = Read-Host "Enter choice [1/2/3] (default: 1)"
if (-not $choice) { $choice = "1" }

switch ($choice) {
    "1" {
        $dest = $prevDest
    }
    "2" {
        $dest = $defaultDest
        # Save this as the new previous location
        $dest | Set-Content $prefFile
        Add-Content $logFile "Destination chosen: $dest (set as previous location)"
    }
    "3" {
        Add-Type -AssemblyName System.Windows.Forms
        $folderDialog = New-Object System.Windows.Forms.OpenFileDialog
        $folderDialog.Title = "Select destination folder for shortcuts"
        $folderDialog.InitialDirectory = [Environment]::GetFolderPath("Desktop")
        $folderDialog.Filter = "Folder|."
        $folderDialog.CheckFileExists = $false
        $folderDialog.CheckPathExists = $true
        $folderDialog.FileName = "Select Folder"
        if ($folderDialog.ShowDialog() -eq "OK") {
            $dest = [System.IO.Path]::GetDirectoryName($folderDialog.FileName)
            $dest | Set-Content $prefFile
            Add-Content $logFile "Destination chosen: $dest"
        } else {
            Write-Host "`nNo folder selected. Script aborted."
            Read-Host "Press ENTER to exit..."
            exit
        }
    }
}

Write-Host "`nUsing destination: $dest"
if (!(Test-Path $dest)) { New-Item -ItemType Directory -Path $dest | Out-Null }

# --- Shortcut creation ---
$wshell = New-Object -ComObject WScript.Shell
$created = 0

Write-Host "`nCreating shortcuts..."
foreach ($path in $paths) {
    $target = $path.Trim([char]34)
    $name   = [System.IO.Path]::GetFileNameWithoutExtension($target)
    $link   = "$dest\$name.lnk"

    try {
        $shortcut = $wshell.CreateShortcut($link)
        $shortcut.TargetPath = $target
        $shortcut.Save()
        Add-Content $logFile "CREATED: $link -> $target"
        $created++
    } catch {
        Add-Content $logFile "FAILED: $target"
    }
}

# --- Done ---
Write-Host "`nShortcuts created: $created"
Write-Host "Log saved to: $logFile"
Add-Content $logFile "`nTotal shortcuts created: $created"
Add-Content $logFile "=== Script completed successfully ==="
Write-Host "`nPress ENTER to exit..."
[void][System.Console]::ReadLine()

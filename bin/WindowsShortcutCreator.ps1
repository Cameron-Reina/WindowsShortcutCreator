# WindowsShortcutCreator.ps1
# Version 3.0 - Full PowerShell Rewrite
# Made by: Cameron Reina

# --- Set Window Title and Banner ---
$Host.UI.RawUI.WindowTitle = "WindowsShortcutCreator - Version 3.0"

# Global variable to track WSCPaths file for cleanup
$global:WSCPathsFile = ""

# Register cleanup event for when PowerShell window is closed
Register-EngineEvent PowerShell.Exiting -Action {
    if ($global:WSCPathsFile -and (Test-Path $global:WSCPathsFile)) {
        try {
            Remove-Item $global:WSCPathsFile -Force
        } catch {
            # Silent cleanup attempt
        }
    }
} | Out-Null

# Function to show dialog in foreground
function Show-DialogInForeground {
    param($Dialog)
    
    # Get the current PowerShell window
    Add-Type -TypeDefinition @"
        using System;
        using System.Runtime.InteropServices;
        public class Win32 {
            [DllImport("user32.dll")]
            public static extern IntPtr GetForegroundWindow();
            [DllImport("user32.dll")]
            public static extern bool SetForegroundWindow(IntPtr hWnd);
            [DllImport("kernel32.dll")]
            public static extern IntPtr GetConsoleWindow();
        }
"@
    
    try {
        # Store current window
        $consoleWindow = [Win32]::GetConsoleWindow()
        
        # Show dialog and bring it to front
        $result = $Dialog.ShowDialog()
        
        # Restore focus to console
        if ($consoleWindow -ne [IntPtr]::Zero) {
            [Win32]::SetForegroundWindow($consoleWindow) | Out-Null
        }
        
        return $result
    } catch {
        # Fallback to normal ShowDialog if Win32 API fails
        return $Dialog.ShowDialog()
    }
}

# Main script function that can be called repeatedly
function Start-ShortcutCreator {
    # Initialize tracking variable for file type
    $usingManualFile = $false
    $wscPathsFile = ""
      try {
        # Clear the screen for a clean start
        Clear-Host

        Write-Host ""
        Write-Host "======================================================"
        Write-Host "       WindowsShortcutCreator - Version 3.0"
        Write-Host "              Made by: Cameron Reina"
        Write-Host "======================================================"
        Write-Host ""

        # --- UAC Elevation Check ---
        $currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
        $isAdmin = $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
        $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent().Name

        if ($isAdmin) {
            Write-Host "[+] Running with administrator privileges..." -ForegroundColor Green
            Write-Host "    Current user: $currentUser" -ForegroundColor Green
            Write-Host "    Public desktop access: Available" -ForegroundColor Green
            Write-Host ""
        } else {
            Write-Host "[!] Running with standard user privileges..." -ForegroundColor Yellow
            Write-Host "    Current user: $currentUser" -ForegroundColor Yellow
            Write-Host "    Public desktop access: Limited" -ForegroundColor Yellow
            Write-Host ""
            
            # Show additional info if running as administrator account but not elevated
            if ($currentUser -match "Administrator|admin") {
                Write-Host "    Note: You're using an administrator account, but the process is not elevated." -ForegroundColor Cyan
                Write-Host "    Even administrator accounts need to 'Run as administrator' for full privileges." -ForegroundColor Cyan
                Write-Host ""
            }
            
            Write-Host "Would you like to restart with administrator privileges?" -ForegroundColor Cyan
            Write-Host ""
            Write-Host "  [Y] Yes - Enable public desktop access (recommended)" -ForegroundColor Green
            Write-Host "  [N] No - Continue with current privileges" -ForegroundColor Yellow
            Write-Host ""
            
            $elevateChoice = Read-Host "Restart as administrator? [Y/N] (default: Y)"
            if (-not $elevateChoice) { $elevateChoice = "Y" }
            
            if ($elevateChoice.ToUpper() -eq "Y") {
                Write-Host ""
                Write-Host "[*] Restarting with administrator privileges..." -ForegroundColor Cyan
                Write-Host "    This will enable access to the public desktop for all users" -ForegroundColor Cyan
                try {
                    # Find the batch file that launched us and restart it with admin privileges
                    $batchFile = $null
                    $parentProcess = Get-WmiObject Win32_Process -Filter "ProcessId = $((Get-WmiObject Win32_Process -Filter "ProcessId = $PID").ParentProcessId)"
                    if ($parentProcess -and $parentProcess.CommandLine -match '\.bat') {
                        # Extract the batch file path from the command line
                        if ($parentProcess.CommandLine -match '"([^"]*\.bat)"') {
                            $batchFile = $matches[1]
                        } elseif ($parentProcess.CommandLine -match '(\S*\.bat)') {
                            $batchFile = $matches[1]
                        }
                    }
                    
                    if ($batchFile -and (Test-Path $batchFile)) {
                        Write-Host "    Restarting batch file: $([System.IO.Path]::GetFileName($batchFile))" -ForegroundColor Cyan
                        Start-Process cmd -ArgumentList "/c `"$batchFile`"" -Verb RunAs
                        Write-Host ""
                        Write-Host "[*] New elevated window should open. This window will close..." -ForegroundColor Green
                        # Force exit this PowerShell session
                        [Environment]::Exit(0)
                    } else {
                        # Fallback: restart this PowerShell script directly
                        Write-Host "    Restarting PowerShell script directly..." -ForegroundColor Cyan
                        Start-Process PowerShell -ArgumentList "-ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
                        Write-Host ""
                        Write-Host "[*] New elevated window should open. This window will close..." -ForegroundColor Green
                        # Force exit this PowerShell session
                        [Environment]::Exit(0)
                    }
                } catch {
                    Write-Host ""
                    Write-Host "[!] Failed to restart as administrator. Continuing with current privileges..." -ForegroundColor Red
                    Write-Host "    You can still use personal desktop and custom folders." -ForegroundColor Yellow
                    Write-Host ""
                }
            } else {
                Write-Host ""
                Write-Host "[+] Continuing with standard privileges..." -ForegroundColor Green
                Write-Host "    Note: Public desktop will not be available" -ForegroundColor Yellow
                Write-Host ""
            }
        }

        # --- Setup Environment ---
        $baseDir   = Split-Path -Parent $PSCommandPath
        $pathsDir  = "$baseDir\paths"
        $logDir    = "$baseDir\logs"
        $inputFile = "$pathsDir\paths.txt"
        $wscPathsFile = "$pathsDir\TemporaryFileSelection.txt"  # Temporary file for GUI selections
        $global:WSCPathsFile = $wscPathsFile  # Set global for cleanup
        $prefFile  = "$pathsDir\shortcut_destination.txt"
        $timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
        $logFile   = "$logDir\shortcut_log_$timestamp.txt"

    # --- Create folders ---
    foreach ($folder in @($pathsDir, $logDir)) {
        if (!(Test-Path $folder)) { New-Item -ItemType Directory -Path $folder | Out-Null }
    }

    # Clean up any existing TemporaryFileSelection.txt from previous runs
    if (Test-Path $wscPathsFile) {
        try {
            Remove-Item $wscPathsFile -Force
            Add-Content $logFile "Cleaned up existing TemporaryFileSelection.txt from previous run"
        } catch {
            Add-Content $logFile "Warning: Could not clean up existing TemporaryFileSelection.txt - $($_.Exception.Message)"
        }
    }

    # --- Logger helper ---
    Add-Content $logFile "=== WindowsShortcutCreator Log - $(Get-Date) ==="
    Add-Content $logFile "Script version: 3.0"
    Add-Content $logFile "PowerShell version: $($PSVersionTable.PSVersion)"
    Add-Content $logFile "User: $($env:USERNAME)"
    Add-Content $logFile "Computer: $($env:COMPUTERNAME)"
    Add-Content $logFile "Administrator privileges: $isAdmin"
    Add-Content $logFile "Script directory: $baseDir"
    Add-Content $logFile "Paths directory: $pathsDir"
    Add-Content $logFile "Log directory: $logDir"
    Add-Content $logFile "Input file: $inputFile"
    Add-Content $logFile "WSC Paths file: $wscPathsFile"
    Add-Content $logFile "Preferences file: $prefFile"
    Add-Content $logFile "------------------------------------------------------"
    
    Write-Host "[+] Environment setup complete" -ForegroundColor Green
    Write-Host "    - Paths directory: $pathsDir"
    Write-Host "    - Log directory: $logDir"
    Write-Host "    - Log file: $logFile"
    Write-Host "    - Manual paths file: paths.txt"
    Write-Host "    - GUI selections file: TemporaryFileSelection.txt (auto-deleted)"
    Write-Host ""

    # --- Get shortcut targets ---
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host "[1] STEP 1: Select Shortcut Targets" -ForegroundColor Cyan
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "How do you want to select shortcut targets?" -ForegroundColor White
    Write-Host ""
    Write-Host "  [1] Select a text file containing target paths (GUI File Picker)"
    Write-Host "  [2] Select target files directly (GUI File Picker)"
    Write-Host ""
    $targetChoice = Read-Host "Enter choice [1/2] (default: 2)"
    if (-not $targetChoice) { $targetChoice = "2" }

    Write-Host ""
    Write-Host "[1] STEP 1: Select Shortcut Targets - Choice: $targetChoice" -ForegroundColor Green
    Write-Host ""

    Add-Content $logFile "User choice for target selection: $targetChoice"

    # Initialize tracking variable for file type
    $usingManualFile = $false

    if ($targetChoice -eq "1") {
        Write-Host "[*] Using text file method..." -ForegroundColor Cyan
        Write-Host "    NOTE: File selection dialog will open. It may appear behind this window initially." -ForegroundColor Yellow
        Add-Content $logFile "User selected text file method"
        
        # GUI file picker for text file selection
        Add-Type -AssemblyName System.Windows.Forms
        $textFileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $textFileDialog.Title = "Select a text file containing shortcut target paths"
        $textFileDialog.InitialDirectory = $pathsDir
        $textFileDialog.Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*"
        $textFileDialog.FilterIndex = 1
        $textFileDialog.Multiselect = $false
        
        if ((Show-DialogInForeground $textFileDialog) -eq "OK") {
            $inputFile = $textFileDialog.FileName
            Write-Host "[+] Selected text file: $([System.IO.Path]::GetFileName($inputFile))" -ForegroundColor Green
            Add-Content $logFile "Text file selected via GUI: $inputFile"
        } else {
            Write-Host ""
            Write-Host "[!] No text file selected. Using default paths.txt..." -ForegroundColor Yellow
            Add-Content $logFile "No text file selected via GUI, using default: $inputFile"
        }
        
        if (!(Test-Path $inputFile)) {
            Write-Host ""
            Write-Host "[!] ERROR: File not found: $inputFile" -ForegroundColor Red
            Write-Host "    Script aborted." -ForegroundColor Red
            Add-Content $logFile "ERROR: Input file not found: $inputFile"
            Write-Host ""
            return
        }
        
        Write-Host "[+] File found and loaded successfully" -ForegroundColor Green
        Add-Content $logFile "SUCCESS: Text file loaded: $inputFile"
        $usingManualFile = $true
    } else {
        Write-Host "[*] Opening target file selection dialog..." -ForegroundColor Cyan
        Write-Host "    NOTE: File selection dialog will open. Look for it if it appears behind this window." -ForegroundColor Yellow
        Add-Content $logFile "Using GUI file picker for target files"
        
        Add-Type -AssemblyName System.Windows.Forms
        $fileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $fileDialog.Title = "Select files to create shortcuts for"
        $fileDialog.InitialDirectory = [Environment]::GetFolderPath("Desktop")
        $fileDialog.Filter = "All files (*.*)|*.*|Executable files (*.exe)|*.exe|Documents (*.pdf;*.doc;*.docx)|*.pdf;*.doc;*.docx"
        $fileDialog.FilterIndex = 1
        $fileDialog.Multiselect = $true
        
        if ((Show-DialogInForeground $fileDialog) -eq "OK") {
            $selectedCount = $fileDialog.FileNames.Count
            Write-Host "[+] Selected $selectedCount file(s)" -ForegroundColor Green
            # Use TemporaryFileSelection.txt for GUI selections - this will be wiped after completion
            $inputFile = $wscPathsFile
            # Add quotes around each path for consistency and robustness
            $quotedPaths = $fileDialog.FileNames | ForEach-Object { "`"$_`"" }
            $quotedPaths | Out-File -FilePath $inputFile -Encoding UTF8
            Add-Content $logFile "SUCCESS: $selectedCount files selected via GUI picker"
            Add-Content $logFile "Files saved to temporary file: $inputFile"
            foreach ($file in $fileDialog.FileNames) {
                Add-Content $logFile "  - `"$file`""
            }
            $usingManualFile = $false
        } else {
            Write-Host ""
            Write-Host "[!] No files selected. Script aborted." -ForegroundColor Red
            Add-Content $logFile "ERROR: No files selected in GUI picker"
            Write-Host ""
            return
        }
    }

    Write-Host ""
    $paths = Get-Content $inputFile -Encoding UTF8 | Where-Object { $_ -and ($_ -notmatch "^#") }
    
    # Remove quotes from paths if they exist (handles both quoted and unquoted paths)
    $cleanedPaths = $paths | ForEach-Object {
        $cleanPath = $_.Trim()
        if ($cleanPath.StartsWith('"') -and $cleanPath.EndsWith('"')) {
            $cleanPath.Substring(1, $cleanPath.Length - 2)
        } else {
            $cleanPath
        }
    }
    
    $pathCount = $cleanedPaths.Count

    # Filter out incompatible file types and separate shortcut files for copying
    $validPaths = @()
    $shortcutFiles = @()

    foreach ($path in $cleanedPaths) {
        $extension = [System.IO.Path]::GetExtension($path).ToLower()
        # Separate .url files and other shortcut types for copying instead of creating new shortcuts
        if ($extension -in @('.url', '.lnk', '.appref-ms')) {
            $shortcutFiles += $path
            Add-Content $logFile "SHORTCUT TO COPY: $path (file type: $extension - will be copied)"
        } else {
            $validPaths += $path
            Add-Content $logFile "TARGET FOR SHORTCUT: $path (file type: $extension - will create .lnk shortcut)"
        }
    }

    Clear-Host
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host "[1] STEP 1: Target Selection Complete" -ForegroundColor Cyan
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "[i] Found $pathCount total target(s)" -ForegroundColor Cyan
    if ($usingManualFile) {
        Write-Host "[i] Using manual paths file: $([System.IO.Path]::GetFileName($inputFile))" -ForegroundColor Cyan
    } else {
        Write-Host "[i] Using temporary GUI selections (will be cleaned up after completion)" -ForegroundColor Cyan
    }
    
    if ($shortcutFiles.Count -gt 0) {
        Write-Host ""
        Write-Host "[*] $($shortcutFiles.Count) existing shortcut file(s) will be copied:" -ForegroundColor Cyan
        foreach ($shortcut in $shortcutFiles) {
            $fileName = [System.IO.Path]::GetFileName($shortcut)
            Write-Host "    - $fileName" -ForegroundColor Cyan
        }
    }
    
    if ($validPaths.Count -gt 0) {
        Write-Host ""
        if ($shortcutFiles.Count -gt 0) {
            Write-Host "[+] $($validPaths.Count) file(s) will have new shortcuts created:" -ForegroundColor Green
        } else {
            Write-Host "[+] $($validPaths.Count) file(s) will have shortcuts created:" -ForegroundColor Green
        }
        foreach ($target in $validPaths) {
            $fileName = [System.IO.Path]::GetFileName($target)
            Write-Host "    - $fileName" -ForegroundColor Green
        }
    }

    Add-Content $logFile "Total paths found: $pathCount"
    Add-Content $logFile "Files for shortcut creation: $($validPaths.Count)"
    Add-Content $logFile "Shortcut files to copy: $($shortcutFiles.Count)"

    if ($validPaths.Count -eq 0 -and $shortcutFiles.Count -eq 0) {
        Write-Host ""
        Write-Host "[!] ERROR: No valid targets found!" -ForegroundColor Red
        Add-Content $logFile "ERROR: No valid targets found"
        Write-Host ""
        return
    }

    Write-Host ""
    Write-Host "Press ENTER to continue to destination selection..." -ForegroundColor Gray
    Read-Host

    # --- Destination setup ---
    Clear-Host
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host "[2] STEP 2: Choose Shortcut Destination" -ForegroundColor Cyan
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host ""
    
    $prevDest = if (Test-Path $prefFile) { Get-Content $prefFile } else { "" }
    $userDesktop = [Environment]::GetFolderPath("Desktop")  # User's personal desktop
    $publicDesktop = [Environment]::GetFolderPath("CommonDesktopDirectory")  # Public desktop (requires admin)

    Add-Content $logFile "Previous destination (if any): $prevDest"
    Add-Content $logFile "User desktop: $userDesktop"
    Add-Content $logFile "Public desktop: $publicDesktop"

    # Destination selection loop (for Option 5 to loop back)
    do {
        $destinationSelected = $false
        
        # Show current preference status
        if ($prevDest) {
            Write-Host "[i] Current saved preference: $prevDest" -ForegroundColor Cyan
            Write-Host ""
        } else {
            Write-Host "[i] No saved preference - will use default location" -ForegroundColor Cyan
            Write-Host ""
        }
        
        Write-Host "Where would you like to save the shortcuts?" -ForegroundColor White
        Write-Host ""
        
        if ($prevDest) {
            Write-Host "  [1] Use saved location: $prevDest" -ForegroundColor Cyan
            
            # Only show other options if they're different from saved location
            $optionNumber = 2
            
            # Show public desktop option if admin and different from saved location
            if ($isAdmin -and $publicDesktop -ne $prevDest) {
                Write-Host "  [$optionNumber] Use public desktop (all users): $publicDesktop" -ForegroundColor Green
                $optionNumber++
            } elseif (!$isAdmin -and $publicDesktop -ne $prevDest) {
                Write-Host "  [$optionNumber] Use public desktop (all users): $publicDesktop" -ForegroundColor DarkGray
                Write-Host "      (Requires administrator privileges - not available)" -ForegroundColor DarkGray
                $optionNumber++
            }
            
            # Show personal desktop option if different from saved location
            if ($userDesktop -ne $prevDest) {
                Write-Host "  [$optionNumber] Use your personal desktop: $userDesktop"
                $optionNumber++
            }
            
            # Always show GUI option
            Write-Host "  [$optionNumber] Select folder via GUI (Folder Browser)"
            $optionNumber++
            
            # Always show clear preference option
            Write-Host "  [$optionNumber] Clear saved preference and choose from all options" -ForegroundColor Yellow
            $maxChoice = $optionNumber
        } else {
            # No saved preference - show cleaner options
            if ($isAdmin) {
                Write-Host "  [1] Use public desktop (all users): $publicDesktop" -ForegroundColor Green
            } else {
                Write-Host "  [1] Use public desktop (all users): $publicDesktop" -ForegroundColor DarkGray
                Write-Host "      (Requires administrator privileges - not available)" -ForegroundColor DarkGray
            }
            Write-Host "  [2] Use your personal desktop: $userDesktop"
            Write-Host "  [3] Select folder via GUI (Folder Browser)"
            $maxChoice = 3
        }
        
        if (-not $isAdmin) {
            Write-Host ""
            if ($prevDest) {
                Write-Host "Note: Public desktop option is disabled (requires administrator privileges)" -ForegroundColor Yellow
            } else {
                Write-Host "Note: Option 1 is disabled (requires administrator privileges)" -ForegroundColor Yellow
            }
        }
        Write-Host ""

        $defaultChoice = "1"
        $choice = Read-Host "Enter choice [1-$maxChoice] (default: $defaultChoice)"
        if (-not $choice) { $choice = $defaultChoice }

        Clear-Host
        Write-Host "======================================================" -ForegroundColor Cyan
        Write-Host "[2] STEP 2: Destination Selection - Choice: $choice" -ForegroundColor Cyan
        Write-Host "======================================================" -ForegroundColor Cyan
        Write-Host ""

        Add-Content $logFile "User choice for destination: $choice"
        Add-Content $logFile "Max choice allowed: $maxChoice"
        Add-Content $logFile "Has saved preference: $($prevDest -ne '')"
        Add-Content $logFile "Public desktop different from saved: $($publicDesktop -ne $prevDest)"
        Add-Content $logFile "User desktop different from saved: $($userDesktop -ne $prevDest)"
        
        # Validate choice is a number
        if (-not [int]::TryParse($choice, [ref]$null)) {
            Write-Host "[!] Invalid choice '$choice'. Please enter a number..." -ForegroundColor Red
            Clear-Host
            Write-Host "======================================================" -ForegroundColor Cyan
            Write-Host "[2] STEP 2: Choose Shortcut Destination" -ForegroundColor Cyan
            Write-Host "======================================================" -ForegroundColor Cyan
            Write-Host ""
            continue
        }
        
        $choiceNum = [int]$choice
        
        # Validate choice is within range
        if ($choiceNum -lt 1 -or $choiceNum -gt $maxChoice) {
            Write-Host "[!] Invalid choice '$choice'. Please enter a number between 1 and $maxChoice..." -ForegroundColor Red
            Add-Content $logFile "Invalid choice: $choice (out of range 1-$maxChoice)"
            Clear-Host
            Write-Host "======================================================" -ForegroundColor Cyan
            Write-Host "[2] STEP 2: Choose Shortcut Destination" -ForegroundColor Cyan
            Write-Host "======================================================" -ForegroundColor Cyan
            Write-Host ""
            continue
        }

        if ($prevDest) {
            # Handle saved preference menu structure
            Add-Content $logFile "Processing choice with saved preference logic"
            
            if ($choiceNum -eq 1) {
                # Option 1 with saved preference = use saved location
                $dest = $prevDest
                Write-Host "[+] Using saved location: $dest" -ForegroundColor Green
                Add-Content $logFile "Using saved destination: $dest"
                $destinationSelected = $true
            } else {
                # Build option map for dynamic menu
                $optionMap = @{}
                $currentOption = 2
                
                # Map public desktop option if it exists
                if ($publicDesktop -ne $prevDest) {
                    $optionMap[$currentOption] = "public"
                    Add-Content $logFile "Option $currentOption mapped to: public desktop"
                    $currentOption++
                }
                
                # Map personal desktop option if it exists
                if ($userDesktop -ne $prevDest) {
                    $optionMap[$currentOption] = "personal"
                    Add-Content $logFile "Option $currentOption mapped to: personal desktop"
                    $currentOption++
                }
                
                # Map GUI option
                $optionMap[$currentOption] = "gui"
                Add-Content $logFile "Option $currentOption mapped to: GUI folder picker"
                $currentOption++
                
                # Map clear preference option
                $optionMap[$currentOption] = "clear"
                Add-Content $logFile "Option $currentOption mapped to: clear preference"
                
                # Process the selected option
                if ($optionMap.ContainsKey($choiceNum)) {
                    $selectedOption = $optionMap[$choiceNum]
                    Add-Content $logFile "User selected option type: $selectedOption"
                    
                    switch ($selectedOption) {
                        "public" {
                            if ($isAdmin) {
                                $dest = $publicDesktop
                                Write-Host "[+] Using public desktop (all users): $dest" -ForegroundColor Green
                                Add-Content $logFile "Using public desktop: $dest"
                            } else {
                                Write-Host "[!] Public desktop requires administrator privileges!" -ForegroundColor Red
                                Write-Host "    Switching to your personal desktop instead..." -ForegroundColor Yellow
                                $dest = $userDesktop
                                Write-Host "[+] Using your personal desktop: $dest" -ForegroundColor Green
                                Add-Content $logFile "Public desktop not available, using user desktop: $dest"
                            }
                            $destinationSelected = $true
                        }
                        "personal" {
                            $dest = $userDesktop
                            Write-Host "[+] Using your personal desktop: $dest" -ForegroundColor Green
                            Add-Content $logFile "Using user desktop: $dest"
                            $destinationSelected = $true
                        }
                        "gui" {
                            Write-Host "[*] Opening folder selection dialog..." -ForegroundColor Cyan
                            Write-Host "    NOTE: Folder dialog will open. Look for it if it appears behind this window." -ForegroundColor Yellow
                            Add-Content $logFile "Using GUI folder picker"
                            
                            Add-Type -AssemblyName System.Windows.Forms
                            $folderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
                            $folderDialog.Description = "Select destination folder for shortcuts"
                            $folderDialog.SelectedPath = [Environment]::GetFolderPath("Desktop")
                            $folderDialog.ShowNewFolderButton = $true
                            
                            if ((Show-DialogInForeground $folderDialog) -eq "OK") {
                                $dest = $folderDialog.SelectedPath
                                Write-Host "[+] Selected folder: $dest" -ForegroundColor Green
                                Add-Content $logFile "Custom destination selected: $dest"
                                $destinationSelected = $true
                            } else {
                                Write-Host ""
                                Write-Host "[!] No folder selected. Please choose again..." -ForegroundColor Yellow
                                Add-Content $logFile "No folder selected in GUI picker, returning to menu"
                                Write-Host ""
                                Clear-Host
                                Write-Host "======================================================" -ForegroundColor Cyan
                                Write-Host "[2] STEP 2: Choose Shortcut Destination" -ForegroundColor Cyan
                                Write-Host "======================================================" -ForegroundColor Cyan
                                Write-Host ""
                            }
                        }
                        "clear" {
                            Write-Host "[*] Clearing saved preference..." -ForegroundColor Cyan
                            if (Test-Path $prefFile) {
                                Remove-Item $prefFile -Force
                            }
                            $prevDest = ""  # Clear the variable so the menu updates
                            Add-Content $logFile "Saved preference cleared, returning to destination menu"
                            Write-Host "[+] Preference cleared. You can now choose from all available options..." -ForegroundColor Green
                            Write-Host ""
                            Clear-Host
                            Write-Host "======================================================" -ForegroundColor Cyan
                            Write-Host "[2] STEP 2: Choose Shortcut Destination" -ForegroundColor Cyan
                            Write-Host "======================================================" -ForegroundColor Cyan
                            Write-Host ""
                        }
                    }
                } else {
                    Write-Host "[!] Invalid choice. Please try again..." -ForegroundColor Red
                    Add-Content $logFile "Invalid choice: $choice (not found in option map)"
                    Clear-Host
                    Write-Host "======================================================" -ForegroundColor Cyan
                    Write-Host "[2] STEP 2: Choose Shortcut Destination" -ForegroundColor Cyan
                    Write-Host "======================================================" -ForegroundColor Cyan
                    Write-Host ""
                }
            }
        } else {
            # Handle no saved preference menu structure (original 3-option layout)
            switch ($choiceNum) {
                1 {
                    # Option 1 without saved preference = public desktop
                    if ($isAdmin) {
                        $dest = $publicDesktop
                        Write-Host "[+] Using public desktop (all users): $dest" -ForegroundColor Green
                        Add-Content $logFile "Using public desktop: $dest"
                    } else {
                        Write-Host "[!] Public desktop requires administrator privileges!" -ForegroundColor Red
                        Write-Host "    Switching to your personal desktop instead..." -ForegroundColor Yellow
                        $dest = $userDesktop
                        Write-Host "[+] Using your personal desktop: $dest" -ForegroundColor Green
                        Add-Content $logFile "Public desktop not available, using user desktop: $dest"
                    }
                    $destinationSelected = $true
                }
                2 {
                    # Option 2 without saved preference = personal desktop
                    $dest = $userDesktop
                    Write-Host "[+] Using your personal desktop: $dest" -ForegroundColor Green
                    Add-Content $logFile "Using user desktop: $dest"
                    $destinationSelected = $true
                }
                3 {
                    # Option 3 without saved preference = GUI folder picker
                    Write-Host "[*] Opening folder selection dialog..." -ForegroundColor Cyan
                    Write-Host "    NOTE: Folder dialog will open. Look for it if it appears behind this window." -ForegroundColor Yellow
                    Add-Content $logFile "Using GUI folder picker"
                    
                    Add-Type -AssemblyName System.Windows.Forms
                    $folderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
                    $folderDialog.Description = "Select destination folder for shortcuts"
                    $folderDialog.SelectedPath = [Environment]::GetFolderPath("Desktop")
                    $folderDialog.ShowNewFolderButton = $true
                    
                    if ((Show-DialogInForeground $folderDialog) -eq "OK") {
                        $dest = $folderDialog.SelectedPath
                        Write-Host "[+] Selected folder: $dest" -ForegroundColor Green
                        Add-Content $logFile "Custom destination selected: $dest"
                        $destinationSelected = $true
                    } else {
                        Write-Host ""
                        Write-Host "[!] No folder selected. Please choose again..." -ForegroundColor Yellow
                        Add-Content $logFile "No folder selected in GUI picker, returning to menu"
                        Write-Host ""
                        Clear-Host
                        Write-Host "======================================================" -ForegroundColor Cyan
                        Write-Host "[2] STEP 2: Choose Shortcut Destination" -ForegroundColor Cyan
                        Write-Host "======================================================" -ForegroundColor Cyan
                        Write-Host ""
                    }
                }
                default {
                    Write-Host "[!] Invalid choice. Please try again..." -ForegroundColor Red
                    Clear-Host
                    Write-Host "======================================================" -ForegroundColor Cyan
                    Write-Host "[2] STEP 2: Choose Shortcut Destination" -ForegroundColor Cyan
                    Write-Host "======================================================" -ForegroundColor Cyan
                    Write-Host ""
                }
            }
        }
    } while (-not $destinationSelected)

    Clear-Host
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host "[2] STEP 2: Destination Confirmed" -ForegroundColor Cyan
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "[>] Destination: $dest" -ForegroundColor Green
    Write-Host ""
    
    # Ask if user wants to save this destination as preference (only if not using saved preference)
    if (!$prevDest -or $dest -ne $prevDest) {
        Write-Host "Would you like to save this destination as your default preference?" -ForegroundColor Cyan
        Write-Host ""
        Write-Host "  [Y] Yes - Save as default (recommended)" -ForegroundColor Green
        Write-Host "  [N] No - Use only for this session" -ForegroundColor Yellow
        Write-Host ""
        
        $saveChoice = Read-Host "Save as default? [Y/N] (default: N)"
        if (-not $saveChoice) { $saveChoice = "N" }
        
        if ($saveChoice.ToUpper() -eq "Y") {
            $dest | Set-Content $prefFile
            Write-Host ""
            Write-Host "[+] Destination saved as default preference!" -ForegroundColor Green
            Add-Content $logFile "Destination saved as preference: $dest"
        } else {
            Write-Host ""
            Write-Host "[+] Using destination for this session only" -ForegroundColor Green
            Add-Content $logFile "Destination used for session only: $dest"
        }
    } else {
        Add-Content $logFile "Using existing saved preference: $dest"
    }

    # Create destination directory if it doesn't exist
    if (!(Test-Path $dest)) { 
        New-Item -ItemType Directory -Path $dest | Out-Null 
        Write-Host "[+] Created destination directory" -ForegroundColor Green
        Add-Content $logFile "Created destination directory: $dest"
    } else {
        Add-Content $logFile "Destination directory already exists: $dest"
    }
    Write-Host ""

    # --- Shortcut creation ---
    Clear-Host
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host "[3] STEP 3: Processing Files" -ForegroundColor Cyan
    Write-Host "======================================================" -ForegroundColor Cyan
    Write-Host ""
    
    $totalToProcess = $validPaths.Count + $shortcutFiles.Count
    Write-Host "Processing $totalToProcess target(s)..." -ForegroundColor White
    Write-Host ""

    $wshell = New-Object -ComObject WScript.Shell
    $created = 0
    $copied = 0
    $failed = 0
    $skipped = 0

    Add-Content $logFile "Starting file processing"
    Add-Content $logFile "COM object created: WScript.Shell"
    Add-Content $logFile "Files to create shortcuts for: $($validPaths.Count)"
    Add-Content $logFile "Shortcut files to copy: $($shortcutFiles.Count)"

    # First, create shortcuts for regular files
    if ($validPaths.Count -gt 0) {
        Write-Host "[*] Creating shortcuts for files..." -ForegroundColor Cyan
        Write-Host ""
    }
    
    # Variables for "Apply to All" functionality for shortcuts
    $shortcutApplyToAll = $false
    $shortcutApplyAction = ""
    
    $currentIndex = 0
    foreach ($path in $validPaths) {
        $currentIndex++
        $target = $path.Trim([char]34)
        $name   = [System.IO.Path]::GetFileNameWithoutExtension($target)
        
        # Sanitize shortcut name to remove illegal characters
        $illegalChars = [System.IO.Path]::GetInvalidFileNameChars()
        $sanitizedName = $name
        foreach ($char in $illegalChars) {
            $sanitizedName = $sanitizedName.Replace($char, '_')
        }
        # Also replace question marks which can appear from encoding issues
        $sanitizedName = $sanitizedName.Replace('?', '_')
        
        $link   = "$dest\$sanitizedName.lnk"

        Write-Host "  Creating shortcut: $name..." -NoNewline
        if ($name -ne $sanitizedName) {
            Write-Host ""
            Write-Host "    [i] Name contains special characters, sanitizing..." -ForegroundColor Yellow
            Write-Host "    [i] $name -> $sanitizedName" -ForegroundColor Yellow
            Write-Host "  Creating sanitized shortcut: $sanitizedName..." -NoNewline
        }
        
        Add-Content $logFile "Creating shortcut for target: $target"
        Add-Content $logFile "  - Original name: $name"
        if ($name -ne $sanitizedName) {
            Add-Content $logFile "  - Sanitized name: $sanitizedName"
        }
        Add-Content $logFile "  - Shortcut path: $link"

        try {
            # Verify target exists
            if (!(Test-Path $target)) {
                Write-Host " [X] Target not found" -ForegroundColor Red
                Add-Content $logFile "  - ERROR: Target file does not exist"
                $failed++
                continue
            }

            # Test write permissions before attempting to save
            $testFile = "$dest\test_permissions_$(Get-Random).tmp"
            try {
                [System.IO.File]::WriteAllText($testFile, "test")
                Remove-Item $testFile -Force
            } catch {
                Write-Host " [X] No write permission" -ForegroundColor Red
                Add-Content $logFile "  - ERROR: No write permission to destination folder"
                $failed++
                continue
            }
            
            # Check if shortcut already exists (name-based) or functionally duplicates existing shortcuts
            $shouldSkip = $false
            $duplicateType = ""
            $duplicateName = ""
            
            if (Test-Path $link) {
                $duplicateType = "exact"
                $duplicateName = "$sanitizedName.lnk"
            } else {
                # Check for functional duplicates - shortcuts in destination that point to same target
                $existingShortcuts = Get-ChildItem "$dest\*.lnk" -ErrorAction SilentlyContinue
                foreach ($existingShortcut in $existingShortcuts) {
                    try {
                        $existingWshShortcut = $wshell.CreateShortcut($existingShortcut.FullName)
                        if ($existingWshShortcut.TargetPath -eq $target) {
                            $duplicateType = "functional"
                            $duplicateName = $existingShortcut.Name
                            break
                        }
                    } catch {
                        # Skip if we can't read the existing shortcut
                        continue
                    }
                }
            }
            
            if ($duplicateType -ne "") {
                Write-Host ""  # Complete the initial "Creating shortcut..." line
                if ($duplicateType -eq "exact") {
                    Write-Host "    [!] Shortcut already exists: $duplicateName" -ForegroundColor Yellow
                } else {
                    Write-Host "    [!] Functional duplicate found: $duplicateName" -ForegroundColor Yellow
                    Write-Host "        (Both shortcuts point to the same target: $([System.IO.Path]::GetFileName($target)))" -ForegroundColor Gray
                }
                
                # Check if we should apply a previous "Apply to All" choice
                if ($shortcutApplyToAll) {
                    $duplicateChoice = $shortcutApplyAction
                    Write-Host "    [*] Applying previous choice to all: $duplicateChoice" -ForegroundColor Magenta
                } else {
                    Write-Host "    What would you like to do?" -ForegroundColor Cyan
                    Write-Host "      [O] Overwrite the existing shortcut" -ForegroundColor Green
                    Write-Host "      [S] Skip this shortcut" -ForegroundColor Yellow
                    Write-Host "      [R] Rename with timestamp" -ForegroundColor Cyan
                    Write-Host ""
                    $duplicateChoice = Read-Host "    Choose action [O/S/R] (default: O)"
                    if (-not $duplicateChoice) { $duplicateChoice = "O" }
                    
                    # Ask if they want to apply to all remaining (only if there are multiple items total and remaining)
                    $totalItems = $validPaths.Count + $shortcutFiles.Count
                    $remainingInCurrentSection = $validPaths.Count - $currentIndex
                    $totalRemainingItems = $remainingInCurrentSection + $shortcutFiles.Count
                    if ($totalItems -gt 1 -and $totalRemainingItems -gt 0) {
                        $actionText = switch ($duplicateChoice.ToUpper()) {
                            "O" { "OVERWRITE" }
                            "S" { "SKIP" }
                            "R" { "RENAME" }
                            default { "OVERWRITE" }
                        }
                        Write-Host ""
                        $applyToAllChoice = Read-Host "    Apply '$actionText' to all $totalRemainingItems remaining items? [Y/N] (default: N)"
                        if ($applyToAllChoice.ToUpper() -eq "Y") {
                            $shortcutApplyToAll = $true
                            $shortcutApplyAction = $duplicateChoice.ToUpper()
                            # Also update shared variables for file copying section
                            $applyToAll = $true
                            $applyAction = $duplicateChoice.ToUpper()
                            Write-Host "    [*] Will $actionText all remaining items" -ForegroundColor Magenta
                        }
                    }
                }
                
                switch ($duplicateChoice.ToUpper()) {
                    "O" {
                        if ($duplicateType -eq "exact") {
                            Write-Host "    [*] Overwriting existing shortcut..." -ForegroundColor Green
                            Add-Content $logFile "  - User chose to overwrite existing shortcut: $duplicateName"
                        } else {
                            Write-Host "    [*] Overwriting functional duplicate..." -ForegroundColor Green
                            Add-Content $logFile "  - User chose to overwrite functional duplicate: $duplicateName"
                            # For functional duplicates, use the existing shortcut's path instead of creating new one
                            $link = "$dest\$duplicateName"
                            Write-Host "    [*] Will update existing shortcut: $duplicateName" -ForegroundColor Gray
                            Add-Content $logFile "  - Will update existing shortcut path: $link"
                        }
                    }
                    "S" {
                        if ($duplicateType -eq "exact") {
                            Write-Host "    [*] Skipping shortcut..." -ForegroundColor Yellow
                            Add-Content $logFile "  - User chose to skip existing shortcut: $duplicateName"
                        } else {
                            Write-Host "    [*] Skipping functional duplicate..." -ForegroundColor Yellow
                            Add-Content $logFile "  - User chose to skip functional duplicate: $duplicateName"
                        }
                        Write-Host " [S] Skipped" -ForegroundColor Yellow
                        $shouldSkip = $true
                        $skipped++
                    }
                    "R" {
                        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
                        $link = "$dest\${sanitizedName}_$timestamp.lnk"
                        Write-Host "    [*] Renaming to: $([System.IO.Path]::GetFileName($link))" -ForegroundColor Cyan
                        Add-Content $logFile "  - User chose to rename shortcut, using: $link"
                    }
                    default {
                        Write-Host "    [*] Invalid choice, overwriting existing shortcut..." -ForegroundColor Green
                        Add-Content $logFile "  - Invalid choice, defaulting to overwrite shortcut"
                    }
                }
                
                if ($shouldSkip) {
                    continue
                }
                
                # Show the correct shortcut name after duplicate handling
                $shortcutDisplayName = [System.IO.Path]::GetFileNameWithoutExtension($link)
                Write-Host "  Creating shortcut: $shortcutDisplayName..." -NoNewline
            }
            
            # Create the shortcut with the correct path (after duplicate handling)
            $shortcut = $wshell.CreateShortcut($link)
            $shortcut.TargetPath = $target
            
            # Set working directory to the target's directory if it's a file
            if ((Get-Item $target) -is [System.IO.FileInfo]) {
                $shortcut.WorkingDirectory = [System.IO.Path]::GetDirectoryName($target)
            }
            
            # Only create shortcut if not skipped
            
            $shortcut.Save()
            
            # Brief pause to ensure file system sync
            Start-Sleep -Milliseconds 100
            
            # Verify shortcut was created
            if (Test-Path $link) {
                Write-Host " [+] Created" -ForegroundColor Green
                Add-Content $logFile "  - SUCCESS: Shortcut created successfully"
                $created++
            } else {
                Write-Host " [X] Failed to create" -ForegroundColor Red
                Add-Content $logFile "  - ERROR: Shortcut file not found after creation"
                $failed++
            }
        } catch {
            Write-Host ""
            Write-Host " [X] ERROR DETAILS:" -ForegroundColor Red
            Write-Host "     $($_.Exception.Message)" -ForegroundColor Red
            Write-Host ""
            Write-Host "     Press ENTER to continue or CTRL+C to abort..." -ForegroundColor Yellow
            Read-Host
            Add-Content $logFile "  - ERROR: $($_.Exception.Message)"
            Add-Content $logFile "  - STACK TRACE: $($_.Exception.StackTrace)"
            $failed++
        }
    }

    # Then, copy existing shortcut files
    if ($shortcutFiles.Count -gt 0) {
        if ($validPaths.Count -gt 0) {
            Write-Host ""
        }
        Write-Host "[*] Copying existing shortcut files..." -ForegroundColor Cyan
        Write-Host ""
    }
    
    # Variables for "Apply to All" functionality (shared across both shortcuts and file copying)
    $applyToAll = $shortcutApplyToAll
    $applyAction = $shortcutApplyAction
    
    $currentFileIndex = 0
    foreach ($path in $shortcutFiles) {
        $currentFileIndex++
        $remainingFileCount = $shortcutFiles.Count - $currentFileIndex
        $source = $path.Trim([char]34)
        $fileName = [System.IO.Path]::GetFileName($source)
        
        # Sanitize filename to remove illegal characters
        $illegalChars = [System.IO.Path]::GetInvalidFileNameChars()
        $sanitizedFileName = $fileName
        foreach ($char in $illegalChars) {
            $sanitizedFileName = $sanitizedFileName.Replace($char, '_')
        }
        # Also replace question marks which can appear from encoding issues
        $sanitizedFileName = $sanitizedFileName.Replace('?', '_')
        
        $destination = "$dest\$sanitizedFileName"

        # Check if source and destination are the same
        $sourceFullPath = [System.IO.Path]::GetFullPath($source)
        $destinationFullPath = [System.IO.Path]::GetFullPath($destination)
        if ($sourceFullPath -eq $destinationFullPath) {
            Write-Host ""
            Write-Host "    [i] Source and destination are the same, skipping..." -ForegroundColor Yellow
            Add-Content $logFile "  - SKIPPED: Source and destination paths are identical"
            continue
        }

        Write-Host "  Copying shortcut: $fileName..." -NoNewline
        if ($fileName -ne $sanitizedFileName) {
            Write-Host ""
            Write-Host "    [i] Filename contains special characters, sanitizing..." -ForegroundColor Yellow
            Write-Host "    [i] $fileName -> $sanitizedFileName" -ForegroundColor Yellow
            Write-Host "  Copying sanitized shortcut: $sanitizedFileName..." -NoNewline
        }
        
        Add-Content $logFile "Copying shortcut file: $source"
        Add-Content $logFile "  - Source: $source"
        Add-Content $logFile "  - Original filename: $fileName"
        if ($fileName -ne $sanitizedFileName) {
            Add-Content $logFile "  - Sanitized filename: $sanitizedFileName"
        }
        Add-Content $logFile "  - Destination: $destination"

        try {
            # Verify source exists
            if (!(Test-Path $source)) {
                Write-Host " [X] Source not found" -ForegroundColor Red
                Add-Content $logFile "  - ERROR: Source file does not exist"
                $failed++
                continue
            }

            # Test write permissions before attempting to copy
            $testFile = "$dest\test_copy_permissions_$(Get-Random).tmp"
            try {
                [System.IO.File]::WriteAllText($testFile, "test")
                Remove-Item $testFile -Force
            } catch {
                Write-Host " [X] No write permission" -ForegroundColor Red
                Add-Content $logFile "  - ERROR: No write permission to destination folder"
                $failed++
                continue
            }

            # Check if destination already exists (name-based) or functionally duplicates existing shortcuts
            $shouldSkipFile = $false
            $duplicateType = ""
            $duplicateName = ""
            
            if (Test-Path $destination) {
                $duplicateType = "exact"
                $duplicateName = $sanitizedFileName
            } else {
                # Check for functional duplicates - get target of source shortcut and compare
                try {
                    $sourceShortcut = $wshell.CreateShortcut($source)
                    $sourceTarget = $sourceShortcut.TargetPath
                    
                    # Check existing shortcuts in destination that point to same target
                    $existingShortcuts = Get-ChildItem "$dest\*.lnk" -ErrorAction SilentlyContinue
                    foreach ($existingShortcut in $existingShortcuts) {
                        try {
                            $existingWshShortcut = $wshell.CreateShortcut($existingShortcut.FullName)
                            if ($existingWshShortcut.TargetPath -eq $sourceTarget) {
                                $duplicateType = "functional"
                                $duplicateName = $existingShortcut.Name
                                break
                            }
                        } catch {
                            # Skip if we can't read the existing shortcut
                            continue
                        }
                    }
                } catch {
                    # If we can't read the source shortcut, just proceed with normal filename check
                }
            }
            
            if ($duplicateType -ne "") {
                Write-Host ""
                if ($duplicateType -eq "exact") {
                    Write-Host "    [!] File already exists: $duplicateName" -ForegroundColor Yellow
                } else {
                    Write-Host "    [!] Functional duplicate found: $duplicateName" -ForegroundColor Yellow
                    Write-Host "        (Both shortcuts point to the same target)" -ForegroundColor Gray
                }
                
                # Check if we should apply a previous "Apply to All" choice
                if ($applyToAll) {
                    $duplicateChoice = $applyAction
                    Write-Host "    [*] Applying previous choice to all: $duplicateChoice" -ForegroundColor Magenta
                } else {
                    Write-Host "    What would you like to do?" -ForegroundColor Cyan
                    Write-Host "      [O] Overwrite the existing file" -ForegroundColor Green
                    Write-Host "      [S] Skip this file" -ForegroundColor Yellow
                    Write-Host "      [R] Rename with timestamp" -ForegroundColor Cyan
                    Write-Host ""
                    $duplicateChoice = Read-Host "    Choose action [O/S/R] (default: O)"
                    if (-not $duplicateChoice) { $duplicateChoice = "O" }
                    
                    # Ask if they want to apply to all remaining (only if there are multiple items total and remaining)
                    $totalItems = $validPaths.Count + $shortcutFiles.Count
                    if ($totalItems -gt 1 -and $remainingFileCount -gt 0) {
                        $actionText = switch ($duplicateChoice.ToUpper()) {
                            "O" { "OVERWRITE" }
                            "S" { "SKIP" }
                            "R" { "RENAME" }
                            default { "OVERWRITE" }
                        }
                        Write-Host ""
                        $applyToAllChoice = Read-Host "    Apply '$actionText' to all $remainingFileCount remaining files? [Y/N] (default: N)"
                        if ($applyToAllChoice.ToUpper() -eq "Y") {
                            $applyToAll = $true
                            $applyAction = $duplicateChoice.ToUpper()
                            # Also update shortcut variables in case there are more shortcuts after files
                            $shortcutApplyToAll = $true
                            $shortcutApplyAction = $duplicateChoice.ToUpper()
                            Write-Host "    [*] Will $actionText all remaining files" -ForegroundColor Magenta
                        }
                    }
                }
                
                switch ($duplicateChoice.ToUpper()) {
                    "O" {
                        if ($duplicateType -eq "exact") {
                            Write-Host "    [*] Overwriting existing file..." -ForegroundColor Green
                            Add-Content $logFile "  - User chose to overwrite existing file: $duplicateName"
                        } else {
                            Write-Host "    [*] Overwriting functional duplicate..." -ForegroundColor Green
                            Add-Content $logFile "  - User chose to overwrite functional duplicate: $duplicateName"
                            # For functional duplicates, replace the existing shortcut
                            $destination = "$dest\$duplicateName"
                            Write-Host "    [*] Will replace existing shortcut: $duplicateName" -ForegroundColor Gray
                            Add-Content $logFile "  - Will replace existing shortcut path: $destination"
                        }
                    }
                    "S" {
                        if ($duplicateType -eq "exact") {
                            Write-Host "    [*] Skipping file..." -ForegroundColor Yellow
                            Add-Content $logFile "  - User chose to skip existing file: $duplicateName"
                        } else {
                            Write-Host "    [*] Skipping functional duplicate..." -ForegroundColor Yellow
                            Add-Content $logFile "  - User chose to skip functional duplicate: $duplicateName"
                        }
                        $shouldSkipFile = $true
                        $skipped++
                    }
                    "R" {
                        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
                        $nameWithoutExt = [System.IO.Path]::GetFileNameWithoutExtension($sanitizedFileName)
                        $extension = [System.IO.Path]::GetExtension($sanitizedFileName)
                        $destination = "$dest\${nameWithoutExt}_$timestamp$extension"
                        Write-Host "    [*] Renaming to: $([System.IO.Path]::GetFileName($destination))" -ForegroundColor Cyan
                        Add-Content $logFile "  - User chose to rename, using: $destination"
                    }
                    default {
                        Write-Host "    [*] Invalid choice, overwriting existing file..." -ForegroundColor Green
                        Add-Content $logFile "  - Invalid choice, defaulting to overwrite"
                    }
                }
                
                if ($shouldSkipFile) {
                    continue
                }
                
                Write-Host "  Copying shortcut: $([System.IO.Path]::GetFileName($destination))..." -NoNewline
            }

            # Only copy if not skipped

            Copy-Item -Path $source -Destination $destination -Force
            
            # Verify copy was successful
            if (Test-Path $destination) {
                Write-Host " [+] Copied" -ForegroundColor Green
                Add-Content $logFile "  - SUCCESS: Shortcut copied successfully"
                $copied++
            } else {
                Write-Host " [X] Failed to copy" -ForegroundColor Red
                Add-Content $logFile "  - ERROR: Copied file not found after operation"
                $failed++
            }
        } catch {
            Write-Host ""
            Write-Host " [X] ERROR DETAILS:" -ForegroundColor Red
            Write-Host "     $($_.Exception.Message)" -ForegroundColor Red
            Write-Host ""
            Write-Host "     Press ENTER to continue or CTRL+C to abort..." -ForegroundColor Yellow
            Read-Host
            Add-Content $logFile "  - ERROR: $($_.Exception.Message)"
            Add-Content $logFile "  - STACK TRACE: $($_.Exception.StackTrace)"
            $failed++
        }
    }

    Write-Host ""

    # --- Results Summary ---
    Clear-Host
    Write-Host "======================================================"
    Write-Host "              OPERATION COMPLETE" -ForegroundColor Green
    Write-Host "======================================================"
    Write-Host ""
    Write-Host "  [+] Shortcuts created: $created" -ForegroundColor Green
    if ($copied -gt 0) {
        Write-Host "  [+] Shortcuts copied: $copied" -ForegroundColor Green
    }
    if ($skipped -gt 0) {
        Write-Host "  [-] Skipped: $skipped" -ForegroundColor Yellow
    }
    if ($failed -gt 0) {
        Write-Host "  [X] Failed: $failed" -ForegroundColor Red
    }
    Write-Host "  [#] Total processed: $($created + $copied + $skipped + $failed)"
    Write-Host ""
    Write-Host "  [>] Destination: $dest" -ForegroundColor Cyan
    Write-Host "  [*] Log file: $logFile" -ForegroundColor Gray
    Write-Host ""
    Write-Host "======================================================"
    Write-Host ""

    Add-Content $logFile "OPERATION SUMMARY:"
    Add-Content $logFile "  - Total shortcuts created: $created"
    Add-Content $logFile "  - Total shortcuts copied: $copied"
    Add-Content $logFile "  - Total skipped: $skipped"
    Add-Content $logFile "  - Total failures: $failed"
    Add-Content $logFile "  - Total processed: $($created + $copied + $skipped + $failed)"
    Add-Content $logFile "  - Destination folder: $dest"
    Add-Content $logFile "=== Script completed at $(Get-Date) ==="

    $totalSuccess = $created + $copied
    if ($totalSuccess -gt 0) {
        Write-Host "[*] Success! Your shortcuts are ready to use." -ForegroundColor Green
        if ($created -gt 0 -and $copied -gt 0) {
            Write-Host "    Created $created new shortcuts and copied $copied existing shortcuts." -ForegroundColor Green
        } elseif ($created -gt 0) {
            Write-Host "    Created $created new shortcuts." -ForegroundColor Green
        } else {
            Write-Host "    Copied $copied existing shortcuts." -ForegroundColor Green
        }
        if ($skipped -gt 0) {
            Write-Host "    Skipped $skipped duplicate shortcuts." -ForegroundColor Yellow
        }
    } elseif ($skipped -gt 0 -and $failed -eq 0) {
        Write-Host "[i] All selected shortcuts were skipped (already existed)." -ForegroundColor Yellow
    } elseif ($failed -gt 0) {
        Write-Host "[!] Some issues occurred. Check the log file for details." -ForegroundColor Yellow
        if ($skipped -gt 0) {
            Write-Host "    Also skipped $skipped duplicate shortcuts." -ForegroundColor Yellow
        }
    } else {
        Write-Host "[i] No shortcuts were processed." -ForegroundColor Cyan
    }

    Write-Host ""
    
    } finally {
        # Clean up temporary TemporaryFileSelection.txt file if it exists
        if ($wscPathsFile -and (Test-Path $wscPathsFile)) {
            try {
                Remove-Item $wscPathsFile -Force
                Write-Host "[*] Cleaned up temporary paths file" -ForegroundColor Gray
            } catch {
                Write-Host "[!] Note: Could not clean up temporary paths file" -ForegroundColor Yellow
            }
        }
    }
}

# Main script execution with restart loop
do {
    # Clean up any TemporaryFileSelection.txt file before starting
    $wscCleanupPath = "$((Split-Path -Parent $PSCommandPath))\paths\TemporaryFileSelection.txt"
    if (Test-Path $wscCleanupPath) {
        try {
            Remove-Item $wscCleanupPath -Force
        } catch {
            # Silent cleanup attempt
        }
    }
    
    # Run the main script function
    Start-ShortcutCreator
    
    # Ask user if they want to restart or exit
    Write-Host "======================================================"
    Write-Host "Would you like to run the shortcut creator again?" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "  [Y] Yes - Create more shortcuts (restart)" -ForegroundColor Green
    Write-Host "  [N] No - Exit the program" -ForegroundColor Yellow
    Write-Host ""
    
    $restartChoice = Read-Host "Restart program? [Y/N] (default: Y)"
    if (-not $restartChoice) { $restartChoice = "Y" }
    
    if ($restartChoice.ToUpper() -eq "Y") {
        Write-Host ""
        Write-Host "[*] Restarting WindowsShortcutCreator..." -ForegroundColor Cyan
        Write-Host ""
        Start-Sleep -Seconds 2
        # Continue the loop
    } else {
        Write-Host ""
        Write-Host "[+] Thank you for using WindowsShortcutCreator!" -ForegroundColor Green
        Write-Host ""
        
        # Final cleanup of TemporaryFileSelection.txt
        if ($global:WSCPathsFile -and (Test-Path $global:WSCPathsFile)) {
            try {
                Remove-Item $global:WSCPathsFile -Force
            } catch {
                # Silent cleanup attempt
            }
        }
        
        Write-Host "[*] Exiting program..." -ForegroundColor Gray
        Start-Sleep -Seconds 1
        
        # Signal successful completion to batch file and close window
        [Environment]::Exit(0)
    }
} while ($true)

# Windows Shortcut Creator

**Created by: Cameron Reina**  
**STU TECH**

**Version 3.0** - Create desktop shortcuts quickly and easily

---

## What it does

Creates desktop shortcuts for programs and files. Works on single files or multiple files at once. Can create shortcuts for your desktop or for all users.

## How to use it

1. Run `WindowsShortcutCreator Starter.bat`
2. Choose your files (either pick them or use a text file list)
3. Choose where to put the shortcuts
4. Done

## Requirements

- Windows 10/11
- PowerShell (already installed on Windows)
- Administrator rights (only needed for "all users" shortcuts)

## Installation

1. Download or clone this repository
2. Extract the files
3. Run `WindowsShortcutCreator Starter.bat`

That's it. No setup required.

## Features

- Pick files with a file browser
- Use a text file list for batch processing
- Create shortcuts on your desktop or for all users
- Handles duplicate shortcuts automatically
- Logs all operations
- Works with or without administrator rights

## File locations

```
WindowsShortcutCreator/
|-- WindowsShortcutCreator Starter.bat    # Run this file
|-- bin/
|   |-- WindowsShortcutCreator.ps1        # Main script
|   |-- paths/
|   |   |-- paths.txt                     # Put file paths here (optional)
|   |   +-- shortcut_destination.txt      # Remembers your choice
|   +-- logs/
|       +-- shortcut_log_[timestamp].txt  # Operation logs
+-- README.txt                            # This file
```

## Using a file list (paths.txt)

Create a text file with one file path per line:
```
"C:\Program Files\Notepad++\notepad++.exe"
"C:\Program Files\Chrome\chrome.exe"
"C:\Tools\myprogram.exe"
```

Put quotes around paths with spaces.

## Troubleshooting

**Permission errors**
- Run as Administrator for "all users" shortcuts
- Check that you can write to the destination folder

**File not found errors**
- Make sure file paths are correct
- Use quotes around paths with spaces

**Can't see dialog windows**
- Check your taskbar
- Click on the PowerShell window

**Need more help**
- Check the log files in `bin/logs/`
- Log files show exactly what happened

## When to use Administrator mode

- **Your desktop only**: No administrator needed
- **All users desktop**: Administrator required
- **Custom folder**: Depends on folder permissions

## Common uses

- Set up shortcuts on new computers
- Deploy program shortcuts to multiple users
- Organize desktop shortcuts
- Batch create shortcuts from a list

## Version history

- **v3.0**: Complete rewrite with better duplicate handling
- **v2.0**: Added administrator support and better errors
- **v1.0**: Basic shortcut creation

## Author

Cameron Reina

---

## Tips

- Run as Administrator for company-wide deployments
- Save file lists in paths.txt for repeated use
- Check log files if something goes wrong
- The program cleans up temporary files automatically
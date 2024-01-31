<#
.SYNOPSIS
This script manages Microsoft Teams updates by renaming update.exe and squirrel.exe, and updating the shortcut.

.DESCRIPTION
The script renames the update.exe and squirrel.exe files in the Microsoft Teams directory
and updates the shortcut target path to point to the renamed update.exe.

If at least one of the specified shortcuts exists, it updates the current existing shortcut.
Otherwise, it creates a new shortcut.

.NOTES
File Name      : Manage-classicTeamsUpdates.ps1
Author         : 361b3rn3@github
Prerequisite   : PowerShell
Version        : 4.0
#>

# Function to log messages
function Log-Message {
    param (
        [string]$message,
        [string]$logFile = "C:\Path\To\Log\Manage-classic-Teams-Updates.log"
    )

    # Create the log directory if it doesn't exist
    $logDirectory = Split-Path -Path $logFile
    if (-not (Test-Path $logDirectory -PathType Container)) {
        New-Item -ItemType Directory -Path $logDirectory | Out-Null
    }

    # Log the message to the file
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $logEntry = "$timestamp - $message"
    Add-Content -Path $logFile -Value $logEntry
}

# Function to rename and check the existence of files
function RenameAndCheckFile {
    param (
        [string]$filePath,
        [string]$backupSuffix
    )

    $backupPath = $filePath -replace '\.exe$', ('_' + $backupSuffix + '.exe')

    if (Test-Path $filePath -PathType Leaf) {
        Rename-Item -Path $filePath -NewName $backupPath
        Log-Message "$($filePath.Split('\')[-1]) renamed to $($backupPath.Split('\')[-1])"
    } else {
        Log-Message "$($filePath.Split('\')[-1]) not found. Skipping rename."
    }
}

# Step 1: Rename update.exe
$teamsUpdatePath = "$env:LOCALAPPDATA\Microsoft\Teams\update.exe"
RenameAndCheckFile -filePath $teamsUpdatePath -backupSuffix (Get-Date -Format 'yyyyMMddHHmmss')

# Step 2: Rename squirrel.exe
$teamsSquirrelPath = "$env:LOCALAPPDATA\Microsoft\Teams\current\squirrel.exe"
RenameAndCheckFile -filePath $teamsSquirrelPath -backupSuffix (Get-Date -Format 'yyyyMMddHHmmss')

# Step 3: Update or create shortcut
$shortcutPaths = @(
    "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\Microsoft Teams.lnk",
    "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\Microsoft Teams classic.lnk",
    "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\Microsoft Teams classic (work or school).lnk"
)

$newTargetPath = "$env:LOCALAPPDATA\Microsoft\Teams\current\teams.exe"

# Check if at least one shortcut exists before updating
$shortcutExists = $false
foreach ($shortcutPath in $shortcutPaths) {
    if (Test-Path $shortcutPath -PathType Leaf) {
        $shortcutExists = $true
        break
    }
}

if ($shortcutExists) {
    # Update the current existing shortcut
    try {
        $shortcutPath = $shortcutPaths | Where-Object { Test-Path $_ -PathType Leaf }
        $shortcut = (New-Object -ComObject WScript.Shell).CreateShortcut($shortcutPath)
        $shortcut.TargetPath = $newTargetPath
        $shortcut.Save()
        Log-Message "$($shortcutPath.Split('\')[-1]) updated to point to $($newTargetPath.Split('\')[-1])"
    } catch {
        Log-Message "Error updating $($shortcutPath.Split('\')[-1]): $_"
    }
} else {
    # Create a new shortcut if none exist
    try {
        $shortcut = (New-Object -ComObject WScript.Shell).CreateShortcut($shortcutPaths[0])
        $shortcut.TargetPath = $newTargetPath
        $shortcut.Save()
        Log-Message "$($shortcutPaths[0].Split('\')[-1]) created and set to $($newTargetPath.Split('\')[-1])"
    } catch {
        Log-Message "Error creating $($shortcutPaths[0].Split('\')[-1]): $_"
    }
}

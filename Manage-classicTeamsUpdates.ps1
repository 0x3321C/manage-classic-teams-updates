<#
.SYNOPSIS
This script manages Microsoft Teams updates by renaming update.exe and squirrel.exe, and updating the shortcut.

.DESCRIPTION
The script renames the update.exe and squirrel.exe files in the Microsoft Teams directory
and updates the shortcut target path to point to the renamed update.exe.

If at least one of the specified shortcuts exists and has the correct target path, it updates the current existing shortcut.
Otherwise, it creates a new shortcut.

.NOTES
File Name      : Manage-classicTeamsUpdates.ps1
Author         : 361b3rn3@github
Prerequisite   : PowerShell
Version        : 4.1
#>

# Function to log messages
function Write-log {
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
function Rename-File {
    param (
        [string]$filePath,
        [string]$backupSuffix
    )

    $backupPath = $filePath -replace '\.exe$', ('_' + $backupSuffix + '.exe')

    if (Test-Path $filePath -PathType Leaf) {
        Rename-Item -Path $filePath -NewName $backupPath
        Write-log "$($filePath.Split('\')[-1]) renamed to $($backupPath.Split('\')[-1])"
    } else {
        Write-log "$($filePath.Split('\')[-1]) not found. Skipping rename."
    }
}

# Function to check if the target path of a shortcut is correct
function Test-TargetPath {
    param (
        [string]$shortcutPath,
        [string]$correctTargetPath
    )

    $shortcut = (New-Object -ComObject WScript.Shell).CreateShortcut($shortcutPath)
    return ($shortcut.TargetPath -eq $correctTargetPath)
}

# Step 1: Rename update.exe
$teamsUpdatePath = "$env:LOCALAPPDATA\Microsoft\Teams\update.exe"
Rename-File -filePath $teamsUpdatePath -backupSuffix (Get-Date -Format 'yyyyMMddHHmmss')

# Step 2: Rename squirrel.exe
$teamsSquirrelPath = "$env:LOCALAPPDATA\Microsoft\Teams\current\squirrel.exe"
Rename-File -filePath $teamsSquirrelPath -backupSuffix (Get-Date -Format 'yyyyMMddHHmmss')

# Step 3: Update or create shortcut
$shortcutPaths = @(
    "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\Microsoft Teams.lnk",
    "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\Microsoft Teams classic.lnk",
    "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\Microsoft Teams classic (work or school).lnk"
)

$newTargetPath = "$env:LOCALAPPDATA\Microsoft\Teams\current\teams.exe"

# Check if at least one correct shortcut exists before updating
$correctShortcutExists = $false
foreach ($shortcutPath in $shortcutPaths) {
    if (Test-Path $shortcutPath -PathType Leaf -and (Test-TargetPath -shortcutPath $shortcutPath -correctTargetPath $newTargetPath)) {
        $correctShortcutExists = $true
        break
    }
}

if ($correctShortcutExists) {
    # Update the current existing correct shortcut
    try {
        $shortcutPath = $shortcutPaths | Where-Object { Test-Path $_ -PathType Leaf -and (Test-TargetPath -shortcutPath $_ -correctTargetPath $newTargetPath) }
        $shortcut = (New-Object -ComObject WScript.Shell).CreateShortcut($shortcutPath)
        $shortcut.TargetPath = $newTargetPath
        $shortcut.Save()
        Write-log "$($shortcutPath.Split('\')[-1]) updated to point to $($newTargetPath.Split('\')[-1])"
    } catch {
        Write-log "Error updating $($shortcutPath.Split('\')[-1]): $_"
    }
} else {
    # Create a new shortcut if no correct shortcut exists
    try {
        $shortcut = (New-Object -ComObject WScript.Shell).CreateShortcut($shortcutPaths[0])
        $shortcut.TargetPath = $newTargetPath
        $shortcut.Save()
        Write-log "$($shortcutPaths[0].Split('\')[-1]) created and set to $($newTargetPath.Split('\')[-1])"
    } catch {
        Write-log "Error creating $($shortcutPaths[0].Split('\')[-1]): $_"
    }
}

<#
.SYNOPSIS
This script manages Microsoft Teams updates by renaming update.exe and squirrel.exe, and updating the shortcut.

.DESCRIPTION
The script renames the update.exe and squirrel.exe files in the Microsoft Teams directory
and updates the shortcut target path to point to the renamed update.exe.

.NOTES
File Name      : Manage-TeamsUpdates.ps1
Author         : Your Name
Prerequisite   : PowerShell
version        : 1.0

#>

# Step 1: Rename update.exe
$teamsUpdatePath = "$env:LOCALAPPDATA\Microsoft\Teams\update.exe"
$backupUpdatePath = "$env:LOCALAPPDATA\Microsoft\Teams\update.exe_backup"

# Check if update.exe exists before renaming
if (Test-Path $teamsUpdatePath -PathType Leaf) {
    Rename-Item -Path $teamsUpdatePath -NewName $backupUpdatePath
    Write-Host "update.exe renamed to update.exe_backup"
} else {
    Write-Host "update.exe not found. Skipping rename."
}

# Step 2: Rename squirrel.exe
$teamsSquirrelPath = "$env:LOCALAPPDATA\Microsoft\Teams\current\squirrel.exe"
$backupSquirrelPath = "$env:LOCALAPPDATA\Microsoft\Teams\current\squirrel.exe_backup"

# Check if squirrel.exe exists before renaming
if (Test-Path $teamsSquirrelPath -PathType Leaf) {
    Rename-Item -Path $teamsSquirrelPath -NewName $backupSquirrelPath
    Write-Host "squirrel.exe renamed to squirrel.exe_backup"
} else {
    Write-Host "squirrel.exe not found. Skipping rename."
}

# Step 3: Update shortcut target path
$shortcutLegacyPath = "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\Microsoft Teams.lnk"
$shortcutNewPath = "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\Microsoft Teams classic.lnk"
$shortcutWorkPath = "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\Microsoft Teams classic (work or school).lnk"
$newTargetPath = "$env:LOCALAPPDATA\Microsoft\Teams\current\teams.exe"

$shortcutPaths = @()
$shortcutPaths = $shortcutLegacyPath, $shortcutNewPath, $shortcutWorkPath

# Check if the shortcut exists before updating
foreach($shortcutPath in $shortcutPaths){

if (Test-Path $shortcutPath -PathType Leaf) {
    $shortcut = (New-Object -ComObject WScript.Shell).CreateShortcut($shortcutPath)
    $shortcut.TargetPath = $newTargetPath
    $shortcut.Save()
    Write-Host "Shortcut updated to point to teams.exe"
} else {
    Write-Host "Shortcut not found. Skipping update."
}
}

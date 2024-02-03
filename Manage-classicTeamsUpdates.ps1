<#
.SYNOPSIS
This script manages Microsoft Teams updates by renaming update.exe and squirrel.exe, updating the shortcut, and removing a specific registry key value.

.DESCRIPTION
The script renames the update.exe and squirrel.exe files in the Microsoft Teams directory,
updates the shortcut target path to point to the renamed update.exe, and removes a specific registry key value.

If a shortcut with the correct target path already exists, it skips both the update and creation steps.

.NOTES
File Name      : Manage-classicTeamsUpdates.ps1
Author         : 0x3M321C@github
Prerequisite   : PowerShell
Version        : 5.1.0
#>

# Function to log messages
function Write-Log {
    param (
        [string]$message,
        [string]$logFile = "C:\temp\log\Manage-classic-Teams-Updates.log"
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

    try {
        if (Test-Path $filePath -PathType Leaf) {
            Rename-Item -Path $filePath -NewName $backupPath -ErrorAction Stop
            Write-Log "$($filePath.Split('\')[-1]) renamed to $($backupPath.Split('\')[-1])"
        } else {
            Write-Log "$($filePath.Split('\')[-1]) not found. Skipping rename."
        }
    }
    catch {
        Write-Log "Error renaming $($filePath.Split('\')[-1]): $_"
    }
}

# Function to check if the target path of a shortcut is correct
function Test-TargetPath {
    param (
        [string]$shortcutPath,
        [string]$correctTargetPath
    )

    try {
        $shortcut = (New-Object -ComObject WScript.Shell).CreateShortcut($shortcutPath)

        if (Test-Path $shortcutPath -PathType Leaf) {
            Write-Log "$($shortcutPath.Split('\')[-1]) exists."

            $targetPath = $shortcut.TargetPath
            return $targetPath -eq $correctTargetPath
        }
        else {
            Write-Log "$($shortcutPath.Split('\')[-1]) does not exist."
            $targetPath = ""
            return $targetPath -eq $correctTargetPath
        }
    }
    catch {
        Write-Log "Error accessing shortcut properties: $_"
        return $false
    }
    finally {
        if ($shortcut -ne $null) {
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($shortcut) | Out-Null
            Remove-Variable -Name shortcut -Force
        }
    }
}

# Function to remove a specific registry key value
function Remove-RegistryValue {
    param (
        [string]$keyPath,
        [string]$valueName
    )

    try {
        if (Test-Path $keyPath) {
            Remove-ItemProperty -Path $keyPath -Name $valueName -ErrorAction Stop
            Write-Log "Registry key value '$valueName' removed from '$keyPath'"
        } else {
            Write-Log "Registry key '$keyPath' not found. Skipping removal of value '$valueName'."
        }
    }
    catch {
        Write-Log "Error removing registry key value '$valueName' from '$keyPath': $_"
    }
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

# Check if a shortcut with the correct path already exists
$correctShortcutExists = $false

foreach ($shortcutPath in $shortcutPaths) {
    $correctShortcutExists = Test-TargetPath -shortcutPath $shortcutPath -correctTargetPath $newTargetPath
    if ($correctShortcutExists) {
        Write-Log "The shortcut, $($shortcutPath.Split('\')[-1]) already has a correct target path. Skipping update or creation."
        Break
    }
}

if (-not $correctShortcutExists) {
    # Update or create a new shortcut
    try {
        # Attempt to update the current existing shortcut or create a new shortcut
        $existingShortcut = $false

        foreach ($shortcutPath in $shortcutPaths) {
            # Search for an existing shortcut
            $existingShortcut = Test-Path $shortcutPath -PathType Leaf
            if ($existingShortcut) {
                Write-Log "The shortcut, $($shortcutPath.Split('\')[-1]) should be updated."

                # Update the current existing shortcut
                $shortcut = (New-Object -ComObject WScript.Shell).CreateShortcut($shortcutPath)
                $shortcut.Arguments = ""
                $shortcut.TargetPath = $newTargetPath
                $shortcut.Save()
                Write-Log "$($shortcutPath.Split('\')[-1]) updated to point to $($newTargetPath.Split('\')[-1])"
                Break
            }
        }

        if (-not $existingShortcut) {   
            # Create a new shortcut if no correct shortcut exists
            Write-Log "No existing shortcut with the target path $($newTargetPath.Split('\')[-1])"

            $shortcut = (New-Object -ComObject WScript.Shell).CreateShortcut($shortcutPaths[0])
            $shortcut.TargetPath = $newTargetPath
            $shortcut.Save()
            Write-Log "$($shortcutPaths[0].Split('\')[-1]) created and set to $($newTargetPath.Split('\')[-1])"
        }           
    }   
    catch {
        Write-Log "Error updating or creating shortcut: $_"
    }
    finally {
        if ($shortcut -ne $null) {
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($shortcut) | Out-Null
            Remove-Variable -Name shortcut -Force
        }
    }
}

# Step 4: Remove the specified registry key value
$registryKeyPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run"
$registryValueName = "com.squirrel.Teams.Teams"

Remove-RegistryValue -keyPath $registryKeyPath -valueName $registryValueName

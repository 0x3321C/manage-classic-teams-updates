<#
.SYNOPSIS
This script manages Microsoft Teams updates by renaming update.exe and squirrel.exe, and updating the shortcut.

.DESCRIPTION
The script renames the update.exe and squirrel.exe files in the Microsoft Teams directory
and updates the shortcut target path to point to the renamed update.exe.

If a shortcut with the correct target path already exists, it skips both the update and creation steps.

.NOTES
File Name      : Manage-classicTeamsUpdates.ps1
Author         : 0x3M321C@github
Prerequisite   : PowerShell
Version        : 5.0.0
#>

# Function to log messages
function Write-log {
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

    if (Test-Path $filePath -PathType Leaf) {
        Rename-Item -Path $filePath -NewName $backupPath
        Write-log "$($filePath.Split('\')[-1]) renamed to $($backupPath.Split('\')[-1])"
    }
    else {
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

    $targetPath=$shortcut.TargetPath 

    if (Test-Path $shortcutPath -PathType Leaf) {
       
        Write-log "$($shortcutPath.Split('\')[-1]) exists."

         return $targetPath -eq $correctTargetPath
     
    }
    else {

        Write-log "$($shortcutPath.Split('\')[-1]) does not exists"
        $targetPath = ""
         return $targetPath -eq $correctTargetPath
    }  

    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($shortcut)
    Remove-Variable shortcut
    #return $targetPath -eq $correctTargetPath
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
    if ($correctShortcutExists ) {
        Write-log "The shortcut, $($shortcutPath.Split('\')[-1]) already has a correct target path. Skipping update or creation."
        Break
    }
}



if (-not $correctShortcutExists) {
    # Update or create a new shortcut

    try {
        # Attempt to update the current existing shortcut or create a new shortcut
        $existingShortcut = $false

        foreach ($shortcutPath in $shortcutPaths) {
            # Search for existing shortcut
            $existingShortcut = Test-Path $shortcutPath -PathType Leaf
            if ($existingShortcut) {
             
                Write-log "The shortcut, $($shortcutPath.Split('\')[-1]) should be updated."

                # Update the current existing shortcut
                $shortcut = (New-Object -ComObject WScript.Shell).CreateShortcut($shortcutPath)
                $shortcut.Arguments = ""
                $shortcut.TargetPath = $newTargetPath
                $shortcut.Save()
                Write-log "$($shortcutPath.Split('\')[-1]) updated to point to $($newTargetPath.Split('\')[-1])"
                
                
                Break
            }
        }

        if (-not $existingShortcut) {   
            # Create a new shortcut if no correct shortcut exists
            Write-log "No existing shorcut with the target path $($newTargetPath.Split('\')[-1])"

            $shortcut = (New-Object -ComObject WScript.Shell).CreateShortcut($shortcutPaths[0])
            $shortcut.TargetPath = $newTargetPath
            $shortcut.Save()
            Write-log "$($shortcutPaths[0].Split('\')[-1]) created and set to $($newTargetPath.Split('\')[-1])"
        }           
        
    }   
    catch {
        Write-log "Error updating or creating shortcut, $_ "
    }
    finally {
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($shortcut)
        Remove-Variable shortcut
    }
}


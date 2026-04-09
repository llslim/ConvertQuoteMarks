# Copyright (c) LL Slim LLC

# Define paths
$InstallDir = Join-Path $env:LOCALAPPDATA "ConvertQuoteMarks"
$SettingsDir = Join-Path $env:USERPROFILE ".ConvertQuoteMarks"

$DesktopPath = [Environment]::GetFolderPath("Desktop")
$ShortcutFile = Join-Path $DesktopPath "ConvertQuoteMarks.lnk"

$StartMenuPath = [Environment]::GetFolderPath("Programs")
$StartMenuShortcutFile = Join-Path $StartMenuPath "ConvertQuoteMarks.lnk"

# 1. Remove Shortcuts
if (Test-Path $ShortcutFile) {
    Remove-Item -Path $ShortcutFile -Force
    Write-Host "Removed Desktop Shortcut."
}

if (Test-Path $StartMenuShortcutFile) {
    Remove-Item -Path $StartMenuShortcutFile -Force
    Write-Host "Removed Start Menu Shortcut."
}

# 2. Remove Application Files
if (Test-Path $InstallDir) {
    Remove-Item -Path $InstallDir -Recurse -Force
    Write-Host "Removed Application Directory: $InstallDir"
}

# 3. Remove User Settings
if (Test-Path $SettingsDir) {
    Remove-Item -Path $SettingsDir -Recurse -Force
    Write-Host "Removed Settings Directory: $SettingsDir"
}

Write-Host "Uninstallation Complete!"
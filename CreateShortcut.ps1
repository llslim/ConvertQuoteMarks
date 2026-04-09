$ScriptPath = Join-Path $PSScriptRoot "ConvertQuoteMarks_WPF.ps1"
$DesktopPath = [Environment]::GetFolderPath("Desktop")
$ShortcutFile = Join-Path $DesktopPath "ConvertQuoteMarks.lnk"

$WshShell = New-Object -ComObject WScript.Shell
$Shortcut = $WshShell.CreateShortcut($ShortcutFile)

$Shortcut.TargetPath = "powershell.exe"
# -WindowStyle Hidden ensures the black console window doesn't stay open
$Shortcut.Arguments = "-WindowStyle Hidden -ExecutionPolicy Bypass -File `"$ScriptPath`""
$Shortcut.Description = "ConvertQuoteMarks Converter"
$Shortcut.Save()

Write-Host "Shortcut created successfully: $ShortcutFile"
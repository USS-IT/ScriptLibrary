@echo off
REM Switch to Ivanti VPN AlwaysOn as standard user
REM Author: Matt Carras (mcarras8)
REM Created: 6-26-2026
powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "Write-EventLog -LogName 'USS-EventLog' -Source 'SwitchVPNProfile-AlwaysOn' -EventID 1000 -EntryType Information -Message 'Trigger VPN switch to AlwaysOn'"
powershell.exe -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File "%~dp0Show-Toast.ps1" -Text "Attempting to switch to VPN AlwaysOn profile. This may take a minute or two." -Title "VPN Switch" -LauncherID "%~f0"
@echo off
REM Switch to Ivanti VPN General as standard user
REM Requires installing:
REM - VPN Utilities
REM - Install-Tasks-SwitchVPNProfile
REM Author: Matt Carras (mcarras8)
REM Created: 6-26-2026

powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "Write-EventLog -LogName 'USS-EventLog' -Source 'SwitchVPNProfile-AlwaysOn' -EventID 1000 -EntryType Information -Message 'Trigger VPN switch to AlwaysOn'"

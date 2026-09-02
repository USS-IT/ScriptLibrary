@echo off
SET "_LOGFILE=C:\USS\Logs\User\NetworkUtilities\%~nx0.log"
(
	powershell.exe -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File "%~dp0Show-Toast.ps1" -Text "Adding Public DNS servers to online adapters. This may take a minute." -Title "Add DNS" -LauncherID "%~f0"
	powershell.exe -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File "%~dp0Install-Tasks-User.ps1"
	powershell.exe -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -Command "Write-EventLog -LogName 'USS-EventLog' -Source 'SwitchDNS-Add' -EventID 1000 -EntryType Information -Message 'Trigger Switch-DNS -Mode Add'"
) > "%_LOGFILE%" 2>&1
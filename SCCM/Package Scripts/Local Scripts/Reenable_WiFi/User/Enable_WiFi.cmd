@echo off
SET "_LOGFILE=C:\USS\Logs\User\Reenable_WiFi\%~nx0.log"
(
	REM Show toast notification	
	powershell.exe -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File "%~dp0Show-Toast.ps1" -Text "Attempting to re-enable the WiFi adapter. This may take a minute or two." -Title "WiFi re-enable" -LauncherID "%~f0"
	powershell.exe -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File "%~dp0Register-Tasks-NotifyReenableWiFi.ps1"
	powershell.exe -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -Command "Write-EventLog -LogName 'USS-EventLog' -Source 'Enable-WiFi' -EventID 1000 -EntryType Information -Message 'Trigger Re-enable WiFi Adapters'"
) > "%_LOGFILE%" 2>&1
@echo off
SET "_LOGFILE=C:\USS\Logs\User\Reenable_WiFi\%~nx0.log"
(
	powershell.exe -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File "%~dp0Register-Tasks-NotifyReenableWiFi.ps1"
	powershell.exe -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -Command "Write-EventLog -LogName 'USS-EventLog' -Source 'Enable-WiFi' -EventID 1000 -EntryType Information -Message 'Trigger Re-enable WiFi Adapters'"
) > "%_LOGFILE%" 2>&1
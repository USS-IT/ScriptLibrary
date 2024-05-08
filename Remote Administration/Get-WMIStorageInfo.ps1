# Little snippets for getting real-time software info remotely queried through WMI. System must be on the Hopkins network.
# Account used must have either local admin or remote admin rights on system.
# Can sometimes fail with "RPC server not available" error depending on system configuration or if the computer is offline.

# Queries disk info.
$comp = Read-Host "Enter Computer Name"
# Check if computer is online (try up to three times).
If((Test-Connection -ComputerName $comp -Count 1 -Quiet) -Or (Test-Connection -ComputerName $comp -Count 1 -Quiet) -Or (Test-Connection -ComputerName $comp -Count 1 -Quiet)) {
	Write-Host "[$comp] appears to be online, querying..."
} else {
	Write-Warning "[$comp] did not respond to any ping attempts, querying anyway..."
}
$owmi = Get-WmiObject -Class Win32_LogicalDisk -ComputerName $comp | ? {$_. DriveType -eq 3} | select DeviceID, {$_.Size /1GB}, {$_.FreeSpace /1GB}, VolumeName
if($owmi) {
	$owmi | Format-Table | Out-String | Write-Host
	$owmi2 = Get-WmiObject -Query "Select * from Win32_diskdrive" -ComputerName $comp
	$owmi2 | Select ($owmi2.Properties | foreach {$_.Name}) | Format-List | Out-String | Write-Host
}
Read-Host "Press enter to exit" | Out-Null

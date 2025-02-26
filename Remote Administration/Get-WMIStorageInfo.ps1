<#
	.SYNOPSIS
	Uses WMI to query an online computer's disk info (free space, type, etc.)
	
	.DESCRIPTION
	Uses WMI to query an online computer's disk info (free space, type, etc.)
	
	.NOTES
	Computer must be on the Hopkins network.
	Must be run from an account with local admin or Remote WMI Admin privileges on machine.
	Firewall must be set to allow WMI queries.
	
	Author: mcarras8
#>
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

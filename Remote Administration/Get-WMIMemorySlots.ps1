<#
	.SYNOPSIS
	Uses WMI to query info on an online computer's installed and free memory slots.
	
	.DESCRIPTION
	Uses WMI to query info on an online computer's installed and free memory slots.
	
	.NOTES
	Adapted from: http://www.powershellpro.com/dimm-witt/200/
	
	Computer must be on the Hopkins network.
	Must be run from an account with local admin or Remote WMI Admin privileges on remote machine.
	Firewall must be set to allow WMI queries.

	Created: 9-19-23
	Author: mcarras8
#>
$comp = Read-Host "Enter Computer Name"
If((Test-Connection -ComputerName $comp -Count 1 -Quiet) -Or (Test-Connection -ComputerName $comp -Count 1 -Quiet) -Or (Test-Connection -ComputerName $comp -Count 1 -Quiet)) {
	Write-Host "[$comp] appears to be online, querying..."
} else {
	Write-Warning "[$comp] did not respond to any ping attempts, querying anyway..."
}
Get-WmiObject -Class "win32_PhysicalMemoryArray" -namespace "root\CIMV2" -computerName $comp | % {
	"Total Number of DIMM Slots: " + $_.MemoryDevices
}
Get-WmiObject -Class "win32_PhysicalMemory" -namespace "root\CIMV2" -computerName $comp | % {
     "Memory Installed: " + $_.DeviceLocator
     "Memory Size: " + ($_.Capacity / 1GB) + " GB"
}
Read-Host "Press enter to exit" | Out-Null

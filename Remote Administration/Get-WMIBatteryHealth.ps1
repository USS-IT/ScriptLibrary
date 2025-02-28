<#
	.SYNOPSIS
	Uses WMI to query info on an online computer's battery health and current capacity.
	
	.DESCRIPTION
	Uses WMI to query info on an online computer's battery health and current capacity.
	
	.NOTES
	Computer must be on the Hopkins network.
	Must be run from an account with local admin or Remote WMI Admin privileges on remote machine.
	Firewall must be set to allow WMI queries.

	Author: mcarras8
#>
$comp = Read-Host "Enter Computer Name"
If((Test-Connection -ComputerName $comp -Count 1 -Quiet) -Or (Test-Connection -ComputerName $comp -Count 1 -Quiet) -Or (Test-Connection -ComputerName $comp -Count 1 -Quiet)) {
	Write-Host "[$comp] appears to be online, querying..."
} else {
	Write-Warning "[$comp] did not respond to any ping attempts, querying anyway..."
}
$current = (Get-WmiObject -Class "BatteryFullChargedCapacity" -Namespace "ROOT\WMI" -ComputerName $comp).FullChargedCapacity
$design = (Get-WmiObject -Class "BatteryStaticData" -Namespace "ROOT\WMI" -ComputerName $comp).DesignedCapacity
"Current Capacity: ${current}mW, Design Capacity: ${design}mW, Overall Health: {0}%" -f [math]::Round($current / $design * 100, 2)
Read-Host "Press enter to exit" | Out-Null

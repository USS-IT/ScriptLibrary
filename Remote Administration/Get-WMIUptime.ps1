<#
	.SYNOPSIS
	Uses WMI to query an online computer's uptime.
	
	.DESCRIPTION
	Uses WMI to query an online computer's uptime. Note Power > Shutdown will not reset this value.
	
	.NOTES
	Computer must be on the Hopkins network.
	Must be run from an account with local admin or Remote WMI Admin privileges on machine.
	Firewall must be set to allow WMI queries.
	
	Author: mcarras8
#>
$comp = Read-Host "Enter Computer Name"
If((Test-Connection -ComputerName $comp -Count 1 -Quiet) -Or (Test-Connection -ComputerName $comp -Count 1 -Quiet) -Or (Test-Connection -ComputerName $comp -Count 1 -Quiet)) {
	Write-Host "[$comp] appears to be online, querying..."
} else {
	Write-Warning "[$comp] did not respond to any ping attempts, querying anyway..."
}
"Last Restart for [${comp}]: " + [Management.ManagementDateTimeConverter]::ToDateTime((Get-WmiObject Win32_OperatingSystem -ComputerName $comp).LastBootUpTime)
Read-Host "Press enter to exit" | Out-Null

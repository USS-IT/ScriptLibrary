<#
	.SYNOPSIS
	Uses WMI to query an online computer's operating system info. If the system is offline, checks AD instead (requires RSAT tools).
	
	.DESCRIPTION
	Uses WMI to query an online computer's operating system info. If the system is offline, checks AD instead (requires RSAT tools).
	
	.NOTES
	Computer must be on the Hopkins network.
	Must be run from an account with local admin or Remote WMI Admin privileges on remote machine.
	Firewall must be set to allow WMI queries.
	
	Author: mcarras8
#>

$comp = Read-Host "Enter Computer Name"
Write-Host "Checking if [$comp] is online, please wait..."
# Check if computer is online (try up to three times).
If((Test-Connection -ComputerName $comp -Count 1 -Quiet) -Or (Test-Connection -ComputerName $comp -Count 1 -Quiet) -Or (Test-Connection -ComputerName $comp -Count 1 -Quiet)) {
	Write-Host "[$comp] appears to be online, querying WMI..."
	Get-WmiObject -Class Win32_OperatingSystem -ComputerName $comp | Select PSComputerName, Caption, OSArchitecture, Version, BuildNumber
} else {
	Write-Warning "[$comp] did not respond to any ping attempts, querying AD instead"
	Get-ADComputer $comp -Properties OperatingSystemVersion
}
Read-Host "Press enter to exit" | Out-Null

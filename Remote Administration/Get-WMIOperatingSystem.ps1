<#
	.SYNOPSIS
	Uses WMI to query an online computer's operating system info.
	
	.DESCRIPTION
	Uses WMI to query an online computer's operating system info.
	
	.NOTES
	EOLVER must be updated periodically in this script.
	
	Computer must be on the Hopkins network.
	Must be run from an account with local admin or Remote WMI Admin privileges on remote machine.
	Firewall must be set to allow WMI queries.
	
	Author: mcarras8
#>
# Windows 11, 22H2
$EOLVER=22631

$comp = Read-Host "Enter Computer Name"
# Check if computer is online (try up to three times).
If((Test-Connection -ComputerName $comp -Count 1 -Quiet) -Or (Test-Connection -ComputerName $comp -Count 1 -Quiet) -Or (Test-Connection -ComputerName $comp -Count 1 -Quiet)) {
	Write-Host "[$comp] appears to be online, querying..."
} else {
	Write-Warning "[$comp] did not respond to any ping attempts, querying anyway..."
}
Get-WmiObject -Class Win32_OperatingSystem -ComputerName $comp | Select PSComputerName, Caption, OSArchitecture, Version, BuildNumber, @{N="EOL Version"; Expression={ $EOLVER }}, @{N="Windows End of Life"; Expression={ $_.BuildNumber -le $EOLVER }}
Read-Host "Press enter to exit" | Out-Null

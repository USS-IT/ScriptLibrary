<#
	.SYNOPSIS
	Uses WMI to query an online computer's installed software.
	
	.DESCRIPTION
	Uses WMI to query an online computer's installed software. This should match Add/Remove Programs.
	
	.NOTES
	Computer must be on the Hopkins network.
	Must be run from an account with local admin or Remote WMI Admin privileges on remote machine.
	Firewall must be set to allow WMI queries.
	
	Author: mcarras8
#>

# Example querying all software.
$comp = Read-Host "Enter Computer Name"
# Check if computer is online (try up to three times).
If((Test-Connection -ComputerName $comp -Count 1 -Quiet) -Or (Test-Connection -ComputerName $comp -Count 1 -Quiet) -Or (Test-Connection -ComputerName $comp -Count 1 -Quiet)) {
	Write-Host "[$comp] appears to be online, querying..."
} else {
	Write-Warning "[$comp] did not respond to any ping attempts, querying anyway..."
}
Get-WmiObject -Class Win32_Product -ComputerName $comp | Select Name, Version, Vendor, InstallDate, InstallLocation | Sort Name | Out-GridView 
Read-Host "Press enter to exit" | Out-Null

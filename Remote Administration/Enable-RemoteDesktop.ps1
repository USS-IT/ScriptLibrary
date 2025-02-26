<#
	.SYNOPSIS
	Remotely enables Remote Desktop using WMI.
	
	.DESCRIPTION
	Remotely enables Remote Desktop for the target computer using WMI.
	
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
$RDP = Get-WmiObject -Class Win32_TerminalServiceSetting -Namespace root\CIMV2\TerminalServices -Authentication 6 -ComputerName $comp
$RDP.AllowTSConnections
$RDP.SetAllowTsConnections(1,1)
Read-Host "Press enter to exit" | Out-Null

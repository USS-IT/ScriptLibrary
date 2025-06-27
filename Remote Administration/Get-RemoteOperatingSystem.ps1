<#
	.SYNOPSIS
	Queries a system's operating system version using both WMI (if it's online) and AD. Requires RSAT tools.
	
	.DESCRIPTION
	Queries a system's operating system version using both WMI (if it's online) and AD. Requires RSAT tools.
	
	.NOTES
	Computer must be on the Hopkins network.
	Must be run from an account with local admin or Remote WMI Admin privileges on remote machine.
	Firewall must be set to allow WMI queries.
	
	Author: mcarras8
#>

$comp = Read-Host "Enter Computer Name"

Write-Host "Getting computer info from AD..."
$adComp = Get-ADComputer $comp -Properties OperatingSystemVersion

$wmiOS = $null
$wmiDisk = $null
$lastBootUp = $null
Write-Host "Checking if [$comp] is online, please wait..."
If((Test-Connection -ComputerName $comp -Count 1 -Quiet) -Or (Test-Connection -ComputerName $comp -Count 1 -Quiet) -Or (Test-Connection -ComputerName $comp -Count 1 -Quiet)) {
	Write-Host "[$comp] appears to be online, querying WMI..."
	$wmiOS = Get-WmiObject -Class Win32_OperatingSystem -ComputerName $comp
	if ($wmiOS) {
		$lastBootUp = [Management.ManagementDateTimeConverter]::ToDateTime($wmiOS.LastBootUpTime)
	}
	$wmiDisk = Get-WmiObject -Class Win32_LogicalDisk -ComputerName $comp | where {$_.DriveType -eq 3 -and $_.DeviceID -eq "C:"} | Select @{N="Size"; Expression={[Math]::Round($_.Size /1GB, 2)}}, @{N="FreeSpace"; Expression={[Math]::Round($_.FreeSpace /1GB, 2)}}
} else {
	Write-Warning "[$comp] did not respond to any ping attempts"
}
$o = [ordered]@{
	"AD Name"=$adComp.Name
	"AD Enabled"=$adComp.Enabled
	"AD OS Version"=$adComp.OperatingSystemVersion
}
if ($wmiOS) {
	$o["WMI OS Version*"]=$wmiOS.Version
	$o["WMI Last Restart"]=$lastBootUp
} else {
	$o["WMI OS Version*"]="<Unable to query WMI>"
	$o["WMI Last Restart"]="<Unable to query WMI>"
}
if ($wmiDisk) {
	$o["WMI Free Space"]="$($wmiDisk.FreeSpace)GB free ($($wmiDisk.Size)GB total)"
} else {
	$o["WMI Free Space"]="<Unable to query WMI>"
}
[PSCustomObject]$o
Write-Host "* WMI OS Version will be most current. Only accessible if the system is online."
Read-Host "Press enter to exit" | Out-Null

<#
	.SYNOPSIS
	Get OperatingSystemVersion reported by AD for given computer name.
	
	.DESCRIPTION
	Get OperatingSystemVersion reported by AD for given computer name. 
	
	.NOTES
	Requires RSAT tools.
	
	Created: 5-16-24
	Author: mcarras8
#>

$comp = Read-Host "Enter Computer Name"
Get-ADComputer $comp -Properties OperatingSystemVersion
Read-Host "Press enter to exit" | Out-Null

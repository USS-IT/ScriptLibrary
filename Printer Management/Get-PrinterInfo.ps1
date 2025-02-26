<#
    .SYNOPSIS
    Gets installed printer info for an online computer.
    
	.DESCRIPTION
    Gets installed printer info for an online computer. Attempts to resolve WSD to IP addresses.
    
    .NOTES
    Author: mcarras8
#>
$comp = Read-Host "Enter Computer Name"
Write-Host "** Getting printer info from [$comp], please wait...if this takes too long the computer may be offline"
Get-Printer -ComputerName $comp | Select Name, ComputerName, DriverName, PortName, @{N="PrinterHostAddress"; Expression={if ($_.PortName -LIKE "WSD*") { Get-PrinterPort -Name $_.PortName -ComputerName $_.ComputerName | Select -ExpandProperty DeviceURL } else { Get-PrinterPort -Name $_.PortName -ComputerName $_.ComputerName | Select -ExpandProperty PrinterHostAddress }} }
Read-Host "Press enter to exit" | Out-Null

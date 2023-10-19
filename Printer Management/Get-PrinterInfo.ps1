# Returns printers for the specified computer. Attempts to resolve WSD addresses to IP.
$comp = Read-Host "Enter Computer Name"
Write-Host "** Getting printer info from [$comp], please wait...if this takes too long the computer may be offline"
Get-Printer -ComputerName $comp | Select Name, ComputerName, DriverName, PortName, @{N="PrinterHostAddress"; Expression={if ($_.PortName -LIKE "WSD*") { Get-PrinterPort -Name $_.PortName -ComputerName $_.ComputerName | Select -ExpandProperty DeviceURL } else { Get-PrinterPort -Name $_.PortName -ComputerName $_.ComputerName | Select -ExpandProperty PrinterHostAddress }} }
Read-Host "Press enter to exit" | Out-Null

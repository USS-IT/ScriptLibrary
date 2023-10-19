# Adapted from:  http://www.powershellpro.com/dimm-witt/200/
# MJC 9-19-23
$comp = Read-Host "Enter Computer Name"
If((Test-Connection -ComputerName $comp -Count 1 -Quiet) -Or (Test-Connection -ComputerName $comp -Count 1 -Quiet) -Or (Test-Connection -ComputerName $comp -Count 1 -Quiet)) {
	Write-Host "[$comp] appears to be online, querying..."
} else {
	Write-Warning "[$comp] did not respond to any ping attempts, querying anyway..."
}
Get-WmiObject -Class "win32_PhysicalMemoryArray" -namespace "root\CIMV2" -computerName $comp | % {
	"Total Number of DIMM Slots: " + $_.MemoryDevices
}
Get-WmiObject -Class "win32_PhysicalMemory" -namespace "root\CIMV2" -computerName $comp | % {
     "Memory Installed: " + $_.DeviceLocator
     "Memory Size: " + ($_.Capacity / 1GB) + " GB"
}

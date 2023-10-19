# Returns computer uptime since last Restart.
# Note Power > Shutdown will not reset this.
$comp = Read-Host "Enter Computer Name"
If((Test-Connection -ComputerName $comp -Count 1 -Quiet) -Or (Test-Connection -ComputerName $comp -Count 1 -Quiet) -Or (Test-Connection -ComputerName $comp -Count 1 -Quiet)) {
	Write-Host "[$comp] appears to be online, querying..."
} else {
	Write-Warning "[$comp] did not respond to any ping attempts, querying anyway..."
}
"Last Restart for [${comp}]: " + [Management.ManagementDateTimeConverter]::ToDateTime((Get-WmiObject Win32_OperatingSystem -ComputerName $comp).LastBootUpTime)
Read-Host "Press enter to exit" | Out-Null

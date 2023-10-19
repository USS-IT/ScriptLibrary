# Little snippets for getting real-time software info remotely queried through WMI. System must be on the Hopkins network.
# Account used must have either local admin or remote admin rights on system.
# Can sometimes fail with "RPC server not available" error depending on system configuration or if the computer is offline.

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

# Example with filtered out results.
#$comp = "HW-AD-17C6613"
#Get-WmiObject -Class Win32_Product -ComputerName $comp | where {$_.Name -like "Pulse*"} | Select Name, Version | Sort Name
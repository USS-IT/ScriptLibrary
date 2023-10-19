# Little snippets for getting real-time OS info remotely queried through WMI. System must be on the Hopkins network.
# Account used must have either local admin or remote admin rights on system.
# Can sometimes fail with "RPC server not available" error depending on system configuration or if the computer is offline.
$EOLVER=19042

$comp = Read-Host "Enter Computer Name"
# Check if computer is online (try up to three times).
If((Test-Connection -ComputerName $comp -Count 1 -Quiet) -Or (Test-Connection -ComputerName $comp -Count 1 -Quiet) -Or (Test-Connection -ComputerName $comp -Count 1 -Quiet)) {
	Write-Host "[$comp] appears to be online, querying..."
} else {
	Write-Warning "[$comp] did not respond to any ping attempts, querying anyway..."
}
Get-WmiObject -Class Win32_OperatingSystem -ComputerName $comp | Select PSComputerName, Caption, OSArchitecture, Version, BuildNumber, @{N="EOL Version"; Expression={ $EOLVER }}, @{N="Windows End of Life"; Expression={ $_.BuildNumber -le $EOLVER }}
Read-Host "Press enter to exit" | Out-Null

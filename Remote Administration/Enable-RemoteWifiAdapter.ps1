# MJC 1-26-23
# Enable wireless adapter remotely using WMI.
# Should return 0 if successful.
$computername = Read-Host "** Enter computer name"
# If any of 3 pings work, continue
If( (Test-Connection $computername -Count 1 -Quiet) -Or (Test-Connection $computername -Count 1 -Quiet) -Or (Test-Connection $computername -Count 1 -Quiet) ) {
	$results = Get-WmiObject win32_networkadapter -computerName $computername | where {$_.Name -like "*Wireless*" -OR $_.Name -like "*WiFi*" -OR $_.Name -like "*Wi-Fi*"}
	if (-Not [string]::IsNullOrEmpty($results.DeviceID)) {
		if ($results.DeviceID.Count -gt 1) {
			$results
			Write-Host "** More than 1 wireless device found. If you get an error after entering device ID, try again."
			$deviceid = Read-Host "** Enter Device ID # for Wireless Card"
			if ([string]::IsNullOrEmpty($deviceid) -Or ($deviceid -as [int]) -isnot [int]) {
				Write-Error "Device ID must be a number"
			}
		} else {
			$deviceid = $results.DeviceID | Select -First 1
		}
		if ($nic = Get-WMIObject win32_networkadapter -computerName $computername -Filter "DeviceID = $deviceid") {
			$nic.Enable()
		} else {
			Write-Error "Cannot find NIC with device ID [$deviceid]"
		}
	} else {
		$results
		Write-Warning "** No WiFi adapters found."
	}
}
# Comment out line below to run unattended
Read-Host "** Press enter to exit" | Out-Null
# Attempts to re-enable the Wireless adapter if it's down.
# Should be safe to run at startup.
# Original Author: mcarras8
# Last Updated: 2-13-25 mcarras8

# First check if we already have a connection (likely Ethernet).
$upcount = Get-NetAdapter -Physical | where {$_.Status -eq "Up"} | Measure-Object | Select -ExpandProperty Count
if ( $upcount -eq 0 -or $upcount -eq $null) {
	# No connection found.
	
	# Try looking for InterfaceType 71 (802.11 Wireless)
	$adapter = Get-NetAdapter -Physical | where {$_.InterfaceType -eq 71 -And ($_.Status -eq 'Disabled' -Or $_.Status -eq 'Not Present')} | Select -First 1

	# If no valid results, try using a broader Get-NetAdapter query.
	If ([string]::IsNullOrEmpty($adapter.ifIndex)) {
		# Try using Net-Adapter to narrow down the choices.
		$adapter = Get-NetAdapter -Physical | where {($_.Status -eq 'Disabled' -Or $_.Status -eq 'Not Present') -And ($_.Name -like 'Wi-Fi' -Or  $_.Name -like 'WiFi' -Or $_.Name -like 'Wireless' -Or $_.InterfaceDescription -like 'Wi-Fi' -Or  $_.InterfaceDescription -like 'WiFi' -Or $_.InterfaceDescription -like 'Wireless')} | Select -First 1
	}

	# If we have a valid result, try to enable the adapter two different ways.
	If (-Not [string]::IsNullOrEmpty($adapter.ifIndex)) {
		$adapter | Enable-NetAdapter -Confirm:$false
		$adapter = Get-WMIObject win32_networkadapter -Filter ("InterfaceIndex = '{0}'" -f $adapter.ifindex)
		$adapter.Enable()
	}
}

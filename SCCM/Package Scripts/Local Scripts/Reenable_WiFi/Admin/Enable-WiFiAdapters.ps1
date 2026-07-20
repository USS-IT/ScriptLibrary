<#
.SYNOPSIS
    Re-enables all disabled wireless adapters.
.DESCRIPTION
    Re-enables all disabled wireless adapters.
.PARAMETER Force
	Attempt to re-enable all wireless adapters, regardless if they're up or not.
.PARAMETER Notify
	Invokes an event-based notification script to notify the current user.
.PARAMETER LogDir
	Directory for logging.
.NOTES
	Author: Matt Carras (mcarras8)
	Created: 06-30-2026
	
	Requires Administrator privileges.
#>
#Requires -RunAsAdministrator
param(
	[switch] $Force,
	
	[switch] $Notify,
	
	[string] $LogDir = "C:\USS\Logs\Packages\Reenable_WiFi"
)

try {
	$_scriptName = Split-Path -Leaf $PSCommandPath
} catch {
	$_scriptName = "Enable-WiFiAdapters.ps1"
}
$LogPath = "$LogDir\$($_scriptName).log"

# Create the log path if it doesn't already exist
if (-not (Test-Path $LogDir -PathType Container)) {
	$null = New-Item -Path $LogDir -ItemType Directory
}
Start-Transcript $LogPath -Force

$adapters = Get-NetAdapter -Physical | where {$_.Status -eq 'Up'}
Write-Host ("Network adapters currently up: " + ($adapters.InterfaceDescription -join ","))
	
$doFallback = $true
try {
	# Try using WMI first.
	$filter = "PhysicalAdapter = True AND ((Name like '%Wi-Fi%' OR Name like '%WiFi%' OR Name like '%Wireless%') OR (Description like '%Wi-Fi%' OR Description like '%WiFi%' OR Description like '%Wireless%') OR NetConnectionID = 'Wi-Fi' OR NetConnectionID = 'WiFi' OR NetConnectionID = 'Wireless')"
	# If -Force is not given, only include ones reporting they're disabled.
	if (-Not $Force) {
		$filter += " AND (NetEnabled = False OR NetConnectionStatus = 0)"
	}
	$adapters = Get-CimInstance Win32_NetworkAdapter -Filter $filter -ErrorAction Stop
	If ($adapters -eq $null -Or ($adapters | Measure).Count -le 0) {
		Write-Host "No disabled wireless adapters found from WMI."
	} else {
		$doRecheck = $false
		$adapter = $null
		# Make sure we have only one result.
		if (($adapters | Measure).Count -eq 1) {
			$adapter = $adapters
		} else {
			Write-Host "WMI query returned more than one result. Attempting to narrow down results..."
			$adapter = $adapters | where {$_.NetEnabled -eq $false}
			if (($adapter | Measure).Count -ne 1) {
				$adapter = $adapters | where {$_.NetConnectionStatus -eq 0}
				if (($adapter | Measure).Count -ne 1) {
					$adapter = $null
				}
			}
		}
		if ($adapter -ne $null -And (Invoke-CimMethod -InputObject $adapter -MethodName Enable -Arguments @{} -ErrorAction Stop)) {
			Write-Host "[$($adapter.Name)] [$($adapter.InterfaceIndex)] enabled by WMI. Waiting and re-checking..."
			$counter = 0
			$sleepSec = 2
			$sleepMax = 20
			while ($counter -le $sleepMax) {
				Start-Sleep -Seconds $sleepSec
				$counter += $sleepSec
				$adapter = Get-CimInstance Win32_NetworkAdapter -Filter "InterfaceIndex = '$($adapter.InterfaceIndex)'"
				if ($adapter.NetEnabled -eq $true) {
					Write-Host "[$($adapter.Name)] [$($adapter.InterfaceIndex)] successfully re-enabled by WMI"
					$doFallback = $false
					if ($Notify) {
						Write-Host "Sending notification"
						Write-EventLog -LogName 'USS-EventLog' -Source 'Notify-Reenable_WiFi-Success' -EventID 1000 -EntryType Information -Message 'Notify of successful WiFi reenable'
					}
					break
				}
			}
			if ($doFallback) {
				Write-Host "[$($adapter.Name)] [$($adapter.InterfaceIndex)] failed to re-enable by WMI"
			}
		}
	}
} catch {
	Write-Error $_
}

# Fallback to Enable-NetAdapter if needed. This can also re-enable multiple adapters.
if ($doFallback) {
	Write-Host "Falling back to Get-NetAdapter"
	$adapters = Get-NetAdapter -Physical | where {($_.NdisPhysicalMedium -eq '802.11' -Or $_.InterfaceDescription -match 'Wi-?Fi|Wireless' -or $_.Name -match 'Wi-?Fi|Wireless')}
	if (-Not $Force) {
		$adapters = $adapters | where {($_.Status -eq 'Disabled' -Or $_.Status -eq 'Not Present')}
	}
	if (($adapters | Measure).Count -eq 0) {
		if ($Force) {
			Write-Host "No wireless adapters found"
			if ($Notify) {
				Write-Host "Sending notification"
				Write-EventLog -LogName 'USS-EventLog' -Source 'Notify-Reenable_WiFi-Missing' -EventID 1000 -EntryType Information -Message 'Notify no WiFi adapters found'
			}
		} else {
			Write-Host "No disabled wireless adapters found"
			if ($Notify) {
				Write-Host "Sending notification"
				Write-EventLog -LogName 'USS-EventLog' -Source 'Notify-Reenable_WiFi-Missing' -EventID 1000 -EntryType Information -Message 'Notify no disabled WiFi adapters found'
			}
		}
	} else {
		foreach ($adapter in $adapters) {
			$doFallback = $false
			try {
				Write-Host "Calling Enable-NetAdapter for [$($adapter.InterfaceDescription)] [$($adapter.ifindex)], then re-checking after waiting"
				$adapter | Enable-NetAdapter -Confirm:$false -ErrorAction Stop
				$counter = 0
				$sleepSec = 2
				$sleepMax = 20
				while ($counter -le $sleepMax) {
					Start-Sleep -Seconds $sleepSec
					$counter += $sleepSec
					$adapter = Get-NetAdapter -InterfaceIndex $adapter.ifindex
					if ($adapter.Status -eq 'Up') {
						Write-Host "[$($adapter.InterfaceDescription)] [$($adapter.ifindex)] successfully re-enabled by Enable-NetAdapter"
						if ($Notify) {
							Write-Host "Sending notification"
							Write-EventLog -LogName 'USS-EventLog' -Source 'Notify-Reenable_WiFi-Success' -EventID 1000 -EntryType Information -Message 'Notify of successful WiFi reenable'
						}
						break
					}
				}
				if ($adapter.Status -eq 'Disabled' -Or $adapter.Status -eq 'Not Present') {
					Write-Host "Get-NetAdapter still lists [$($adapter.InterfaceDescription)] [$($adapter.ifindex)] as disabled after re-checking"
					$doFallback = $true
				}
			} catch {
				Write-Error $_
				$doFallback = $true
			}
			if ($doFallback) {
				Write-Host "Attempting to re-enable [$($adapter.InterfaceDescription)] [$($adapter.ifindex)] by WMI fallback"
				$adapter = Get-CimInstance Win32_NetworkAdapter -Filter "InterfaceIndex = '$($adapter.ifindex)'"
				if ($adapter -ne $null -And (Invoke-CimMethod -InputObject $adapter -MethodName Enable -Arguments @{})) {
					Write-Host "[$($adapter.Name)] [$($adapter.interfaceindex)] enabled by WMI fallback. Waiting and re-checking..."
					$counter = 0
					$sleepSec = 2
					$sleepMax = 20
					$isSuccess = $false
					while ($counter -le $sleepMax) {
						Start-Sleep -Seconds $sleepSec
						$counter += $sleepSec
						$adapter = Get-CimInstance Win32_NetworkAdapter -Filter "InterfaceIndex = '$($adapter.InterfaceIndex)'"
						if ($adapter.NetEnabled -eq $true) {
							Write-Host "[$($adapter.Name)] [$($adapter.InterfaceIndex)] successfully re-enabled by WMI fallback"
							$isSuccess = $true
							if ($Notify) {
								Write-Host "Sending notification"
								Write-EventLog -LogName 'USS-EventLog' -Source 'Notify-Reenable_WiFi-Success' -EventID 1000 -EntryType Information -Message 'Notify of successful WiFi reenable'
							}
							break
						}
					}
					if (-Not $isSuccess) {
						Write-Host "[$($adapter.Name)] [$($adapter.interfaceindex)] still disabled after re-checking."
						if ($Notify) {
							Write-Host "Sending notification"
							Write-EventLog -LogName 'USS-EventLog' -Source 'Notify-Reenable_WiFi-Failure' -EventID 1000 -EntryType Information -Message 'Notify of failed WiFi reenable'
						}
					}
				}
			}
		}
	}
}
	
Stop-Transcript | Out-Null

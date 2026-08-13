# Get windows events from one or more online systems.
# FilterTable
# ComputerName - ComputerName to query. Required.
# Result - Result this computer had with an upgrade event. Optional.
# LogName - One or more logs to search, e.g. 'System'
# ProviderName - One or more providers/sources, e.g. Microsoft-Windows-TPM-WMI
# Id - One or more event IDs to filter. Optional.
# StartTime - Start datetime to filter. Optional.
# EndTime - End datetime to filter. Optional.
# At least ComputerName, LogName, or ProviderName is required.
#
# FocusID - Check these IDs without datetime restrictions.
#

param(
	[Parameter(Mandatory=$true)]
    [PSObject[]] $FilterTable,

	[int[]] $FocusID,
	
	[string] $CSV,
	
	[switch] $NoPause
)

$queried = @{}
$results = foreach ($o in $FilterTable) {
	$systemName = $o.ComputerName
	$prevResult = $o.Result
	
	if ([string]::IsNullOrEmpty($systemName)) {
		Write-Warning "Blank SystemName. Skipping."
		continue
	}

	if ($queried.ContainsKey($systemName)) {
		Write-Host "'$systemName' queried previously"
		
		if (-Not [string]::IsNullOrEmpty($prevResult) -And -Not $queried[$systemName].ContainsKey($prevResult)) {
			$queried[$systemName].Results[$prevResult] = $true
		}
		
		continue
	}
	
	Write-Host "** Checking if '$systemName' is online..."
	if (-Not (Test-Connection -ComputerName $systemName -Count 1 -Quiet) -And 
		-Not (Test-Connection -ComputerName $systemName -Count 1 -Quiet) -And 
		-Not (Test-Connection -ComputerName $systemName -Count 1 -Quiet)) {
		
		Write-Warning "'$systemName' did not respond to any ping attempts. Skipping."
		
		$queried[$systemName] = @{
			'Results' = @{}
			'Events' = $null
		}
		if (-Not [string]::IsNullOrEmpty($prevResult)) {
			$queried[$systemName].Results[$prevResult] = $true
		}
		
		continue
	}

	try {
		Write-Host "'$systemName' appears to be online. Querying..."
		
		$startDate = $o.StartTime -as [datetime]
		$endDate = $o.EndTime -as [datetime]
		if ($startDate -And -not $endDate) {
			$endDate = (Get-Date)
		}
		
		$filterHash = @{}
		if ($o.LogName) {
			$filterHash.Add('LogName', $o.LogName)
		}
		if ($o.ProviderName) {
			$filterHash.Add('ProviderName', $o.ProviderName)
		}
		if ($o.Id) {
			$filterHash.Add('Id', $o.Id)
		}
		if ($startDate) {
			$filterHash.Add('StartTime', $startDate)
		}
		if ($endDate) {
			$filterHash.Add('EndTime', $endDate)
		}

		$events = Get-WinEvent -ComputerName $systemName `
							   -FilterHashtable $filterHash `
							   -ErrorAction Stop
							   
		$queried[$systemName] = @{
			'Results' = @{}
			'Events' = $events
		}
		if (-Not [string]::IsNullOrEmpty($prevResult)) {
			$queried[$systemName].Results[$prevResult] = $true
		}
	
		if ($events.Count -eq 0) {
			Write-Host "No entries matching filter were found on '$systemName'."
		} else {
			# Also add ID if given
			if ($FocusID) {
				Write-Host "Also querying for IDs [$($FocusID -join ',')]..."
				$extraIds = @()
				foreach ($id in $FocusID) {
					if ($id -in $events.Id) {
						$extraIds += $id
					}
				}
				if ($extraIds.Count -gt 0) {
					$extraEvents = $null
					if ($filterHash.ContainsKey('StartTime')) {
						$filterHash.Remove('StartTime')
					}
					if ($filterHash.ContainsKey('EndTime')) {
						$filterHash.Remove('EndTime')
					}
					$filterHash['Id'] = $extraIds
					$extraEvents = Get-WinEvent -ComputerName $systemName `
								   -FilterHashtable $filterHash `
								   -ErrorAction SilentlyContinue
					if ($extraEvents -And $extraEvents.Count -gt 0) {
						$events += $extraEvents
					}
				}
			}
			
			Write-Host "Found $($events.Count) event(s) on '$systemName' matching filters."
			$events | Select-Object @{N="System"; E={$systemName}},
									TimeCreated,
									Id,
									@{N="Source"; E={$_.ProviderName}},
									@{N="Log";    E={$_.LogName}},
									LevelDisplayName,
									Message
		}
	} catch [System.Exception] {
		if ($_.Exception.Message -like "*No events were found*") {
			Write-Host "ERROR: No entries matching filter were found on '$systemName'."
		} else {
			Write-Error "Failed to query '$systemName': $($_.Exception.Message)"
		}
	}
}

if ($results) {
	$results = $results | Select System,
								 @{N="Results"; E={if ($queried.ContainsKey($_.System)) { ($queried[$_.System].Results.Keys | foreach { $_ }) -join ", " }}},
								 TimeCreated,
								 Id,
								 Source,
								 Log,
								 LevelDisplayName,
								 Message
}

if (-not [string]::IsNullOrEmpty($CSV)) {
	Write-Host "Exporting results to [$CSV]..."
	$results | Export-CSV -NoTypeInformation $CSV
} else {
	$results
	
	if (-Not $NoPause) {
		$_ = Read-Host ":Press any key to exit"
	}
}


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

param(
	[Parameter(Mandatory=$true)]
    [PSObject[]] $FilterTable,

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
	
	Write-Host "Checking if '$systemName' is online..."
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
		if (-Not [string]::IsNullOrEmpty($prevResult)) {
			Write-Host "'$systemName' appears to be online. Logged result was [$prevResult]. Querying..."
		} else {
			Write-Host "'$systemName' appears to be online. Querying..."
		}

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
			Write-Host "No entries matching filter were found on '$systemName' for $($startDate.ToString('yyyy-MM-dd'))."
		} else {
			Write-Host "Found $($events.Count) event(s) on '$systemName' for $($startDate.ToString('yyyy-MM-dd')):"
			$events | Select-Object TimeCreated,
									Id,
									@{N="Source"; E={$_.ProviderName}},
									@{N="Log";    E={$_.LogName}},
									LevelDisplayName,
									Message
		}
	} catch [System.Exception] {
		if ($_.Exception.Message -like "*No events were found*") {
			Write-Host "ERROR: No entries matching filter were found on '$systemName' for $($startDate.ToString('yyyy-MM-dd'))."
		} else {
			Write-Error "Failed to query '$systemName': $($_.Exception.Message)"
		}
	}
}

if (-not [string]::IsNullOrEmpty($CSV)) {
	Write-Host "Writing results [$CSV]..."
	$results | Export-CSV -NoTypeInformation $CSV
} else {
	$results
}

if (-Not $NoPause) {
	$_ = Read-Host ":Press any key to exit"
}
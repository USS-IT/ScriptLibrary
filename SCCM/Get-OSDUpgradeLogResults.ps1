# Exports Windows upgrade results parsed from the osdlogs folder.
param(
	[Parameter(Mandatory=$false)]
    [string] $LogsDir = '\\win.ad.jhu.edu\Data\osdlogs$\USS-ECI Upgrade Win11 25H2 x64 2026.02-V1.1',
	
	[Parameter(Mandatory=$true)]
	[string] $OutCSV
)

Write-Host "Searching [$LogsDir] for result directories..."
$results = Get-ChildItem -Path (Join-Path $LogsDir '\*') -Directory |
    ForEach-Object {
		try {
			$foldername = $_.FullName -split '\\' | Select -Last 1
			if (-Not $foldername) {
				Write-Warning "Problem parsing [$($_.FullName)]"
				continue
			}
		} catch {
			Write-Warning "Error parsing [$($_.FullName)]: $($_.Exception.Message)"
		}
	
		try {
			$foldername = $_.FullName -split '\\' | Select -Last 1
			if (-Not $foldername) {
				Write-Warning "Problem parsing [$($_.FullName)]"
				continue
			}
		} catch {
			Write-Warning "Error parsing [$($_.FullName)]: $($_.Exception.Message)"
		}
		
		if ($foldername -match "([^\.]+)\.([\d\.]+)@([\d\s\.]+)\.(.+)`$") {
			$systemName = $Matches[1]
			$logDate = $Matches[2]
			$logTime = $Matches[3]
			$logResult = $Matches[4]
		
			$diagfp = Join-Path $_.FullName 'setupdiagresults.log'
			if (-Not (Test-Path $diagfp )) {
				Write-Warning "[$diagfp] does not exist, not parsing"
			} else {
				$content = Get-Content $diagfp -Raw
				if ($content -match "Last Operation = (.+)") {
					$lastop = $Matches[1]
				} else {
					$lastop = $null
				}
				if ($content -match "Error = (.+)") {
					$lastop_error = $Matches[1]
				} else {
					$lastop_error = $null
				}
				if ($content -match "UpdateStartTime = (.+)") {
					$upgradeStart = $Matches[1]
				}
				if ($content -match "UpdateEndTime = (.+)") {
					$upgradeEnd = $Matches[1]
				}
				if ($content -match "RollbackStartTime = (.+)") {
					$rollbackStart = $Matches[1]
				}
				if ($content -match "RollbackEndTime = (.+)") {
					$rollbackEnd = $Matches[1]
				}
			}
			
			[PSCustomObject] @{
				'System' = $systemName
				'Result' = $logResult
				'Date' = $logDate
				'Time' = $logTime
				'Error' = $lastop_error
				'Last Op before Error' = $lastop
				'Upgrade Start Time' = $upgradeStartTime
				'Upgrade End Time' = $upgradeEndTime
				'Rollback Start Time' = $rollbackStartTime
				'Rollback End Time' = $rollbackEndTime
				'Log Path' = $_.FullName
			}
		} else {
			Write-Warning "Problem parsing folder name [$foldername] for [$($_.FullName)]"
		}
    }

if ($results) {
	Write-Warning "No results to export."
} else {
	$results | Export-CSV -NoTypeInformation $OutCSV
}
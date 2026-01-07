<#
	.SYNOPSIS
	Cross-references an export from various sources with the latest Snipe-It export.
	
	.DESCRIPTION
	Cross-references an export from various sources with the latest Snipe-It export. Currently, this only supports the Defender Missing Patch export.

	.PARAMETER InputCSV
	The exported CSV source file to cross-reference with Snipe-It.
	
	.PARAMETER OutputCSV
	Optional. The name for the cross-referenced output CSV. Default: snipeit_crossreference.csv
	
	.PARAMETER ExportType
	Optional. Determines the primary matching column and which columns was exported from source. Only option currently is "MissingPatch".
	
	.NOTES	
	Created: 5-20-25
	Author: Matt Carras (mcarras8)
#>
 [CmdletBinding(DefaultParameterSetName = 'Default')]
param(
	[Parameter(Position=0,Mandatory=$true)]
	[Alias("CSV")]
	[string] $InputCSV,
	
	[Parameter(Mandatory=$false)]
	[string] $OutputCSV="snipeit_crossreference.csv",
	
	[Parameter(ParameterSetName="Default",Mandatory=$false)]
	[ValidateSet("MissingPatch")]
	[string] $ExportType="MissingPatch"
)

# -- START CONFIGURATION --
# Path to latest Snipe-It export CSV
$SNIPEIT_EXPORT_CSV = "\\win.ad.jhu.edu\cloud\hsa$\ITServices\Reports\SnipeIt\Exports\assets_snipeit_latest.csv"
# Root of the Snipe-it instance for links
$SNIPEIT_ROOT_URL = "https://jh-uss.snipe-it.io"
# -- END CONFIGURATION

$defcomps = Import-CSV $InputCSV | where {-Not [string]::IsNullOrEmpty($_.Computer)}
$defcomps_count = ($defcomps | Measure).Count
Write-Host("Loaded [$defcomps_count] assets from source export [$InputCSV]")
$spcomps = Import-CSV $SNIPEIT_EXPORT_CSV
Write-Host("Loaded [{0}] assets from snipe-it export [$SNIPEIT_EXPORT_CSV]" -f ($spcomps | Measure).Count)
Write-Host("Cross-referencing...")
$counter = 0
$comps = $defcomps | % { 
	$counter++
	if ($counter -eq 1 -Or ($counter % 50) -eq 0) {
		Write-Host("Processed: $counter out of $defcomps_count")
	}
	$name = $_.Computer;
	if(-Not [string]::IsNullOrWhitespace($Name)) { 
		$comp = $spcomps | where {$_.name -eq $name} | Select -First 1; 
		if (-Not [string]::IsNullOrEmpty($comp.name)) {
			[PSCustomObject][ordered]@{
				name = $name
				Earliest_Missing_Patch = $_.Earliest_Missing_Patch
				Months = $_.Months
				Month_Group = $_.Month_Group
				asset_tag = $comp.asset_tag
				Status = $comp.status_label
				Department = $comp.Department
				assigned_to = $comp.assigned_to
				"Primary Users" = $comp."Primary Users"
				"AD LastLogonTime" = $comp."AD LastLogonTime"
				"SCCM LastActiveTime" = $comp."SCCM LastActiveTime"
				"OS Version" = $comp."OS Version"
				created_at = $comp.created_at
				model = $comp.model
				URL = "$SNIPEIT_ROOT_URL/hardware/bytag?assetTag=$($comp.asset_tag)"
			}
		} else {
			[PSCustomObject][ordered]@{
				name = $name
				Earliest_Missing_Patch = $_.Earliest_Missing_Patch
				Months = $_.Months
				Month_Group = $_.Month_Group
				asset_tag = "<NOT FOUND>"
				Status = ""
				Department = ""
				assigned_to = ""
				"Primary Users" = ""
				"AD LastLogonTime" = ""
				"SCCM LastActiveTime" = ""
				"OS Version" = ""
				created_at = ""
				model = ""
				URL = ""
			}
		}
	}
}

$comps | Export-CSV -NoTypeInformation -Force $OutputCSV
Write-Host("Exported [{0}] results to [$OutputCSV]" -f ($comps | Measure).Count)

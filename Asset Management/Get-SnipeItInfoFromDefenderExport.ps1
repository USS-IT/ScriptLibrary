<#
	.SYNOPSIS
	Cross-references an export from the Defender Health Report with the latest Snipe-It export.
	
	.DESCRIPTION
	Cross-references an export from the Defender Health Report with the latest Snipe-It export.

	.PARAMETER CSV
	The Defender export as a CSV file.
	
	.PARAMETER ExportType
	Optional. Determines the primary matching column and which columns was exported from source. Only option currently is "MissingPatch".
	
	.NOTES	
	Created: 5-20-25
	Author: Matt Carras (mcarras8)
#>
param(
	[DefaultParameterSetName="Default"]
	
	[Parameter(Position=0,Mandatory=$true)]
	[string] $CSV,
	
	[Parameter(ParameterSetName="Default")]
	[ValidateSet("MissingPatch")]
	[string] $ExportType="MissingPatch"
)

# -- START CONFIGURATION --
# Path to latest Snipe-It export CSV
$SNIPEIT_EXPORT_CSV = "\\win.ad.jhu.edu\cloud\hsa$\ITServices\Reports\SnipeIt\Exports\assets_snipeit_latest.csv"
# Root of the Snipe-it instance for links
$SNIPEIT_ROOT_URL = "https://jh-uss.snipe-it.io"
# -- END CONFIGURATION

$defcomps = Import-CSV $CSV | where {-Not [string]::IsNullOrEmpty($_.Computer)}
$spcomps = Import-CSV $SNIPEIT_EXPORT_CSV
$comps = $defcomps | % { 
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

$comps | Export-CSV -NoTypeInformation "Defender_Missing_Patches_with_SnipeItLookup.csv"

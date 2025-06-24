<#
	.SYNOPSIS
	Cross-references an export from the Defender Health Report with the latest Snipe-It export.
	
	.DESCRIPTION
	Cross-references an export from the Defender Health Report with the latest Snipe-It export.

	.NOTES	
	Created: 5-20-25
	Author: mcarras8
#>
$defcomps = Import-CSV ".\Hidden-Defender Missing Patch Month.csv" | where {-Not [string]::IsNullOrEmpty($_.Computer)}
$spcomps = Import-CSV "\\win.ad.jhu.edu\cloud\hsa$\ITServices\Reports\SnipeIt\Exports\assets_snipeit_latest.csv" | where {$_.name -in $defcomps.Computer}
$comps = $defcomps | % { 
	$name = $_.Computer;
	if(-Not [string]::IsNullOrWhitespace($Name)) { 
		$comp = $spcomps | where {$_.name -eq $name} | Select -First 1; 
		if (-Not [string]::IsNullOrEmpty($comp.name)) {
			[PSCustomObject]@{
				name = $name
				Earliest_Missing_Patch = $_.Earliest_Missing_Patch
				Months = $_.Months
				Month_Group = $_.Month_Group
				asset_tag = $comp.asset_tag
				Department = $comp.Department
				assigned_to = $comp.assigned_to
				"Primary Users" = $comp."Primary Users"
				"AD LastLogonTime" = $comp."AD LastLogonTime"
				"SCCM LastActiveTime" = $comp."SCCM LastActiveTime"
				created_at = $comp.created_at
				Status = $comp.status_label
				model = $comp.model
				URL = "https://jh-uss.snipe-it.io/hardware/bytag?assetTag=$($comp.asset_tag)"
			}
		} else {
			[PSCustomObject]@{
				name = $name
				Earliest_Missing_Patch = $_.Earliest_Missing_Patch
				Months = $_.Months
				Month_Group = $_.Month_Group
				asset_tag = ""
				Department = ""
				assigned_to = ""
				"Primary Users" = ""
				"AD LastLogonTime" = ""
				"SCCM LastActiveTime" = ""
				created_at = ""
				Status = ""
				model = ""
				URL = ""
			}
		}
	}
}

$comps | Export-CSV -NoTypeInformation "Defender_Missing_Patches_with_SnipeItLookup.csv"

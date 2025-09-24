<#
	.SYNOPSIS
	Exports comments and status on each GPO under given OUs to a CSV file.
	
	.DESCRIPTION
	Exports comments and status on each GPO under given OUs to a CSV file. Exports into your OneDrive Documents folder by default.
	
	.NOTES
	Author: Matthew Carras
	
	Requirements:
	* RSAT: Active Directory PowerShell module
#>

# May give multiple search bases.
$SEARCHBASES = @("OU=USS,DC=win,DC=ad,DC=jhu,DC=edu")
# Will export to OneDrive Documents folder by default (if it exists), otherwise current directory.
$exportFP = $null

# Can possibly get more info by using Get-GPOReport -Guid <guid> -ReportType XML and parsing the results
Write-Host "Exporting GPOs, please wait..."
$gpos = $SEARCHBASES | foreach { 
	Write-Host "Exporting GPOs for [$_]..."
	Get-ADOrganizationalUnit -Filter * -Searchbase $_ -Searchscope subtree | foreach { 
		$ou = $_.distinguishedname
		$linkedCount = ($_.LinkedGroupPolicyObjects | Measure).Count
		if ($linkedCount -gt 0) {
			Write-Host("Processing [{0}] linked GPOs in [$ou]" -f $linkedCount)
			foreach( $guidlink in $_.LinkedGroupPolicyObjects ) { 
				if($guidlink -match "cn=\{([^}]+)\}," -And -Not [string]::IsNullOrWhitespace($Matches[1])) { 
					$gpo = Get-GPO -Guid $Matches[1]
					[PSCustomObject]@{
						"Linked OU" = $ou
						"GPO Name" = $gpo.DisplayName
						"GPO Status" = $gpo.GpoStatus
						"GPO Description" = $gpo.Description
					}
				} 
			}
		}
	}
} | Group-Object -Property "GPO Name" | foreach {
	[PSCustomObject]@{
		"GPO Name" = $_.Group[0]."GPO Name" 
		"GPO Status" = $_.Group[0]."GPO Status" 
		"GPO Description" = $_.Group[0]."GPO Description" 
		"Linked OUs" = $_.Group."Linked OU" -join "; "
	}
}

if ([string]::IsNullOrEmpty($exportFP)) {
	$exportFP = "${ENV:OneDrive}\Documents\uss-gpos.csv"
	if (-Not (Test-Path "${ENV:OneDrive}\Documents" -PathType Container)) {
		$exportFP = "uss-gpos.csv"
	}
}
$gpos | Export-CSV -NoTypeInformation $exportFP
Write-Host "Exported GPO info to [$exportFP]"

$_ = Read-Host "Press enter to exit"
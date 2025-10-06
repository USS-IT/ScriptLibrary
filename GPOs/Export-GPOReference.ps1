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
# Fix a bug with Trustee output from Get-GPPermission for PowerShell 7.x
if ($PSVersionTable.PSVersion.Major -ge 7) {
	Import-Module -Name GroupPolicy -SkipEditionCheck -Force
} else {
	Import-Module -Name GroupPolicy
}

# May give multiple search bases.
$SEARCHBASES = @("OU=USS,DC=win,DC=ad,DC=jhu,DC=edu")
# Will export to OneDrive Documents folder by default (if it exists), otherwise current directory.
$exportFP = $null
# Default GPO permission groups which are ignored for exports.
$GPODEFAULTTRUSTEE = @("SYSTEM","Enterprise Admins","Domain Admins","ENTERPRISE DOMAIN CONTROLLERS")

# Can possibly get more info by using Get-GPOReport -Guid <guid> -ReportType XML and parsing the results
Write-Host "Exporting GPOs, please wait..."
$perm_map = @{}
$gpo_map = @{}
$gpos = $SEARCHBASES | foreach { 
	Write-Host "Exporting GPOs for [$_]..."
	Get-ADOrganizationalUnit -Filter * -Searchbase $_ -Searchscope subtree | foreach { 
		$ou = $_.distinguishedname
		$linkedCount = ($_.LinkedGroupPolicyObjects | Measure).Count
		if ($linkedCount -gt 0) {
			Write-Host("Processing [{0}] linked GPOs in [$ou]" -f $linkedCount)
			foreach( $guidlink in $_.LinkedGroupPolicyObjects ) { 
				if($guidlink -match "cn=\{([^}]+)\}," -And -Not [string]::IsNullOrWhitespace($Matches[1])) {
					try {
						# First, check the hash table.
						$gpo = $gpo_map[$Matches[1]]
						if ([string]::IsNullOrEmpty($gpo.DisplayName)) {
							$gpo = Get-GPO -Guid $Matches[1] -ErrorAction Stop
							$gpo_map[$Matches[1]] = $gpo
						}
						$perms = $perm_map[$Matches[1]]
						if ([string]::IsNullOrEmpty(($perms | Select -First 1 -ExpandProperty Trustee).Name)) {
							$perms = Get-GPPermission -Guid $Matches[1] -All -ErrorAction Stop | where {-Not [string]::IsNullOrEmpty($_.Trustee.Name) -And $_.Trustee.Name -notin $GPODEFAULTTRUSTEE}
							$perm_map[$Matches[1]] = $perms
						}
						$permsEdit = ($perms | where {$_.Permission -eq "GpoEditDeleteModifySecurity"} | % { $_.Trustee.Name }) -join "; "
						$permsApply = ($perms | where {$_.Permission -eq "GpoApply"} | % { $_.Trustee.Name }) -join "; "
						$permsCustom = ($perms | where {$_.Permission -eq "GpoCustom"} | % { $_.Trustee.Name }) -join "; "
						
						[PSCustomObject]@{
							"Linked OU" = $ou
							"GPO Name" = $gpo.DisplayName
							"GPO Status" = $gpo.GpoStatus
							"GPO Description" = $gpo.Description
							"GPO Perms - Edit" = $permsEdit
							"GPO Perms - Apply" = $permsApply
							"GPO Perms - Custom (Deny)" = $permsCustom
						}
					} catch {
						Write-Error $_
						[PSCustomObject]@{
							"Linked OU" = "Error"
							"GPO Name" = $Matches[1]
							"GPO Status" = "Error"
							"GPO Description" = $_.Exception.Message
							"GPO Perms - Edit" = ""
							"GPO Perms - Apply" = ""
							"GPO Perms - Custom (Deny)" = ""
						}
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
		"GPO Perms - Edit" = $_.Group[0]."GPO Perms - Edit"
		"GPO Perms - Apply" = $_.Group[0]."GPO Perms - Apply"
		"GPO Perms - Custom (Deny)" = $_.Group[0]."GPO Perms - Custom (Deny)"
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
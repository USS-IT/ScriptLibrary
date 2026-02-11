<#
	.SYNOPSIS
	Returns 0 (success) or 1 (error) depending on whether the given software's uninstall key can be found in the registry.
	
	.DESCRIPTION
	Returns 0 (success) or 1 (error) depending on whether the given software's uninstall key can be found in the registry.
    
	.PARAMETER SoftwareName
	Required. The display name of the software to check installation status. This should be the same as the one from Add/Remove Programs. This will be a partial match unless -ExactMatch is given.
	
	.PARAMETER Publisher
	Optional publisher to match along with display name. This will be a partial match unless -ExactMatch is given.

	.PARAMETER Version
	Optional version to match. This is always an exact match.

	.PARAMETER ExactMatch
	Only accept exact matches for the given software and publisher. 
	
	.PARAMETER Regex
	Treat SoftwareName and Publisher as regex patterns. Overriden by -ExactMatch switch.
	
	.EXAMPLE
	powershell.exe -NoProfile -Windowstyle Hidden -ExecutionPolicy Bypass -File "InstallationChecker.ps1" -SoftwareName "PowerFAIDS"

	.NOTES
	Author: mcarras8
	
	Changelog
	1-9-2026 mcarras8 Script creation
#>

param(
	
	[parameter(Mandatory=$true, Position=0)]
	[string] $SoftwareName,
	
	[string] $Publisher,

	[string] $Version,

	[switch] $ExactMatch,
	
	[switch] $Regex
)

$_scriptName = split-path $PSCommandPath -Leaf

# Escape given strings unless we're using -ExactMatch or -Regex switches
$_displayname = $SoftwareName
$_publisher = $Publisher
if (-Not $ExactMatch -And -Not $Regex) {
	$_displayname = [Regex]::Escape($_displayname)
	$_publisher = [Regex]::Escape($_publisher)
}

# Check 32-bit, then 64-bit registry nodes.
$uninstallString = $null
$installKey = gci "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall" | foreach { gp $_.PSPath } | ? { ( $ExactMatch -And $_.DisplayName -eq $_displayname -And ( [string]::IsNullOrEmpty($_publisher) -Or $_.Publisher -eq $_publisher ) ) -Or ( -Not $ExactMatch -And $_ -match $_displayname -And ( [string]::IsNullOrEmpty($_publisher) -Or $_.Publisher -match $_publisher ) ) -And ( [string]::IsNullOrEmpty($Version) -Or $_.DisplayVersion -eq $Version ) }
if ($installKey -And ($installKey | Get-Member -Type NoteProperty | ? {$_.Name -eq 'UninstallString'} | Measure).Count -gt 0) {
	$uninstallString = $installKey | Select UninstallString | Select -ExpandProperty UninstallString
}
# If missing or blank UninstallString, check 64-bit registry
if ([string]::IsNullOrWhitespace($uninstallString)) {
	$installKey = gci "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall" | foreach { gp $_.PSPath } | ? { ( $ExactMatch -And $_.DisplayName -eq $_displayname -And ( [string]::IsNullOrEmpty($_publisher) -Or $_.Publisher -eq $_publisher ) ) -Or ( -Not $ExactMatch -And $_ -match $_displayname -And ( [string]::IsNullOrEmpty($_publisher) -Or $_.Publisher -match $_publisher ) ) -And ( [string]::IsNullOrEmpty($Version) -Or $_.DisplayVersion -eq $Version ) }
	if ($installKey -And ($installKey | Get-Member -Type NoteProperty | ? {$_.Name -eq 'UninstallString'} | Measure).Count -gt 0) {
		$uninstallString = $installKey | Select UninstallString | Select -ExpandProperty UninstallString
	}
}

if ([string]::IsNullOrWhitespace($uninstallString)) {
	Write-Warning "[$_scriptName] Unable to find uninstall string for [$SoftwareName] in Registry"
	exit 1
}

Write-Verbose "[$_scriptName] Found uninstallation key for [$SoftwareName] in Registry: $uninstallString"
exit 0
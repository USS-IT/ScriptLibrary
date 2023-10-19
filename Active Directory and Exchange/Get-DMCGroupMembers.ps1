<#
    .SYNOPSIS
    Show DMC group memberships for given domain user.
    
	.DESCRIPTION
    Show DMC group memberships for given domain user.
    
	.PARAMETER JHED
    The JHED of the user to search for. Prompted if not given.
	
	.EXAMPLE
	Get-DMCGroupMembers.ps1 mcarras8
	
	.NOTES
    Author: MJC 11-2-2022
#>
param(
	[Parameter(Mandatory=$false, Position=0)]
	[ValidateScript({-Not [string]::IsNullOrWhitespace($_)})]
	[string]$JHED
)

# Restrict groups to those in the given OU.
$groupOU = "OU=DMC Equipment Borrowing,OU=Groups,OU=USS,DC=win,DC=ad,DC=jhu,DC=edu"

# Default domain controller to query.
# Will attempt to get current DC if left blank.
$DC = "DC=win,DC=ad,DC=jhu,DC=edu"

# Title for prompt.
$title = 'DMC Equipment Group Search'

[void][System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')

if ([string]::IsNullOrEmpty($DC)) {
	try {
		$root = [ADSI]"LDAP://RootDSE"
		$rootDC = $root.rootDomainNamingContext
	} catch {
		[Microsoft.VisualBasic.Interaction]::MsgBox("ERROR: Could not retrieve DC", "OKOnly,Critical", "Error") | Out-Null
		return -1
	}
}

# Create mappings of DN = group name
$groups = @{}
try {
	$searcher = New-Object System.DirectoryServices.DirectorySearcher
	$searcher.SearchScope = 'Subtree'
	$searcher.SearchRoot = [ADSI]"LDAP://$groupOU"
	$searcher.filter = '(objectClass=group)'
	$searcher.FindAll() | foreach {
		$groups[($_.Properties.distinguishedname | Select -First 1)] = $_.Properties.name | Select -First 1
	}
} catch {
	[Microsoft.VisualBasic.Interaction]::MsgBox("ERROR: Unable to retrieve groups from [$groupOU]", "OKOnly,Critical", "Error - Get-DMCGroupMembers") | Out-Null
	return -2
}

# Sanity check
if ($groups.Count -eq 0) {
	[Microsoft.VisualBasic.Interaction]::MsgBox("ERROR: [$groupOU] returned no groups", "OKOnly,Critical", "Error - Get-DMCGroupMembers") | Out-Null
	return -3
}

# Grab user memberships
$memberOf = $null
$loopUntilValid = $false
do {
	if (-Not [string]::IsNullOrWhitespace($JHED)) {
		$name = $JHED
	} else {
		$name = [Microsoft.VisualBasic.Interaction]::InputBox('Please enter JHED:',$title)
		$loopUntilValid = $true
	}
	if ([string]::IsNullOrEmpty($name)) {
		# Cancel pressed
		$loopUntilValid = $false
		break
	}
	if (-Not [string]::IsNullOrWhitespace($name)) {
		try {
			$searcher.SearchRoot = [ADSI]"LDAP://$DC"
			$searcher.filter = "(&(objectClass=user)(objectCategory=person)(name=$name))"
			$memberOf = ($searcher.FindOne() | foreach { $_ }).properties.memberof
		} catch {
			[Microsoft.VisualBasic.Interaction]::MsgBox("ERROR retrieving user [$name]", "OKOnly,Critical", "Error - Get-DMCGroupMembers") | Out-Null
			return -4
		}
		if ($memberOf.Count -eq 0) {
			[Microsoft.VisualBasic.Interaction]::MsgBox("User [$name] not found or invalid, try again", "OKOnly,Critical", "Error - Get-DMCGroupMembers") | Out-Null
		} else {
			$loopUntilValid = $false
			break
		}
	}
} while($loopUntilValid)

if ($memberOf.Count -gt 0) {
	# Restrict to DMC groups.
	$memberOf = $memberOf | where {$_ -in $groups.Keys}
	if ($memberOf.Count -eq 0) {
		[Microsoft.VisualBasic.Interaction]::MsgBox("User [$name] is not a member of any DMC Equipment groups.", "OKOnly", "Get-DMCGroupMembers")  | Out-Null
	} else {
		$memberOf | foreach { $groups[$_] } | Out-GridView -Title "DMC Equipment Group Membership for [$name]"
	}
}
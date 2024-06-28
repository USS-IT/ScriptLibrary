<#
    .SYNOPSIS
    Search for primary and alias email addresses in AD and display access group information for a service account mailbox or distribution group, if applicable.
    
	.DESCRIPTION
    Search for accounts with primary mail and aliases using given wildcard, showing information about any associated mail management (OLGroups) found.
	
	For example, "*communications*" will show accounts with the word "communications" included in their primary or alias email addresses.
    
	.PARAMETER Email
    Required. Partial or full email address to search for.
	
	.PARAMETER OnlyUsers
    Only return user accounts, excluding mail-enabled groups from the search (distribution and M365 groups).
	
	.PARAMETER Silent
	Only output the result to console.
	
	.EXAMPLE
	Find-ADEmailInfo.ps1 "hsaitservices*"
	
	.NOTES
	Requires RSAT AD PowerShell module.
	
    Author: MJC 10-12-2022
#>
param(
	[Parameter(Mandatory=$true, Position=0)]
	[string]$Email, 
	
	[switch]$OnlyUsers,
	
	[switch]$Silent
)

function Find-ADEmailInfo {
	param (
		# Email address to search for in ProxyAddresses. Accepts * wildcards (Ex: "wellness*"). Note using two wildcards like "*address*" will take a bit longer.
		[parameter(Mandatory=$true,
				   ValueFromPipeline=$true,
				   ValueFromPipelineByPropertyName=$true)]
		[string]$Email,
		# Only return user accounts, excluding mail-enabled groups from the search (distribution and M365 groups).
		[switch]$OnlyUsers,
		# Only output the result to console.
		[switch]$Silent
	)
	# Add "@*" to end if "@" is missing and given search string doesn't end with a wildcard
	if ($Email.IndexOf('@') -lt 0 -And $Email.LastIndexOf('*') -ne ($Email.length-1)) {
		$Email += "@*"
	}
	if ($OnlyUsers) {
		if (-Not $Silent) {
			Write-Host "Searching users for string [$Email] in mail and ProxyAddresses..."
		}
		$filter = "(|(mail=$Email)(proxyAddresses=smtp:$Email))"
		$results = Get-ADUser -LDAPFilter $filter -Properties SamAccountName,mail,ProxyAddresses,Manager,DistinguishedName,UserAccountControl,UserPrincipalName
	} else {
		if (-Not $Silent) {
			Write-Host "Searching users and mail-enabled groups for string [$Email] in mail and ProxyAddresses..."
		}
		$filter = "(&(|(objectClass=group)(objectClass=user))(|(mail=$Email)(proxyAddresses=smtp:$Email)))"
		$results = Get-ADObject -LDAPFilter $filter -Properties SamAccountName,mail,ProxyAddresses,Manager,DistinguishedName,ManagedBy,msExchCoManagedByLink,UserAccountControl,UserPrincipalName
	}
	$results | foreach-object { 
		$group = $null
		$olgroupname = $null
		$olgroupmember = $null
		$olgroupowner = $null
		$olgroupcoowners = $null
		$enabled = $null
		$sam = $_.SamAccountName
		$upn = $_.UserPrincipalName
		$mail = $_.mail

		if ($_.ObjectClass -eq "group") {
			$group = $_
		} else {
			$groupfound = $false
			# Attempt to find mailbox management groups in the format "grp-SamAccountName", ignoring errors
			try { 
				$group = Get-ADGroup "grp-$sam" -Properties DistinguishedName,ManagedBy,msExchCoManagedByLink -ErrorAction SilentlyContinue
				if (-Not [string]::IsNullOrEmpty($group.ManagedBy) -Or $group.msExchCoManagedByLink.Count -gt 0 -Or $group.Members.Count -gt 0) {
					$groupfound = $true
				}
			} catch {}
			# Try again with format "grp-NameFromUPN"
			if (-Not $groupfound) {
				try { 
					$group = Get-ADGroup ("grp-{0}" -f ($upn -split "@" | Select -First 1)) -Properties DistinguishedName,ManagedBy,msExchCoManagedByLink -ErrorAction SilentlyContinue
				} catch {}
				if (-Not [string]::IsNullOrEmpty($group.ManagedBy) -Or $group.msExchCoManagedByLink.Count -gt 0 -Or $group.Members.Count -gt 0) {
					$groupfound = $true
				}
			}
			# Try one final time with format "grp-NameFromMail"
			if (-Not $groupfound) {
				try {
					$group = Get-ADGroup ("grp-{0}" -f ($mail -split "@" | Select -First 1)) -Properties DistinguishedName,ManagedBy,msExchCoManagedByLink -ErrorAction SilentlyContinue
				} catch {}
				if (-Not [string]::IsNullOrEmpty($group.ManagedBy) -Or $group.msExchCoManagedByLink.Count -gt 0 -Or $group.Members.Count -gt 0) {
					$groupfound = $true
				}
			}
		}
		# Collect OLGroup information, if it exists
		if($group) {
			$olgroupname = $group.Name;
			$olgroupmem = (Get-ADGroupMember $group.DistinguishedName).Name -join "; "
			if ($group.ManagedBy) {
				$olgroupowner = (Get-ADObject $group.ManagedBy -ErrorAction SilentlyContinue).Name
			}
			if ($group.msExchCoManagedByLink){ 
				$olgroupcoowners = ($group.msExchCoManagedByLink | foreach {(Get-ADObject $_ -ErrorAction SilentlyContinue).Name}) -join "; "
			}
			if ($_.ObjectClass -eq "group") {
				$type = "DistributionGroup"
				$enabled = "N/A"
			} else {
				$type = "DelegatedMailbox (User)"
			}
		} elseif ($_.ObjectClass -eq "user") {
			$type = "User"
		} else {
			# Shouldn't ever get here.
			if (-Not $Silent) {
				Write-Host "Encountered an unexpected ObjectClass: {0}" -f $_.ObjectClass
			}
			$type = "<ObjectClass: {0}>" -f $_.ObjectClass
		}
		if ($enabled -eq $null -And $_.UserAccountControl -ne $null) {
			# Compute from UserAccountControl property (bitmask not 2)
			$enabled = ($_.UserAccountControl -band 2) -ne 2
		}
		if ($_.Manager) {
			$manager = (Get-ADObject $_.Manager -ErrorAction SilentlyContinue).Name
		} else {
			$manager = $null
		}
		[PSCustomObject]@{
			"Name" = $_.Name
			"mail" = $_.mail
			"UserPrincipalName" = $_.UserPrincipalName
			"SamAccountName" = $_.SamAccountName
			"Type" = $type
			"Enabled" = $enabled
			"Manager" = $manager
			"OLGroup" = $olgroupname
			"OLGroupManagedBy" = $olgroupowner
			"OLGroupComanagedBy" = $olgroupcoowners
			"OLGroupMembers" = $olgroupmem
			"ProxyAddresses" = (($_.ProxyAddresses | foreach { if($_ -imatch "smtp:(.+)") { $matches[1] }}) -join "; ")
			"DistinguishedName" = $_.DistinguishedName
		} 
	}
}

# If command-line parameters are given
if (-Not [string]::IsNullOrWhitespace($Email)) {
	$results = Find-ADEmailInfo -Email $Email -OnlyUsers:$OnlyUsers -Silent:$Silent
    $results
    if (-Not $Silent) {
        $_ = Read-Host "Press enter to exit"
    }
}
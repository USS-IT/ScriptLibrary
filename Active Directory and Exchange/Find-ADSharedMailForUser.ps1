<#
    .SYNOPSIS
    Find information about shared mailboxes a user has access to.
    
	.DESCRIPTION
    Find information about shared mailboxes a user has access to.
    
	.PARAMETER Username
    Required. Username to search for.

	.PARAMETER Silent
	Only output the result to console.
	
	.EXAMPLE
	Find-ADEmailInfoForUser.ps1 mcarras8
	
	.NOTES
	Requires RSAT AD PowerShell module.
	
    Author: MJC 10-12-2022
#>
param(
	[Parameter(Mandatory=$true, Position=0)]
	[string]$Username,
	
	[switch]$Silent
)

function Find-ADSharedMailForUser {
	param (
		[parameter(Mandatory=$true,
				   ValueFromPipeline=$true,
				   ValueFromPipelineByPropertyName=$true)]
		[string]$Username,
		
		[switch]$Silent
	)
	try {
		if (-Not $Silent) {
			Write-Host "Getting groups for [$Username]..."
		}
		$user = Get-ADUser $Username -Properties distinguishedname,memberof -ErrorAction Stop
		$userGroups = $user | Get-ADUser -Properties tokenGroups -ErrorAction Stop | Select -ExpandProperty tokenGroups | Get-ADGroup -Properties distinguishedname,msExchCoManagedByLink,managedby -ErrorAction Stop
		$ownerGroups = Get-ADGroup -LDAPFilter "(|(msExchCoManagedByLink:=$($user.distinguishedname))(managedby=$($user.distinguishedname)))" -Properties distinguishedname,msExchCoManagedByLink,managedby
		$groups = $userGroups + $ownerGroups
	} catch {
		Write-Error $_
		return 
	}
	$olGroups = $groups | where {$_.Name -like "grp-*"}
	
	Write-Verbose "OL Groups found: $($olGroups -join ', ')"
	
	foreach ($group in $olGroups) {
		$olgroupname = $null
		$olgroupmembers = $null
		$olgroupowner = $null
		$olgroupcoowners = $null
		$enabled = $null
		$mail = $null
		$sam = $group.Name -replace 'grp-',''
		
		if ([string]::IsNullOrEmpty($sam)) {
			Write-Host "ERROR Unable to parse SAMAccountName for [$($group.Name)]"
		} else {
			# Lookup the mailbox by sam account name
			try {
				Write-Host "Checking for mailbox or distribution group with name [$sam] for group [$($group.Name)]"
				$mail = Get-ADObject -LDAPFilter "(|((SamAccountName=$sam)(Name=$sam)(proxyAddresses=smtp:$sam@jh.edu)))" -Properties SamAccountName,mail,ProxyAddresses,Manager,DistinguishedName,ManagedBy,msExchCoManagedByLink,UserAccountControl,UserPrincipalName
			} catch {
				Write-Host "ERROR getting mail-enabled AD object for group [$($group.Name)] with SAMAccountName [$sam]. Error message: $_"
			}
			
			# Collect OLGroup information, if it exists
			$olgroupname = $group.Name;
			$olgroupmembers = (Get-ADGroupMember $group.DistinguishedName).Name -join "; "
			if ($group.ManagedBy) {
				$olgroupowner = (Get-ADObject $group.ManagedBy -ErrorAction SilentlyContinue).Name
			}
			if ($group.msExchCoManagedByLink){ 
				$olgroupcoowners = ($group.msExchCoManagedByLink | foreach {(Get-ADObject $_ -ErrorAction SilentlyContinue).Name}) -join "; "
			}
			if ($mail) {
				if ($mail.ObjectClass -eq "group") {
					$type = "DistributionGroup"
					$enabled = "N/A"
				} else {
					$type = "DelegatedMailbox (User)"
				}
				if ($enabled -eq $null -And $mail.UserAccountControl -ne $null) {
					# Compute from UserAccountControl property (bitmask not 2)
					$enabled = ($mail.UserAccountControl -band 2) -ne 2
				}
				if ($mail.Manager) {
					$manager = (Get-ADObject $mail.Manager -ErrorAction SilentlyContinue).Name
				} else {
					$manager = $null
				}
			}
			[PSCustomObject]@{
				"Name" = $mail.Name
				"mail" = $mail.mail
				"UserPrincipalName" = $mail.UserPrincipalName
				"SamAccountName" = $mail.SamAccountName
				"Type" = $type
				"Enabled" = $enabled
				"Manager" = $manager
				"OLGroup" = $olgroupname
				"OLGroupManagedBy" = $olgroupowner
				"OLGroupComanagedBy" = $olgroupcoowners
				"OLGroupMembers" = $olgroupmembers
				"ProxyAddresses" = (($mail.ProxyAddresses | foreach { if($_ -imatch "smtp:(.+)") { $matches[1] }}) -join "; ")
				"DistinguishedName" = $mail.DistinguishedName
			} 
		}
	}
}

# If command-line parameters are given
$results = Find-ADSharedMailForUser -Username $Username -Silent:$Silent
$results
if (-Not $Silent) {
	$_ = Read-Host ":Press enter to exit"
}
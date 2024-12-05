<#
	.SYNOPSIS
	Checks and manages access to SSO access groups for Mental Health Services.
	
	.DESCRIPTION
	Checks and manages access to SSO access groups for Mental Health Services's TimelyCare and SilverCloud online platforms. If modifying access, the user running this script must have permission to manage these groups.
        
#>
	
# -- START CONFIGURATION --

# Various AD OU settings.
$DC_ROOT = "DC=win,DC=ad,DC=jhu,DC=edu"
$OU_USER = "OU=PEOPLE,$DC_ROOT"

# Dynamic AD groups granting TimelyCare SSO access.
$TIMELYCARE_ADGROUPS = @(
	"BSPH_Student_Dynamic", 
	"CBS_Student_Dynamic", 
	"KSAS_Student_Dynamic", 
	"Peabody_Student_Dynamic", 
	"SAIS_Student_Dynamic", 
	"SOE_Student_Dynamic", 
	"SOM_Student_Dynamic", 
	"SON_Student_Dynamic", 
	"WSE_Student_Dynamic"
)

# Ad-hoc AD access groups used for TimelyCare.
# These should be distinguished names.
$TIMELYCARE_ADHOC_ADGROUPS = @(
	"CN=BSPH_Student_AdHoc,OU=USS-TimelyCare_Access,OU=AD-Hoc-Access-Groups,OU=Groups,OU=USS,$DC_ROOT", 
	"CN=CBS_Student_AdHoc,OU=USS-TimelyCare_Access,OU=AD-Hoc-Access-Groups,OU=Groups,OU=USS,$DC_ROOT", 
	"CN=KSAS_Student_AdHoc,OU=USS-TimelyCare_Access,OU=AD-Hoc-Access-Groups,OU=Groups,OU=USS,$DC_ROOT", 
	"CN=Peabody_Student_AdHoc,OU=USS-TimelyCare_Access,OU=AD-Hoc-Access-Groups,OU=Groups,OU=USS,$DC_ROOT", 
	"CN=SAIS_Student_AdHoc,OU=USS-TimelyCare_Access,OU=AD-Hoc-Access-Groups,OU=Groups,OU=USS,$DC_ROOT", 
	"CN=SOE_Student_AdHoc,OU=USS-TimelyCare_Access,OU=AD-Hoc-Access-Groups,OU=Groups,OU=USS,$DC_ROOT", 
	"CN=SOM_Student_AdHoc,OU=USS-TimelyCare_Access,OU=AD-Hoc-Access-Groups,OU=Groups,OU=USS,$DC_ROOT", 
	"CN=SON_Student_AdHoc,OU=USS-TimelyCare_Access,OU=AD-Hoc-Access-Groups,OU=Groups,OU=USS,$DC_ROOT", 
	"CN=WSE_Student_AdHoc,OU=USS-TimelyCare_Access,OU=AD-Hoc-Access-Groups,OU=Groups,OU=USS,$DC_ROOT"
)

$TIMELYCARE_ADGROUPS += $TIMELYCARE_ADHOC_ADGROUPS

# Dynamic AD Groups granting SilverCloud SSO access.
$SILVERCLOUD_ADGROUPS = @(
	"SSO_SilverCloud"
)

# Ad-hoc AD access groups used for SilverCloud.
# These should be distinguished names.
$SILVERCLOUD_ADHOC_ADGROUPS = @(
	"CN=HW-HSA-SilverCloudADHoc,OU=AD-Hoc-Access-Groups,OU=Groups,OU=USS,$DC_ROOT"
)

$SILVERCLOUD_ADGROUPS += $SILVERCLOUD_ADHOC_ADGROUPS

# -- END CONFIGURATION --

# -- START FUNCTIONS --

<#
	.SYNOPSIS
	Checks whether a user is a member of a given group using ADSI.
	
	.DESCRIPTION
	Checks whether a user is a member of a given group using ADSI. Returns the group names found.
    
	.PARAMETER Username
	Required username to check.
	
	.PARAMETER Groups
	Required string array of AD group names to check. This can be full distinguished name or name only.
	
	.PARAMETER ReturnFullGroup
	If given, always returns the groups full distinguished names instead of just their name.
	
	.EXAMPLE
	Check-ADSIUserMembership -Username mcarras8 -Groups @("HW-HSA-SilverCloudADHoc")
#>
function Check-ADSIUserMembership {
	param (
		[parameter(Mandatory=$true, Position=0)]
		[ValidateScript({-Not [string]::IsNullOrWhitespace($_)})]
		[string]$Username,
		
		[parameter(Mandatory=$true, Position=1)]
		[array]$Groups,
		
		[switch]$ReturnFullGroup
	)
	
	$results = $false
	
	Write-Host "Searching for user [$Username]..."
	$userObj = [adsi]("LDAP://CN=$Username,$OU_USER")
	If([string]::IsNullOrEmpty($userObj.Properties.name)) {
		Write-Host "** [$username] not found in Active Directory or an error has occurred"
	} else {
		Write-Host "Searching user's group memberships..."
		$results = @()
		foreach($ssoGroup in $Groups) {
			$ssoGroupName = $ssoGroup
			# Convert group DN to name only if needed
			if($ssoGroupName -imatch "^CN=([^,]+),") {
				$ssoGroupName = $Matches[1]
			}
			if([string]::IsNullOrWhitespace($ssoGroupName)) {
				Write-Warning "Invalid group name [$ssoGroupName]"
			} else {
				$userObj.Properties.memberof | ForEach-Object {
					# Convert group DN to name only if needed
					If($_ -imatch "^CN=$ssoGroupName,") {
						if ($ReturnFullGroup) {
							$results += @($ssoGroup)
						} else {
							$results += @($ssoGroupName)
						}
					}
				}
			}
		}
	}
	
	return $results
}

<#
	.SYNOPSIS
	Add/removes a user from a given group using ADSI.
	
	.DESCRIPTION
	Add/removes a user from a given group using ADSI. Returns true if successful.
    
	.PARAMETER Group
	Required AD group to modify. This should always be full distinguished name.
	
	.PARAMETER Username
	Required username to add or remove.
	
	.PARAMETER Add
	Attempts to add the user to the given group if given.
	
	.PARAMETER Remove
	Attempts to remove the user from the given group if given.
	
	.EXAMPLE
	Modify-ADSIGroupMember -Group HW-HSA-SilverCloudADHoc -Username mcarras8 -Remove
#>
function Modify-ADSIGroupMember {
	param (		
		[parameter(Mandatory=$true, ParameterSetName = 'Add', Position=0)]
		[parameter(Mandatory=$true, ParameterSetName = 'Remove', Position=0)]
		[ValidateScript({-Not [string]::IsNullOrWhitespace($_)})]
		[string]$Group,
		
		[parameter(Mandatory=$true, ParameterSetName = 'Add', Position=1)]
		[parameter(Mandatory=$true, ParameterSetName = 'Remove', Position=1)]
		[ValidateScript({-Not [string]::IsNullOrWhitespace($_)})]
		[string]$Username,
		
		[parameter(Mandatory=$true, ParameterSetName = 'Add')]
		[switch]$Add,
		
		[parameter(Mandatory=$true, ParameterSetName = 'Remove')]
		[switch]$Remove
	)
	
	$results = $false
	
	Write-Host "Searching for user [$Username]..."
	$userObj = [adsi]("LDAP://CN=$Username,$OU_USER")
	If([string]::IsNullOrEmpty($userObj.Properties.name)) {
		Write-Host "** [$username] not found in Active Directory or an error has occurred"
	} else {		
		Write-Host "Searching for group [$Group]..."
		$groupObj = [adsi]("LDAP://$Group")
		If([string]::IsNullOrEmpty($groupObj.Properties.name)) {
			Write-Host "** [$Group] not found in Active Directory or an error has occurred"
		} else {
			If ($Add) {
				$groupObj.member.Add("CN=$Username,$OU_USER")
			} elseif ($Remove) {
				$groupObj.member.Remove("CN=$Username,$OU_USER")
			}
			$groupObj.CommitChanges()
		}
	}
	
	return $results
}

# -- END FUNCTIONS --

# -- START MAIN SCRIPT --

# Do/while loop
$doLoop = $true
while($doLoop) {
	$choice = $null
	Write-Host "** TimelyCare and SilverCloud SSO Access Management **"
	Write-Host "1. Check a user's TimelyCare SSO access"
	Write-Host "2. Check a user's SilverCloud SSO access"
	Write-Host "3. Add a user to TimelyCare ad-hoc SSO access group"
	Write-Host "4. Add a user to SilverCloud ad-hoc SSO access group"
	Write-Host "5. Remove a user from a TimelyCare ad-hoc SSO access group"
	Write-Host "6. Remove a user from a SilverCloud ad-hoc SSO access group"
	Write-Host " "

	$choice = Read-Host "** Enter an option or press enter to exit"
	if ([string]::IsNullOrEmpty($choice)) {
		# BREAK OUT OF LOOP
		$doLoop = $false
		break
	}

	$username = $null
	while([string]::IsNullOrEmpty($username)) {
		Write-Host " "
		$username = Read-Host "** Please enter the user's JHED"
	}
	
	Write-Host " "
	switch($choice) {
		# Check TimelyCare SSO group membership
		"1" { 
			$results = $null
			$results = Check-ADSIUserMembership $username $TIMELYCARE_ADGROUPS
			if($results -ne $false) {
				if($results.Count -eq 0) {
					Write-Host "** [$username] was not found in any of the groups granting SSO access to TimelyCare"
				} else {
					Write-Host "** Access found for [$username]"
					Write-Host ("** User [$username] is in the following groups granting SSO access to TimelyCare: {0}" -f ($results -join ", "))
				}
			}
		}
		
		# Check SilverCloud SSO group membership
		"2" { 
			$results = $null
			$results = Check-ADSIUserMembership $username $SILVERCLOUD_ADGROUPS
			if($results -ne $false) {
				if($results.Count -eq 0) {
					Write-Host "** [$username] was not found in any of the groups granting SSO access to SilverCloud"
				} else {
					Write-Host "** Access found for [$username]"
					Write-Host ("** User [$username] is in the following groups granting SSO access to SilverCloud: {0}" -f ($results -join ", "))
				}
			}
		}
		
		# Add user to TimelyCare SSO ad-hoc group
		"3" {
			$results = $null
			$allgroups = $TIMELYCARE_ADGROUPS
			$adHocGroups = $TIMELYCARE_ADHOC_ADGROUPS
			$results = Check-ADSIUserMembership $username $allgroups
			if($results -ne $false) {
				if($results.Count -ne 0) {
					Write-Host ("** [$username] is already in the following groups granting SSO access to TimelyCare: {0}" -f ($results -join ", "))
				} else {
					Write-Host " "
					$ssoGroups = $adHocGroups
					$count = 1
					foreach($ssoGroup in $ssoGroups) {
						$ssoGroupName = $ssoGroup
						# Convert group DN to name only if needed
						if($ssoGroupName -imatch "^CN=([^,]+),") {
							$ssoGroupName = $Matches[1]
						}
						Write-Host "$count. $ssoGroupName"
						$count++
					}
					
					$choice = $null
					while([string]::IsNullOrEmpty($choice)) {
						$choice = Read-Host "Enter the number of the group to add [$username]"
						# Check whether choice is empty, not an integer, or out of range
						if([string]::IsNullOrEmpty($choice) -Or ($choice -as [int]) -isnot [int] -Or ($choice -as [int]) -ge $count -Or ($choice -as [int]) -le 0) {
							Write-Host "** Invalid entry [$choice]. Try again."
							Write-Host " "
							$choice = $null
						}
					}
					$group = ([array]$ssoGroups)[$choice-1]
					$groupName = $group
					# Convert group DN to name only if needed
					if($groupName -imatch "^CN=([^,]+),") {
						$groupName = $Matches[1]
					}
					Write-Host "** Adding [$username] to [$groupName]"
					$results = $null
					$results = Modify-ADSIGroupMember $group $username -Add
					Write-Host "** Verifying membership has been updated..."
					$results2 = Check-ADSIUserMembership $username @($group)
					if ($results2 -eq $false -Or $results2.Count -eq 0) {
						Write-Warning "Group Modify Failed -- [$username] does not appear to be a member of [$groupName]"
					} else {
						Write-Host "** [$username] added successfully to [$groupName]"
					}
				}
			}
		}
		
		# Add user to SilverCloud SSO ad-hoc group
		"3" {
			$results = $null
			$allgroups = $SILVERCLOUD_ADGROUPS
			$adHocGroups = $SILVERCLOUD_ADHOC_ADGROUPS
			$results = Check-ADSIUserMembership $username $allgroups
			if($results -ne $false) {
				if($results.Count -ne 0) {
					Write-Host ("** [$username] is already in the following groups granting SSO access to SilverCloud: {0}" -f ($results -join ", "))
				} else {
					Write-Host " "
					$ssoGroups = $adHocGroups
					$count = 1
					foreach($ssoGroup in $ssoGroups) {
						$ssoGroupName = $ssoGroup
						# Convert group DN to name only if needed
						if($ssoGroupName -imatch "^CN=([^,]+),") {
							$ssoGroupName = $Matches[1]
						}
						Write-Host "$count. $ssoGroupName"
						$count++
					}
					
					$choice = $null
					while([string]::IsNullOrEmpty($choice)) {
						$choice = Read-Host "Enter the number of the group to add [$username]"
						# Check whether choice is empty, not an integer, or out of range
						if([string]::IsNullOrEmpty($choice) -Or ($choice -as [int]) -isnot [int] -Or ($choice -as [int]) -ge $count -Or ($choice -as [int]) -le 0) {
							Write-Host "** Invalid entry [$choice]. Try again."
							Write-Host " "
							$choice = $null
						}
					}
					$group = ([array]$ssoGroups)[$choice-1]
					$groupName = $group
					# Convert group DN to name only if needed
					if($groupName -imatch "^CN=([^,]+),") {
						$groupName = $Matches[1]
					}
					Write-Host "** Adding [$username] to [$groupName]"
					$results = $null
					$results = Modify-ADSIGroupMember $group $username -Add
					Write-Host "** Verifying membership has been updated..."
					$results2 = Check-ADSIUserMembership $username @($group)
					if ($results2 -eq $false -Or $results2.Count -eq 0) {
						Write-Warning "Group Modify Failed -- [$username] does not appear to be a member of [$groupName]"
					} else {
						Write-Host "** [$username] added successfully to [$groupName]"
					}
				}
			}
		}
		
		# Remove a user from a TimelyCare SSO ad-hoc group
		"5" {
			$results = $null
			$allgroups = $TIMELYCARE_ADGROUPS
			$adHocGroups = $TIMELYCARE_ADHOC_ADGROUPS
			$results = Check-ADSIUserMembership $username $allgroups -ReturnFullGroup
			if($results -ne $false) {
				if($results.Count -eq 0) {
					Write-Host ("** [$username] not found in any TimelyCare SSO access groups")
				} else {
					Write-Host " "
					$ssoGroups = $results
					$count = 1
					foreach($ssoGroup in $ssoGroups) {
						$ssoGroupName = $ssoGroup
						# Convert group DN to name only if needed
						if($ssoGroupName -imatch "^CN=([^,]+),") {
							$ssoGroupName = $Matches[1]
						}
						Write-Host "$count. $ssoGroupName"
						$count++
					}
					
					$choice = $null
					while([string]::IsNullOrEmpty($choice)) {
						$choice = Read-Host "Enter the number of the group which should not contain [$username]"
						# Check whether choice is empty, not an integer, or out of range
						if([string]::IsNullOrEmpty($choice) -Or ($choice -as [int]) -isnot [int] -Or ($choice -as [int]) -ge $count -Or ($choice -as [int]) -le 0) {
							Write-Host "** Invalid entry [$choice]. Try again."
							Write-Host " "
							$choice = $null
						}
					}
					$group = ([array]$ssoGroups)[$choice-1]
					$groupName = $group
					# Convert group DN to name only if needed
					if($groupName -imatch "^CN=([^,]+),") {
						$groupName = $Matches[1]
					}
					Write-Host "** Removing [$username] from [$groupName]"
					$results = $null
					$results = Modify-ADSIGroupMember $group $username -Remove					
					Write-Host "** Verifying membership has been updated..."
					$results2 = Check-ADSIUserMembership $username @($group)
					if ($results2.Count -ge 1) {
						Write-Warning "Group Modify Failed -- [$username] is still a member of [$groupName]"
					} else {
						Write-Host "** [$username] removed successfully from [$groupName]"
					}
				}
			}
		}
		
		# Remove a user from a SilverCloud SSO ad-hoc group
		"6" {
			$results = $null
			$allgroups = $SILVERCLOUD_ADGROUPS
			$adHocGroups = $SILVERCLOUD_ADHOC_ADGROUPS
			$results = Check-ADSIUserMembership $username $allgroups -ReturnFullGroup
			if($results -ne $false) {
				if($results.Count -eq 0) {
					Write-Host ("** [$username] not found in any SilverCloud SSO access groups")
				} else {
					Write-Host " "
					$ssoGroups = $results
					$count = 1
					foreach($ssoGroup in $ssoGroups) {
						$ssoGroupName = $ssoGroup
						# Convert group DN to name only if needed
						if($ssoGroupName -imatch "^CN=([^,]+),") {
							$ssoGroupName = $Matches[1]
						}
						Write-Host "$count. $ssoGroupName"
						$count++
					}
					
					$choice = $null
					while([string]::IsNullOrEmpty($choice)) {
						$choice = Read-Host "Enter the number of the group which should not contain [$username]"
						# Check whether choice is empty, not an integer, or out of range
						if([string]::IsNullOrEmpty($choice) -Or ($choice -as [int]) -isnot [int] -Or ($choice -as [int]) -ge $count -Or ($choice -as [int]) -le 0) {
							Write-Host "** Invalid entry [$choice]. Try again."
							Write-Host " "
							$choice = $null
						}
					}
					$group = ([array]$ssoGroups)[$choice-1]
					$groupName = $group
					# Convert group DN to name only if needed
					if($groupName -imatch "^CN=([^,]+),") {
						$groupName = $Matches[1]
					}
					Write-Host "** Removing [$username] from [$groupName]"
					$results = $null
					$results = Modify-ADSIGroupMember $group $username -Remove					
					Write-Host "** Verifying membership has been updated..."
					$results2 = Check-ADSIUserMembership $username @($group)
					if ($results2.Count -ge 1) {
						Write-Warning "Group Modify Failed -- [$username] is still a member of [$groupName]"
					} else {
						Write-Host "** [$username] removed successfully from [$groupName]"
					}
				}
			}
		}
				
		default {
			Write-Host "** Invalid option [$choice]"
		}
	}
	
	Write-Host " "
}

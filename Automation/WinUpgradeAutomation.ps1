<#
	.SYNOPSIS
	Sends out notifications for required Windows 11 upgrades for AD computer reporting lower than $EOLVER, outputs reports, and manages the group for toast notifications.
	
	.DESCRIPTION
	Sends out notifications for required Windows 11 upgrades for AD computer reporting lower than $EOLVER, outputs reports, and manages the group for toast notifications.

	.PARAMETER ManageGroupOnly
	Only manage the GPO filtering notification group. Do not send any emails.
	
	.PARAMETER DryRun
	Only output results. Do not send emails or make any other changes (-WhatIf).
	
	.PARAMETER DebugEmailOverride
	Send an email to the given address instead of the contact user. This overrides -DryRun for emailing.
	
	.PARAMETER DebugEmailLimit
	Limit the number of emails sent (Default: 1).
	
	.PARAMETER Verbose
	Enable additional verbose/debugging output.
	
	.NOTES
	Requirements:
	* RSAT AD Tools
	
	EOLVER and UPGRADEVER must be updated periodically in this script. UPGRADEVER is only used in the report names.
	
	Email will be sent out to the assigned user if its valid (contains "@"). If it doesn't exist, then the first USS staff member not a member of USS IT will be used instead, searching LastLogonUser then Primary Users in that order.
	
	Created: 3-14-25
	Author: mcarras8
	
	Changelog
	06-23-25 - mcarras8 - Switched back to 23H2. Other tweaks. Ready for prod.
	03-26-25 - mcarras8 - Fixed Group management. Added Department to output and -Verbose support. Other fixes/tweaks.
	03-14-25 - mcarras8 - Initial upload.
#>
param(
	[Parameter(Mandatory=$false)]
	[switch]$ManageGroupOnly,
	
	[Parameter(Mandatory=$false)]
	[switch]$DryRun,
	
	[Parameter(Mandatory=$false)]
	[string]$DebugEmailOverride,
	
	[Parameter(Mandatory=$false)]
	[string]$DebugEmailLimit=1
)
		
# -- START CONFIGURATION --
# This should be the previous version before the latest build version.
$EOLVER=22621		# Windows 11, 22H2
# This should be the version offered in the upgrade.
$UPGRADEVER=22631	# Windows 11, 23H2
# $UPGRADEVER=26100   # Windows 11, 24H2
# The searchbase to search for matching computers in AD.
$SEARCHBASE = "OU=Computers,OU=USS,DC=win,DC=ad,DC=jhu,DC=edu"
# Array of OUs to always exclude
# Note disabled systems are excluded automatically.
$EXCLUDE_OUS = @()
# AD Computer Properties/Attributes used in logic. Assumes the following:
# extensionAttribute1 = asset tag
# extensionAttribute2 = assigned user by userprinciplname (used first for "Contact User")
# extensionAttribute5 = System Form Factor (Laptop, Tablet, etc.)
# extensionAttribute10 = LastLogonUser from query or SCCM export (used second for "Contact User" if no assigned user)
# extensionAttribute3 = Primary Users from SCCM export, deliminated by semi-colon (used last for "Contact User" if nothing else matches)
# extensionAttribute8 = Last successful sync with Snipe-It
$COMP_PROPS = @{
	"operatingsystemversion"="operatingsystemversion"
	"LastLogonDate"="LastLogonDate"
	"AssetTag"="extensionAttribute1"
	"AssignedUser"="extensionAttribute2"
	"PrimaryUsers"="extensionAttribute3"
	"LastLogonUser"="extensionAttribute10"
	"FormFactor"="extensionAttribute5"
	"LastSyncDate"="extensionAttribute8"
}
# Array of company/companies users must be a part of when searching for alternate contacts in the LastLogonUser or PrimaryUsers fields.
$CONTACTUSER_COMPANIES = @("JHU; University Student Services")
# If given, contact users must be a member of the given groups to be considered valid. This includes the assigned user field.
# If this is set, it should include the main sync group for the SOR.
$CONTACTUSER_INCLUDE_GROUPS = @("USS-IT-SnipeItUsersAll")
# If given, these groups will be excluded from contact / notifications.
# $CONTACTUSER_EXCLUDE_GROUPS = @("USS-VIP")
# Array of group(s) containing members of IT.
# They will only be contacted if there's no other valid contact users.
$CONTACTUSER_EXCLUDE_ITGROUPS = @("USS-IT-JHEDs")
# Always exclude these contact users.
$CONTACTUSER_EXCLUDE = @("local_users")
# Exclude contact users matching patterns.
$CONTACTUSER_EXCLUDE_REGEX = "SC\-"
# User domain if not set. This is the domain appended for all AD lookups and emails (if needed).
# Some attributes like LastLogonUser and PrimaryUsers won't have domain.
$USER_DOMAIN = "@jh.edu"
# OU containing all enabled users.
$USER_OU = "OU=PEOPLE,DC=win,DC=ad,DC=jhu,DC=edu"
# If set, only display asset tag if its numeric.
$ASSET_TAG_IS_NUMERIC=$true
# Only send email notifications when the system has a valid assigned user.
$CONTACT_ONLY_WHEN_ASSIGNED = $false
# Group for Toast Notification GPO
$NOTIFICATION_GROUP = "USS-GPO-Win11UpgradeToast"
# Add a warning if LastLogonDate is X number of days ago.
$STALE_PC_DAYS = 30
# The System-Of-Record URL to display in the exports, appending the asset tag #.
$SORURL_ASSETTAG = "https://jh-uss.snipe-it.io/hardware/bytag?assetTag="

# Email settings.
$EMAIL_SMTP = 'smtp.johnshopkins.edu'
$EMAIL_FROM = 'Jerome.Powell@jhu.edu'
# CC a copy of each email to the given address (make sure rules are enabled).
$EMAIL_CC = @('Jerome.Powell@jhu.edu','mcarras8@jhu.edu')
# Optional BCC.
$EMAIL_BCC = 'ussitservices@jhu.edu'
$EMAIL_SUBJECT = "[USS-IT] Windows 11 Upgrade Required"
# This is the first part of each email. Allows for HTML.
# Each email will be in the format of either:
# $EMAIL_INTRO_HTML - where {0} is replaced by ", SYSTEM_NAME (asset #), "
# $EMAIL_FOOTER_HTML
# ...or, if the user has multiple systems...
# $EMAIL_INTRO_HTML
# <Table with system info>
# <System list items>
# $EMAIL_FOOTER_HTML
$EMAIL_INTRO_HTML = @"
<p>This is an automated message.</p>
<p>You are receiving this email because your system{0} is currently running Windows 10 or an older version of Windows 11, and it needs to be upgraded.</p>
 
<p>Starting on <b>August 12th</b>, Central IT will block updates for these versions. Please install the new version yourself by following the instructions in the follow link: <a href="https://t.jh.edu/USS-WindowsUpgrade">https://t.jh.edu/USS-WindowsUpgrade</a>.</p>

<p>If you encounter any issues with the update, please <a href="https://johnshopkins.service-now.com/serviceportal?id=report_problem&sys_id=3f1dd0320a0a0b99000a53f7604a2ef9">open a helpdesk ticket</a>.</p>
 
<p>Thank you for your cooperation.</p>
"@
# Version of the email intro that will only be used if all the systems are desktops/aios.
# If blank, it will use the default intro text for all systems.
# {1} will be replaced by the date set in $UPGRADE_DESKTOP_REQUIRED_DATE.
<#
$EMAIL_INTRO_HTML_DESKTOPS = @"<p>This is an automated message.</p>
<p>You are receiving this email because your system{0} is currently running Windows 10 or an older version of Windows 11, and it needs to be upgraded. Your system has been set to upgrade automatically on or after <b>{1}</b>.</p>

<p>If you'd like to upgrade sooner you may follow the instructions in the following link: <a href="https://t.jh.edu/USS-WindowsUpgrade">https://t.jh.edu/USS-WindowsUpgrade</a>.</p>

<p>If you encounter any issues with the update, please <a href="https://johnshopkins.service-now.com/serviceportal?id=report_problem&sys_id=3f1dd0320a0a0b99000a53f7604a2ef9">open a helpdesk ticket</a>.</p>
 
<p>Thank you for your cooperation.</p>
"@
#>
# If given, emails will reference this date for automatic upgrades for desktops/AIOs.
# $UPGRADE_DESKTOP_REQUIRED_DATE = "5-25-2025"
# Footer that will be displayed in the email, after the system table and notes sections. This can be blank.
$EMAIL_FOOTER_HTML = ""
# If set, only show a detailed table if there's more than one referenced system in the email.
$EMAIL_DETAILED_TABLE_MULTIPLE_SYSTEMS_ONLY = $true
# If set, do not email referencing only desktops/aios.
$EMAIL_USERS_WITH_ONLY_DESKTOPS_AIOS = $false
# If set, show more detailed notes (default: false).
$EMAIL_DETAILED_NOTES = $false
# Amount of time in seconds to sleep between emails.
$EMAIL_SLEEP_SECS = 8
# Number of successful emails to send before sleeping longer (e.g. 10 for every 10 emails).
# The $EMAIL_SLEEP_EXTRA_SECS will also be used if any emails fail.
$EMAIL_SLEEP_EXTRA_MOD=10
$EMAIL_SLEEP_EXTRA_SECS = 30

# Email a report at the end.
$EMAIL_REPORT_FROM = 'USS IT Services <ussitservices@jhu.edu>'
#$EMAIL_REPORT_TO = 'ussitservices@jhu.edu'
$EMAIL_REPORT_TO = @("USS-IT-JHEDs")
$EMAIL_REPORT_SUBJECT = "Weekly Results from Windows 11 Upgrade Campaign"

# Path to systems which report being ineligible for upgrade.
$IMPORT_INELIGIBLE_SYSTEMS_PATH = "\\win.ad.jhu.edu\cloud\hsa$\ITServices\Reports\WinUpgrade\win11_incompatible_systems_3_14_25.csv"
# Whether to skip emailing/notifying ineligible systems (and include them in skip report).
$SKIP_INELIGIBLE_SYSTEMS = $true

# Path to save reports to.
$EXPORT_ALL_SYSTEMS_PATH = "\\win.ad.jhu.edu\cloud\hsa$\ITServices\Reports\WinUpgrade\win11_${UPGRADEVER}_upgrade_systems.csv"
$EXPORT_SKIPPED_SYSTEMS_PATH = "\\win.ad.jhu.edu\cloud\hsa$\ITServices\Reports\WinUpgrade\win11_${UPGRADEVER}_upgrade_systems_skipped.csv"
$EXPORT_PROCESSED_SYSTEMS_PATH = "\\win.ad.jhu.edu\cloud\hsa$\ITServices\Reports\WinUpgrade\win11_${UPGRADEVER}_upgrade_systems_with_contact.csv"
$EXPORT_EMAILED_SYSTEMS_PATH = "\\win.ad.jhu.edu\cloud\hsa$\ITServices\Reports\WinUpgrade\win11_${UPGRADEVER}_upgrade_systems_emailed.csv"

# Path and prefix for the Start-Transcript logfiles.
$LOGFILE_PATH = "\\win.ad.jhu.edu\cloud\hsa$\ITServices\Reports\Logs\WinUpgradeAutomation"
$LOGFILE_PREFIX = "winupgradeautomation"
# Maximum number of days before rotating logfile.
$LOGFILE_ROTATE_DAYS = 180
# -- END CONFIGURATION --

# -- FUNCTION START --
$_ADUSERS=@{}
function Get-ADUserCached {
	<#
		.SYNOPSIS
		Returns AD user object for the given username/identity.
		
		.DESCRIPTION
		Returns AD user object for the given username/identity.
		
		.PARAMETER User
		The AD user name or identity.
		
		.PARAMETER Domain
		Optional Domain to append if needed for caching purposes.
		
		.PARAMETER Properties
		Optional properties to return (default: Company,Department).
		
		.OUTPUTS
		The AD user object.
		
		.NOTES
		Saves a cache to $_ADUSERS.
	#>

	param(
		[Parameter(Mandatory=$true,Position=0)]
		[ValidateNotNullOrEmpty()]
		[string]$User,
		
		[Parameter(Mandatory=$false,Position=1)]
		[string]$Domain,
		
		[Parameter(Mandatory=$false,Position=2)]
		[string[]]$Properties=@("Company","Department")
	)
	
	$UPN = $User
	$isDN = $User -like "CN=*"
	if (-Not [string]::IsNullOrEmpty($Domain) -And -Not $isDN -And $UPN -notmatch "@") {
		$UPN += $Domain
	}
	$u = $_ADUSERS.$UPN
	if ([string]::IsNullOrEmpty($u.distinguishedname)) {
		try {
			# If the given user is a distinguishedname, use that instead.
			if ($isDN) {
				$u = Get-ADUser $UPN -Properties $Properties
			} else {
				$u = Get-ADUser -LDAPFilter "(|(SamAccountName=$UPN)(UserPrincipalName=$UPN))" -Properties $Properties			
			}
			$_ADUSERS[$UPN] = $u
		} catch {
			throw $_
		}
	}
	return $u
}

function Get-ADUsersByGroup {
	<#
		.SYNOPSIS
		Collect all AD users from given target group(s), filtering the results.
		
		.DESCRIPTION
		Collect all AD users from given target group(s), filtering the results. If you want to check all users give a global group like Domain Users.
		
		.PARAMETER TargetGroup
        Required. The AD Group(s) to check.
		
		.PARAMETER ADProperties
        The AD properties to return with each user.
		
		.PARAMETER ADPropertyFilter
        A filterscript to use on the results. Use backticks for property references. E.g. "`$_.distinguishedname -like '*,OU=Users,*'"
		
		.PARAMETER Nested
		Will recurse over groups if given. This may take a while with large groups.
		
        .PARAMETER IncludeDisabled
        If true include disabled users.
		
		.PARAMETER ExitOnError
		Exit on error fetching group membership.
		
		.PARAMETER RecurseLoopCount
		This is used when the function is called recursively.
		
		.OUTPUTS
		The returned users from AD.
		
		.Example
		PS> Get-ADUsersByGroup "Domain Users" -ADProperties @("department","company","title","manager")
	#>
	param (		
		[parameter(Mandatory=$true,
					Position = 0,
					ValueFromPipeline = $true,
					ValueFromPipelineByPropertyName=$true)]
		[string[]]$TargetGroup,
		
		[parameter(Mandatory=$false)]
        [AllowEmptyCollection()]
		[string[]]$ADProperties = @("givenname","surname","department","company","title","manager","physicaldeliveryofficename","mail"),
		
		[parameter(Mandatory=$false)]
		[string]$ADPropertyFilter,
		
		[parameter(Mandatory=$false)]
		[switch]$Nested,

        [parameter(Mandatory=$false)]
		[switch]$IncludeDisabled,
		
		[parameter(Mandatory=$false)]
		[switch]$ExitOnError,
		
		[parameter(Mandatory=$false)]
		[int]$RecurseLoopCount=0
	)
	
	$ad_users = $null
	$props = $ADProperties
	if ($props -ne $null -And -Not $props -is [array]) {
		$props = @($props)
	}
	# We'll use the memberof property to determine if we already got this user.
	$props += @("distinguishedname","memberof") | Select -Unique
	Write-Debug "[Get-ADUsersByGroup] Properties: $props"
		
	foreach ($group in $TargetGroup) {
		# Get all users from AD
		Write-Verbose ("[Get-ADUsersByGroup] Collecting all users from AD group [$group] (Nested=$Nested, With Filter={0})..." -f (-not [string]::IsNullOrEmpty($ADPropertyFilter)))
		
		if ($Nested) {
			try {
				# May not work with >5000 results
				$ad_users += Get-ADGroupMember $group -Recursive -ErrorAction Stop | where {$_.objectClass -eq 'user'}
			} catch [System.TimeoutException],[TimeoutException] {
				Write-Warning ("[Get-ADUsersByGroup] Timeout detected. Trying again, recursing over each member. Please wait...")
				# If we have a timeout, try again recursing over each nested group found.
				# If we have a very high recurse count, assume we're in an infinite loop and throw an error.
				if ($RecurseLoopCount -gt 20) {
					$errorMsg = "Recurse count is too high ($RecurseLoopCount), may be infinite loop, aborting"
					if ($ExitOnError) {
						Write-Error $errorMsg
						exit -1
					} else {
						throw $errorMsg
					}
				}
				try {
					# Manually recurse over nested groups.
					# An alternative is using LDAP_MATCHING_RULE_IN_CHAIN, but it's quite slower.
					# Get the group info.
					$adgroup = Get-ADGroup $group
					# Get all user members of this group.
					$childUsers = Get-ADUser -LDAPFilter "(&(objectCategory=user)(samAccountName=*)(memberOf:=$($adgroup.distinguishedname)))" -Properties $props -ErrorAction Stop
					Write-Debug("[Get-ADUsersByGroup] [group=$group] Found $($childUsers.Count) users")
					# Get all nested groups.
					$childGroups = Get-ADGroup -LDAPFilter "(&(objectCategory=group)(samAccountName=*)(memberOf:=$($adgroup.distinguishedname)))" -ErrorAction Stop | Select -ExpandProperty Name
					Write-Debug("[Get-ADUsersByGroup] [group=$group] Found $($childGroups.Count) groups")
					# Call this function recursively for all groups found.
					if (($childGroups | Measure-Object).Count -gt 0) {
						$ad_users += Get-ADUsersByGroup -TargetGroup $childGroups -ADProperties $ADProperties -Nested -IncludeDisabled:$IncludeDisabled -ExitOnError:$ExitOnError -RecurseLoopCount ($RecurseLoopCount + 1)
					}
				} catch {
					if ($ExitOnError) {
						Write-Error $_
						exit -1
					} else {
						throw
					}
				}
			} catch {
				if ($ExitOnError) {
					Write-Error $_
					exit -1
				} else {
					throw
				}
			}
		} else {
			# No nested groups.
			try {
				$adgroup = Get-ADGroup $group
				$ad_users += Get-ADUser -LDAPFilter "(&(objectCategory=user)(samAccountName=*)(memberOf:=$($adgroup.distinguishedname)))" -Properties $props -ErrorAction Stop
			} catch {
				if ($ExitOnError) {
					Write-Error $_
					exit -1
				} else {
					throw
				}
			}
		}
	}
    if ($ad_users -ne $null) {		
		# Get extra attributes for each user
		Write-Verbose ("[Get-ADUsersByGroup] Getting properties for {0} users..." -f ($ad_users | Measure).Count)
		# Make sure to dedupe users here.
		# Only fetch the user if they are missing the "memberof" property
		try {
			$ad_users = $ad_users | Sort distinguishedname -Unique | foreach { if($_.memberof -ne $null) { $_ } else { Get-ADUserCached $_.distinguishedname -Properties $props } }
		} catch {
			if ($ExitOnError) {
				Write-Error $_
				exit -1
			} else {
				throw
			}
		}
		Write-Debug("[Get-ADUsersByGroup] {0} users after removing duplicates and calling Get-ADUserCached for properties" -f ($ad_users | Measure).Count)
		
		$filterscript = $ADPropertyFilter
		if (-Not $IncludeDisabled) {
			if (-Not [string]::IsNullOrWhitespace($filterscript)) {
				$filterscript += ' -AND '
			}
			$filterscript += "`$_.Enabled -eq `$true"
		}
	    Write-Debug "[Get-ADUsersByGroup] AD Group Filter: $filterscript"
	    if (-Not [string]::IsNullOrWhitespace($filterscript)) {
		    $ad_users = $ad_users | Where-Object -FilterScript ([scriptblock]::create($filterscript))
	    }
    }
	Write-Verbose ("[Get-ADUsersByGroup] Total filtered AD users collected: {0}" -f $ad_users.Count)
	
	return $ad_users
}
# -- FUNCTION END --

# -- START --
$dateStart = Get-Date
$error_count = 0
$_scriptName = split-path $PSCommandPath -Leaf

# Rotate log files
if ($LOGFILE_ROTATE_DAYS -is [int] -And $LOGFILE_ROTATE_DAYS -gt 0) {
	Get-ChildItem "${LOGFILE_PATH}\${LOGFILE_PREFIX}_*.log" | Where-Object { $_.CreationTime -lt (Get-Date).AddDays(-$LOGFILE_ROTATE_DAYS) } | Remove-Item -Force
}

# Start logging
$_logfilepath = "${LOGFILE_PATH}\${LOGFILE_PREFIX}_$(get-date -f yyyy-MM-dd)"
try {
	$_logfilepath = "${_logfilepath}.log"
	Start-Transcript -Path $_logfilepath -Append
} catch {
	# If we get any error, try again with .1 appended in case it's a file lock.
	$_logfilepath = "${_logfilepath}.1.log"
	Start-Transcript -Path $_logfilepath -Append
}

# List parameters given.
if($ManageGroupOnly) {
	Write-Host("[{0}] -ManageGroupOnly switch given. Emails will not be sent." -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"))
}
if($DryRun) {
	Write-Host("[{0}] -DryRun switch given. Emails will not be sent (unless -DebugEmailOverride is given) and groups will not be managed. Reports will still be exported." -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"))
}
if(-Not [string]::IsNullOrEmpty($DebugEmailOverride)) {
	Write-Host("[{0}] -DebugEmailOverride [$DebugEmailOverride] with limit of [$DebugEmailLimit]" -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"))
}

# Get enabled computers matching $eolver or lower.
$_props = $COMP_PROPS.Values | % { $_ }
$comps = Get-ADComputer -Searchbase $searchbase -Filter {Enabled -eq $true} -Properties $_props | where {$_.OperatingSystemVersion -match "10.0 \((\d+)\)" -and $Matches.1 -ne $null -and ($Matches.1 -as [int]) -is [int] -And $Matches.1 -le $eolver -And $_.distinguishedname -notin $EXCLUDE_OUS}

# Convert the hashtable map into a dynamic select array before exporting.
$selectarray = @("distinguishedname")
foreach ($k in $COMP_PROPS.Keys) {
	$selectarray += @(@{"N"=$k; "Expression"=[Scriptblock]::Create("`$_.'$($COMP_PROPS.$k)'")})
}
$comps | Select $selectarray | Export-CSV -NoTypeInformation -Force $EXPORT_ALL_SYSTEMS_PATH
Write-Host("[{0}] Exported {1} systems requiring Windows 11 upgrade collected from AD to [{2}]." -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"), ($comps | Measure).Count, $EXPORT_ALL_SYSTEMS_PATH)

# Get excluded contact users.
$adUserBlacklist = $null
if (($CONTACTUSER_EXCLUDE_GROUPS | Measure).Count -gt 0) {
	$adUserBlacklist = Get-ADUsersByGroup $CONTACTUSER_EXCLUDE_GROUPS -Nested -Verbose | Select -Unique
}
$adUserBlacklistCount = ($adUserBlacklist | Measure).Count
$itUsers = $null
if (($CONTACTUSER_EXCLUDE_ITGROUPS | Measure).Count -gt 0) {
	$itUsers = Get-ADUsersByGroup $CONTACTUSER_EXCLUDE_ITGROUPS -Nested -Verbose | Select -Unique
}
$itUsersCount = ($itUsers | Measure).Count

# Get included contact users.
# These should be DN only.
$adUserWhitelist = $null
$adUserWhitelistCount = 0
if (($CONTACTUSER_INCLUDE_GROUPS | Measure).Count -gt 0) {
	Write-Host("[{0}] Collecting AD group members to include from {1}..." -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"), ($CONTACTUSER_INCLUDE_GROUPS -join ", "))
	$aduserWhitelist = Get-ADUsersByGroup $CONTACTUSER_INCLUDE_GROUPS -Nested -Verbose
	$adUserWhitelistCount = ($aduserWhitelist | Measure).Count
}

# Import a previously exported list of all incompatible systems.
# We'll check against this and set the IsIncompatible flag if it's listed here.
$incompatible_systems = $null
if((Test-Path $IMPORT_INELIGIBLE_SYSTEMS_PATH -PathType Leaf)) {
	$incompatible_systems = Import-CSV $IMPORT_INELIGIBLE_SYSTEMS_PATH
}

# Collect all systems for each contact user.
$skipped_systems = @()
$counter = 0
$processed_systems = foreach($comp in $comps) {
	$contactuser = $comp.($COMP_PROPS.AssignedUser)
	$aduser = $null
	$is_sharedsystem = $false
	$is_assigned = $true
	$is_incompatible = ($comp.Name -in $incompatible_systems.Name)
	$lastlogondate = $comp.($COMP_PROPS.LastLogonDate)
	if ($lastlogondate -and ((Get-Date) - $lastlogondate).Days -gt $STALE_PC_DAYS) {
		$is_stale = $true
	} else {
		$is_stale = $false
	}
	$assettag = $comp.($COMP_PROPS.AssetTag)
	$link_url = $SORURL_ASSETTAG + $assettag
	$skip_reason = $null
	Write-Verbose("[{0}] [{1}] Initial contact user: {2}" -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"), $comp.Name, $contactuser)
	if ($contactuser -match "@") {
		try {
			$aduser = Get-ADUserCached -User $contactuser -Domain $USER_DOMAIN
			Write-Verbose("[{0}] [{1}] Got back AD user info: DN={2}, Enabled={3}" -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"), $comp.Name, $aduser.distinguishedname, $aduser.Enabled)
			if ($adUserWhitelistCount -gt 0 -And $aduser.distinguishedname -notin $aduserWhitelist.distinguishedname) {
				Write-Warning("[{0}] [{1}] Contact user [{2}] not found in user groups" -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"),  $comp.Name, $contactuser)
				$contactuser = $null
			}
		} catch {
			Write-Error $_
			$error_count++
		}
	} else {
		if ([string]::IsNullOrEmpty($contactuser) -And $CONTACT_ONLY_WHEN_ASSIGNED) {
			$skip_reason = "No Assigned User"
			Write-Warning("[{0}] [{1}] - SKIPPING: No assigned user and CONTACT_ONLY_WHEN_ASSIGNED is set." -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"), $comp.Name)
		} else {
			if ([string]::IsNullOrEmpty($contactuser)) {
				Write-Host("[{0}] [{1}] - No valid assigned user [{2}]. Checking other attributes. System Is Stale: {3}" -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"), $comp.Name, $contactuser, $is_stale)
				$is_assigned = $false
			} else {
				# If the system is assigned to a user who doesn't have a valid username then assume its a shared system.
				$is_sharedsystem = $true
			}
			# Check LastLastLogonUser if invalid AssignedUser
			# This should always be samaccountname.
			$contactuser = $comp.($COMP_PROPS.LastLogonUser)
			$is_contactuser_it = $false
			if ([string]::IsNullOrWhitespace($contactuser) -Or $contactuser -in $CONTACTUSER_EXCLUDE -Or $contactuser -match $CONTACTUSER_EXCLUDE_REGEX) {
				$contactuser = $null
				$aduser = $null
				Write-Verbose("[{0}] [{1}] Discarding LastLogonUser [{2}] - Empty or in exclusion list" -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"), $comp.Name, $contactuser)
			} else {
				# Double-check user is valid by checking AD.
				# Basically, if a user is disabled or outside of the default $USER_OU then we assume they're not a valid "User".
				# Also check the user's company to see if they match our supported companies.
				# Finally, check if the user is a member of IT.
				try {
					$aduser = Get-ADUserCached -User $contactuser -Domain $USER_DOMAIN
					Write-Verbose("[{0}] [{1}] Got back AD user info for LastLogonUser: DN={2}, Company={3}, Enabled={4}" -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"), $comp.Name, $aduser.distinguishedname, $aduser.Company, $aduser.Enabled)
					if (-Not $aduser.Enabled -Or $aduser.distinguishedname -ne "CN=$contactuser,$USER_OU") {
						Write-Verbose("[{0}] [{1}] Discarding LastLogonUser [{2}] - Not enabled or invalid OU" -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"), $comp.Name, $contactuser)
						$contactuser = $null
						$aduser = $null
					} elseif (-Not [string]::IsNullOrWhitespace($aduser.Company) -And ($CONTACTUSER_COMPANIES | Measure).Count -gt 0 -And $aduser.Company -notin $CONTACTUSER_COMPANIES) {
						Write-Verbose("[{0}] [{1}] Discarding LastLogonUser [{2}] - Invalid company [{3}]" -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"), $comp.Name, $contactuser, $aduser.Company)
						$contactuser = $null
					} else {
						if ($adUserWhitelistCount -gt 0 -And $aduser.distinguishedname -notin $adUserWhitelist.distinguishedname) {
							Write-Verbose("[{0}] [{1}] Discarding LastLogonUser [{2}] - Not found in user groups" -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"),  $comp.Name, $contactuser)
							$contactuser = $null
						} elseif ($aduser.distinguishedname -ne $null) {
							$is_contactuser_it = $aduser.distinguishedname -in $itUsers.distinguishedname
						}
					}
				} catch {
					Write-Error $_
					$error_count++
				}
			}
			# Check primary users if we don't have a valid contact from LastLogonUser
			# Exclude users not matching criteria, including those not matching a valid company
			# If the Primary User is a member of IT the logic will fall back to use the LastLogonUser regardless
			if ([string]::IsNullOrWhitespace($contactuser) -Or $is_contactuser_it) {
				# PrimaryUsers field is delimited by "; "
				# This should always be samaccountname.
				foreach ($u in ($comp.($COMP_PROPS.PrimaryUsers) -split "; ")) {
					# Remove the domain from each user
					$user = $u -replace "[^\\]+\\",""
					Write-Verbose("[{0}] [{1}] Checking primary user [{2}]" -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"), $comp.Name, $user)
					if (-Not [string]::IsNullOrWhitespace($user) -and $user -notin $CONTACTUSER_EXCLUDE -and $user -notmatch $CONTACTUSER_EXCLUDE_REGEX) {
						try {
							$aduserTemp = Get-ADUserCached -User $user -Domain $USER_DOMAIN
							Write-Verbose("[{0}] [{1}] Got back AD user info: DN={2}, Company={3}, Enabled={4}" -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"), $comp.Name, $aduserTemp.distinguishedname, $aduserTemp.Company, $aduserTemp.Enabled)
							# Exclude IT Users, if set.
							if ($itUsersCount -le 0 -Or $aduserTemp.distinguishedname -notin $itUsers.distinguishedname) {
								if ($aduserTemp.Enabled -And $aduserTemp.distinguishedname -eq "CN=$user,$USER_OU") {
									# Check user whitelist, if set.
									if ($adUserWhitelistCount -gt 0 -And $aduserTemp.distinguishedname -notin $adUserWhitelist.distinguishedname) {
										Write-Verbose("[{0}] [{1}] Discarding primary user [{2}] - Not found in user groups whitelist" -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"),  $comp.Name, $user)
									# Check company, if set.
									} elseif (($CONTACTUSER_COMPANIES | Measure).Count -eq 0 -Or [string]::IsNullOrWhitespace($aduserTemp.Company) -Or $aduserTemp.Company -in $CONTACTUSER_COMPANIES) {
										# If everything else is valid save this user.
										$contactuser = $user
										$aduser = $aduserTemp
										break
									} else {
										Write-Verbose("[{0}] [{1}] Discarding primary user [{2}] - Invalid company [{3}]" -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"), $comp.Name, $user, $aduserTemp.Company)
									}
								} else {
									Write-Verbose("[{0}] [{1}] Discarding primary user [{2}] - Not enabled or invalid OU" -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"), $comp.Name, $user)
								}
							}
						} catch {
							Write-Error $_
							$error_count++
						}
					}
				}
			}
		}
	}
	# Only continue if we have a valid contact user
	if ([string]::IsNullOrWhitespace($contactuser)) {
		$skip_reason = "No Valid Contact User"
		Write-Warning("[{0}] [{1}] - SKIPPING: No Valid Contact User" -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"), $comp.Name)
	# Check blacklist of users to exclude from processing.
	} elseif ($adUserBlacklistCount -gt 0 -And $aduser.distinguishedname -ne $null -And $aduser.distinguishedname -in $aduserBlacklist.distinguishedname) {
		$skip_reason = "Excluded User"
		Write-Warning("[{0}] [{1}] - SKIPPING: Assigned contact user is in excluded user groups (blacklist)" -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"), $comp.Name)
	# Check if we need to skip system due to being ineligible for upgrade.
	} elseif ($is_incompatible -And $SKIP_INELIGIBLE_SYSTEMS) {
		$skip_reason = "Incompatible System"
		Write-Warning("[{0}] [{1}] - SKIPPING: System is in incompatible system list" -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"), $comp.Name)
	}
	
	# We skipped this system.
	If (-not [string]::IsNullOrEmpty($skip_reason)) {
		$skipped_systems += @(
			[PSCustomObject]@{
				Name=$comp.Name
				AssetTag=$comp.($COMP_PROPS.AssetTag)
				AssignedUser=$comp.($COMP_PROPS.AssignedUser)
				FormFactor=$comp.($COMP_PROPS.FormFactor)
				LastLogonDate=$lastlogondate
				IsIncompatible=$is_incompatible
				IsShared=$is_sharedsystem
				IsAssigned=$is_assigned
				IsStale=$is_stale
				OS=$comp.operatingsystemversion
				SkippedReason=$skip_reason
				Link=$link_url
			}
		)
	} else {
		# Add the system to our list.
		Write-Verbose("[{0}] [{1}] Contact User is valid" -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"), $comp.Name)
		$contactuser = $contactuser.ToLower()
		if ($contactuser -notmatch "@") {
			$contactuser = $contactuser + $USER_DOMAIN
		}
		# Note: ADUser is not used in any exports.
		[PSCustomObject]@{
			Name=$comp.Name
			ContactUser=$contactuser
			Department=$aduser.Department
			AssetTag=$comp.($COMP_PROPS.AssetTag)
			FormFactor=$comp.($COMP_PROPS.FormFactor)
			LastLogonDate=$lastlogondate
			IsIncompatible=$is_incompatible
			IsShared=$is_sharedsystem
			IsAssigned=$is_assigned
			IsStale=$is_stale
			OS=$comp.operatingsystemversion
			Link=$link_url
			ADUser = $aduser
		}
	}
	$counter++
}
# Group the results by each contactuser.
$contactusers_systems = $processed_systems | Group-Object -Property ContactUser

# Export systems reports.
$processed_systems | Select * -ExcludeProperty ADUser | Export-CSV -NoTypeInformation -Force $EXPORT_PROCESSED_SYSTEMS_PATH
Write-Host("[{0}] Processing {1} users referencing a combined {2} systems ({3} processed total). Exported report to [{4}]." -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"), ($contactusers_systems | Measure).Count, ($processed_systems | Measure).Count, $counter, $EXPORT_PROCESSED_SYSTEMS_PATH)
$skipped_systems | Export-CSV -NoTypeInformation -Force $EXPORT_SKIPPED_SYSTEMS_PATH
Write-Host("[{0}] Exported report of {1} skipped systems to [{2}]." -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"), ($skipped_systems | Measure).Count, $EXPORT_SKIPPED_SYSTEMS_PATH)

# Loop over each user and compose an email (unless -DryRun is given).
$success_email_count = 0
$emailed_users = $null
if ($ManageGroupOnly) {
	Write-Host("[{0}] Skipping emails due to -ManageGroupOnly switch." -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"))
} else {
	$counter = 0
	$emailed_users = foreach ($o in $contactusers_systems) {
		$user = $o.Name
		$department = $o.Group.Department | Select -Unique -First 1
		$systems = $o.Group
		
		# -- System Table --
		$has_shared_system = $false
		$has_stale_system = $false
		$has_unassigned_system = $false
		$has_incompatible_system = $false
		$desktop_count = 0
		$msgSystemTable = ""
		$msgSystem = ""
		foreach($systeminfo in $systems) {
			$shared_system = ""
			if ($systeminfo.IsSharedSystem) {
				$shared_system = "<b>Yes</b>"
				$has_shared_system = $true
			}
			$compatible = "Yes"
			if ($systeminfo.IsIncompatible) {
				$compatible = "<b>NO</b>"
				$has_incompatible_system = $true
			}
			$assettag = $systeminfo.AssetTag
			if( $ASSET_TAG_IS_NUMERIC -And -Not $assettag -match "^\d+$" ) {
				$assettag = $null
			}
			$formFactor = $systeminfo.FormFactor
			if ($systeminfo.FormFactor -eq "Desktop" -Or $systeminfo.FormFactor -eq "All-In-One") {
				$formFactor += "*"
				$desktop_count++
			}
			<#
			$msgSystemTable += @"
	<tr>
		<td>{0}</td><td>{1}</td><td>{2}</td><td>{3}</td><td>{4}</td><td>{5}</td>
	</tr>
"@ -f $systeminfo.Name, $assettag, $formFactor, $systeminfo.LastLogonDate, $shared_system, $compatible
#>
			$msgSystemTable += @"
				<tr>
					<td>{0}</td><td>{1}</td><td>{2}</td><td>{3}</td>
				</tr>
"@ -f $systeminfo.Name, $assettag, $formFactor, $systeminfo.LastLogonDate
			# We can use these values to change the message.
			if (-Not $systeminfo.IsAssigned) {
				$has_unassigned_system = $true
			}
			if ($systeminfo.IsStale) {
				$has_stale_system = $true
			}
			
			# Set the system message for contacts with a single computer.
			if ($EMAIL_DETAILED_TABLE_MULTIPLE_SYSTEMS_ONLY -And [string]::IsNullOrEmpty($msgSystem)) {
				if ($assettag -ne $null) {
					$msgSystem = (", {0} (asset #{1})," -f $systeminfo.Name, $assettag)
				} else {
					$msgSystem = (", {0}," -f $systeminfo.Name)
				}
			}
		}
		
		# -- HEADER --
		# Only add the system name in the first sentence if we aren't showing a table.
		if(-Not $EMAIL_DETAILED_TABLE_MULTIPLE_SYSTEMS_ONLY -Or $o.Count -gt 1) {
			$msgSystem = ""
		}
		# Change header depending on whether they only have a desktop or not.
		if($desktop_count -eq $o.Count -And -Not [string]::IsNullOrWhitespace($EMAIL_INTRO_HTML_DESKTOPS)) {
			$msgRequiredDate = ""
			if ($UPGRADE_DESKTOP_REQUIRED_DATE) {
				$msgRequiredDate = (Get-Date $UPGRADE_DESKTOP_REQUIRED_DATE -Format "dddd, MMMM dd")
			}
			$msgHtml = ($EMAIL_INTRO_HTML_DESKTOPS -f $msgSystem,$msgRequiredDate)
		} else {
			$msgHtml = ($EMAIL_INTRO_HTML -f $msgSystem)
		}
		
		# -- BODY - SYSTEM TABLE --
		# Only show the table if we have more than 1 system or $EMAIL_DETAILED_TABLE_MULTIPLE_SYSTEMS_ONLY is not set.
		if(-Not $EMAIL_DETAILED_TABLE_MULTIPLE_SYSTEMS_ONLY -Or $o.Count -gt 1) {
<#
			$msgHtml += @"
			
	<table border=1>
		<tr>
			<td>Name</td><td>Asset Tag</td><td>Type</td><td>Last Active Date</td><td>Shared</td><td>Compatible</td>
		</tr>
	$msgSystemTable
		</table>
"@
#>
			$msgHtml += @"
			
	<table border=1>
		<tr>
			<td>Name</td><td>Asset Tag</td><td>Type</td><td>Last Active Date</td>
		</tr>
	$msgSystemTable
		</table>
"@
			# -- BODY - NOTES --
			# Set $EMAIL_DETAILED_NOTES to display more detailed notes.
			if (-Not $EMAIL_DETAILED_NOTES) {
				if ($desktop_count -gt 0 -And ($desktop_count -lt $o.Count -Or [string]::IsNullOrWhitespace($EMAIL_INTRO_HTML_DESKTOPS))) {
					$msgHtml += @"
			<p>* Desktops and all-in-ones will be upgraded automatically pending next restart.</p>
"@
				}
			} elseif ($has_shared_system -Or $has_stale_system -Or $has_unassigned_system -Or $has_incompatible_system) {
				$msghtml += @"
	<br />
	<b>Notes</b>
	<ul>
"@
				if ($has_incompatible_system) {
					$msgHtml += @"
		<li>One or more of these systems may not be compatible with Windows 11. You may reply to this email for more information.</li>
"@
				}
				if ($desktop_count -gt 0 -And ($desktop_count -lt $o.Count -Or [string]::IsNullOrWhitespace($EMAIL_INTRO_HTML_DESKTOPS))) {
					$msgHtml += @"
		<li>Desktops and all-in-ones will be upgraded automatically pending next restart.</li>
"@
				}
				if ($has_shared_system) {
					$msgHtml += @"
		<li>One or more of these systems may be a shared system.</li>
"@
				}
				if ($has_stale_system) {
					$msgHtml += @"
		<li>One or more of these systems have not been active for over $STALE_PC_DAYS days.</li>
"@
				}
			if ($has_unassigned_system) {
					$msgHtml += @"
		<li>One or more of these systems are currently not assigned in our asset system. Please reply to this email to help us confirm this is your assigned system.</li>
"@
				}
				$msgHtml += @"
	</ul>
"@
			}
#>
		}
		
		# -- FOOTER --
		if (-Not [string]::IsNullOrEmpty($EMAIL_FOOTER_HTML)) {
			$msgHtml += $EMAIL_FOOTER_HTML
		}
		if (-Not [string]::IsNullOrEmpty($DebugEmailOverride)) {
			$msgHtml += "<p>-DebugEmailOverride enabled. This email would have been sent to [$user].</p>"
		}

		# -DebugEmailOverride will override -DryRun.
		# If -DebugEmailOverride and -DebugEmailLimit is given, only email up to the limit.
		$email_success = $false
		if (-Not $DryRun -Or (-Not [string]::IsNullOrEmpty($DebugEmailOverride) -And (-Not $DebugEmailLimit -Or $success_email_count -le $DebugEmailLimit))) {
			# If all systems are desktops and $EMAIL_USERS_WITH_ONLY_DESKTOPS_AIOS is set to $false.
			if (-Not $EMAIL_USERS_WITH_ONLY_DESKTOPS_AIOS -And $desktop_count -eq $o.Count) {
				Write-Host("[{0}] Skipping email for {1} - only has desktops/aios and `$EMAIL_USERS_WITH_ONLY_DESKTOPS_AIOS is set to false" -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"), $user)
			} else {
				$emailParams = @{
					From = $EMAIL_FROM
					To = $user
					CC = $EMAIL_CC
					Subject = $EMAIL_SUBJECT
					Body = $msgHtml
					Priority = "High"
					DeliveryNotificationOption = @("OnSuccess", "OnFailure")
					SmtpServer = $EMAIL_SMTP
				}
				# Override used for debugging purposes.
				if (-Not [string]::IsNullOrEmpty($DebugEmailOverride)) {
					$emailParams["To"] = $DebugEmailOverride
				}
				# Also send to BCC if set.
				if (-Not [string]::IsNullOrEmpty($EMAIL_BCC)) {
					$emailParams["BCC"] = $EMAIL_BCC
				}
				$sleep_secs = $EMAIL_SLEEP_SECS
				try {
					Send-MailMessage @emailParams -BodyAsHtml
					$email_success = $true
					$success_email_count++
					
					if ($EMAIL_SLEEP_EXTRA_MOD -And ($success_email_count % $EMAIL_SLEEP_EXTRA_MOD) -eq 0) {
						$sleep_secs = $EMAIL_SLEEP_EXTRA_SECS
					}
				} catch {
					Write-Error $_
					$error_count++
					$sleep_secs = $EMAIL_SLEEP_EXTRA_SECS
				}
				# Wait until sending out the next email.
				Start-Sleep -Seconds $sleep_secs
			}
		}
		
		$counter++
		
		[PSCustomObject]@{
			ContactUser = $user
			Department = $department
			SystemCount = ($systems | Measure).Count
			Systems = ($systems | Select @{N="Name"; Expression={$name = $_.Name; if($_.IsIncompatible) { $name += " (!)" }; $name}}).Name  -join ","
			SharedSystems = $has_shared_system
			StaleSystems = $has_stale_system
			UnassignedSystems = $has_unassigned_system
			DesktopCount = $desktop_count
			IncompatibleSystems = $has_incompatible_system
			EmailSent = $email_success
		}
	}
	
	$emailed_users | Export-CSV -NoTypeInformation -Force $EXPORT_EMAILED_SYSTEMS_PATH
	if ($DryRun) {
		Write-Host("[{0}] Would have emailed {1} users (-DryRun enabled). See saved report [{2}]." -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"), ($emailed_users | Measure).Count, $EXPORT_EMAILED_SYSTEMS_PATH)
	} else {
		Write-Host("[{0}] Emailed {1} out of {2} users. Saved report to [{3}]." -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"), $success_email_count, ($emailed_users | Measure).Count, $EXPORT_EMAILED_SYSTEMS_PATH)
	}
}

# Next, we'll also remove all users from the notification GPO and re-add only the ones we're sending emails to.
# Make sure to exclude all the blacklisted users as well.
$users = $contactusers_systems | % { $_.Group.ADUser | Select -ExpandProperty distinguishedname -Unique -First 1 }
if (-Not [string]::IsNullOrEmpty($NOTIFICATION_GROUP)) {
	if (($users | Measure).Count -gt 0) {
		Write-Host("[{0}] Processing notification GPO filter group [{1}]" -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"), $NOTIFICATION_GROUP)
		Write-Verbose("[{0}] First user to process: {1}" -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"), $users | Select -First 1)
		try {
			$group = Get-ADGroup $NOTIFICATION_GROUP
			if (-Not [string]::IsNullOrEmpty($group.distinguishedname)) {
				$groupMembers = Get-ADUser -LDAPFilter "(memberOf=$($group.distinguishedname))"
				if (($groupMembers | Measure).Count -le 0) {
					Write-Host("[{0}] Group is currently empty." -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"))
				} else {
					$removeGroupMembers = $groupMembers | where {$_.distinguishedname -notin $users}
					Write-Host("[{0}] Removing {1} out of {2} members from group [{3}]" -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"), ($removeGroupMembers | Measure).Count, ($groupMembers | Measure).Count, $NOTIFICATION_GROUP)
					if (($removeGroupMembers | Measure).Count -gt 0) {
						Remove-ADGroupMember $NOTIFICATION_GROUP -Members $removeGroupMembers -Confirm:$false -WhatIf:$DryRun
					}
					$users = $users | where {$_ -notin $groupMembers.distinguishedname}
				}
				Write-Host("[{0}] Adding {1} members to group [{2}]" -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"), ($users | Measure).Count, $NOTIFICATION_GROUP)
				Add-ADGroupMember $NOTIFICATION_GROUP -Members $users -Confirm:$false -WhatIf:$DryRun
			}
		} catch {
			Write-Error $_
			$error_count++
		}
	} else {
		Write-Host("[{0}] No users to process for notification GPO filter group." -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"))
	}
}

# Send a report of results.
if ($DryRun) {
	Write-Host("[{0}] Skipping email report due to -DryRun switch." -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"))
} elseif ($EMAIL_REPORT_TO -ne $null) {
	$emailParams = @{
		From = $EMAIL_REPORT_FROM
		Subject = $EMAIL_REPORT_SUBJECT
		#Priority = "High"
		DeliveryNotificationOption = @("OnSuccess", "OnFailure")
		SmtpServer = $EMAIL_SMTP
	}
	if ($EMAIL_REPORT_TO -is [array]) {
		$report_users = $null
		$report_users = Get-ADUsersByGroup $EMAIL_REPORT_TO -ADProperties mail -Nested -Verbose
		$emailParams["To"] = $report_users | Select -ExpandProperty mail -Unique
	} else {
		$emailParams["To"] = $EMAIL_REPORT_TO
	}
	$emailParams["Body"] = @"
<p>Sent [$success_email_count] emails for [{0}] users referencing [{1}] systems. See [$EXPORT_EMAILED_SYSTEMS_PATH] for more info.</p>

<p><b>Skipped processing [{2}] systems (no valid contact, ineligible, or otherwise excluded). Please check [$EXPORT_SKIPPED_SYSTEMS_PATH] and email / update manually as needed.</b></p>

<p>There were [$error_count] caught errors from [$_scriptName] running on [${ENV:COMPUTERNAME}]. See [$_logfilepath] for more details.</p>
"@ -f ($emailed_users | Measure).Count, ($processed_systems | Measure).Count, ($skipped_systems | Measure).Count

	Send-MailMessage @emailParams -BodyAsHtml
}

Write-Host("[0] Errors encountered: {1}" -f ((Get-Date).toString("yyyy/MM/dd HH:mm:ss")), $errorCount)
	
$runtimeDiff = ((Get-Date) - $dateStart)
Write-Host("[{0}] Total Runtime: {1} hours {2} minutes ({3} total minutes)" -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"), $runtimeDiff.Hours, $runtimeDiff.Minutes, $runtimeDiff.TotalMinutes)

# Stop logging
Stop-Transcript -ErrorAction SilentlyContinue | Out-Null

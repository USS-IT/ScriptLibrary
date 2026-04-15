<#
	.SYNOPSIS
	Stale computer cleanup script
	
	.DESCRIPTION
	Script designed to scan an OU for computer objects that are inactive based on network connectivity and their last logon date. Inactive (stale) computers will be moved to a "retirement" OU, where they wille eventually be disabled, then deleted from AD. Results are recorded in a dated CSV on the HSA network share. Assigned users will be emailed warnings before action is taken.

	.PARAMETER DryRun
	Only output/log results. Do not make any changes or send any emails (-WhatIf).
	
	.PARAMETER DisableWarningEmail
	Do not send any warning emails to end-users for inactive systems.
	
	.PARAMETER WarningDays
	Number of days prior to current date which will result in end-users getting a warning email. Default: 30+ days stale
	
	.PARAMETER RetirementDays
	Number of days prior to current date which will result in system getting disabled from AD. Default: 90+ days stale
	
	.PARAMETER DeletionDays
	Number of days prior to current date which will result in system getting deleted from AD. Default: 180+ days stale
	
	.PARAMETER LogLevel
	Level of verbose console output. Give greater than 0 for more output.
	
	.NOTES
	Requirements:
	* RSAT AD Tools
	
	To run on-demand, use the "Days" parameters to tweak the various thresholds. Make sure to use -DryRun to check the actions the script would perform. Use -DisableWarningEmail if you want to avoid sending out extra warning emails to end-users for each on-demand run and aren't using -DryRun.
	
	Saves LAPS PW in exported logs in case it's needed.
	
	Authors:
	Daniel Anderson - dander83@jhu.edu
	Jerome Powell - Jerome.Powell@jhu.edu
	Matthew Carras - mcarras8@jhu.edu
	
	Changelog
	04-15-26 - mcarras8 - Fixed blank model. Added support for excluding emailing users. Added -LogLevel switch. Changes to emailed report.
	02-10-26 - mcarras8 - Slight change to result email, skipped email for deleted computers, fix asset tag in emails
	12-23-25 - mcarras8 - Moved thresholds to parameters, added support for saving LAPS PW, changed -DryRun output
	12-19-25 - mcarras8 - Added -DisableWarningEmail parameter
	11-03-25 - mcarras8 - Fixed typos. Added CC to report email.
	08-19-25 - mcarras8 - Minor tweaks
	07-21-25 - mcarras8 - Retry failed emails, show model info in email, additional logging
	07-16-25 - mcarras8 - Fix for Send-MailMessage not retrying as intended
						- Fix for VIP Users
	07-09-25 - mcarrasu - Added support for marking VIP users in reports
	07-03-25 - mcarras8 - Added shared contact mapping support
	04-10-25 - mcarras8 - Revamped script
#>
param(
	[Parameter(Mandatory=$false)]
	[switch]$DryRun,
	
	[Parameter(Mandatory=$false)]
	[switch]$DisableWarningEmail,
	
	[Parameter(Mandatory=$false)]
	[int]$WarningDays = 30,
	
	[Parameter(Mandatory=$false)]
	[int]$RetirementDays = 90,
	
	[Parameter(Mandatory=$false)]
	[int]$DeletionDays = 180,
	
	[Parameter(Mandatory=$false)]
	[int]$LogLevel = 0
)

#Import AD module for earlier versions of PowerShell
Import-Module ActiveDirectory

# -- START CONFIGURATION --
# Additional properties we need from AD.
$COMP_PROPS = @{
	# If given, save the current LAPS password in the export. Set to null or empty to skip.
	"LAPSPW" = "ms-Mcs-AdmPwd"
	# Set the attribute synced from SOR for computer assignment.
	# This field will be emailed if they are past $warningDays inactive.
	"Assignment" = "extensionAttribute2"
	# Additional attributes synced from SOR. Used in the email sent to the user and reports.
	"FormFactor" = "extensionAttribute5"
	"AssetTag" = "extensionAttribute1"
	"Model" = "extensionAttribute7"
}

# The OU containing all contactable users.
$OU_USER = "OU=PEOPLE,DC=win,DC=ad,DC=jhu,DC=edu"

# Main searchbase
$OU_COMPUTERS = 'OU=Computers,OU=USS,DC=win,DC=ad,DC=jhu,DC=edu'
# OU used to move retired computers to
$OU_RETIREMENT = 'OU=USS-Retired,OU=Computers,OU=USS,DC=win,DC=ad,DC=jhu,DC=edu'
# List of OUs to exclude from processing.
$OU_EXCLUDE = @('OU=USS-VPS,OU=Computers,OU=USS,DC=win,DC=ad,DC=jhu,DC=edu','OU=USS-DMG,OU=USS-DMC,OU=Computers,OU=USS,DC=win,DC=ad,DC=jhu,DC=edu','OU=USS-STARS,OU=Computers,OU=USS,DC=win,DC=ad,DC=jhu,DC=edu')
# Computers in this group will also be excluded.
$COMP_GROUP_EXCLUDE = 'USS-StalePCExcludeComps'
# If true, still delete the systems in $COMP_GROUP_EXCLUDE after $DATE_REMOVAL has passed.
$COMP_GROUP_EXCLUDE_Delete = $true
# Computers assigned to users in this group will also be excluded. Requires $COMP_PROPS.Assignment to be set and valid.
$ASSIGNED_USER_GROUP_EXCLUDE = 'USS-StalePCExcludeUsers'
# If set, still email assigned users of excluded systems.
$EMAIL_EXCLUDED_SYSTEMS = $true
# Users in this group will not be emailed. Requires $COMP_PROPS.Assignment to be set and valid.
$ASSIGNED_USER_GROUP_EMAIL_EXCLUDE = 'USS-StalePCExcludeEmailUsers'
# Users in the these groups will be marked as "VIP" in reports.
$VIP_USER_GROUPS = @("USS-VIP")

# Fallback for shared systems and systems missing contact info. Matches on DistinguishedName.
# Header: Pattern,Username
# To match a name, start with "CN=". To match on an OU, use ",OU=<ou>,"
# The script will check if the username still exists in AD.
$CONTACTUSER_MAPPING_FALLBACK_FP = "ContactMappingFallback.csv"

# Location and filename for storing CSV results
$CSV_RESULTS_PATH = "\\win.ad.jhu.edu\cloud\HSA$\ITServices\Reports\StalePCs"
$CSV_RESULTS_FP = "$CSV_RESULTS_PATH\StalePCs-{0}.csv" -f (Get-Date -format 'MM-dd-yyyy')
$CSV_HEADER = @("Name","LastLogonDate","PingResult","Action","AssignedUser","Emailed","VIP","FormFactor","AssetTag","LAPS PW")

# Automated email settings.
$EMAIL_ASSIGNEDUSER = $true
$EMAIL_SMTP = 'smtp.johnshopkins.edu'
$EMAIL_FROM = 'mcarras8@jhu.edu'
$EMAIL_CC = @('Jerome.Powell@jhu.edu','mcarras8@jhu.edu')
$EMAIL_BCC = 'ussitservices@jhu.edu'
$EMAIL_SUBJECT = "[USS-IT] Inactive System Alert"
$EMAIL_INTRO_HTML = @"
<p>This is an automated message.</p>
<p>You are receiving this email because one or more systems assigned to you have been offline for an extended period of time. To prevent future complications please login to your system as soon as possible. If are working remotely, you may need to leave the system connected to its charger and the internet overnight to fully update.</p>

<p>If you are no longer using this system, or think you may have received this email in error, please reply back to this email to help update our records.</p>

<p>Thank you for your cooperation.</p>
"@
# Amount of time in seconds to sleep between emails.
$EMAIL_SLEEP_SECS = 10
# Number of successful emails to send before sleeping longer (e.g. 10 for every 10 emails).
# The $EMAIL_SLEEP_EXTRA_SECS will also be used if any emails fail.
$EMAIL_SLEEP_EXTRA_MOD=10
$EMAIL_SLEEP_EXTRA_SECS = 60
# Attempt to send the email again on failure after waiting $EMAIL_SLEEP_EXTRA_SECS.
# Set to 0 to disable.
$EMAIL_RETRY_LIMIT = 4

# Email a report at the end.
$EMAIL_REPORT_FROM = 'USS IT Services <ussitservices@jhu.edu>'
$EMAIL_REPORT_TO = @("ussitservices@jhu.edu")
# Overrides $EMAIL_REPORT_TO
#$EMAIL_REPORT_TO_GROUPS = $null
#$EMAIL_REPORT_CC = $null
# Overrides $EMAIL_REPORT_CC
$EMAIL_REPORT_CC_GROUPS = @("USS-IT-StalePCReports")
$EMAIL_REPORT_SUBJECT = "Results from Stale PC Cleaner script"

# Path and prefix for the Start-Transcript logfiles.
$LOGFILE_PATH = "\\win.ad.jhu.edu\cloud\hsa$\ITServices\Reports\Logs\StalePCCleaner"
$LOGFILE_PREFIX = "stalepccleaner"
# Maximum number of days before rotating logfile.
$LOGFILE_ROTATE_DAYS = 90
# -- END CONFIGURATION --

# Dates to check LastLogonDate against. These are set in the parameters.
# Change the value after AddDays to customize the timeframes
# Date threshold to warn assigned users of possible pending action
$DATE_WARNING = (Get-Date).AddDays((-1 * $WarningDays))
# Date threshold to move system to retirement OU
# If system is already in retirement OU, it will be disabled instead
$DATE_RETIREMENT = (Get-Date).AddDays((-1 * $RetirementDays))
# Date threshold to delete system out of AD entirely
# If not set or $null this action will always be skipped
$DATE_REMOVAL = (Get-Date).AddDays((-1 * $DeletionDays))

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

if ($DryRun) {
	Write-Host("[{0}] -DryRun set. Only output results to file/console. Using -WhatIf or otherwise skipping actions." -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"))
}

# Get a list of contact fallback mappings.
$contactFallbackMappings = $null
if (-Not [string]::IsNullOrEmpty($CONTACTUSER_MAPPING_FALLBACK_FP)) {
	if (-Not (Test-Path $CONTACTUSER_MAPPING_FALLBACK_FP -PathType Leaf)) {
		Write-Warning "Contact Fallback Mapping file [$CONTACTUSER_MAPPING_FALLBACK_FP] not found"
	} else {
		$contactFallbackMappings = Import-CSV $CONTACTUSER_MAPPING_FALLBACK_FP
	}
}

# Collect VIP users. These will only be used for reports.
$VIPUsers = $null
$VIPUsersCount = 0
if (($VIP_USER_GROUPS | Measure).Count -gt 0) {
	Write-Host("[{0}] Collecting VIP users from groups: {1}" -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"), ($VIP_USER_GROUPS -join ", "))
	try {
		$VIPUsers = Get-ADUsersByGroup $VIP_USER_GROUPS -ADProperties "mail" -Nested -Verbose
		$VIPUsersCount = ($VIPUsers | Measure).Count
	} catch {
		Write-Error $_
		$error_count++
	}
}

# Add assignment to AD properties.
# Compute assigned users excluded from processing and/or email.
$excludedUsers = $null
$excludedUsersCount = 0
$excludedEmailUsers = $null
$excludedEmailUsersCount = 0
$props = @("Name","LastLogonDate")
if (-Not [string]::IsNullOrWhitespace($COMP_PROPS['Assignment'])) {
	# If we also have groups to exclude
	if (-Not [string]::IsNullOrWhitespace($ASSIGNED_USER_GROUP_EXCLUDE)) {
		Write-Host("[{0}] Collecting excluded users from group: {1}" -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"), $ASSIGNED_USER_GROUP_EXCLUDE)
		try {
			$excludedUsers = Get-ADUsersByGroup $ASSIGNED_USER_GROUP_EXCLUDE -ADProperties "mail" -Nested -Verbose
			$excludedUsersCount = ($excludedUsers | Measure).Count
			Write-Host("[{0}] Collected {1} users to exclude" -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"), $excludedUsersCount)
		} catch {
			Write-Error $_
			$error_count++
		}
	}
	if (-Not [string]::IsNullOrWhitespace($ASSIGNED_USER_GROUP_EMAIL_EXCLUDE)) {
		Write-Host("[{0}] Collecting users excluded from being emailed from group: {1}" -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"), $ASSIGNED_USER_GROUP_EMAIL_EXCLUDE)
		try {
			$excludedEmailUsers = Get-ADUsersByGroup $ASSIGNED_USER_GROUP_EMAIL_EXCLUDE -ADProperties "mail" -Nested -Verbose
			$excludedEmailUsersCount = ($excludedEmailUsers | Measure).Count
			Write-Host("[{0}] Collected {1} users to exclude from emails" -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"), $excludedEmailUsersCount)
		} catch {
			Write-Error $_
			$error_count++
		}
	}
}

# Compute additional properties for Get-ADComputer.
foreach ($item in $COMP_PROPS.GetEnumerator()) {
	if (-Not [string]::IsNullOrWhitespace($item.Value)) {
		$props += @($item.Value)
	}
}

# Get list of excluded computers.
# These may still be deleted if $COMP_GROUP_EXCLUDE_Delete is set to true.
$excludedComps = $null
$excludedCompsCount = 0
if (-not [string]::IsNullOrWhitespace($COMP_GROUP_EXCLUDE)) {
	Write-Host("[{0}] Collected excluded computers from group {1}..." -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"), $COMP_GROUP_EXCLUDE)
	try {
		$excludedComps = Get-ADGroupMember $COMP_GROUP_EXCLUDE -Recursive | where {$_.objectClass -eq "computer"}
		$excludedCompsCount = ($excludedComps | Measure).Count
		Write-Host("[{0}] Collected {1} computers to exclude from group {2}" -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"), $excludedCompsCount, $COMP_GROUP_EXCLUDE)
	} catch {
		Write-Error $_
		$error_count++
	}
}

# Scan Computers OU (SearchBase) for systems that have not been logged in since $warningDays.
# First ping the computers up to 3 times. If any pass, skip all other checks.
# Computers with LastLogonDate older than $DATE_RETIREMENT will be moved to the Retired OU if they haven't already.
# If they are already in the Retired OU, they will be disabled.
# If they are already disabled, and if $DATE_REMOVAL is set, they will be deleted out of AD.
	
# Hash table of users to email.
$contactUserSystems = @{}
# Hash table of systems to add messages for.
$logSystems = @{}
# Stats
$movedSystemCount = 0
$disabledSystemCount = 0
$deletedSystemCount = 0
# Grab the systems.
$comps = Get-ADComputer -Property $props -Filter * -SearchBase $OU_COMPUTERS | where {$_.LastLogonDate -isnot [datetime] -Or $_.LastLogonDate -lt $DATE_WARNING}
Write-Host("[{0}] Collected {1} computers from AD matching criteria" -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"), ($comps | Measure).Count)
$comps | ForEach-Object {
	# Get the OU from the DistinguishedName
	$ou = $null
	if ($_.distinguishedname -match "CN=[^,]+,(.+)" -And -Not [string]::IsNullOrEmpty($Matches.1)) {
		$ou = $Matches[1]
	}
	Write-Host("[{0}] Pinging {1} with LastLogonDate={2}, OU=[{3}]" -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"), $_.Name, $_.LastLogonDate, $ou)
	$pingResult = ""
	$actionTaken = ""
	$contactEmail = $null
	$assignedUser = $null
	$lapsPW = ""
	# Attempt to ping the system up to 3 times.
	# If it responds, stop all other processing.
	if (((Test-Connection $_.name -Count 1 -ErrorAction SilentlyContinue) -Or 
	    (Test-Connection $_.name -Count 1 -ErrorAction SilentlyContinue) -Or 
		(Test-Connection $_.name -Count 1 -ErrorAction SilentlyContinue))) {
		Write-Host("[{0}] Ping success for {1}" -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"), $_.Name)
		$pingResult = "Success"
		Write-Host("[{0}] Ping success from {1}, no further action" -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"), $_.Name)
	} Else {
		$pingResult = "Fail"
		Write-Host("[{0}] No ping response from {1}" -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"), $_.Name)
		
		$skipProcessing = $false
		# Get contact email and check if assigned user is excluded from processing.
		if (-Not [string]::IsNullOrWhitespace($COMP_PROPS['Assignment']) -And $_.($COMP_PROPS['Assignment']) -ne $null) {
			$assignedUser = $_.($COMP_PROPS['Assignment'])
			if ($assignedUser -match "@") {
				try {
					Write-Host("[{0}] Looking up assigned user [{1}] in AD" -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"), $assignedUser)
					$u = Get-ADUserCached $assignedUser -Properties "mail"
					# If user is enabled and in a valid user OU
					if (-Not $u.Enabled -Or $u.distinguishedname -notlike "CN=*,$OU_USER") {
						Write-Host("[{0}] Excluding user contact - not enabled or invalid user OU " -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"))
					} else {
						# Check if the user is in one of the excluded user groups
						if ($excludedUsersCount -gt 0 -And $u.distinguishedname -in $excludedUsers.distinguishedname) {
							Write-Host("[{0}] Skipping action on [{1}] due to assigned user [{2}] found in excluded user group" -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"), $_.DistinguishedName, $assignedUser)
							$actionTaken = "Assigned User Excluded (Skipped)"
							$skipProcessing = $true
						}
						# Add the contact email unless this system is being excluded and $EMAIL_EXCLUDED_SYSTEMS is not set.
						if (-Not [string]::IsNullOrWhitespace($u.mail) -And (-Not $skipProcessing -Or $EMAIL_EXCLUDED_SYSTEMS)) {
							$contactEmail = $u.mail
						}
					}
				} catch {
					Write-Error $_
					$error_count++
				}
			} elseif (-Not [string]::IsNullOrEmpty($assignedUser)) {
				Write-Host("[{0}] - [{1}] has possible departmental user assignment [{2}]" -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"), $_.Name, $assignedUser)
				
				# Check the fallback mappings for shared systems.
				# This is mostly for shared systems.
				if($contactFallbackMappings) {
					foreach($m in $contactFallbackMappings) {
						if(-Not [string]::IsNullOrWhitespace($m.Pattern) -And $comp.distinguishedname -match $m.Pattern) {
							$fallbackContact = $m.Username
							if ([string]::IsNullOrWhitespace($m.Username)) {
								Write-Error("[{0}] [{1}] - Matched fallback contact pattern [{2}] but Username column is blank" -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"), $_.Name, $m.Pattern)
								$error_count++
							} else {
								Write-Host("[{0}] [{1}] - Found fallback contact [{2}]" -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"), $_.Name, $fallbackContact)
								
								$u = Get-ADUserCached -User $fallbackContact -Properties "mail"
								if (-Not $u.Enabled -Or $u.distinguishedname -notlike "CN=*,$OU_USER") {
									Write-Warning("[{0}] [{1}] Fallback contact user [{2}] is disabled or not found in user OU" -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"),  $_.Name, $fallbackContact)
								} elseif (-Not [string]::IsNullOrWhitespace($u.mail)) {
									$contactEmail = $u.mail
								}
							}
							break
						}
					}
				}
			}
		}
		
		# Check if system is in an excluded OU.
		# Using -match in this way should allow matching sub-OUs.
		If (-Not [string]::IsNullOrEmpty($ou) -And ($OU_EXCLUDE.Where({$ou -match $_}) | Measure-Object).Count -gt 0) {
			Write-Host("[{0}] Skipping action on [{1}] due to excluded OU" -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"), $_.DistinguishedName)
			$actionTaken = "OU Excluded (Skipped)"
			$skipProcessing = $true
		}
		
		# If the system is past retirementDate.
		# Assume a null LastLogonDate is the same as being past all retirement dates.
		# Skip action on this item if its assigned to an excluded user or in an excluded OU.
		if ($skipProcessing) {
			# If $EMAIL_EXCLUDED_SYSTEMS is not set, then don't contact this user.
			if(-Not $EMAIL_EXCLUDED_SYSTEMS) {
				Write-Host("[{0}] Removing contact email for [{1}] due to excluded user/OU and `$EMAIL_EXCLUDED_SYSTEMS being set to $EMAIL_EXCLUDED_SYSTEMS" -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"), $_.DistinguishedName)
				$contactEmail = $null
			}
		} elseif ($_.LastLogonDate -isnot [datetime] -Or $_.LastLogonDate -le $DATE_RETIREMENT) {
			$onlyDelete = $false
			# Check if we have any excluded computers. If so, 
			If ($excludedCompsCount -gt 0 -And $_.DistinguishedName -in $excludedComps.DistinguishedName) {
				Write-Host("[{0}] Computer [{1}] found in exclusion group [$COMP_GROUP_EXCLUDE], only deleting if matching criteria (`$COMP_GROUP_EXCLUDE_Delete set to $COMP_GROUP_EXCLUDE_Delete)" -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"), $_.DistinguishedName)
				$contactEmail = $null
				$onlyDelete = $COMP_GROUP_EXCLUDE_Delete
			}
			
			# If the computer has not already been moved.
			if ($_.DistinguishedName -notlike "CN=*,$OU_RETIREMENT" -And -Not $onlyDelete) {
			  try {
				If ($DryRun) {
					Write-Host("[{0}] Would move [{1}] to [{2}] (-DryRun)" -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"), $_.DistinguishedName, $OU_RETIREMENT)
				} else {
					Write-Host("[{0}] Moving [{1}] to [{2}]" -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"), $_.DistinguishedName, $OU_RETIREMENT)
				}
				Move-ADObject $_.DistinguishedName -TargetPath $OU_RETIREMENT -WhatIf:$DryRun
				if ($DryRun) {
					$actionTaken = "Pending Move (-DryRun)"
				} else {
					$actionTaken = "Moved"
					$movedSystemCount++
				}
			  } catch {
				  Write-Error $_
				  $error_count++
			  }
			} else {
				# If computer has already been moved to the Retirement OU.
				if ($_.Enabled -And -Not $onlyDelete) {
					try {
						If ($DryRun) {
							Write-Host("[{0}] Would disable [{1}] (-DryRun)" -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"), $_.DistinguishedName)
						} else {
							Write-Host("[{0}] Disabling [{1}]" -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"), $_.DistinguishedName)
						}
						Disable-ADAccount $_.DistinguishedName -WhatIf:$DryRun
						if ($DryRun) {
							$actionTaken = "Pending Disable (-DryRun)"
						} else {
							$actionTaken = "Disabled"
							$disabledSystemCount++
						}
					} catch {
						Write-Error $_
						$error_count++
					}
				# If computer has already been disabled and we have $DATE_REMOVAL set.
				} ElseIf ($DATE_REMOVAL -is [datetime] -And ($_.LastLogonDate -isnot [datetime] -Or $_.LastLogonDate -le $DATE_REMOVAL)) {
					try {
						If (-Not [string]::IsNullOrWhitespace($COMP_PROPS['LAPSPW'])) {
							$lapsPW = $_.($COMP_PROPS['LAPSPW'])
							Write-Host("[{0}] Saving LAPS PW for [{1}] before deletion" -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"), $_.DistinguishedName)
						}
						If ($DryRun) {
							Write-Host("[{0}] Would DELETE [{1}] (-DryRun)" -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"), $_.DistinguishedName)
						} else {
							Write-Host("[{0}] DELETING [{1}]" -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"), $_.DistinguishedName)
						}
						Remove-ADObject $_.DistinguishedName -Confirm:$false -WhatIf:$DryRun
						if ($DryRun) {
							$actionTaken = "Pending Deletion (-DryRun)"
						} else {
							$actionTaken = "Deleted"
							$deletedSystemCount++
						}
					} catch {
						Write-Error $_
						$error_count++
					}
				} elseif (-Not $_.Enabled) {
					Write-Host("[{0}] Already disabled [{1}]" -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"), $_.DistinguishedName)
					$actionTaken = "Disabled already"
				}
			}
		}
	}
	
	# Add to list of systems to add to results file.
	$logSystems[$_.Name] = @{
		LastLogonDate = $_.LastLogonDate
		PingResult = $pingResult
		Action = $actionTaken
	}
	# Add optional attributes.
	if (-Not [string]::IsNullOrWhitespace($COMP_PROPS['AssetTag'])) {
		$logSystems[$_.Name]["AssetTag"] = $_.($COMP_PROPS['AssetTag'])
	}
	if (-Not [string]::IsNullOrWhitespace($COMP_PROPS['FormFactor'])) {
		$logSystems[$_.Name]["FormFactor"] = $_.($COMP_PROPS['FormFactor'])
	}
	if (-Not [string]::IsNullOrWhitespace($COMP_PROPS['Model'])) {
		$logSystems[$_.Name]["Model"] = $_.($COMP_PROPS['Model'])
	}
	if (-Not [string]::IsNullOrWhitespace($COMP_PROPS['Assignment']) -Or $contactEmail -ne $null) {
		$logSystems[$_.Name]["ContactEmail"] = $contactEmail
		$logSystems[$_.Name]["AssignedUser"] = $assignedUser
		$logSystems[$_.Name]["Emailed"] = ""
		$logSystems[$_.Name]["VIP"] = ""
	}	
	If (-Not [string]::IsNullOrWhitespace($COMP_PROPS['LAPSPW'])) {
		$logSystems[$_.Name]["LAPS PW"] = $lapsPW
	}

	# If we have a valid contact email, and the system hasn't been deleted, add it to the list of users to email.
	if ($contactEmail -match "@" -And 
		$actionTaken -ne "Deleted" -And 
		($excludedEmailUsersCount -gt 0 -And 
		 (($assignedUser.distinguishedname -And $assignedUser.distinguishedname -in $excludedEmailUsers.distinguishedname) -Or 
		  (-Not $assignedUser.distinguishedname -And $contactEmail -in $excludedEmailUsers.mail))
		 )
	) {
		Write-Host("[{0}] Adding contact email [{1}] for system [{2}]" -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"), $contactEmail, $_.Name)
		if ($contactUserSystems[$contactEmail] -eq $null) {
			$contactUserSystems[$contactEmail] = @($_.Name)
		} else {
			$contactUserSystems[$contactEmail] += @($_.Name)
		}
	}
}

# Email all the users we collected earlier.
$success_email_count = 0
$failed_email_count = 0
$VIPuser_contact_count = 0
$_email_retry_limit = 0
if ($EMAIL_RETRY_LIMIT -ne $null) {
	$_email_retry_limit = $EMAIL_RETRY_LIMIT
}
if (-Not $EMAIL_ASSIGNEDUSER -Or $DisableWarningEmail) {
	Write-Host("[{0}] Would have [{1}] users to email, however emailing users is disabled" -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"), $contactUserSystems.Count)
} else {
	Write-Host("[{0}] Setting up [{1}] emails" -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"), $contactUserSystems.Count)
	foreach($ht in $contactUserSystems.GetEnumerator()) {
		$email = $ht.Name
		$systems = $ht.Value
		if ($email -match "@") {
			$isVIPUser = $false
			$msgHtml = $EMAIL_INTRO_HTML
			$msgSystemTable = ""
			
			# -- BODY - SYSTEM TABLE --
			$validSystems = $false
			foreach($systemName in $systems) {
				if (-Not [string]::IsNullOrWhitespace($systemName)) {
					$system = $logSystems[$systemName]
					if ($system -ne $null) {
						$validSystems = $true
						# If asset tag is not numeric, null it out.
						$assetTag = $system.AssetTag
						if ($assetTag -notmatch "^\d+") {
							$assetTag = ""
						}
						$msgSystemTable += @"
			<tr>
				<td>$systemName</td><td>$assetTag</td><td>$($system.FormFactor)</td><td>$($system.Model)</td><td>$($system.LastLogonDate)</td>
			</tr>
"@
					}
				}
			}
			
			# -- BODY --
			$msgHtml += @"		
	
	<table border=1>
		<tr>
			<td>Name</td><td>Asset Tag</td><td>Type</td><td>Model</td><td>Last Active Date</td>
		</tr>
		$msgSystemTable
	</table>
"@

			if ($LogLevel -gt 0) {
				Write-Host("To: $email")
				Write-Host($msgHTML)
			}
			
			# Only send the email if we have at least one valid system.
			$email_success = $false
			if ($validSystems) {				
				$emailUser = $email
				$emailParams = @{
					From = $EMAIL_FROM
					To = $emailUser
					CC = $EMAIL_CC
					Subject = $EMAIL_SUBJECT
					Body = $msgHtml
					#Priority = "High"
					DeliveryNotificationOption = @("OnSuccess", "OnFailure")
					SmtpServer = $EMAIL_SMTP
				}
				if (-Not [string]::IsNullOrEmpty($EMAIL_BCC)) {
					$emailParams["BCC"] = $EMAIL_BCC
				}
			
				$email_retry_count=0
				while($email_retry_count -le $_email_retry_limit -And -Not $email_success) {
					
					$sleep_secs = $EMAIL_SLEEP_SECS
					try {
						if ($DryRun) {
							Write-Host("[{0}] Would email [{1}] for systems: {2} (-DryRun)" -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"), $email, ($systems -join ", "))
						} else {
							Write-Host("[{0}] Emailing [{1}] for systems: {2}" -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"), $email, ($systems -join ", "))
							Send-MailMessage @emailParams -BodyAsHtml -ErrorAction Stop
						}
						$email_success = $true
						
						if ($EMAIL_SLEEP_EXTRA_MOD -And ($success_email_count % $EMAIL_SLEEP_EXTRA_MOD) -eq 0) {
							$sleep_secs = $EMAIL_SLEEP_EXTRA_SECS
						}
					} catch {
						Write-Error $_
						$sleep_secs = $EMAIL_SLEEP_EXTRA_SECS
						$email_success = $false
						if ($EMAIL_RETRY_LIMIT) {
							Write-Warning("[{0}] Failed to send to [{1}]. Total sent so far: {2}. Retry count {3} of {4}. " -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"), $emailParams["To"], $success_email_count, $email_retry_count, $EMAIL_RETRY_LIMIT)
						} else {
							Write-Warning("[{0}] Failed to send to [{1}]. Total sent so far: {2}" -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"), $emailParams["To"], $success_email_count)
						}
						
						$email_retry_count++
						# Only increment the failed count if we've reached the limit without any successfully sent emails
						if ($email_retry_count -gt $_email_retry_limit) {
							Write-Host("[{0}] ERROR: Over retry limit for emailing [{1}]." -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"), $emailParams["To"])
							$failed_email_count++
							$error_count++
						}
					}
					# Wait until sending out the next email.
					Start-Sleep -Seconds $sleep_secs
				}
				if ($email_success) {
					$success_email_count++
				}
			}
			# If we successfully sent an email, make sure to log it.
			if ($email_success) {
				# Check if this is an VIP user.
				$isVIPUser = $VIPUsersCount -gt 0 -And ($email -in $VIPUsers.mail)
				if ($isVIPUser) {
					$VIPuser_contact_count++
				}
				foreach($systemName in $systems) {
					$logSystems[$systemName]["Emailed"] = $email
					if ($VIPUsersCount -gt 0) {
						$logSystems[$systemName]["VIP"] = $isVIPUser
					}
				}
			}
		}
	}
	
	if ($DryRun) {
		Write-Host("[{0}] Would have emailed [{1}] users (-DryRun)" -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"), $success_email_count)
	} else {
		Write-Host("[{0}] Emailed [{1}] users (including {2} VIPs) with {3} failed emails." -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"), $success_email_count, $VIPuser_contact_count, $failed_email_count)
	}
}

# Log all the systems to the results file, converting the nested hashtable to a PSCustomObject first.
$logSystemsObj = $logSystems.GetEnumerator() | foreach { $o = $_.Value; $o.Add("Name", $_.Name); [PSCustomObject]$o }
Write-Host("[{0}] Saving results for {1} systems to [{2}]" -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"), ($logSystemsObj | Measure).Count, $CSV_RESULTS_FP)
if ($logSystemsObj -ne $null) {
	$logSystemsObj | Select $CSV_HEADER | Export-CSV $CSV_RESULTS_FP -NoTypeInformation -Force
}

# Send a report of results.
if (-Not $DryRun -And (($EMAIL_REPORT_TO | Measure).Count -gt 0 -Or ($EMAIL_REPORT_TO_GROUPS | Measure).Count -gt 0)) {
	$emailParams = @{
		From = $EMAIL_REPORT_FROM
		Subject = $EMAIL_REPORT_SUBJECT
		#Priority = "High"
		DeliveryNotificationOption = @("OnSuccess", "OnFailure")
		SmtpServer = $EMAIL_SMTP
	}
	Write-Host("-- Email Message --")
	# Add To field
	if (($EMAIL_REPORT_TO_GROUPS | Measure).Count -gt 0) {
		$users = Get-ADUsersByGroup $EMAIL_REPORT_TO_GROUPS -ADProperties mail -Nested
		$emailParams["To"] = $users | Select -ExpandProperty mail -Unique
	} else {
		$emailParams["To"] = $EMAIL_REPORT_TO
	}
	Write-Host("To: " + ($emailParams["To"] -join ", "))
	
	# Add CC field (optional)
	if (($EMAIL_REPORT_CC_GROUPS | Measure).Count -gt 0) {
		$users = Get-ADUsersByGroup $EMAIL_REPORT_CC_GROUPS -ADProperties mail -Nested
		$emailParams["CC"] = $users | Select -ExpandProperty mail -Unique
	} elseif (($EMAIL_REPORT_CC | Measure).Count -gt 0) {
		$emailParams["CC"] = $EMAIL_REPORT_CC 
	}
	if ($emailParams["CC"] -ne $null) {
		Write-Host("CC: " + ($emailParams["CC"] -join ", "))
	}
	
	$emailParams["Body"] = @"
Processed {0} inactive systems (excluding {1}) out of {2} total.<br />
Sent [$success_email_count] emails to [{3}] users (including [$VIPuser_contact_count] VIPs).<br />
Failed to email [$failed_email_count] users.<br />
Moved Systems: $movedSystemCount<br />
Disabled Systems: $disabledSystemCount<br />
Deleted Systems: $deletedSystemCount<br />
<p>See [$CSV_RESULTS_FP] for more info on each action taken.</p>

<p>There were [$error_count] caught errors from [$_scriptName] running on [${ENV:COMPUTERNAME}]. See [$_logfilepath] for more details.</p>
"@ -f ($logSystemsObj | Measure).Count, ($logSystemsObj | where {$_.Action -match "Skipped"} | Measure).Count, ($comps | Measure).Count, ($contactUserSystems.Keys | Measure).Count
	Write-Host($emailParams.Body)
	Write-Host("-- End Email Message --")
	
	Send-MailMessage @emailParams -BodyAsHtml
}

Write-Host("[{0}] Errors encountered: {1}" -f ((Get-Date).toString("yyyy/MM/dd HH:mm:ss")), $error_count)

$runtimeDiff = ((Get-Date) - $dateStart)
Write-Host("[{0}] Total Runtime: {1} hours {2} minutes ({3} total minutes)" -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"), $runtimeDiff.Hours, $runtimeDiff.Minutes, $runtimeDiff.TotalMinutes)

# Stop logging
Stop-Transcript -ErrorAction SilentlyContinue | Out-Null
<#
	.SYNOPSIS
	Stale computer cleanup script
	
	.DESCRIPTION
	Script designed to scan an OU for computer objects that are inactive based on network connectivity and their last logon date. Inactive (stale) computers will be moved to a "retirement" OU, where they wille eventually be disabled, then deleted from AD. Results are recorded in a dated CSV on the HSA network share. Assigned users will be emailed warnings before action is taken.

	.PARAMETER DryRun
	Only output/log results. Do not make any changes or send any emails (-WhatIf).
	
	.PARAMETER Verbose
	Enable additional verbose/debugging output.
	
	.NOTES
	Requirements:
	* RSAT AD Tools
	
	Logs saved to .\Logs\stalepccleaner-<date>.log
	
	Authors:
	Daniel Anderson - dander83@jhu.edu
	Jerome Powell - Jerome.Powell@jhu.edu
	Matthew Carras - mcarras8@jhu.edu
	
	Changelog
	04-10-25 - mcarras8 - Revamped script
#>
param(
	[Parameter(Mandatory=$false)]
	[switch]$DryRun
)

#Import AD module for earlier versions of PowerShell
Import-Module ActiveDirectory

# -- START CONFIGURATION --
# Dates to check LastLogonDate against.
# Change the value after AddDays to customize the timeframes
# Date threshold to warn assigned users of possible pending action
$DATE_WARNING = (Get-Date).AddDays(-30)
# Date threshold to move system to retirement OU
# If system is already in retirement OU, it will be disabled instead
$DATE_RETIREMENT = (Get-Date).AddDays(-90)
# Date threshold to delete system out of AD entirely
# If not set or $null this action will always be skipped
#$DATE_REMOVAL = (Get-Date).AddDays(-180)

# Set the attribute synced from SOR for computer assignment.
# This field will be emailed if they are past $warningDays inactive.
$PROP_ASSIGNMENT = "extensionAttribute2"
# AD attribute for system form factor (Laptop, Desktop, etc.)
$PROP_FORMFACTOR = "extensionAttribute5"
# AD attribute for asset tag
$PROP_ASSETTAG = "extensionAttribute1"

# The OU containing all contactable users.
$OU_USER = "OU=PEOPLE,DC=win,DC=ad,DC=jhu,DC=edu"

# Main searchbase
$OU_COMPUTERS = 'OU=Computers,OU=USS,DC=win,DC=ad,DC=jhu,DC=edu'
# OU used to move retired computers to
$OU_RETIREMENT = 'OU=USS-Retired,OU=Computers,OU=USS,DC=win,DC=ad,DC=jhu,DC=edu'
# List of OUs to exclude from processing.
$OU_EXCLUDE = @('OU=USS-VPS,OU=Computers,OU=USS,DC=win,DC=ad,DC=jhu,DC=edu','OU=USS-DMG,OU=USS-DMC,OU=Computers,OU=USS,DC=win,DC=ad,DC=jhu,DC=edu')
# Computers in this group will also be excluded.
# $OU_EXCLUDEGroups = @('')
# Computers assigned to users in this group will also be excluded. Requires $PROP_ASSIGNMENT to be set and valid.
$ASSIGNED_USER_GROUPS_EXCLUDE = @("USS-VIP")
# If set, still email assigned users of excluded systems.
$EMAIL_EXCLUDED_SYSTEMS = $true

# Location and filename for storing CSV results
$CSV_RESULTS_PATH = "\\win.ad.jhu.edu\cloud\HSA$\ITServices\Reports\StalePCs"
$CSV_RESULTS_FP = "$CSV_RESULTS_PATH\StalePCs-{0}.csv" -f (Get-Date -format 'MM-dd-yyyy')
$CSV_HEADER = @("Name","LastLogonDate","PingResult","Action","AssignedUser","Emailed","FormFactor","AssetTag")

# Automated email settings.
$EMAIL_ASSIGNEDUSER = $true
$EMAIL_SMTP = 'smtp.johnshopkins.edu'
$EMAIL_FROM = 'Jerome.Powell@jhu.edu'
$EMAIL_CC = @('Jerome.Powell@jhu.edu','mcarras8@jhu.edu')
$EMAIL_SUBJECT = "[USS-IT] Inactive System Alert"
$EMAIL_INTRO_HTML = @"
<p>This is an automated message.</p>
<p>You are receiving this email because one or more systems assigned to you have been offline for an extended period of time. To prevent future complications please login to your system as soon as possible. If are working remotely, you may need to leave the system connected to its charger and the internet overnight to fully update.</p>

<p>If you are no longer using this system, or think you may have received this email in error, please reply back to this email to help update our records.</p>

<p>Thank you for your cooperation.</p>
"@
# Number of seconds to sleep in-between each email.
$EMAIL_SLEEP_SECS = 5

# Email a report at the end.
$EMAIL_REPORT_FROM = 'USS IT Services <ussitservices@jhu.edu>'
$EMAIL_REPORT_TO = @("ussitservices@jhu.edu")
#$EMAIL_REPORT_TO_GROUPS = @("USS-IT-JHEDs")
$EMAIL_REPORT_SUBJECT = "Results from Stale PC Cleaner script"

# Path and prefix for the Start-Transcript logfiles.
$LOGFILE_PATH = ".\Logs"
$LOGFILE_PREFIX = "stalepccleaner"
# Maximum number of days before rotating logfile.
$LOGFILE_ROTATE_DAYS = 90
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
$error_count = 0
$_scriptName = split-path $PSCommandPath -Leaf

# Rotate log files
if ($LOGFILE_ROTATE_DAYS -is [int] -And $LOGFILE_ROTATE_DAYS -gt 0) {
	Get-ChildItem "${LOGFILE_PATH}\${LOGFILE_PREFIX}_*.log" | Where-Object { $_.CreationTime -lt (Get-Date).AddDays(-$LOGFILE_ROTATE_DAYS) } | Remove-Item -Force
}

# Start logging
$_logfilepath = "${LOGFILE_PATH}\${LOGFILE_PREFIX}_$(get-date -f yyyy-MM-dd).log"
Start-Transcript -Path $_logfilepath -Append

if ($DryRun) {
	Write-Host("[{0}] -DryRun set. Only output results to file/console. Using -WhatIf or otherwise skipping actions." -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"))
}

# Scan Computers OU (SearchBase) for systems that have not been logged in since $warningDays.
# First ping the computers up to 3 times. If any pass, skip all other checks.
# Computers with LastLogonDate older than $DATE_RETIREMENT will be moved to the Retired OU if they haven't already.
# If they are already in the Retired OU, they will be disabled.
# If they are already disabled, and if $DATE_REMOVAL is set, they will be deleted out of AD.
$props = @("Name","LastLogonDate")
$excludedUsers = $null
$excludedUsersCount = 0
if (-Not [string]::IsNullOrWhitespace($PROP_ASSIGNMENT)) {
	$props += @($PROP_ASSIGNMENT)
	# If we also have groups to exclude
	if (($ASSIGNED_USER_GROUPS_EXCLUDE | Measure).Count -gt 0) {
		Write-Host("[{0}] Collecting excluded users from groups: {1}" -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"), ($ASSIGNED_USER_GROUPS_EXCLUDE -join ", "))
		try {
			$excludedUsers = Get-ADUsersByGroup $ASSIGNED_USER_GROUPS_EXCLUDE -ADProperties "mail" -Nested -Verbose
			$excludedUsersCount = ($excludedUsers | Measure).Count
		} catch {
			Write-Error $_
			$error_count++
		}
	}
}
if (-Not [string]::IsNullOrWhitespace($PROP_ASSETTAG)) {
	$props += @($PROP_ASSETTAG)
}
if (-Not [string]::IsNullOrWhitespace($PROP_FORMFACTOR)) {
	$props += @($PROP_FORMFACTOR)
}
# Hash table of users to email.
$contactUserSystems = @{}
# Hash table of systems to add messages for.
$logSystems = @{}
$comps = Get-ADComputer -Property $props -Filter * -SearchBase $OU_COMPUTERS | where {$_.LastLogonDate -isnot [datetime] -Or $_.LastLogonDate -lt $DATE_WARNING}
Write-Host("[{0}] Collected {1} computers from AD" -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"), ($comps | Measure).Count)
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
	# Attempt to ping the system up to 3 times.
	# If it responds, stop all other processing.
	if (((Test-Connection $_.name -Count 1 -ErrorAction SilentlyContinue) -Or 
	    (Test-Connection $_.name -Count 1 -ErrorAction SilentlyContinue) -Or 
		(Test-Connection $_.name -Count 1 -ErrorAction SilentlyContinue))) {
		Write-Host("[{0}] Ping success for {1}" -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"), $_.Name)
		$pingResult = "Success"
	} Else {
		$pingResult = "Fail"
		
		$skipProcessing = $false
		# Get contact email and check if assigned user is excluded from processing.
		if (-Not [string]::IsNullOrWhitespace($PROP_ASSIGNMENT) -And $_.$PROP_ASSIGNMENT -ne $null) {
			$assignedUser = $_.$PROP_ASSIGNMENT
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
			}
		}
		# Check if system is in an excluded OU.
		If (-Not [string]::IsNullOrEmpty($ou) -And $ou -in $OU_EXCLUDE) {
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
			# If the computer has not already been moved.
			if ($_.DistinguishedName -notlike "CN=*,$OU_RETIREMENT") {
			  try {
				Write-Host("[{0}] Moving [{1}] to [{2}]" -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"), $_.DistinguishedName, $OU_RETIREMENT)
				# Move-ADObject $_.DistinguishedName -TargetPath $OU_RETIREMENT -WhatIf:$DryRun
				$actionTaken = "Moved"
			  } catch {
				  Write-Error $_
				  $error_count++
			  }
			} else {
				# If computer has already been moved to the Retirement OU.
				if ($_.Enabled) {
					try {
						Write-Host("[{0}] Disabling [{1}]" -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"), $_.DistinguishedName)
						# Disable-ADAccount $_.DistinguishedName -WhatIf:$DryRun
						$actionTaken = "Disabled"
					} catch {
						Write-Error $_
						$error_count++
					}
				# If computer has already been disabled and we have $DATE_REMOVAL set.
				} ElseIf ($DATE_REMOVAL -is [datetime] -And ($_.LastLogonDate -isnot [datetime] -Or $_.LastLogonDate -le $DATE_REMOVAL)) {
					try {
						Write-Host("[{0}] DELETING [{1}]" -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"), $_.DistinguishedName)
						# Remove-ADObject $_.DistinguishedName -Confirm:$false -WhatIf:$DryRun
						$actionTaken = "Deleted"
					} catch {
						Write-Error $_
						$error_count++
					}
				} else {
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
	if (-Not [string]::IsNullOrWhitespace($PROP_ASSETTAG)) {
		# Only add assettag if its numeric.
		$assettag = $_.$PROP_ASSETTAG
		$logSystems[$_.Name]["AssetTag"] = $assettag
	}
	if (-Not [string]::IsNullOrWhitespace($PROP_FORMFACTOR)) {
		$logSystems[$_.Name]["FormFactor"] = $_.$PROP_FORMFACTOR
	}
	if (-Not [string]::IsNullOrWhitespace($PROP_ASSIGNMENT) -Or $contactEmail -ne $null) {
		$logSystems[$_.Name]["ContactEmail"] = $contactEmail
		$logSystems[$_.Name]["AssignedUser"] = $assignedUser
		$logSystems[$_.Name]["Emailed"] = ""
	}

	# If we have a valid contact email, add to the list of users to email.
	if ($contactEmail -match "@") {
		Write-Host("[{0}] Adding contact email [{1}] for system [{2}]" -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"), $contactEmail, $_.Name)
		if ($contactUserSystems[$contactEmail] -eq $null) {
			$contactUserSystems[$contactEmail] = @($_.Name)
		} else {
			$contactUserSystems[$contactEmail] += @($_.Name)
		}
	}
}

# Email all the users we collected earlier.
if (-Not $EMAIL_ASSIGNEDUSER) {
	Write-Host("[{0}] Would have [{1}] users to email, however `$EMAIL_ASSIGNEDUSER is set to $EMAIL_ASSIGNEDUSER" -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"), ($contactUserSystems.Keys | Measure).Count)
} else {
	Write-Host("[{0}] Setting up [{1}] emails" -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"), ($contactUserSystems.Keys | Measure).Count)
	$success_email_count = 0
	foreach($ht in $contactUserSystems.GetEnumerator()) {
		$email = $ht.Name
		$systems = $ht.Value
		if ($email -match "@") {
			Write-Host("[{0}] Emailing [{1}] for systems: {2}" -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"), $email, ($systems -join ", "))
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
						if ($_.$PROP_ASSETTAG -notmatch "^\d+") {
							$assetTag = ""
						}
						$msgSystemTable += @"
			<tr>
				<td>$systemName</td><td>{0}</td><td>$($system.FormFactor)</td><td>$($system.LastLogonDate)</td>
			</tr>
"@ -f $assetTag
					}
				}
			}
			
			# -- BODY --
			$msgHtml += @"		
	
	<table border=1>
		<tr>
			<td>Name</td><td>Asset Tag</td><td>Type</td><td>Last Active Date</td>
		</tr>
		$msgSystemTable
	</table>
"@

			Write-Verbose($msgHTML)
			
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
			
				try {
					if (-Not $DryRun) {
						#Send-MailMessage @emailParams -BodyAsHtml
					}
					$email_success = $true
					$success_email_count++
				} catch {
					Write-Error $_
					$error_count++
				}
				Start-Sleep -Seconds $EMAIL_SLEEP_SECS
			}
			# If we successfully sent an email, make sure to log it.
			if ($email_success) {
				foreach($systemName in $systems) {
					$logSystems[$systemName]["Emailed"] = $email
				}
			}
		}
	}
}
if ($DryRun) {
	Write-Host("[{0}] Would have emailed [{1}] users (-DryRun)" -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"), $success_email_count)
} else {
	Write-Host("[{0}] Emailed [{1}] users" -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"), $success_email_count)
}

# Log all the systems to the results file, converting the nested hashtable to a PSCustomObject first.
$logSystemsObj = $logSystems.GetEnumerator() | foreach { $o = $_.Value; $o.Add("Name", $_.Name); [PSCustomObject]$o }
Write-Host("[{0}] Saving results for {1} systems to [{2}]" -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"), ($logSystemsObj | Measure).Count, $CSV_RESULTS_FP)
if ($logSystemsObj -ne $null) {
	$logSystemsObj | Select $CSV_HEADER | Export-CSV $CSV_RESULTS_FP -NoTypeInformation -Force
}

# Send a report of results.
if (($EMAIL_REPORT_TO | Measure).Count -gt 0 -Or ($EMAIL_REPORT_TO_GROUPS | Measure).Count -gt 0) {
	$emailParams = @{
		From = $EMAIL_REPORT_FROM
		Subject = $EMAIL_REPORT_SUBJECT
		#Body = $msgHtml
		#Priority = "High"
		DeliveryNotificationOption = @("OnSuccess", "OnFailure")
		SmtpServer = $EMAIL_SMTP
	}
	if (($EMAIL_REPORT_TO_GROUPS | Measure).Count -gt 0) {
		$users = Get-ADUsersByGroup $EMAIL_REPORT_TO_GROUPS -Properties mail -Nested
		$emailParams["To"] = $users | Select -ExpandProperty mail -Unique
	} else {
		$emailParams["To"] = $EMAIL_REPORT_TO
	}
	$emailParams["Body"] = @"
<p>Sent [$success_email_count] emails for [{0}] users referencing [{1}] systems. Skipped processing [{2}] systems due to exclusions. See [$CSV_RESULTS_FP] for more info.</p>

<p>There were [$error_count] caught errors from [$_scriptName] running on [${ENV:COMPUTERNAME}]. See [$_logfilepath] for more details.</p>
"@ -f ($contactUserSystems.Keys | Measure).Count, ($logSystemsObj | Measure).Count, ($logSystemsObj | where {$_.Action -match "Skipped"} | Measure).Count

	Write-Host($emailParams.Body)
	#Send-MailMessage @emailParams -BodyAsHtml
}

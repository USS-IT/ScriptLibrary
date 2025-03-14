<#
	.SYNOPSIS
	Sends out notifications for required Windows 11 upgrades for AD computer reporting lower than $EOLVER.
	
	.DESCRIPTION
	Sends out notifications for required Windows 11 upgrades for AD computer reporting lower than $EOLVER. Also manages the group for toast notifications.

	.PARAMETER ManageGroupOnly
	Only manage the GPO notification group. Do not send any emails.
	
	.PARAMETER DryRun
	Only output results. Do not send emails or make any other changes.
	
	.NOTES
	Requirements:
	* RSAT AD Tools
	
	EOLVER and UPGRADEVER must be updated periodically in this script. UPGRADEVER is only used in the report names.
	
	Email will be sent out to the assigned user if its valid (contains "@"). If it doesn't exist, then the first USS staff member not a member of USS IT will be used instead, searching LastLogonUser then Primary Users in that order.
	
	Created: 3-14-25
	Author: mcarras8
	
	Changelog
	03-14-25 - MJC - Initial upload.
#>
param(
	[Parameter(Mandatory=$false)]
	[switch]$ManageGroupOnly,
	
	[Parameter(Mandatory=$false)]
	[switch]$DryRun
)
		
# -- START CONFIGURATION --
# This should be the previous version before the latest build version.
$EOLVER=22621		# Windows 11, 22H2
# This should be the version offered in the upgrade.
$UPGRADEVER=22631	# Windows 11, 23H2
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
$PROPS = @{
	"operatingsystemversion"="operatingsystemversion"
	"LastLogonDate"="LastLogonDate"
	"AssetTag"="extensionAttribute1"
	"AssignedUser"="extensionAttribute2"
	"PrimaryUsers"="extensionAttribute3"
	"LastLogonUser"="extensionAttribute10"
	"FormFactor"="extensionAttribute5"
	"LastSyncDate"="extensionAttribute8"
}
# Array of company/companies to limit when searching for users in other fields.
$CONTACTUSER_COMPANIES = @("JHU; University Student Services")
# Array of group(s) containing members of IT.
# They will only be contacted if there's no other valid contact users.
$CONTACTUSER_EXCLUDE_ITGROUPS = @("USS-IT-JHEDs")
# Always exclude these contact users.
$CONTACTUSER_EXCLUDE = @("local_users")
# Exclude contact users matching patterns.
$CONTACTUSER_EXCLUDE_REGEX = "SC\-"
# Contact user domain if not set. This is the domain appended for emails (if needed).
$CONTACTUSER_DOMAIN = "@jh.edu"
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
$EMAIL_FROM = 'USS IT Services <ussitservices@jhu.edu>'
$EMAIL_SUBJECT = "[USS-IT] Windows 11 Upgrade Required"
# This is the first part of each email. Allows for HTML.
# Each email will be in the format of:
# $EMAIL_INTRO_HTML
# <Table with system info>
# <System list items>
# $EMAIL_FOOTER_HTML
$EMAIL_INTRO_HTML = "<p>You are receiving this email because the following systems are missing a critical upgrade to the latest version of Windows 11...etc </p>"
$EMAIL_FOOTER_HTML = "<p>This is an automated email sent from a mailbox that is not actively monitored. Please submit a Help Desk ticket if you have any questions.</p>"
# Amount of time in seconds to sleep between emails.
$EMAIL_SLEEP_SECS = 0
#$EMAIL_SLEEP_SECS = 60

# Path to systems which report being ineligible for upgrade.
$IMPORT_INELIGIBLE_SYSTEMS_PATH = "\\win.ad.jhu.edu\cloud\hsa$\ITServices\Reports\WinUpgrade\win11_incompatible_systems_3_14_25.csv"

# Path to save reports to.
$EXPORT_ALL_SYSTEMS_PATH = "\\win.ad.jhu.edu\cloud\hsa$\ITServices\Reports\WinUpgrade\win11_${UPGRADEVER}_upgrade_systems.csv"
$EXPORT_SKIPPED_SYSTEMS_PATH = "\\win.ad.jhu.edu\cloud\hsa$\ITServices\Reports\WinUpgrade\win11_${UPGRADEVER}_upgrade_systems_skipped.csv"
$EXPORT_PROCESSED_SYSTEMS_PATH = "\\win.ad.jhu.edu\cloud\hsa$\ITServices\Reports\WinUpgrade\win11_${UPGRADEVER}_upgrade_systems_with_contact.csv"
$EXPORT_EMAILED_SYSTEMS_PATH = "\\win.ad.jhu.edu\cloud\hsa$\ITServices\Reports\WinUpgrade\win11_${UPGRADEVER}_upgrade_systems_emailed.csv"

# Path and prefix for the Start-Transcript logfiles.
$LOGFILE_PATH = ".\Logs"
$LOGFILE_PREFIX = "winupgradeautomation"
# Maximum number of days before rotating logfile.
$LOGFILE_ROTATE_DAYS = 90
# -- END CONFIGURATION --

# -- FUNCTION START --
<#
	.SYNOPSIS
	Returns the AD account type by matching to the given -OU.
	
	.DESCRIPTION
	Returns the AD account type by matching to the given -OU.
	
	.PARAMETER User
	The AD user name or object.
	
	.PARAMETER OU
	The OU containing all enabled users.
	
	.OUTPUTS
	"User" - Regular user account
	"Service" - Service account
	"Disabled" - Disabled account
	
	.NOTES
	Saves a cache to $_ADUSER_TYPES.
#>
$_ADUSERS = @{}
function Get-ADAccountType {
	param(
		[Parameter(Mandatory=$true,Position=0)]
		$User,
		
		[Parameter(Mandatory=$true,Position=1)]
		[string]$OU
	)

	$usertype = $_ADUSERS.$User.UserType
	if ([string]::IsNullOrEmpty($usertype)) {
		try {
			$aduser = Get-ADUser $User -Properties Company			
			if(-Not $aduser.Enabled) {
				$usertype="Disabled"
			} elseif($aduser.distinguishedname -eq "CN=$($aduser.Name),$OU") {
				$usertype="User"
			} elseif(-Not [string]::IsNullOrEmpty($aduser.distinguishedname)) {
				$usertype="Service"
			} else {
				$usertype="Error2"
			}
			# Save result to cache.
			$_ADUSERS[$User] = @{
				UserType = $usertype
				Company = $aduser.Company
			}	
		} catch {
			$usertype="Error1"
		}
	}
	return $usertype
}
<#
	.SYNOPSIS
	Returns the Company for the given AD user.
	
	.DESCRIPTION
	Returns the Company for the given AD user.
	
	.PARAMETER User
	The AD user name or object.
	
	.OUTPUTS
	Company for given user
	
	.NOTES
	Saves a cache to $_ADUSERS.
#>
function Get-ADCompany {
	param(
		[Parameter(Mandatory=$true,Position=0)]
		$User
	)

	$company = $_ADUSERS.$User.Company
	if ([string]::IsNullOrEmpty($company)) {
		try {
			$aduser = Get-ADUser $User -Properties Company				
			$_ADUSERS[$User] = @{
				UserType = $null
				Company = $aduser.Company
			}
		} catch {
			$company = $null
		}
	}
	return $company
}
# -- FUNCTION END --

$error_count = 0

# Rotate log files
if ($LOGFILE_ROTATE_DAYS -is [int] -And $LOGFILE_ROTATE_DAYS -gt 0) {
	Get-ChildItem "${LOGFILE_PATH}\${LOGFILE_PREFIX}_*.log" | Where-Object { $_.CreationTime -lt (Get-Date).AddDays(-$LOGFILE_ROTATE_DAYS) } | Remove-Item -Force
}

# Start logging
$_logfilepath = "${LOGFILE_PATH}\${LOGFILE_PREFIX}_$(get-date -f yyyy-MM-dd).log"
Start-Transcript -Path $_logfilepath -Append
	
# Get enabled computers matching $eolver or lower.
$_props = $PROPS.Values | % { $_ }
$comps = Get-ADComputer -Searchbase $searchbase -Filter {Enabled -eq $true} -Properties $_props | where {$_.OperatingSystemVersion -match "10.0 \((\d+)\)" -and $Matches.1 -ne $null -and ($Matches.1 -as [int]) -is [int] -And $Matches.1 -le $eolver -And $_.distinguishedname -notin $EXCLUDE_OUS}

# Convert the hashtable map into a dynamic select array before exporting.
$selectarray = @("distinguishedname")
foreach ($k in $PROPS.Keys) {
	$selectarray += @(@{"N"=$k; "Expression"=[Scriptblock]::Create("`$_.'$($PROPS.$k)'")})
}
$comps | Select $selectarray | Export-CSV -NoTypeInformation -Force $EXPORT_ALL_SYSTEMS_PATH
Write-Host("[{0}] Exported {1} systems requiring Windows 11 upgrade collected from AD to [{2}]." -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"), ($comps | Measure).Count, $EXPORT_ALL_SYSTEMS_PATH)

# Get excluded contact users.
$itusers = $CONTACTUSER_EXCLUDE_ITGROUPS | % { Get-ADGroupMember $_ -Recursive } | Select -ExpandProperty Name -Unique

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
	$contactuser = $comp.($PROPS.AssignedUser)
	$is_sharedsystem = $false
	$is_assigned = $true
	$is_incompatible = ($comp.Name -in $incompatible_systems.Name)
	$lastlogondate = $comp.($PROPS.LastLogonDate)
	if ($lastlogondate -and ((Get-Date) - $lastlogondate).Days -gt $STALE_PC_DAYS) {
		$is_stale = $true
	} else {
		$is_stale = $false
	}
	$assettag = $comp.($PROPS.AssetTag)
	$link_url = $SORURL_ASSETTAG + $assettag
	if ($contactuser -notmatch "@") {
		if ([string]::IsNullOrEmpty($contactuser) -And $CONTACT_ONLY_WHEN_ASSIGNED) {
			$skipped_systems += @(
				[PSCustomObject]@{
					Name=$comp.Name
					AssetTag=$comp.($PROPS.AssetTag)
					FormFactor=$comp.($PROPS.FormFactor)
					LastLogonDate=$lastlogondate
					IsIncompatible=$is_incompatible
					IsShared=$null
					IsAssigned=$false
					IsStale=$is_stale
					OS=$comp.operatingsystemversion
					SkippedReason="No Assigned User"
					Link=$link_url
				}
			)
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
			$contactuser = $comp.($PROPS.LastLogonUser)
			$is_contactuser_it = $false
			if ([string]::IsNullOrWhitespace($contactuser) -Or $contactuser -in $CONTACTUSER_EXCLUDE -Or $contactuser -match $CONTACTUSER_EXCLUDE_REGEX) {
				$contactuser = $null
			} else {
				# Double-check user is valid by checking AD.
				# Basically, if a user is disabled or outside of the default $USER_OU then we assume they're not a valid "User".
				# Also check the user's company to see if they match our supported companies.
				# Finally, check if the user is a member of IT.
				if ((Get-ADAccountType $contactuser $USER_OU) -ne "User") {
					$contactuser = $null
				} else {
					$company = Get-ADCompany $user
					if (-Not [string]::IsNullOrWhitespace($company) -And ($CONTACTUSER_COMPANIES | Measure).Count -gt 0 -And $company -notin $CONTACTUSER_COMPANIES) {
						$contactuser = $null
					} else {
						$is_contactuser_it = $contactuser -in $itusers
					}
				}
			}
			# Check primary users if we don't have a valid contact from LastLogonUser
			# Exclude users not matching criteria, including those not matching a valid company
			# If the Primary User is a member of IT the logic will fall back to use the LastLogonUser regardless
			if ([string]::IsNullOrWhitespace($contactuser) -Or $is_contactuser_it) {
				foreach ($u in ($comp.($PROPS.PrimaryUsers) -split ",")) {
					$user = $u -replace "[^\\]+\\",""
					if (-Not [string]::IsNullOrWhitespace($user) -and $user -notin $itusers -and $user -notin $CONTACTUSER_EXCLUDE -and $user -notmatch $CONTACTUSER_EXCLUDE_REGEX -and (Get-ADAccountType $user $USER_OU) -eq "User") {
						$company = Get-ADCompany $user
						if ($CONTACTUSER_COMPANIES.Count -eq 0 -Or [string]::IsNullOrWhitespace($company) -Or $company -in $CONTACTUSER_COMPANIES) {
							$contactuser = $user
							break
						}
					}
				}
			}
		}
	}
	# Only continue if we have a valid contact user
	if ([string]::IsNullOrWhitespace($contactuser)) {
		Write-Warning("[{0}] [{1}] - SKIPPING: No valid contact user." -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"), $comp.Name)
		$skipped_systems += @(
			[PSCustomObject]@{
				Name=$comp.Name
				AssetTag=$comp.($PROPS.AssetTag)
				FormFactor=$comp.($PROPS.FormFactor)
				LastLogonDate=$lastlogondate
				IsIncompatible=$is_incompatible
				IsShared=$is_sharedsystem
				IsAssigned=$is_assigned
				IsStale=$is_stale
				OS=$comp.operatingsystemversion
				SkippedReason="No Valid Contact"
				Link=$link_url
			}
		)
	} else {
		# Add the system to our list.
		$contactuser = $contactuser.ToLower()
		[PSCustomObject]@{
			Name=$comp.Name
			ContactUser=$contactuser
			AssetTag=$comp.($PROPS.AssetTag)
			FormFactor=$comp.($PROPS.FormFactor)
			LastLogonDate=$lastlogondate
			IsIncompatible=$is_incompatible
			IsShared=$is_sharedsystem
			IsAssigned=$is_assigned
			IsStale=$is_stale
			OS=$comp.operatingsystemversion
			Link=$link_url
		}
	}
	$counter++
}
# Group the results by each contactuser.
$contactusers_systems = $processed_systems | Group-Object -Property ContactUser

# Export systems reports.
$processed_systems | Export-CSV -NoTypeInformation -Force $EXPORT_PROCESSED_SYSTEMS_PATH
Write-Host("[{0}] Processing {1} users referencing a combined {2} systems ({3} processed total). Exported report to [{4}]." -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"), ($contactusers_systems | Measure).Count, ($processed_systems | Measure).Count, $counter, $EXPORT_PROCESSED_SYSTEMS_PATH)
$skipped_systems | Export-CSV -NoTypeInformation -Force $EXPORT_SKIPPED_SYSTEMS_PATH
Write-Host("[{0}] Exported report of {1} skipped systems to [{2}]." -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"), ($skipped_systems | Measure).Count, $EXPORT_SKIPPED_SYSTEMS_PATH)

# Loop over each user and compose an email (unless -DryRun is given).
$emailed_users = $null
if (-Not $ManageGroupOnly ) {
	$emailed_users = foreach ($o in $contactusers_systems) {
		$user = $o.Name
		$contactuser = $user + $CONTACTUSER_DOMAIN
		$systems = $o.Group
		# DEBUG: Below is a placeholder message. The loop also breaks out early.
		$msgHtml = @"
$EMAIL_INTRO_HTML
<table border=0>
	<tr>
		<td>Name</td><td>Asset Tag</td><td>Form Factor</td><td>Last Active Date</td><td>Shared?</td><td>Compatible?</td>
	</tr>
"@
		$has_shared_system = $false
		$has_stale_system = $false
		$has_unassigned_system = $false
		$has_incompatible_system = $false
		foreach($systeminfo in $systems) {
			if ($systeminfo.IsSharedSystem) {
				$shared_system = "Yes"
			} else {
				$shared_system = ""
			}
			$assettag = $systeminfo.AssetTag
			if( $ASSET_TAG_IS_NUMERIC -And -Not $assettag -match "/d+" ) {
				$assettag = $null
			}
			$msgHtml += @"
	<tr>
		<td>{0}</td><td>{1}</td><td>{2}</td><td>{3}</td>{4}</td>
	</tr>
"@ -f $systeminfo.Name, $assettag, $systeminfo.FormFactor, $systeminfo.LastLogonDate, $shared_system
			# We can use these values to change the message.
			if ($systeminfo.IsSharedSystem) {
				$has_shared_system = $true
			}
			if (-Not $systeminfo.IsAssigned) {
				$has_unassigned_system = $true
			}
			if ($systeminfo.IsIncompatible) {
				$has_incompatible_system = $true
			}
			if ($systeminfo.IsStale) {
				$has_stale_system = $true
			}
		}
		$msgHtml += "</table>"
		if ($has_shared_system -Or $has_stale_system -Or $has_unassigned_system -Or $has_incompatible_system) {
			$msghtml += @"
<ul>
"@
			if ($has_incompatible_system) {
				$msgHtml += @"
	<li>One or more of these systems may not be compatible with Windows 11. Please reach out to your LAN Admin for more information.</li>
"@
			}
			if ($has_shared_system) {
				$msgHtml += @"
	<li>One or more of these systems may be shared systems.</li>
"@
			}
			if ($has_stale_system) {
				$msgHtml += @"
	<li>One or more of these systems have not been active for over $STALE_PC_DAYS days.</li>
"@
			}
			if ($has_unassigned_system) {
				$msgHtml += @"
	<li>One or more of these systems are currently not assigned in our asset system. Please reach out to your LAN Admin to help confirm who owns this system.</li>
"@
			}
			$msgHtml += @"
</ul>
"@
		}
		
		$msgHtml += $EMAIL_FOOTER_HTML

		$email_success = $false
		if (-Not $DryRun) {
			<#
			$email_success = $true
			$emailParams = @{
				From = $EMAIL_FROM
				To = "mcarras8@jh.edu"
				Subject = $EMAIL_SUBJECT
				Body = $msgHtml
				Priority = "High"
				DeliveryNotificationOption = @("OnSuccess", "OnFailure")
				SmtpServer = $EMAIL_SMTP
				BodyAsHtml = $true
			}
			try {
				Send-MailMessage @emailParams
			} catch {
				Write-Error $_
				$error_count++
				$email_success = $false
			}
			#>
			Start-Sleep -Seconds $EMAIL_SLEEP_SECS
		}
		
		[PSCustomObject]@{
			ContactUser = $contactuser
			Systems = ($systems | Select @{N="Name"; Expression={$name = $_.Name; if($_.IsIncompatible) { $name += " (!)" }; $name}}).Name  -join ","
			SharedSystems = $has_shared_system
			StaleSystems = $has_stale_system
			UnassignedSystem = $has_unassigned_system
			IncompatibleSystems = $has_incompatible_system
			EmailSent = $email_success
		}
	}
	
	$emailed_users | Export-CSV -NoTypeInformation -Force $EXPORT_EMAILED_SYSTEMS_PATH
	Write-Host("[{0}] Emailed {1} users. Saved report to [{2}]." -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"), ($emailed_users | Measure).Count, $EXPORT_EMAILED_SYSTEMS_PATH)
}

# Next, we'll also remove all users from the notification GPO and re-add only the ones we're sending emails to.
<#
$users = $contactusers_systems.Keys | % { $_ }
if (-Not $DryRun -And -Not [string]::IsNullOrEmpty($NOTIFICATION_GROUP) -And ($users | Measure).Count -gt 0) {
	try {
		Get-ADGroupMember $NOTIFICATION_GROUP | ForEach-Object { Remove-ADGroupMember $NOTIFICATION_GROUP $_ -Confirm:$false }
		Add-ADGroupMember $NOTIFICATION_GROUP -Members $users -Confirm:$false
	} catch {
		Write-Error $_
		$error_count++
	}
}
#>

# Stop logging
Stop-Transcript -ErrorAction SilentlyContinue | Out-Null

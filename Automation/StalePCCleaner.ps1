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
$warningDate = (Get-Date).AddDays(-30)
# Date threshold to move system to retirement OU
# If system is already in retirement OU, it will be disabled instead
$retirementDate = (Get-Date).AddDays(-90)
# Date threshold to delete system out of AD entirely
# If not set or $null this action will always be skipped
#$removalDate = (Get-Date).AddDays(-365)

# Set the attribute synced from SOR for computer assignment.
# This field will be emailed if they are past $warningDays inactive.
$propAssignment = "extensionAttribute2"
# AD attribute for system form factor (Laptop, Desktop, etc.)
$propFormFactor = "extensionAttribute5"
# AD attribute for asset tag
$propAssetTag = "extensionAttribute1"

# The OU containing all contactable users.
$USER_OU = "OU=PEOPLE,DC=win,DC=ad,DC=jhu,DC=edu"

# Main searchbase
$OUComputers = 'OU=Computers,OU=USS,DC=win,DC=ad,DC=jhu,DC=edu'
# OU used to move retired computers to
$OUretirement = 'OU=USS-Retired,OU=Computers,OU=USS,DC=win,DC=ad,DC=jhu,DC=edu'
# List of OUs to exclude from processing.
$OUExclude = @('OU=USS-VPS,OU=Computers,OU=USS,DC=win,DC=ad,DC=jhu,DC=edu')

# The current date to use for the output file. Do not change!
$CurrentDate = ((Get-Date).ToString('MM-dd-yyyy'))
# Location and filename for storing CSV results
$CSVResultPath = "\\win.ad.jhu.edu\cloud\HSA$\ITServices\Reports\StalePCs"
$CSVResultFP = "$CSVResultPath\StalePCs-$CurrentDate.csv"
$CSVHeader = @("Name","LastLogonDate","PingResult","Action","Emailed","FormFactor","AssetTag")

# Automated email settings.
$EMAIL_ASSIGNEDUSER = $true
$EMAIL_SMTP = 'smtp.johnshopkins.edu'
$EMAIL_FROM = 'Jerome.Powell@jhu.edu'
$EMAIL_CC = @('Jerome.Powell@jhu.edu','mcarras8@jhu.edu')
$EMAIL_SUBJECT = "[USS-IT] Inactive System Alert"
$EMAIL_INTRO_HTML = @"
<p>This is an automated message.</p>
<p>You are receiving this email because one or more systems assigned to you have been offline for an extended period of time. To prevent future complications please login to your system as soon as possible. If are working remotely, you may need to leave the system connected to its charger and the internet overnight to fully update.</p>

<p>If you are no longer using this system, or think you may have received this email in error, please reply back to this email.</p>

<p>Thank you for your cooperation.</p>
"@
# Number of seconds to sleep in-between each email.
$EMAIL_SLEEP_SECS = 5

# Path and prefix for the Start-Transcript logfiles.
$LOGFILE_PATH = ".\Logs"
$LOGFILE_PREFIX = "stalepccleaner"
# Maximum number of days before rotating logfile.
$LOGFILE_ROTATE_DAYS = 90
# -- END CONFIGURATION --

# -- FUNCTION START --
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
$_ADUSERS=@{}
function Get-ADUserCached {
	param(
		[Parameter(Mandatory=$true,Position=0)]
		[string]$User,
		
		[Parameter(Mandatory=$false,Position=1)]
		[string]$Domain,
		
		[Parameter(Mandatory=$false,Position=2)]
		[string[]]$Properties=@("Company","Department")
	)
	
	$UPN = $User
	if (-Not [string]::IsNullOrEmpty($Domain) -And $UPN -notmatch "@") {
		$UPN += $Domain
	}
	$u = $_ADUSERS.$UPN
	if ([string]::IsNullOrEmpty($u.distinguishedname)) {
		try {
			$u = Get-ADUser -LDAPFilter "(|(SamAccountName=$UPN)(UserPrincipalName=$UPN))" -Properties $Properties			
			$_ADUSERS[$UPN] = $u
		} catch {
			throw $_
		}
	}
	return $u
}

<#
	.SYNOPSIS
	Returns a valid contact user given the assigned user attribute value.
	
	.DESCRIPTION
	Returns a valid contact user given the assigned user attribute value.
	
	.PARAMETER AssignedUser
	The assigned user as a string.
	
	.PARAMETER UserOU
	The OU containing all valid users.
	
	.OUTPUTS
	The email address for the assigned user.
#>
function Get-ValidContactEmail {
	param(
		[Parameter(Mandatory=$true,Position=0)]
		[AllowNull()]
		[string]$AssignedUser,
		
		[Parameter(Mandatory=$true,Position=1)]
		[string]$UserOU
	)
	
	$email = $null
	if($AssignedUser -match "@") {
		try {
			$u = Get-ADUserCached $AssignedUser -Properties "mail"
			if ($u.Enabled -And $u.distinguishedname -like "CN=*,$UserOU") {
				$email = $u.mail
			}
		} catch {
			throw $_
		}
	}
	return $email
}

# -- FUNCTION END --

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

$csvformat | Set-Content $CSVResultFP

# Scan Computers OU (SearchBase) for systems that have not been logged in since $warningDays.
# First ping the computers up to 3 times. If any pass, skip all other checks.
# Computers with LastLogonDate older than $retirementDate will be moved to the Retired OU if they haven't already.
# If they are already in the Retired OU, they will be disabled.
# If they are already disabled, and if $removalDate is set, they will be deleted out of AD.
$error_count = 0
$props = @("Name","LastLogonDate")
if (-Not [string]::IsNullOrWhitespace($propAssignment)) {
	$props += @($propAssignment)
}
if (-Not [string]::IsNullOrWhitespace($propAssetTag)) {
	$props += @($propAssetTag)
}
if (-Not [string]::IsNullOrWhitespace($propFormFactor)) {
	$props += @($propFormFactor)
}
# Hash table of users to email.
$contactUserSystems = @{}
# Hash table of systems to add messages for.
$logSystems = @{}
$comps = Get-ADComputer -Property $props -Filter {lastlogondate -lt $warningDate} -SearchBase $OUComputers
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
	if ((Test-Connection $_.name -Count 1 -ErrorAction SilentlyContinue) -Or (Test-Connection $_.name -Count 1 -ErrorAction SilentlyContinue) -Or (Test-Connection $_.name -Count 1 -ErrorAction SilentlyContinue)) {
		Write-Host("[{0}] Ping success for {1}" -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"), $_.Name)
		$pingResult = "Success"
	} Else {
		$pingResult = "Fail"
		If ($_.LastLogonDate -isnot [datetime]) {
			Write-Host("[{0}] Invalid LastLogonDate for [{1}], skipping" -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"), $_.DistinguishedName)
			$actionTaken = "Invalid LastLogonDate (Skipped)"
		} ElseIf (-Not [string]::IsNullOrEmpty($ou) -And $ou -in $OUExclude) {
			Write-Host("[{0}] Skipping [{1}] due to excluded OU" -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"), $_.DistinguishedName)
			$actionTaken = "OU Excluded (Skipped)"
		} Else {
			# Get contact email, if set.
			if (-Not [string]::IsNullOrWhitespace($propAssignment) -And -Not [string]::IsNullOrWhitespace($_.$propAssignment)) {
				try {
					$contactEmail = Get-ValidContactEmail $_.$propAssignment $USER_OU
				} catch {
					Write-Error $_
					$error_count++
				}
			}
			# If the system is past retirementDate.
			If ($_.LastLogonDate -le $retirementDate) {	
				# If the computer has not already been moved.
				if ($_.DistinguishedName -notlike "CN=*,$OUretirement") {
				  try {
					Write-Host("[{0}] Moving [{1}] to [{2}]" -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"), $_.DistinguishedName, $OUretirement)
					# Move-ADObject $_.DistinguishedName -TargetPath $OUretirement -WhatIf:$DryRun
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
					# If computer has already been disabled and we have $removalDate set.
					} ElseIf ($removalDate -is [datetime] -And $_.LastLogonDate -le $removalDate) {
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
	}
	
	# Add to list of systems to add to results file.
	$logSystems[$_.Name] = @{
		LastLogonDate = $_.LastLogonDate
		PingResult = $pingResult
		Action = $actionTaken
		ContactEmail = $contactEmail
		Emailed = ""
	}
	# Add optional attributes.
	if (-Not [string]::IsNullOrWhitespace($propAssetTag)) {
		# Only add assettag if its numeric.
		$assettag = ""
		if ($_.$propAssetTag -match "^\d+") {
			$assettag = $_.$propAssetTag
		}
		$logSystems[$_.Name]["AssetTag"] = $assettag
	}
	if (-Not [string]::IsNullOrWhitespace($propFormFactor)) {
		$logSystems[$_.Name]["FormFactor"] = $_.$propFormFactor
	}
	if (-Not [string]::IsNullOrWhitespace($propAssignment)) {
		$logSystems[$_.Name]["AssignedUser"] = $_.$propAssignment
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
			Write-Host("[{0}] Emailing [{1}] for systems: {2}" -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"), $email, $systems -join ", ")
			$msgHtml = $EMAIL_INTRO_HTML
			$msgSystemTable = ""
			
			# -- BODY - SYSTEM TABLE --
			$validSystems = $false
			foreach($systemName in $systems) {
				if (-Not [string]::IsNullOrWhitespace($systemName)) {
					$system = $logSystems[$systemName]
					if ($system -ne $null) {
						$validSystems = $true
						$msgSystemTable += @"
			<tr>
				<td>$systemName</td><td>$($system.AssetTag)</td><td>$($system.FormFactor)</td><td>$($system.LastLogonDate)</td>
			</tr>
"@
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
Write-Host("[{0}] Saving results for {1} systems to [{2}]" -f (Get-Date -Format "yyyy/MM/dd HH:mm:ss"), ($logSystemsObj | Measure).Count, $CSVResultFP)
if ($logSystemsObj -ne $null) {
	$logSystemsObj | Select $CSVHeader | Export-CSV $CSVResultFP -NoTypeInformation -Force
}

#After the USS OUs are scanned and any stale PCs moved to the retiring OU, the script will scan the Retiring OU to detect any computers that should not be in there, and disable and delete old objects
#If a computer can be pinged, the CSV will be updated and the LAN Admin should investigate.
#If a computer is not reachable but it has been logged into within 90 days, the CSV will be updated and the LAN Admin should investigate.
#If a computer has not been logged into for 365 days (the $removal date), AND the lastlogondate field is not null, it will be deleted from AD
#If a computer has not been logged into for 90 days and it is still enabled in AD, it will be disabled.
#If a computer is already disabled, it will be marked as "disabled" on the CSV
#Anything that does not fit the criteria above will be marked on the CSV as "Unknown"

# Send-MailMessage -From 'HSA IT Services <hsaitservices@jhu.edu>' -To 'HSA IT Services <hsaitservices@jhu.edu>' -Subject 'Stale PC Results' -Body "Results of Stale PC script attached" -Attachments $CSVResultFP -Priority High -DeliveryNotificationOption OnSuccess, OnFailure -SmtpServer 'smtp.johnshopkins.edu' -whatif

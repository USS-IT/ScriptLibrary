<#
	.SYNOPSIS
	Displays a Windows toast notification if the system has been inactive for too long or is otherwise out of compliance.
	
	.DESCRIPTION
	Displays a Windows toast notification if the system has been inactive for too long or is otherwise out of compliance.

	.PARAMETER DaysInactive
	Number of days the computer has not checked in with AD / SCCM. Default is 30.
	
	.PARAMETER IgnoreMCM
	Don't check / notify about the MCM Heartbeat / Discovery Cycle.
#>
Param (
	[Parameter(Mandatory=$false)]
	[int]$DaysInactive=30,
	
	[Parameter(Mandatory=$false)]
	[switch]$IgnoreMCM
)

# -- START FUNCTIONS ---
<#
	.SYNOPSIS
	Displays a Windows toast notification for the current user.
	
	.DESCRIPTION
	Displays a Windows toast notification for the current user.

	.PARAMETER Text
	Text to display in toast notification. Max ~175 characters.
	
	.PARAMETER Title
	Title to use for toast notification. Default is "Alert from IT"
	
	.PARAMETER ClickableLink
	Optional URL link to assign to a clickable button.
	
	.PARAMETER ClickableLinkText
	Text for the clickable link button.
	
	.PARAMETER LauncherID
	Required AppID to use for the notification. You can use any from Get-StartApps. Default is "Microsoft.SoftwareCenter.DesktopToasts".
	
	.PARAMETER ShowSnoozeTimer
	Determines whether we will display a snooze timer and use the Reminder scenario (persistent toast).
	
	.PARAMETER Duration
	Duration in minutes before automatically dismissing the toast. Only used when ShowSnoozeTimer is NOT set. Default is 15 minutes.
	
	.PARAMETER ClearOldNotifications
	Automatically clear old notifications for the given LauncherID from the user's toast history in the Notification Center.
#>
function Show-Toast {
	[cmdletbinding(DefaultParametersetName='None')]
	Param (
		[Parameter(Mandatory=$true,
			ValueFromPipeline=$true,
			Position=0)]
		[string]$Text,
		[string]$Title = "Alert from IT",
		[Parameter(Mandatory=$false,
			ParameterSetName="ClickableLink")]
		[string]$ClickableLink,
		[Parameter(Mandatory=$true,
			ParameterSetName="ClickableLink")]
		[string]$ClickableLinkText,
		[string]$LauncherID = "Microsoft.SoftwareCenter.DesktopToasts",
		[switch]$ShowSnoozeTimer,
		[uint32]$Duration = 15,
		[switch]$ClearOldNotifications
	)

	[Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime] | Out-Null
	[Windows.Data.Xml.Dom.XmlDocument, Windows.Data.Xml.Dom.XmlDocument, ContentType = WindowsRuntime] | Out-Null

	if ([string]::IsNullOrWhitespace($Text)) {
		throw "Toast text is all whitespace or empty"
	}
	
	# Clear out old notifications, if set.
	If($ClearOldNotifications) {
		try {
			$ToastHistory = [Windows.UI.Notifications.ToastNotificationManager]::History
			$ToastHistory.Clear($LauncherID)
		} catch {
			Write-Error $_
		}
	}

	# NOT cast to XML
	if ($ShowSnoozeTimer) {
		$Actions = @"
			<input id="snoozeTime" type="selection" defaultInput="60">
				<selection id="60" content="Snooze for 1 hour"/>
				<selection id="240" content="Snooze for 4 hours"/>
				<selection id="1440" content="Snooze for 1 day"/>
			</input>
			<action activationType="system" arguments="snooze" hint-inputId="snoozeTime" content="" />
"@
		if (-Not [string]::IsNullOrWhitespace($ClickableLink) -And -Not [string]::IsNullOrWhitespace($ClickableLinkText)) {
			$Actions += @"
			
			<action arguments="$ClickableLink" content="$ClickableLinkText" activationType="protocol" />
"@
		}
		
		# Main template
		[xml]$ToastTemplateXml = @"
		<toast scenario="reminder">
			<visual>
				<binding template="ToastGeneric">
					<text id="1">$Title</text>
					<text id="2">$Text</text>
				</binding>
			</visual>
			<actions>
				$Actions
			</actions>
		</toast>
"@
	} else {
		if (-Not [string]::IsNullOrWhitespace($ClickableLink) -And -Not [string]::IsNullOrWhitespace($ClickableLinkText)) {
			# NOT cast to XML
			$Actions = @"
				<action arguments="$ClickableLink" content="$ClickableLinkText" activationType="protocol" />
"@
		} else {
			$Actions = ""
		}
		
		# Main template
		[xml]$ToastTemplateXml = @"
		<toast>
			<visual>
				<binding template="ToastGeneric">
					<text id="1">$Title</text>
					<text id="2">$Text</text>
				</binding>
			</visual>
			<actions>
				$Actions
			</actions>
		</toast>
"@
	}

	$SerializedXml = [Windows.Data.Xml.Dom.XmlDocument]::New()
	$SerializedXml.LoadXml($ToastTemplateXml.OuterXml)

	$Toast = [Windows.UI.Notifications.ToastNotification]::new($SerializedXml)
	if (-Not $ShowSnoozeTimer) {
		$Toast.ExpirationTime = [DateTimeOffset]::Now.AddMinutes($Duration)
	}
	[Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier($LauncherID).Show($Toast)
}
# -- END FUNCTIONS ---

$staleDate = (Get-Date).AddDays(-$DaysInactive)

# Check AD using ADSI
$rootDSE  = [ADSI]"LDAP://RootDSE"
$domainDN = $rootDSE.defaultNamingContext

$computerName = $env:COMPUTERNAME + '$'
try {
	$searcher = New-Object System.DirectoryServices.DirectorySearcher
	$searcher.SearchRoot = [ADSI]"LDAP://$domainDN"
	$searcher.Filter = "(&(objectCategory=computer)(sAMAccountName=$computerName))"
	$searcher.PropertiesToLoad.Add("lastLogonTimestamp") | Out-Null
	# Setup timeout options
	$searcher.ClientTimeout = New-TimeSpan -Seconds 300
	$searcher.ServerTimeLimit = New-TimeSpan -Seconds 300
	$searcher.ReferralChasing = "None"

	$result = $searcher.FindOne()

	if ($result -and $result.Properties["lastLogonTimestamp"].Count -gt 0) {
		$lastSeenDate = [DateTime]::FromFileTimeUtc(
			[Int64]$Result.Properties["lastLogonTimestamp"][0]
		)

		if ($lastSeenDate -ge $staleDate) {
			Write-Output ("ACTIVE in AD last logon " + $lastSeenDate.ToLocalTime())
		} else {
			Write-Output ("STALE in AD last logon " + $lastSeenDate.ToLocalTime())
		}
	} else {
		Write-Output "No lastLogonTimestamp present"
	}
} catch {
	Write-Output "Domain not available or AD timeout"
}

# Check MCM using WMI
if (-Not $IgnoreMCM) {
	try {
		$heartbeat = Get-WmiObject -Namespace root\ccm\invagt -Class InventoryActionStatus -Filter 'InventoryActionID="{00000000-0000-0000-0000-000000000003}"'
		
		if ($heartbeat -and $heartbeat.LastCycleStartedDate) {
			$lastHeartbeatDate = [Management.ManagementDateTimeConverter]::ToDateTime($heartbeat.LastCycleStartedDate)
			
			if ($lastHeartbeatDate -ge $staleDate) {
				Write-Output ("ACTIVE in MCM last heartbeat " + $lastHeartbeatDate.ToLocalTime())
			} else {
				Write-Output ("STALE in MCM last heartbeat " + $lastHeartbeatDate.ToLocalTime())
			}
		} else {
			Write-Output "Heartbeat Discovery has never run"
		}
	}
	catch {
		Write-Output "MCM client not present or WMI unavailable"
	}
}
	
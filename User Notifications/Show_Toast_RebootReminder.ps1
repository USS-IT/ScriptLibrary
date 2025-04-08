<#
	.SYNOPSIS
	Shows a toast notification if the users system has not restarted in the given length of time (7 days by default).
	
	.DESCRIPTION
	Shows a toast notification if the users system has not restarted in the given length of time (7 days by default).

	.PARAMETER RestartAlertDays
	Number of days before an alert is displayed. Default is 7.
	
	.PARAMETER ToastTitle
	Optional title for the toast notification. Default is "Alert from USS IT".
	
	.PARAMETER ToastText
	Optional text for the toast notification. {0} will be substituted with # of days since last restart. Must be <= 175 characters. Default is "Your system requires a restart to keep running smoothly. Last Restart: {0} days ago."
	
	.PARAMETER ClickableLink
	Optional URL link to assign to a clickable button.
	
	.PARAMETER ClickableLinkText
	Required text for the clickable button if -ToastClickableLink is given.
	
	.PARAMETER Duration
	Duration (in minutes) before the alert is automatically dismissed. Default is 60 minutes.
	
	.PARAMETER ShowSnoozeTimer
	Show a snoozable timer instead of automatically dismissing.
	
	.PARAMETER ClearOldNotifications
	Clears old notifications for the default app ID before displaying the new toast. Useful with -ShowSnoozeTimer.
	
	.NOTES
	Must be run under the user's context.
	
	Created: 3-23-23
	Author: mcarras8
#>
[cmdletbinding(DefaultParametersetName='None')]
param(
	[Parameter(Mandatory=$false)]
	[uint64]$RestartAlertDays=7,
	
	[Parameter(Mandatory=$false)]
	[string]$ToastTitle="Alert from USS IT",
	
	[Parameter(Mandatory=$false)]
	[string]$ToastText="Your system requires a restart to keep running smoothly. Last Restart: {0} days ago.",
	
	[Parameter(Mandatory=$false,
	 ParameterSetName="ClickableLink")]
	[string]$ClickableLink,
	
	[Parameter(Mandatory=$true,
	 ParameterSetName="ClickableLink")]
	[string]$ClickableLinkText,
	
	[Parameter(Mandatory=$false)]
	[uint32]$Duration=60,
	
	[Parameter(Mandatory=$false)]
	[switch]$ShowSnoozeTimer,
	
	[Parameter(Mandatory=$false)]
	[switch]$ClearOldNotifications
)

# -- START FUNCTIONS --
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
# -- END FUNCTIONS --

If($ToastText.Length -gt 175) {
	Write-Warning "-ToastText is too long, text may be truncated"
}

try {
	# Grab uptime from WMI. If 
	$uptimeDays = ((Get-Date) - [Management.ManagementDateTimeConverter]::ToDateTime((Get-WmiObject Win32_OperatingSystem | Select -ExpandProperty LastBootUpTime))) | Select -ExpandProperty Days
	if ($uptimeDays -ge $RestartAlertDays) {
		# Add extra parameters before calling the function.
		$extraParams = @{}
		if (-Not [string]::IsNullorWhitespace($ClickableLink)) {
			$extraParams.Add("ClickableLink", $ClickableLink)
			$extraParams.Add("ClickableLinkText", $ClickableLinkText)
		}
		if ($ShowSnoozeTimer) {
			$extraParams.Add("ShowSnoozeTimer", $true)
		} else {
			$extraParams.Add("Duration", $Duration)
		}
		if ($ClearOldNotifications) {
			$extraParams.Add("ClearOldNotifications", $true)
		}
		Show-Toast -Text ($ToastText -f $uptimeDays) @extraParams
	}
} catch {
	Write-Error $_
}
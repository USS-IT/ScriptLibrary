<#
	.SYNOPSIS
	Shows a toast notification if the Windows version is equal or older than -EOLVER.
	
	.DESCRIPTION
	Shows a toast notification if the Windows version is equal or older than -EOLVER.
	
	.PARAMETER EOLVER
	Required build version of Windows to check against (e.g., 22621).
	
	.PARAMETER ToastTitle
	Optional title for the toast notification. Default is "Alert from USS IT".
	
	.PARAMETER ToastText
	Optional text for the toast notification. {0} will be substituted with the current build. Must be <= 175 characters. Default is "The version of Windows on your computer requires an upgrade to avoid being blocked by Central IT. Please click on the Instructions button for more info."
	
	.PARAMETER ToastTextLowSpace
	Optional text to show if it detects less than 23GB of free space to install the upgrade. {0} will be substituted with the amount of free space. Must be <= 175 characters. Default is "Your computer does not have enough free space for a required upgrade. Please click on the Instructions button and see the Troubleshooting section for more info."
	
	.PARAMETER ClickableLink
	Optional URL link to assign to a clickable button. NOTE: THIS URL must be url-encoded! Default: https://t.jh.edu/USS-WindowsUpgrade
	
	.PARAMETER ClickableLinkText
	Required text for the clickable button if -ClickableLink is given. Default: Instructions
	
	.PARAMETER Duration
	Duration (in minutes) before the alert is automatically dismissed and -ShowSnoozeTimer is not given. Default is 15 minutes.
	
	.PARAMETER ShowSnoozeTimer
	Show a snoozable timer instead of automatically dismissing.
	
	.PARAMETER ClearOldNotifications
	Clears old notifications for the default app ID before displaying the new toast. Useful with -ShowSnoozeTimer.
	
	.NOTES
	Must be run under the user's context.
	
	Author: mcarras8
	
	04-09-25 mcarras8 Added more parameters and renamed to "Show_Toast_WinUpgrade".
	03-24-23 mcarras8 Initial creation
#>
[cmdletbinding(DefaultParametersetName='None')]
param(
	[Parameter(Mandatory=$true)]
	[uint64]$EOLVER,
	
	[Parameter(Mandatory=$false)]
	[string]$ToastTitle="Alert from USS IT",
	
	[Parameter(Mandatory=$false)]
	[string]$ToastText="The version of Windows on your computer requires an upgrade to avoid being blocked by Central IT. Please click on the Instructions button for more info.",
	
	[Parameter(Mandatory=$false)]
	[string]$ToastTextLowSpace="Your computer does not have enough free space for a required Windows upgrade. Please open a Help Desk ticket for assistance.",
	
	[Parameter(Mandatory=$false,
	 ParameterSetName="ClickableLink")]
	[string]$ClickableLink="https://t.jh.edu/USS-WindowsUpgrade",
	
	[Parameter(Mandatory=$true,
	 ParameterSetName="ClickableLink")]
	[string]$ClickableLinkText="Instructions",
	
	[Parameter(Mandatory=$false)]
	[uint32]$Duration=15,
	
	[Parameter(Mandatory=$false)]
	[switch]$ShowSnoozeTimer,
	
	[Parameter(Mandatory=$false)]
	[switch]$ClearOldNotifications
)
# -- START CONFIGURATION --
# The amount of GB required for the upgrade.
$UPGRADE_SPACE_REQUIRED_GB = 23
# -- END CONFIGURATION --

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

# Grab version from registry.
$ver = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion" -ErrorAction Stop).CurrentBuild
if (-Not [string]::IsNullOrWhitespace($ver) -And $ver -le $EOLVER) {
	$toastmsg = ($ToastText -f $ver)
	# Display a different message if computer is low on free space.
	try {
		if (-Not [string]::IsNullOrWhitespace($ToastTextLowSpace)) {
			$freespace = Get-Volume -DriveLetter (${ENV:SYSTEMDRIVE} -replace ":","") | Select -ExpandProperty SizeRemaining
			if ($freespace -ne $null) {
				$freespace = [math]::round($freespace /1Gb, 2)
				if ( $freespace -le $UPGRADE_SPACE_REQUIRED_GB) {
					$toastmsg = ($ToastTextLowSpace -f $freespace)
				}
			}
		}
	} catch {}
	
	If($toastmsg.Length -gt 175) {
		Write-Warning "Toast text is > 175 characters, text may be truncated"
	}

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
		
	# Show a dismissable toast notification with a clickable link for the current user.
	Show-Toast -Text $toastmsg @extraParams
}

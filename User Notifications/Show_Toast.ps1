# MJC 4-27-22
# Shows a Windows 10 Toast Notification. Must be run under the user's account.
function Show-Toast {
	[cmdletbinding(DefaultParametersetName='None')]
	Param (
		[Parameter(Mandatory=$true,
			ValueFromPipeline=$true,
			Position=0)]
		[string]$Text,
		[string]$Title = "Alert from IT",
		# Displays button link and text along with Dismiss button
		[Parameter(Mandatory=$false,
			ParameterSetName="ClickableLink")]
		[string]$ClickableLink,
		[Parameter(Mandatory=$true,
			ParameterSetName="ClickableLink")]
		[string]$ClickableLinkText,
		# Launcher ID, such as the AppIDs from Get-StartApps
		[string]$LauncherID = "Microsoft.SoftwareCenter.DesktopToasts",
		# Determines whether we will display a Snooze timer and use the Reminder scenario (persistent toast).
		[switch]$ShowSnoozeTimer,
		# Duration (Minutes) before automatically being dismissed. Only used when ShowSnoozeTimer is NOT set.
		[uint32]$Duration = 15
	)

	[Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime] | Out-Null
	[Windows.Data.Xml.Dom.XmlDocument, Windows.Data.Xml.Dom.XmlDocument, ContentType = WindowsRuntime] | Out-Null

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

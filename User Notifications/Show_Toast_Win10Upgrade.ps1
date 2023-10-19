# MJC 3-24-23
# Shows a toast notification if the version is equal or older than $EOLVER.
# Must be run under the user's context.
$EOLVER=19042
# Link for custom button which redirects to given URL.
# NOTE: THIS URL must be url-encoded!
# Add-Type -AssemblyName System.Web
# [System.Web.HttpUtility]::HtmlEncode($TOAST_CLICKABLE_LINK)
$TOAST_CLICKABLE_LINK="https://livejohnshopkins.sharepoint.com/sites/USS-IT/Windows-EOL-List/Windows%20EOL%20list/Forms/AllItems.aspx?id=%2Fsites%2FUSS%2DIT%2FWindows%2DEOL%2DList%2FWindows%20EOL%20list%2FWindow%20instructions%2FWindows%20Upgrade%20for%20Laptops%2Epdf&amp;parent=%2Fsites%2FUSS%2DIT%2FWindows%2DEOL%2DList%2FWindows%20EOL%20list%2FWindow%20instructions"
# Text for button.
$TOAST_CLICKABLE_LINK_TEXT="Instructions"
# Title for toast window.
$TOAST_TITLE="Alert from USS IT"
# {0} will display the current Windows build #.
# NOTE: Max length of TOAST_TEXT cannot exceed ~175-185 characters.
$TOAST_TEXT="The version of Windows on your computer requires an upgrade to avoid being blocked by Central IT. Please click on the $TOAST_CLICKABLE_LINK_TEXT button for more info."

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

If($TOAST_TEXT.Length -gt 175) {
	Write-Warning "TOAST_TEXT is too long, text may be truncated"
}

# Grab version from registry.
$ver = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion" -ErrorAction Stop).CurrentBuild
if (-Not [string]::IsNullOrWhitespace($ver) -And $ver -le $EOLVER) {
	Show-Toast -Text ($TOAST_TEXT -f $ver) -ShowSnoozeTimer -ClickableLink $TOAST_CLICKABLE_LINK -ClickableLinkText $TOAST_CLICKABLE_LINK_TEXT
}

# MJC 4-6-22
# Your company's tenant ID
$OD_TENANT_ID = "9fa4f438-b1e6-473b-803f-86f8aedf0dec"

# Creates a button with given text which opens in default browser
$NOTIFY_LINK = "https://www.google.com"
$NOTIFY_LINK_TEXT = "More Info"
$NOTIFY_TITLE = "Notification: OneDrive needs your attention"
$NOTIFY_TEXT = "OneDrive does not appear to be setup to automatically backup your files. Please visit $NOTIFY_LINK for more info."
# Assign the toast to OneDrive
$NOTIFY_LAUNCHERID = "Microsoft.SkyDrive.Desktop"

function Show-Toast {
	[cmdletbinding(DefaultParametersetName='None')]
	Param (
		[Parameter(Mandatory=$true,
			ValueFromPipeline=$true,
			Position=0)]
		[string]$Text,
		[string]$Title = "Notification from IT",
		# Displays button link and text along with Dismiss button
		[Parameter(Mandatory=$false,
			ParameterSetName="ClickableLink")]
		[string]$ClickableLink,
		[Parameter(Mandatory=$true,
			ParameterSetName="ClickableLink")]
		[string]$ClickableLinkText,
		# Duration (Minutes) before automatically being dismissed
		[int]$Duration = 1,
		# Launcher ID from Get-StartApps
		[string]$LauncherID = "Microsoft.SoftwareCenter.DesktopToasts"
	)

	[Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime] | Out-Null
	[Windows.Data.Xml.Dom.XmlDocument, Windows.Data.Xml.Dom.XmlDocument, ContentType = WindowsRuntime] | Out-Null
	
	if (-Not [string]::IsNullOrWhitespace($ClickableLink) -And -Not [string]::IsNullOrWhitespace($ClickableLinkText)) {
		# NOT cast to XML
		$ClickableLinkActions = @"
			<action arguments="$ClickableLink" content="$ClickableLinkText" activationType="protocol" />
			<action arguments="dismiss" content="" activationType="system"/>
"@
	} else {
		$ClickableLinkActions = ""
	}
	
	# Main template
	[xml]$ToastTemplateXml = @"
		<toast>
			<visual>
				<binding template="ToastImageAndText03">
					<text id="1">$Title</text>
					<text id="2">$Text</text>
				</binding>
			</visual>
			<actions>
				$ClickableLinkActions
			</actions>
		</toast>
"@

	$SerializedXml = [Windows.Data.Xml.Dom.XmlDocument]::New()
	$SerializedXml.LoadXml($ToastTemplateXml.OuterXml)

	$Toast = [Windows.UI.Notifications.ToastNotification]::new($SerializedXml)
	$Toast.ExpirationTime = [DateTimeOffset]::Now.AddMinutes($Duration)
	[Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier($LauncherID).Show($Toast)
}

# Show a Toast if the KFMState registry key (Known Folder Move) isn't set on the first matching business account
$ip = Get-ChildItem "HKCU:\Software\Microsoft\OneDrive\Accounts" | Get-ItemProperty -ErrorAction SilentlyContinue | Where-Object {$_.Business -eq 1 -And $_.ConfiguredTenantId -eq $OD_TENANT_ID} | Select -First 1
# Set this to ($ip -And -Not $ip.KFMState) if you only want to check when OneDrive is already signed in
if ($ip -Or -Not $ip.KFMState) {
	Show-Toast -Text $NOTIFY_TEXT -Title $NOTIFY_TITLE -ClickableLink $NOTIFY_LINK -ClickableLinkText $NOTIFY_LINK_TEXT -Duration 3 -LauncherID $NOTIFY_LAUNCHERID
}

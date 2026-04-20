<#
	.SYNOPSIS
	Displays task sequence errors from the main SMSTS logfile and other warnings.
	
	.DESCRIPTION
	Displays task sequence errors from the main SMSTS logfile and other warnings, sends an email notification, and logs the results.
	
	.PARAMETER NotificationLogPath
	Path to the local log file to store the notifications. Defaults to OSDFailure_Notifications.log in _SMSTSLogPath.
	
    .PARAMETER PopupTimeout
	Timeout on the pop-up shown at the end if any errors are detected. Default is 0 (never expire).

	.PARAMETER DisablePopup
	Don't show a pop-up message when errors are detected.
	
	.PARAMETER DisableEmail
	Don't email the results to the email address if it's set.
	
	.PARAMETER EmailSender
	Override the default email sender of ussitservices@jhu.edu.
	
	.PARAMETER EmailSmtp
	Override the default email server of smtp.johnshopkins.edu.
	
    .NOTES
	Author: mcarras8
	
	--Custom configurable TS Variables used--
	Email notifications: XOSDNotifyEmail or XOSDNotifyEmailChoice (in that order)
	Ping Address: XCheckNetAddress or "google.com"
#>
param(
    [Parameter(Mandatory=$false)]
	[string]$NotificationLogPath="OSDTS_Notifications.log",
	
    [Parameter(Mandatory=$false)]
	[int]$PopupTimeout=0,
	
	[Parameter(Mandatory=$false)]
	[switch]$DisablePopup,

	[Parameter(Mandatory=$false)]
	[switch]$DisableEmail,
	
	[Parameter(Mandatory=$false)]
	[string]$EmailSender='USS IT Services <ussitservices@jhu.edu>',
	
	[Parameter(Mandatory=$false)]
	[string]$EmailSmtp='smtp.johnshopkins.edu'
)

try {
	$tsenv = New-Object -ComObject Microsoft.SMS.TSEnvironment
	$logpath = $tsenv.Value("_SMSTSLogPath")
	$logfile = $logpath + "\smsts.log"
} catch {
	throw $_
}

if ($NotificationLogPath -notlike '*\*') {
	$_notificationLogPath = "${logpath}\${NotificationLogPath}"
} else {
	$_notificationLogPath = $NotificationLogPath
}
Start-Transcript -Path $_notificationLogPath -Force

# Check logfile for errors.
$errmsgs = @()
$errorCount = 0
Select-String -Path $logfile "<\!\[LOG\[(Failed to run the action:[^\]]+)]" -AllMatches | Foreach-Object {$_.Matches} | Foreach-Object {
	if ($_.Groups.Count -gt 1) {
		$errmsgs += @("* " + $_.Groups[1].Value)
		$errorCount++
	}
}

# Check for internet connectivity.
# Set this TSVar to "false" if we want to skip checking internet connectivity.
$addr = $tsenv.Value("XCheckNetAddress")
if([string]::IsNullOrWhitespace($addr)) {
	$addr = "google.com"
}
if ($addr -ne "false" -And -Not (
	(Test-Connection -ComputerName $addr -Count 1 -Quiet) -Or 
	(Test-Connection -ComputerName $addr -Count 1 -Quiet) -Or 
	(Test-Connection -ComputerName $addr -Count 1 -Quiet))
) {
	$errmsgs += @("* Cannot reach [$addr]. System may be missing MAC registration.")
	$errorCount++
}

$xosdcompleted = $tsenv.Value("XOSDCompleted")
$osdstatus = $tsenv.Value("OSDStatus")
$_smstsinwinpe = $tsenv.Value("_SMSTSInWinPE")
$xosdruntimestart = $tsenv.Value("XOSDRuntimeStart")
$systemName = $tsenv.Value("XFinalComputerName")

$xosdrenamesuccess = $tsenv.Value("XOSDRenameSuccess")
$xosdmovesuccess = $tsenv.Value("XOSDMoveSuccess")
$XHasTSOSDGUIRun = $tsenv.Value("XHasTSOSDGUIRun")
$XHasTSPartRun = $tsenv.Value("XHasTSPartRun")
$XHasTSNetRun = $tsenv.Value("XHasTSNetRun")
$XHasTSSoftwareRun = $tsenv.Value("XHasTSSoftwareRun")
$XHasTSPostSetupRun = $tsenv.Value("XHasTSPostSetupRun")
$XHasTSOSDGUICompleted = $tsenv.Value("XHasTSOSDGUICompleted")
$XHasTSPartCompleted = $tsenv.Value("XHasTSPartCompleted")
$XHasTSNetCompleted = $tsenv.Value("XHasTSNetCompleted")
$XHasTSSoftwareCompleted = $tsenv.Value("XHasTSSoftwareCompleted")
$XHasTSPostSetupCompleted = $tsenv.Value("XHasTSPostSetupCompleted")

# Get the current runtime in hours.
# We use a WebRequest to avoid time syncing issues.
try {
	$xosdruntimestart = $xosdruntimestart -as [DateTime]
	if ($xosdruntimestart -is [DateTime]) {
		$internetDate = (Invoke-WebRequest -UseBasicParsing -Uri "http://johnshopkins.edu" -Method Head -TimeoutSec 120).Headers.Date -as [DateTime]
		if ($internetDate -is [DateTime]) {
			$runTimeHours = [math]::Round(($internetDate - $xosdruntimestart).TotalHours, 2)
		}
	}
} catch {
	Write-Error $_
}

# Show an error if any required TS hasn't run.
if ( $XHasTSOSDGUIRun -ne "true" ) {
	$errmsgs += @("* Child TS Not Run: OSD GUI")
	$errorCount++
}
if ( $XHasTSPartRun -ne "true" ) {
	$errmsgs += @("* Child TS Not Run: Partition and Format")
	$errorCount++
}
<#
if ( $XHasTSNetRun -ne "true" ) {
	$errmsgs += @("* Child TS Not Run: Network Drivers")
	$errorCount++
}
#>
if ( $XHasTSSoftwareRun -ne "true" ) {
	$errmsgs += @("* Child TS Not Run: Software Install")
	$errorCount++
}
if ( $XHasTSPostSetupRun -ne "true" ) {
	$errmsgs += @("* Child TS Not Run: Post-Setup")
	$errorCount++
}

# Formats error messages for the pop-up and email.
$errmsgs = $errmsgs -join "`r`n"
if( $xosdcompleted -eq "true" ) {
    $errmsgs += "`r`n>> NOTE: All other steps should have completed successfully."
} else {
	$errmsgs += "`r`n>> NOTE: May have exited early due to failed actions."
}
if ($runtimeHours) {
	$errmsgs += "`r`n>> Total Runtime Hours: $runTimeHours"
}
	
# Sends an email notificaton if an email address was set.
if (-Not $DisableEmail) {	
	$emailTo = $tsenv.Value("XOSDNotifyEmail")
	Write-Host "XOSDNotifyEmail: $emailTo"
	if ([string]::IsNullOrWhitespace($emailTo)) {
		$xosdnotifyemailchoice = $tsenv.Value("XOSDNotifyEmailChoice")
		Write-Host "XOSDNotifyEmailChoice: $xosdnotifyemailchoice"
		if ($xosdnotifyemailchoice -ne "NONE") {
			$emailTo = $xosdnotifyemailchoice
		}
	}
	if (-Not [string]::IsNullOrWhitespace($emailTo) -And $emailTo -like "*@*" -And -Not [string]::IsNullOrEmpty($EmailSender)) {
		if ($errorCount -le 0 -And $osdstatus -ne "Failure") {
			$emailSubject = "Imaging Success for $systemName"
			$emailPriority = "Normal"
			$body = $emailSubject + ". All required steps have been completed. `r`n`r`nIf driver or software updates are selected, they will be installed prior to completing the operating system deployment task sequence (may take 1+ hours)."
			if ($runtimeHours) {
				$body += "`r`n`r`nTotal Runtime Hours: $runtimeHours"
			}
		} else {
			$_smtspackagename = $tsenv.Value("_SMSTSPackageName")
			$osdlogpath = ('{0}\{1}' -f '\\win.ad.jhu.edu\Data\osdlogs$', $_smtspackagename)
			$emailSubject = "Imaging Failure for $systemName"
			$emailPriority = "High"
			$body = $emailSubject + ".`r`n" + $errmsgs + "`r`n`r`nSystem Name: " + $systemName + "`r`nLog Path: " + $osdlogpath
		}
			  
		$emailParams = @{
			From = $EmailSender
			To = $emailTo
			CC = $EmailSender
			Subject = $emailSubject
			Body = $body
			Priority = $emailPriority
			DeliveryNotificationOption = @("OnSuccess", "OnFailure")
			SmtpServer = $EmailSmtp
		}
		
		try {
			Send-MailMessage @emailParams -ErrorAction Stop
			Write-Host "Email sent to [$emailTo]"
		} catch {
			Write-Error $_
		}
	}
}

# Log all the errors and warnings.
try {
	$currentIP = (Get-NetRoute -DestinationPrefix '0.0.0.0/0' | Sort-Object RouteMetric | Select-Object -First 1 | Get-NetIPAddress -AddressFamily IPv4).IPAddress
} catch {
	Write-Error $_
}
Write-Host "System Final Name: $systemName"
Write-Host "OSDStatus: $osdstatus"
Write-Host "Current IP: $currentIP"
Write-Host "XOSDCompleted: $xosdcompleted"
Write-Host "Total Runtime Hours: $runtimeHours"
Write-Host "Rename Computer Success (XOSDRenameSuccess): $xosdrenamesuccess"
Write-Host "Move Computer Success (XOSDMoveSuccess): $xosdmovesuccess"
Write-Host "OSDGUI TS Completed: $XHasTSOSDGUICompleted"
Write-Host "Partition TS Completed: $XHasTSPartCompleted"
Write-Host "Net Drivers TS Completed: $XHasTSNetCompleted"
Write-Host "Software TS Completed: $XHasTSSoftwareCompleted"
Write-Host "Post-Setup TS Completed: $XHasTSPostSetupCompleted"
Write-Host "Error Count: $errorCount"
Write-Host $errmsgs

# Further logging if we have any errors.
if ($errorCount -gt 0 -Or $osdstatus -eq "Failure") {
	if ($_smstsinwinpe -ne "true") {
		try {
			# Check current WMI consistency
			$repo = & winmgmt /verifyrepository 2>&1
			Write-Host $repo
		} catch {
			Write-Error $_
		}
		
		# Basic WMI query
		try {
			Get-CimInstance Win32_ComputerSystem -ErrorAction Stop | Out-Null
			Write-Host "Basic WMI Query: OK"
		}
		catch {
			Write-Host ("Basic WMI Query Error: " + $_.Exception.Message)
		}

		try {
			# Check and report on WBEM Provider errors
			$wbemErrorCount = Get-WinEvent -FilterHashtable @{
				LogName   = 'Microsoft-Windows-WMI-Activity/Operational'
				Level     = 2
				StartTime = (Get-Date).AddMinutes(-15)
			} -MaxEvents 50 | Where-Object {
			  $_.Message -match 'Provider|WBEM_E'
			} | Measure-Object | Select -ExpandProperty Count
			Write-Host ("Recent WBEM_E (WMI) errors: $wbemErrorCount")
		} catch {
			Write-Error $_
		}
	}
}

# If we have any actual errors then display a pop-up as well.
if ($errorCount -gt 0 -And -Not $DisablePopup) {
  $wshell = New-Object -ComObject Wscript.Shell
  $wshell.Popup($errmsgs, $PopupTimeout, "Task Sequence Errors - $errorCount", 16+4096)
}

Stop-Transcript | Out-Null
[System.Environment]::Exit(0)
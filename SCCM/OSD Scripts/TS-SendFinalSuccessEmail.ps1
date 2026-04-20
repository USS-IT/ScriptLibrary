<#
	.SYNOPSIS
	Sends a final email on imaging completion.
	
	.DESCRIPTION
	Sends a final email on imaging completion.
	
	.PARAMETER EmailSender
	Override the default email sender of ussitservices@jhu.edu.
	
	.PARAMETER EmailSmtp
	Override the default email server of smtp.johnshopkins.edu.
	
    .NOTES
	Author: mcarras8
	
	--Custom configurable TS Variables used--
	Email notifications: XOSDNotifyEmail or XOSDNotifyEmailChoice (in that order)
#>
param(
	[Parameter(Mandatory=$false)]
	[string]$EmailSender='USS IT Services <ussitservices@jhu.edu>',
	
	[Parameter(Mandatory=$false)]
	[string]$EmailSmtp='smtp.johnshopkins.edu'
)

try {
	$tsenv = New-Object -ComObject Microsoft.SMS.TSEnvironment
} catch {
	throw $_
}

$systemName = $tsenv.Value("XFinalComputerName")
$xosdruntimestart = $tsenv.Value("XOSDRuntimeStart")
$xosdcustomdriversuccess = $tsenv.Value("XOSDCustomDriverSuccess")
$xosdsoftwareupdatesuccess = $tsenv.Value("XOSDSoftwareUpdateSuccess")

# Get the current runtime in hours.
# We use a WebRequest to avoid time syncing issues.
try {
  $xosdruntimestart = $tsenv.Value("XOSDRuntimeStart")
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

$emailTo = $tsenv.Value("XOSDNotifyEmail")
if ([string]::IsNullOrWhitespace($emailTo)) {
    $xosdnotifyemailchoice = $tsenv.Value("XOSDNotifyEmailChoice")
    if ($xosdnotifyemailchoice -ne "NONE") {
        $emailTo = $xosdnotifyemailchoice
    }
}
if (-Not [string]::IsNullOrWhitespace($emailTo) -And $emailTo -like "*@*" -And -Not [string]::IsNullOrEmpty($EmailSender)) {
    $body = @"
Imaging Complete for $systemName. All required steps have completed. If custom drivers or software updates were selected, their status will show below.
Custom Driver Install Completed: $xosdcustomdriversuccess
Software Updates Completed: $xosdsoftwareupdatesuccess
Total Runtime Hours: $runtimeHours
"@
    $emailParams = @{
        From = $EmailSender
        To = $emailTo
        CC = $EmailSender
        Subject = "Imaging Complete for $systemName"
        Body = $body
        Priority = "Normal"
        DeliveryNotificationOption = @("OnSuccess", "OnFailure")
        SmtpServer = $EmailSmtp
    }
    
    Send-MailMessage @emailParams
}

[System.Environment]::Exit(0)
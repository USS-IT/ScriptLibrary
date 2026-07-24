# SCCM Removal User UI
# PowerShell 5.1 / Windows Forms
#
# Purpose:
# - Let user enter a numeric removal code
# - Write code to a file for SYSTEM removal task to read
# - Trigger SCCM removal scheduled task
# - Watch transcript log
# - Display only transcript body, excluding Start-Transcript and Stop-Transcript blocks
# - Refresh SCCM status labels after Stop-Transcript end marker is detected
# - Timeout after 30 minutes

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ----------------------------
# Configuration
# ----------------------------

$script:TranscriptLogPath = 'C:\USS\Logs\Remove-MCMClient.ps1.log'
$script:CodeFilePath      = 'C:\TEMP\SCCMRemovalCode.txt'
$script:ScheduledTaskName = 'JHU\SCCM Client Removal'
$script:PollingTimeoutMinutes = 30

$LogDir = "C:\USS\Logs\User\MCMTroubleshooting"

# -- END Configuration --

$script:WatchStarted          = $null
$script:LastDisplayedText     = ''
$script:StopTranscriptSeen    = $false
$script:TaskStarted           = $false

# Create the log path if it doesn't already exist
if (-not (Test-Path $LogDir -PathType Container)) {
	$null = New-Item -Path $LogDir -ItemType Directory
}
try {
	$_scriptName = Split-Path -Leaf $PSCommandPath
} catch {
	$_scriptName = "Start-MCMClientRemoval.ps1"
}
$LogPath = "$LogDir\$($_scriptName).log"
Start-Transcript $LogPath -Force

# Use WScript popup instead of MessageBox static calls
$script:Popup = New-Object -ComObject WScript.Shell

# Row number -> controls in that row
$script:ControlRows = @{}

# ----------------------------
# Helper functions
# ----------------------------

function Register-NewControl {
    param(
		[Parameter(Mandatory=$true)]
        [System.Windows.Forms.Control]$Control,
		
		[Parameter(Mandatory=$true)]
        [int]$Row
    )

    if (-not $script:ControlRows.ContainsKey($Row)) {
        $script:ControlRows[$Row] = New-Object System.Collections.ArrayList
    }

    [void]$script:ControlRows[$Row].Add($Control)
}

function Get-ControlRelativeLocation {
    param(
		[Parameter(Mandatory=$true)]
        [int]$Row,
				
        [int]$Spacing = 10,
		
        [int]$MarginLeft = 10,
		
        [int]$MarginTop = 10
    )

	$X = $MarginLeft
	$Y = $MarginTop
    # First control on first row
    if (-not $script:ControlRows.ContainsKey($Row) -Or $script:ControlRows[$Row].Count -eq 0) {		
		if ($script:ControlRows.ContainsKey($Row - 1)) {
			$PreviousRow = $script:ControlRows[$Row - 1]
			if ($PreviousRow.Count -gt 0) {
				$TallestBottom = (
					$PreviousRow |
					ForEach-Object { $_.Bottom } |
					Measure-Object -Maximum
				).Maximum

				$Y = $TallestBottom + $Spacing
			}
		}
    } else {
		$LastControl =
			$script:ControlRows[$Row][
				$script:ControlRows[$Row].Count - 1
			]
			
		$X = $LastControl.Right + $Spacing
		$Y = $LastControl.Top
	}

    return New-Object System.Drawing.Point($X, $Y)
}

function Show-ErrorDialog {
    param(
        [string]$Message,
        [string]$Title = 'SCCM Client Removal'
    )

    # 16 = Critical/Error icon
	Write-Host "[$(Get-Date -f 'MM-dd-yyyy HH:mm:ss')] Show-ErrorDialog: $Message"
    $null = $script:Popup.Popup($Message, 0, $Title, 16)
}

function Show-WarningDialog {
    param(
        [string]$Message,
        [string]$Title = 'SCCM Client Removal'
    )

    # 48 = Warning icon
	Write-Host "[$(Get-Date -f 'MM-dd-yyyy HH:mm:ss')] Show-WarningDialog: $Message"
    $null = $script:Popup.Popup($Message, 0, $Title, 48)
}

function Add-UiLine {
    param(
        $TextBox,
        [string]$Message
    )

	Write-Host $Message
    $TextBox.AppendText($Message + "`r`n")
    $TextBox.SelectionStart = $TextBox.Text.Length
    $TextBox.ScrollToCaret()
}

function Test-NumericCode {
    param(
        [string]$Code
    )

    if ($null -eq $Code) {
        return $false
    }

    $Code = $Code.Trim()

    if ($Code -match '^\d+$') {
        return $true
    }

    return $false
}

function Get-McmClientVersion {
    $RegPaths = @(
        'HKLM:\SOFTWARE\Microsoft\SMS\Mobile Client',
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\SMS\Mobile Client'
    )

    foreach ($Path in $RegPaths) {
        try {
            $Item = Get-ItemProperty -Path $Path -Name ProductVersion -ErrorAction Stop
            $Version = $Item.ProductVersion

            if ($null -ne $Version) {
                $VersionText = $Version.ToString().Trim()

                if ($VersionText.Length -gt 0) {
                    return $VersionText
                }
            }
        }
        catch {
            # Continue checking other path
        }
    }

    return 'Missing / blank'
}

function Get-CcmExecServiceStatus {
    try {
        $Service = Get-Service -Name CcmExec -ErrorAction Stop

        if ($null -ne $Service) {
            return $Service.Status.ToString()
        }
    }
    catch {
        return 'Missing'
    }

    return 'Missing'
}

function Test-SmsClientExists {
    try {
        $Client = Get-CimInstance -Namespace root\ccm -ClassName SMS_Client -ErrorAction Stop

        if ($null -ne $Client) {
            return $true
        }
    }
    catch {
        return $false
    }

    return $false
}

function Update-ClientStatusLabels {
    param(
        $VersionLabel,
        $ServiceLabel,
        $SmsClientLabel
    )

    $Version = Get-McmClientVersion
    $ServiceStatus = Get-CcmExecServiceStatus
    $SmsClientExists = Test-SmsClientExists

    $VersionLabel.Text = "MCM Client Version: $Version"
    $ServiceLabel.Text = "CcmExec Status: $ServiceStatus"

    if ($SmsClientExists) {
        $SmsClientLabel.Text = 'SMS_Client WMI: Exists'
    }
    else {
        $SmsClientLabel.Text = 'SMS_Client WMI: Missing'
    }
}

function Get-TranscriptInfo {
    param(
        [string]$Path
    )

    $Result = New-Object PSObject
    $Result | Add-Member -MemberType NoteProperty -Name BodyText -Value ''
    $Result | Add-Member -MemberType NoteProperty -Name StopMarkerSeen -Value $false

    if (-not (Test-Path $Path)) {
        return $Result
    }

    try {
        $Lines = Get-Content -Path $Path -ErrorAction Stop
    }
    catch {
        return $Result
    }

    if ($null -eq $Lines) {
        return $Result
    }

    if ($Lines.Count -eq 0) {
        return $Result
    }

    # Find the last transcript start marker.
    # This avoids showing stale content if the log was appended instead of overwritten.
    $StartMarkerIndex = -1

    for ($i = 0; $i -lt $Lines.Count; $i++) {
        if ($Lines[$i] -match 'PowerShell transcript start') {
            $StartMarkerIndex = $i
        }
    }

    if ($StartMarkerIndex -lt 0) {
        $Result.BodyText = ($Lines -join "`r`n")
        return $Result
    }

    # Find the end of the transcript header.
    # Header usually ends at the next line made of asterisks.
    $BodyStartIndex = $StartMarkerIndex + 1

    for ($i = $StartMarkerIndex + 1; $i -lt $Lines.Count; $i++) {
        if ($Lines[$i] -match '^\*{10,}\s*$') {
            $BodyStartIndex = $i + 1
            break
        }
    }

    # Optional cleanup:
    # Remove the informational line emitted by Start-Transcript if it appears
    # immediately after the header.
    while ($BodyStartIndex -lt $Lines.Count) {
        if ($Lines[$BodyStartIndex] -match '^\s*Transcript started, output file is') {
            $BodyStartIndex++
            continue
        }

        if ($Lines[$BodyStartIndex].Trim().Length -eq 0) {
            $BodyStartIndex++
            continue
        }

        break
    }

    # Find Stop-Transcript footer marker.
    $EndMarkerIndex = -1

    for ($i = $BodyStartIndex; $i -lt $Lines.Count; $i++) {
        if ($Lines[$i] -match 'PowerShell transcript end') {
            $EndMarkerIndex = $i
            break
        }
    }

    $BodyEndIndex = $Lines.Count - 1

    if ($EndMarkerIndex -ge 0) {
        $Result.StopMarkerSeen = $true

        # Footer usually begins at the asterisk line immediately before
        # "PowerShell transcript end".
        $FooterStartIndex = $EndMarkerIndex

        for ($i = $EndMarkerIndex - 1; $i -ge $BodyStartIndex; $i--) {
            if ($Lines[$i] -match '^\*{10,}\s*$') {
                $FooterStartIndex = $i
                break
            }
        }

        $BodyEndIndex = $FooterStartIndex - 1
    }

    if ($BodyEndIndex -lt $BodyStartIndex) {
        $Result.BodyText = ''
        return $Result
    }

    $BodyLines = @()

    for ($i = $BodyStartIndex; $i -le $BodyEndIndex; $i++) {
        $Line = $Lines[$i]

        # Do not show Stop-Transcript command output/footer if it slips in.
        if ($Line -match '^\s*Transcript stopped, output file is') {
            continue
        }

        $BodyLines += $Line
    }

    $Result.BodyText = ($BodyLines -join "`r`n")
    return $Result
}

function Update-TranscriptDisplay {
    param(
        $TextBox
    )

    $Info = Get-TranscriptInfo -Path $script:TranscriptLogPath

    if ($Info.BodyText -ne $script:LastDisplayedText) {
        $TextBox.Text = $Info.BodyText
        $TextBox.SelectionStart = $TextBox.Text.Length
        $TextBox.ScrollToCaret()
        $script:LastDisplayedText = $Info.BodyText
    }

    if ($Info.StopMarkerSeen) {
        $script:StopTranscriptSeen = $true
    }
}

function Start-RemovalTask {
    param(
        [string]$Code,
        $TextBox
    )

    try {
        $ParentFolder = Split-Path -Path $script:CodeFilePath -Parent

        if (-not (Test-Path $ParentFolder)) {
            New-Item -Path $ParentFolder -ItemType Directory -Force | Out-Null
        }

        # Write the code to a file instead of putting it on the command line.
        # The elevated removal script should read this file and delete it.
        Set-Content -Path $script:CodeFilePath -Value $Code -Force -Encoding ASCII

        Add-UiLine -TextBox $TextBox -Message "[$(Get-Date -f 'MM-dd-yyyy HH:mm:ss')] ** Wrote removal request."
        Add-UiLine -TextBox $TextBox -Message "[$(Get-Date -f 'MM-dd-yyyy HH:mm:ss')] ** Starting SCCM removal task..."
        Add-UiLine -TextBox $TextBox -Message ""

		<#
        $Arguments = '/Run /TN "' + $script:ScheduledTaskName + '"'

        $Process = Start-Process -FilePath 'schtasks.exe' -ArgumentList $Arguments -WindowStyle Hidden -Wait -PassThru -ErrorAction Stop

        if ($Process.ExitCode -ne 0) {
            Add-UiLine -TextBox $TextBox -Message "ERROR: schtasks.exe returned exit code $($Process.ExitCode)."
            return $false
        }
		#>

        $script:WatchStarted = Get-Date
        $script:StopTranscriptSeen = $false
        $script:TaskStarted = $true

        Add-UiLine -TextBox $TextBox -Message "[$(Get-Date -f 'MM-dd-yyyy HH:mm:ss')] ** SCCM client removal task was started."
        Write-Host "[$(Get-Date -f 'MM-dd-yyyy HH:mm:ss')] Watching log: $script:TranscriptLogPath"
        Add-UiLine -TextBox $TextBox -Message ""

        return $true
    }
    catch {
        Add-UiLine -TextBox $TextBox -Message "[$(Get-Date -f 'MM-dd-yyyy HH:mm:ss')] ** ERROR: Failed to start SCCM removal task."
        Add-UiLine -TextBox $TextBox -Message $_.Exception.Message
        return $false
    }
}

# ----------------------------
# Build UI
# ----------------------------

$form = New-Object System.Windows.Forms.Form
$form.Text = 'SCCM Client Removal Script'
$form.Size = New-Object System.Drawing.Size(700, 650)
$form.StartPosition = 'CenterScreen'
$form.FormBorderStyle = 'FixedDialog'
$form.MaximizeBox = $false
$form.MinimizeBox = $true
$form.KeyPreview = $true

$marginLeft = 10
$marginTop = 10
$spacing = 10
$labelHeight = 22
$textBoxHeight = 24
$buttonHeight = 30

# Row 0
$row = 0
# Client Version
$labelVersion = New-Object System.Windows.Forms.Label
$labelVersion.Location = Get-ControlRelativeLocation -Row $row -Spacing 5
$labelVersion.Size = New-Object System.Drawing.Size(350, $labelHeight)
$labelVersion.Text = 'MCM Client Version: Checking...'
$labelVersion.Font = New-Object System.Drawing.Font(
	$labelVersion.Font,
	[System.Drawing.FontStyle]::Bold
)
$form.Controls.Add($labelVersion)
Register-NewControl -Control $labelVersion -Row $row

# Row 2
$row++
# Service Status
$labelService = New-Object System.Windows.Forms.Label
$labelService.Location = Get-ControlRelativeLocation -Row $row -Spacing 5
$labelService.Size = New-Object System.Drawing.Size(200, $labelHeight)
$labelService.Text = 'CcmExec Status: Checking...'
$labelService.Font = New-Object System.Drawing.Font(
	$labelService.Font,
	[System.Drawing.FontStyle]::Bold
)
$form.Controls.Add($labelService)
Register-NewControl -Control $labelService -Row $row

# Row 3
$row++
# SMS Client WMI Status
$labelSmsClient = New-Object System.Windows.Forms.Label
$labelSmsClient.Location = Get-ControlRelativeLocation -Row $row -Spacing 5
$labelSmsClient.Size = New-Object System.Drawing.Size(200, $labelHeight)
$labelSmsClient.Text = 'SMS_Client WMI: Checking...'
$labelSmsClient.Font = New-Object System.Drawing.Font(
	$labelSmsClient.Font,
	[System.Drawing.FontStyle]::Bold
)
$form.Controls.Add($labelSmsClient)
Register-NewControl -Control $labelSmsClient -Row $row

# Row 4
$row++
# Start and Input controls
$buttonStart = New-Object System.Windows.Forms.Button
$buttonStart.Location = Get-ControlRelativeLocation -Row $row
$buttonStart.Size = New-Object System.Drawing.Size(120, $buttonHeight)
$buttonStart.Text = 'Start Removal'
$form.Controls.Add($buttonStart)
Register-NewControl -Control $buttonStart -Row $row

$labelCode = New-Object System.Windows.Forms.Label
$labelCode.Location = Get-ControlRelativeLocation -Row $row
$labelCode.Size = New-Object System.Drawing.Size(120, $labelHeight)
$labelCode.Text = 'Removal code:'
$labelCode.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
$form.Controls.Add($labelCode)
Register-NewControl -Control $labelCode -Row $row

$textBoxCode = New-Object System.Windows.Forms.TextBox
$textBoxCode.Location = Get-ControlRelativeLocation -Row $row
$textBoxCode.Size = New-Object System.Drawing.Size(160, $textBoxHeight)
$textBoxCode.MaxLength = 12
$form.Controls.Add($textBoxCode)
Register-NewControl -Control $textBoxCode -Row $row

# Row 5
$row++
# Script status label
$labelStatus = New-Object System.Windows.Forms.Label
$labelStatus.Location = Get-ControlRelativeLocation -Row $row
$labelStatus.Size = New-Object System.Drawing.Size(850, $labelHeight)
$labelStatus.Text = 'Enter the removal code, then click Start Removal.'
$labelStatus.Font = New-Object System.Drawing.Font(
	$labelStatus.Font,
	[System.Drawing.FontStyle]::Bold
)
$form.Controls.Add($labelStatus)
Register-NewControl -Control $labelStatus -Row $row

# Row 4
$row++
# Log textbox
$textBoxLog = New-Object System.Windows.Forms.TextBox
$textBoxLog.Location = Get-ControlRelativeLocation -Row $row
$textBoxLog.Size = New-Object System.Drawing.Size(
	($form.ClientSize.Width - $textBoxLog.Location.X - 10), 
	($form.ClientSize.Height - $textBoxLog.Location.Y - 10)
)
$textBoxLog.Multiline = $true
$textBoxLog.ScrollBars = 'Both'
$textBoxLog.ReadOnly = $true
$textBoxLog.Font = New-Object System.Drawing.Font('Consolas', 9)
$textBoxLog.WordWrap = $false
$form.Controls.Add($textBoxLog)
Register-NewControl -Control $textBoxLog -Row $row

# Timer
$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = 1000

# Initial status label update
Update-ClientStatusLabels -VersionLabel $labelVersion -ServiceLabel $labelService -SmsClientLabel $labelSmsClient

# ----------------------------
# Events
# ----------------------------

$timer.Add_Tick({
    if (-not $script:TaskStarted) {
        return
    }

    Update-TranscriptDisplay -TextBox $textBoxLog

    if ($script:StopTranscriptSeen) {
        $timer.Stop()

        $labelStatus.Text = 'Removal script completed. Refreshing client status.'

        Update-ClientStatusLabels -VersionLabel $labelVersion -ServiceLabel $labelService -SmsClientLabel $labelSmsClient

        $buttonStart.Enabled = $true
        $textBoxCode.Enabled = $true

        Add-UiLine -TextBox $textBoxLog -Message ''
        Add-UiLine -TextBox $textBoxLog -Message "[$(Get-Date -f 'MM-dd-yyyy HH:mm:ss')] ** Removal script transcript ended. Refreshing Client labels."

        $script:TaskStarted = $false
        return
    }

    if ($null -ne $script:WatchStarted) {
        $Elapsed = New-TimeSpan -Start $script:WatchStarted -End (Get-Date)

        if ($Elapsed.TotalMinutes -ge $script:PollingTimeoutMinutes) {
            $timer.Stop()

            $labelStatus.Text = 'Timed out waiting for removal transcript to finish.'

            $buttonStart.Enabled = $true
            $textBoxCode.Enabled = $true

            Add-UiLine -TextBox $textBoxLog -Message ''
            Add-UiLine -TextBox $textBoxLog -Message "[$(Get-Date -f 'MM-dd-yyyy HH:mm:ss')] ** ERROR: Timed out waiting for log end marker."

            Update-ClientStatusLabels -VersionLabel $labelVersion -ServiceLabel $labelService -SmsClientLabel $labelSmsClient

            Show-ErrorDialog -Message 'Timed out after 30 minutes waiting for the SCCM removal log to complete. The removal process may still be running, stalled, or may have failed before closing the transcript.'

            $script:TaskStarted = $false
            return
        }

        $Minutes = [int]$Elapsed.TotalMinutes
        $Seconds = $Elapsed.Seconds

        $labelStatus.Text = "Removal task started. Watching transcript log. Elapsed: $Minutes minute(s), $Seconds second(s)."
    }
})

$buttonStart.Add_Click({
    $Code = $textBoxCode.Text.Trim()

    if (-not (Test-NumericCode -Code $Code)) {
        Add-UiLine -TextBox $textBoxLog -Message "[$(Get-Date -f 'MM-dd-yyyy HH:mm:ss')] ** ERROR: Please enter a numeric removal code."
        $labelStatus.Text = 'Invalid code. Numeric values only.'
        Show-WarningDialog -Message 'Please enter a numeric removal code.'
        $textBoxCode.Focus()
        return
    }

    $buttonStart.Enabled = $false
    $textBoxCode.Enabled = $false

    $script:LastDisplayedText = ''
    $script:StopTranscriptSeen = $false
    $script:TaskStarted = $false
    $script:WatchStarted = $null

    $StartedOk = Start-RemovalTask -Code $Code -TextBox $textBoxLog

    if ($StartedOk) {
        $timer.Start()
        $labelStatus.Text = 'Removal started. Watching transcript log.'
    }
    else {
        $buttonStart.Enabled = $true
        $textBoxCode.Enabled = $true
        $labelStatus.Text = 'Failed to start removal task.'
        Show-ErrorDialog -Message 'Failed to start the SCCM removal task. Check the task name, permissions, and local script files.'
    }
})

$textBoxCode.Add_KeyDown({
    if ($_.KeyCode -eq 'Enter') {
        $buttonStart.PerformClick()
    }
})

$form.Add_FormClosing({
    $timer.Stop()
    $timer.Dispose()
	try {
		Stop-Transcript | Out-Null
	} catch {}
})

$form.ShowDialog() | Out-Null

try {
	Stop-Transcript | Out-Null
} catch {}
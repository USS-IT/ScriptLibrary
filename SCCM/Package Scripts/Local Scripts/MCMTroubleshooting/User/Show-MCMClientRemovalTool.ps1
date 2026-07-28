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
#
# Author: Matt Carras (mcarras8)
# Created: 7-24-26

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ----------------------------
# Configuration
# ----------------------------

# File path to temporarily store a code that the removal script will read.
$CodeFilePath      = 'C:\TEMP\SCCMRemovalCode.txt'
# Should match max timeout for removal script.
$PollingTimeoutMinutes = 30

$TranscriptLogPath = 'C:\USS\Logs\Remove-MCMClient.ps1.log'
$ScheduledTaskName = 'JHU\SCCM Client Removal'

$LogDir = "C:\USS\Logs\User\MCMTroubleshooting"

# UI Config
$MainFormTitle = 'SCCM Client Removal Script'
$MainFormWidth = 700
$MainFormHeight = 650

# -- END Configuration --

$WatchStarted          = $null
$LastDisplayedText     = ''
$StopTranscriptSeen    = $false
$TaskStarted           = $false

# Create the log path if it doesn't already exist
if (-not (Test-Path $LogDir -PathType Container)) {
	$null = New-Item -Path $LogDir -ItemType Directory
}
try {
	$_scriptName = Split-Path -Leaf $PSCommandPath
} catch {
	$_scriptName = "Show-MCMClientRemovalTool.ps1"
}
$LogPath = "$LogDir\$($_scriptName).log"
Start-Transcript $LogPath -Force

# Cleanup any leftover code files from last run.
if ((Test-Path $CodeFilePath)) {
	Remove-Item $CodeFilePath -Force
}

# ----------------------------
# Helper functions
# ----------------------------

function Initialize-FormControlRows {
	# Initialize the dictionary of form controls.
	param(
		[Parameter(Mandatory=$true, Position=0)]
		[System.Windows.Forms.Form] $Form
	)
	
	# Initialize dictionary, stored by reference object.
	if ($script:_FormControlRows -eq $null) {
		$script:_FormControlRows = New-Object `
			'System.Collections.Generic.Dictionary[
				System.Windows.Forms.Form,
				System.Collections.ArrayList
			]'
	}
	if ($script:_FormControlRows[$Form] -eq $null) {
		$script:_FormControlRows[$Form] = New-Object System.Collections.ArrayList
		[void]$script:_FormControlRows[$Form].Add((New-Object System.Collections.ArrayList))
	}
}

function Add-FormControl {
	# Add a new control to the given form on the current or given row.
    param(
		[Parameter(Mandatory=$true, Position=0)]
		[System.Windows.Forms.Form] $Form,
		
		[Parameter(Mandatory=$true, Position=1)]
        [System.Windows.Forms.Control] $Control,
		
		[Parameter(Mandatory=$false, Position=2)]
        [Nullable[int]] $Row
    )

	Initialize-FormControlRows $Form | Out-Null
	
	# If Row isn't given, assume the current row.
	if ($Row -eq $null) {
		$Row = $script:_FormControlRows[$Form].Count - 1
	}
	
	if ($Row -ge $script:_FormControlRows[$Form].Count) {
		throw "Invalid row number: $Row. Current row length is $($script:_FormControlRows[$Form].Count), starting at index 0."
	}
	
	Write-Verbose "[Add-FormControl] Adding to row: $Row. New length for row: $($script:_FormControlRows[$Form][$Row].Count)"
    
	# Add the control to our group of form controls.
	($script:_FormControlRows[$Form][$Row]).Add($Control) | Out-Null
	# Add the control to this form.
	[void]$Form.Controls.Add($Control)
}

function New-FormControlRow {
	# Initializes a new row of form controls.
	param(
		[Parameter(Mandatory=$true, Position=0)]
		[System.Windows.Forms.Form] $Form
    )
	
	try {
		if ($script:_FormControlRows -eq $null -Or $script:_FormControlRows[$Form] -eq $null) {
			Initialize-FormControlRows $Form | Out-Null
		} else {			
			[void]$script:_FormControlRows[$Form].Add((New-Object System.Collections.ArrayList))
		}
	} catch {
		throw $_
	}
}
	
function Get-ControlRelativeLocation {
    param(
		[Parameter(Mandatory=$true, Position=0)]
		[System.Windows.Forms.Form] $Form,
	
		[Parameter(Mandatory=$false, Position=1)]
        [Nullable[int]]$Row,
				
        [int]$Spacing = 10,
		
        [int]$MarginLeft = 10,
		
        [int]$MarginTop = 10
    )

	$X = $MarginLeft
	$Y = $MarginTop
	
	try {
		$formrows = $script:_FormControlRows[$Form]
	} catch {
		throw $_
	}
	
	# If very first control on first row, just return margins.
	if ($formrows -ne $null) {
		# If Row isn't given, assume the current row.
		if ($Row -eq $null) {
			$Row = $formrows.Count - 1
		} elseif ($Row -ge $formrows.Count) {
			throw "Invalid row number: $Row. Current row length is $($script:_FormControlRows[$Form].Count), starting at index 0."
		}
	
		# First control on given row
		if ($formrows[$Row].Count -eq 0) {
			if ($Row -gt 0) {
				$PreviousRow = $formrows[$Row - 1]
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
				$formrows[$Row][
					$formrows[$Row].Count - 1
				]
				
			$X = $LastControl.Right + $Spacing
			$Y = $LastControl.Top
		}
	}

    return New-Object System.Drawing.Point($X, $Y)
}

function Show-ErrorDialog {
    param(
		[Parameter(Mandatory=$true)]
        [string]$Message,
		
        [string]$Title = 'SCCM Client Removal'
    )
	
	Write-Host "[$(Get-Date -f 'MM-dd-yyyy HH:mm:ss')] Show-ErrorDialog: $Message"
	[System.Windows.Forms.MessageBox]::Show(
		$Message,
		$Title,
		[System.Windows.Forms.MessageBoxButtons]::OK,
		[System.Windows.Forms.MessageBoxIcon]::Error
	)
}

function Show-WarningDialog {
    param(
		[Parameter(Mandatory=$true)]
        [string]$Message,
		
        [string]$Title = 'SCCM Client Removal'
    )

    # 48 = Warning icon
	Write-Host "[$(Get-Date -f 'MM-dd-yyyy HH:mm:ss')] Show-WarningDialog: $Message"
	[System.Windows.Forms.MessageBox]::Show(
		$Message,
		$Title,
		[System.Windows.Forms.MessageBoxButtons]::OK,
		[System.Windows.Forms.MessageBoxIcon]::Warning
	)
}

function Add-TextBoxLine {
    param(
		[Parameter(Mandatory=$true)]
        $TextBox,
		
		[Parameter(Mandatory=$true)]
        [string]$Message
    )

	Write-Host $Message
    $TextBox.AppendText($Message + "`r`n")
    $TextBox.SelectionStart = $TextBox.Text.Length
    $TextBox.ScrollToCaret()
}

function Test-IsNumericString {
    param(
		[Parameter(Mandatory=$true, Position=0)]
        $Text
	)
	
	if ($Text -eq $null -Or $Text -isnot [string]) {
		return $false
	}
	
	return (($Text.Trim()) -match '^\d+$')
}

function Get-McmClientVersion {
	param()
	
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
	param()
	
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
	param()
	
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

<#
function Test-IsAdministrator {
	# Return whether the current user is running under local admin context.
	param()
	
	try {
		$Identity = [System.Security.Principal.WindowsIdentity]::GetCurrent()

		$Principal = New-Object System.Security.Principal.WindowsPrincipal($Identity) -ErrorAction Stop

		return $Principal.IsInRole(
			[System.Security.Principal.WindowsBuiltInRole]::Administrator
		)
	} catch {
		Write-Error $_
		return $false
	}
}
#>

function Update-ClientStatusLabels {
    param(
		[Parameter(Mandatory=$true)]
        [System.Windows.Forms.Label] $VersionLabel,
		
		[Parameter(Mandatory=$true)]
        [System.Windows.Forms.Label] $ServiceLabel,
		
		[Parameter(Mandatory=$true)]
        [System.Windows.Forms.Label] $SmsClientLabel
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

function Get-InputDialog {
	# Shows an input dialog and returns the result.
	
	param(
		[Parameter(Mandatory=$true)]
		[string]$Title,
		
		[Parameter(Mandatory=$true)]
		[string]$Prompt,
		
		[int]$MaxLength = 255,
		
		[string]$AcceptText = 'OK',
		
		[string]$CancelText = 'Cancel'
	)
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    $form = New-Object System.Windows.Forms.Form
    $form.Text = $Title
    $form.Size = New-Object System.Drawing.Size(350,150)
    $form.StartPosition = 'CenterParent'
    $form.FormBorderStyle = 'FixedDialog'
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false
    $form.ShowInTaskbar = $false

    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(10,15)
    $label.Size = New-Object System.Drawing.Size(120,20)
    $label.Text = $Prompt
    $label.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight

    $textboxInput = New-Object System.Windows.Forms.TextBox
    $textboxInput.Location = New-Object System.Drawing.Point(140,15)
    $textboxInput.Size = New-Object System.Drawing.Size(150,20)
    $textboxInput.MaxLength = $MaxLength

    $buttonOK = New-Object System.Windows.Forms.Button
    $buttonOK.Location = New-Object System.Drawing.Point(70,60)
    $buttonOK.Size = New-Object System.Drawing.Size(80,25)
    $buttonOK.Text = $AcceptText
    $buttonOK.DialogResult = [System.Windows.Forms.DialogResult]::OK

    $buttonCancel = New-Object System.Windows.Forms.Button
    $buttonCancel.Location = New-Object System.Drawing.Point(170,60)
    $buttonCancel.Size = New-Object System.Drawing.Size(80,25)
    $buttonCancel.Text = $CancelText
    $buttonCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel

    $form.Controls.Add($label)
    $form.Controls.Add($textboxInput)
    $form.Controls.Add($buttonOK)
    $form.Controls.Add($buttonCancel)

    $form.AcceptButton = $buttonOK
    $form.CancelButton = $buttonCancel

    [void]$textboxInput.Focus()

	[void]$textboxInput.Add_KeyDown({
		if ($_.KeyCode -eq 'Enter') {
			$buttonOK.PerformClick()
		}
	})

    $Result = $form.ShowDialog()

    if ($Result -ne [System.Windows.Forms.DialogResult]::OK) {
        return $null
    }

    return $textboxInput.Text.Trim()
}

function Get-TranscriptInfo {
	# Returns all of the transcript 
    param(
		[Parameter(Mandatory=$true)]
        [string]$Path
    )

    $Result = [PSCustomObject]@{
		BodyText = ''
		StopMarkerSeen = $false
	}
    
    if (-not (Test-Path $Path)) {
        return $Result
    }

    try {
        $Lines = Get-Content -Path $Path -ErrorAction Stop
    }
    catch {
        return $Result
    }

    if ($Lines -eq $null -Or $Lines.Count -eq 0) {
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

	# If the start marker was not found, return the entire log.
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
		[Parameter(Mandatory=$true)]
        [System.Windows.Forms.TextBox] $TextBox
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
		[Parameter(Mandatory=$true)]
        [string] $Code,
		
		[Parameter(Mandatory=$true)]
        [System.Windows.Forms.TextBox] $TextBox
    )

    try {
        # Write the code to a file.
        # The elevated removal script should read this file and delete it.
		$ParentFolder = Split-Path -Path $script:CodeFilePath -Parent
        if (-not (Test-Path $ParentFolder)) {
            New-Item -Path $ParentFolder -ItemType Directory -Force | Out-Null
        }
		
        Set-Content -Path $script:CodeFilePath -Value $Code -Force -Encoding ASCII -ErrorAction Stop

        Add-TextBoxLine -TextBox $TextBox -Message "[$(Get-Date -f 'MM-dd-yyyy HH:mm:ss')] ** Starting SCCM removal task..."
        Add-TextBoxLine -TextBox $TextBox -Message " "

		Write-EventLog -LogName 'USS-EventLog' -Source 'Remove-MCMClient' -EventID 1000 -EntryType Information -Message 'Trigger Remove-MCMClient with code' -ErrorAction Stop

        $script:WatchStarted = Get-Date
        $script:StopTranscriptSeen = $false
        $script:TaskStarted = $true

        Add-TextBoxLine -TextBox $TextBox -Message "[$(Get-Date -f 'MM-dd-yyyy HH:mm:ss')] ** SCCM client removal task was started."
        Write-Host "[$(Get-Date -f 'MM-dd-yyyy HH:mm:ss')] Watching log: $script:TranscriptLogPath"
        Add-TextBoxLine -TextBox $TextBox -Message " "

        return $true
    }
    catch {
        Add-TextBoxLine -TextBox $TextBox -Message "[$(Get-Date -f 'MM-dd-yyyy HH:mm:ss')] ** ERROR: Failed to start SCCM client removal task."
		Write-Error $_
        return $false
    }
}

function Start-CleanupBeforeExit {
	param()
	
	if ((Test-Path $script:CodeFilePath)) {
		Remove-Item $script:CodeFilePath -Force
	}
	
	try {
		Stop-Transcript | Out-Null
	} catch {}
}
	
# ----------------------------
# Main Form UI
# ----------------------------

$formMain = New-Object System.Windows.Forms.Form
$formMain.Text = $MainFormTitle
$formMain.Size = New-Object System.Drawing.Size($MainFormWidth, $MainFormHeight)
$formMain.StartPosition = 'CenterScreen'
$formMain.FormBorderStyle = 'FixedDialog'
$formMain.MaximizeBox = $false
$formMain.MinimizeBox = $true
$formMain.KeyPreview = $true

$marginLeft = 10
$marginTop = 10
$spacing = 10
$labelHeight = 22
$textBoxHeight = 24
$buttonHeight = 30

# Row 0
# Client Version
$labelVersion = New-Object System.Windows.Forms.Label
$labelVersion.Location = Get-ControlRelativeLocation -Form $formMain -Spacing 5
$labelVersion.Size = New-Object System.Drawing.Size(350, $labelHeight)
$labelVersion.Text = 'MCM Client Version: Checking...'
$labelVersion.Font = New-Object System.Drawing.Font(
	$labelVersion.Font,
	[System.Drawing.FontStyle]::Bold
)
Add-FormControl $formMain $labelVersion

# Row 2
New-FormControlRow $formMain
# Service Status
$labelService = New-Object System.Windows.Forms.Label
$labelService.Location = Get-ControlRelativeLocation -Form $formMain -Spacing 5
$labelService.Size = New-Object System.Drawing.Size(200, $labelHeight)
$labelService.Text = 'CcmExec Status: Checking...'
$labelService.Font = New-Object System.Drawing.Font(
	$labelService.Font,
	[System.Drawing.FontStyle]::Bold
)
Add-FormControl $formMain $labelService

# Row 3
New-FormControlRow $formMain
# SMS Client WMI Status
$labelSmsClient = New-Object System.Windows.Forms.Label
$labelSmsClient.Location = Get-ControlRelativeLocation -Form $formMain -Spacing 5
$labelSmsClient.Size = New-Object System.Drawing.Size(200, $labelHeight)
$labelSmsClient.Text = 'SMS_Client WMI: Checking...'
$labelSmsClient.Font = New-Object System.Drawing.Font(
	$labelSmsClient.Font,
	[System.Drawing.FontStyle]::Bold
)
Add-FormControl $formMain $labelSmsClient

# Row 4
New-FormControlRow $formMain
# Start and Input controls
$buttonStart = New-Object System.Windows.Forms.Button
$buttonStart.Location = Get-ControlRelativeLocation -Form $formMain
$buttonStart.Size = New-Object System.Drawing.Size(120, $buttonHeight)
$buttonStart.Text = 'Start Removal'
Add-FormControl $formMain $buttonStart

# Row 5
New-FormControlRow $formMain
# Script status label
$labelStatus = New-Object System.Windows.Forms.Label
$labelStatus.Location = Get-ControlRelativeLocation -Form $formMain
$labelStatus.Size = New-Object System.Drawing.Size(850, $labelHeight)
$labelStatus.Text = 'Enter the removal code, then click Start Removal.'
$labelStatus.Font = New-Object System.Drawing.Font(
	$labelStatus.Font,
	[System.Drawing.FontStyle]::Bold
)
Add-FormControl $formMain $labelStatus

# Row 4
New-FormControlRow $formMain
# Log textbox
$textBoxLog = New-Object System.Windows.Forms.TextBox
$textBoxLog.Location = Get-ControlRelativeLocation -Form $formMain
$textBoxLog.Size = New-Object System.Drawing.Size(
	($formMain.ClientSize.Width - $textBoxLog.Location.X - 10), 
	($formMain.ClientSize.Height - $textBoxLog.Location.Y - 10)
)
$textBoxLog.Multiline = $true
$textBoxLog.ScrollBars = 'Both'
$textBoxLog.ReadOnly = $true
$textBoxLog.Font = New-Object System.Drawing.Font('Consolas', 9)
$textBoxLog.WordWrap = $false
Add-FormControl $formMain $textBoxLog

# Timer for polling
$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = 1000

# Initial status label update
Update-ClientStatusLabels -VersionLabel $labelVersion -ServiceLabel $labelService -SmsClientLabel $labelSmsClient

# ----------------------------
# Form Events
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

        Add-TextBoxLine -TextBox $textBoxLog -Message ''
        Add-TextBoxLine -TextBox $textBoxLog -Message "[$(Get-Date -f 'MM-dd-yyyy HH:mm:ss')] ** Removal script transcript ended. Refreshing Client labels."

        $script:TaskStarted = $false
        return
    }

    if ($null -ne $script:WatchStarted) {
        $Elapsed = New-TimeSpan -Start $script:WatchStarted -End (Get-Date)

        if ($Elapsed.TotalMinutes -ge $script:PollingTimeoutMinutes) {
            $timer.Stop()

            $labelStatus.Text = 'Timed out waiting for removal transcript to finish.'

            $buttonStart.Enabled = $true

            Add-TextBoxLine -TextBox $textBoxLog -Message ''
            Add-TextBoxLine -TextBox $textBoxLog -Message "[$(Get-Date -f 'MM-dd-yyyy HH:mm:ss')] ** ERROR: Timed out waiting for log end marker."

            Update-ClientStatusLabels -VersionLabel $labelVersion -ServiceLabel $labelService -SmsClientLabel $labelSmsClient

            Show-ErrorDialog -Title $MainFormTitle -Message 'Timed out after waiting 30 minutes for the SCCM removal log to complete. The removal process may still be running, stalled, or may have failed before closing the transcript.'

            $script:TaskStarted = $false
            return
        }

        $Minutes = [int]$Elapsed.TotalMinutes
        $Seconds = $Elapsed.Seconds

        $labelStatus.Text = "Removal task started. Watching transcript log. Elapsed: $Minutes minute(s), $Seconds second(s)."
    }
})

$buttonStart.Add_Click({
	$Code = Get-InputDialog -Title 'Removal Code Required' -Prompt 'Enter numeric code:'
	
	Write-Host "Code: $Code"
	
    if (-not (Test-IsNumericString $Code)) {
        Add-TextBoxLine -TextBox $textBoxLog -Message "[$(Get-Date -f 'MM-dd-yyyy HH:mm:ss')] ** ERROR: Please enter a numeric removal code."
        $labelStatus.Text = 'Invalid code. Numeric values only.'
        Show-WarningDialog -Title $MainFormTitle -Message 'Please enter a numeric removal code.' 
        return
    }

    $buttonStart.Enabled = $false

    $script:LastDisplayedText = ''
    $script:StopTranscriptSeen = $false
    $script:TaskStarted = $false
    $script:WatchStarted = $null

    $StartedOk = Start-RemovalTask -Code $Code -TextBox $textBoxLog

    if ($StartedOk) {
        $timer.Start()
        $labelStatus.Text = 'Removal started. Watching transcript log.'
    } else {
        $buttonStart.Enabled = $true
        $labelStatus.Text = 'Failed to start removal task.'
        Show-ErrorDialog -Title $MainFormTitle -Message 'Failed to start the SCCM removal task.'
    }
})

$formMain.Add_FormClosing({
    $timer.Stop()
    $timer.Dispose()
})

$formMain.ShowDialog() | Out-Null

Start-CleanupBeforeExit | Out-Null
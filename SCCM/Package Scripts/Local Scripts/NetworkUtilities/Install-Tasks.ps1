<#
.SYNOPSIS
    Creates custom event triggers and scheduled tasks for invoking network troubleshooting scripts
.DESCRIPTION
    This script:
    1. Creates a custom shared event log and unique source
    2. Registers a scheduled task triggered by the custom event
    3. The tasks run the applicable scripts
.NOTES
	Author: Matt Carras (mcarras8)
	Created: 09-02-2026
	
	Example trigger (run as any user):
	
	Write-EventLog -LogName 'USS-EventLog' -Source 'SwitchDNS-Add' -EventID 1000 -EntryType Information -Message 'Trigger Switch-DNS -Mode Add'
#>
# Requires Administrator privileges
#Requires -RunAsAdministrator

# Configuration Variables
# Shared by all USS scripts using our custom Event Log
$SharedEventLog = "USS-EventLog"
$LogDir = "C:\USS\Logs\Packages\NetworkUtilities"
$ScriptPath = "C:\USS\Scripts\Admin\NetworkUtilities\Switch-DNS.ps1"

try {
	$_scriptName = Split-Path -Leaf $PSCommandPath
} catch {
	$_scriptName = "Install-Tasks.ps1"
}
$LogPath = "$LogDir\$($_scriptName).log"

<#
 In case we need to setup multiple related tasks in the same script.
 Source is specific to this event log
 Event ID be 1000 unless we have more than one event per source

Parameters for powershell scripts:
Execute = "powershell.exe" `
Argument = "-NoProfile -ExecutionPolicy Bypass -File `"$ScriptPaths`""
#>
$EventsAndTasks = @(
	@{
		EventSource = "SwitchDNS-Toggle"		
		EventID = 1000									
		TaskName = "USS-SwitchDNS-Toggle"
		TaskDescription = "Toggles between adding or removing temporary public DNS servers to the active network adapter. Source: USS NetworkUtilities package. Version: 1.0"
		Execute = "powershell.exe"
		Argument = "-NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$ScriptPath`" -LogDir `"$LogDir`" -Notify -ResetOnStartup -Mode `"Toggle`""
		#WorkingDirectory = ""
	},
	
	@{
		EventSource = "SwitchDNS-Add"		
		EventID = 1000									
		TaskName = "USS-SwitchDNS-Add"
		TaskDescription = "Adds temporary public DNS servers to the active network adapter. Source: USS NetworkUtilities package. Version: 1.0"
		Execute = "powershell.exe"
		Argument = "-NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$ScriptPath`" -LogDir `"$LogDir`" -Notify -ResetOnStartup -Mode `"Add`""
		#WorkingDirectory = ""
	},
	
	@{
		EventSource = "SwitchDNS-Reset"		
		EventID = 1000									
		TaskName = "USS-SwitchDNS-Reset"
		TaskDescription = "Removes any temporary public DNS servers previously added to the active network adapter. Source: USS NetworkUtilities package. Version: 1.0"
		Execute = "powershell.exe"
		Argument = "-NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$ScriptPath`" -LogDir `"$LogDir`" -Notify -Mode `"Reset`""
		#WorkingDirectory = ""
	},
	
	# Custom events for user notifications.
	# We'll need to create the Scheduled Tasks under the user's account later.
	@{
		EventSource = "Notify-SwitchDNS-Changed"		
		EventID = 1000									
	},
	@{
		EventSource = "Notify-SwitchDNS-Failure"		
		EventID = 1000
	},
	@{
		EventSource = "Notify-SwitchDNS-Reverted"		
		EventID = 1000
	},
	@{
		EventSource = "Notify-SwitchDNS-NoAction"		
		EventID = 1000
	}
)

# Create the log path if it doesn't already exist
if (-not (Test-Path $LogDir -PathType Container)) {
	$null = New-Item -Path $LogDir -ItemType Directory
}
Start-Transcript $LogPath -Force

# Step 1: Create the shared event log if it doesn't already exist
try {
    if (-not ([System.Diagnostics.EventLog]::Exists($SharedEventLog))) {
        # Create with a temporary source, then remove it
        New-EventLog -LogName $SharedEventLog -Source "USS-Automation-Init"
        Remove-EventLog -Source "USS-Automation-Init" -ErrorAction SilentlyContinue
        Write-Host "Created shared custom event log: $SharedEventLog"
    } else {
        Write-Host "Custom event log already exists: $SharedEventLog, not creating"
    }
} catch {
    Write-Host "Error with creating custom event log [$SharedEventLog]: $_"
	Stop-Transcript | Out-Null
	exit 1
}

# Iterate over each task
foreach ($task in $EventsAndTasks) {
	Write-Host "Processing [$($task.EventSource)]"

	try {
		Write-Host "Creating custom event source [$($task.EventSource)] if needed..."
		if (-not [System.Diagnostics.EventLog]::SourceExists($task.EventSource)) {

			New-EventLog -LogName $SharedEventLog -Source $task.EventSource
			Write-Host "Event source [$($task.EventSource)] created successfully"
		} else {
			Write-Host "Event source [$($task.EventSource)] already exists"
		}
	} catch {
		Write-Error "Error creating event source [$($task.EventSource)]: $_"
		continue
	}

	if ([string]::IsNullOrEmpty($task.TaskName)) {
		Write-Host "Task Name is empty, skipping task creation"
	} else {
		# Step 3: Create the Scheduled Task
		Write-Host "Creating scheduled task [$($task.TaskName)]..."

		try {
			# Check if scheduled task already exists.	
			if (Get-ScheduledTask -TaskName $task.TaskName -ErrorAction SilentlyContinue) {
				Write-Host "Deleting preexisting scheduled task [$($task.TaskName)]..."
				Unregister-ScheduledTask -TaskName $task.TaskName -Confirm:$false
			}
		
			$params = @{
				Execute = $task.Execute
			}
			if (-Not [string]::IsNullOrEmpty($task.Argument)) {
				$params.Add('Argument', $task.Argument)
			}
			if (-Not [string]::IsNullOrEmpty($task.WorkingDirectory)) {
				$params.Add('WorkingDirectory', $task.WorkingDirectory)
			}
			# Define the action - Run PowerShell script
			$Action = New-ScheduledTaskAction @params

			# Define the event trigger
			# This creates an XPath query to filter for the specific event
			$CIMTriggerClass = Get-CimClass -ClassName MSFT_TaskEventTrigger -Namespace Root/Microsoft/Windows/TaskScheduler:MSFT_TaskEventTrigger
			$Trigger = New-CimInstance -CimClass $CIMTriggerClass -ClientOnly
			$Trigger.Subscription = @"
<QueryList>
  <Query Id="0" Path="$SharedEventLog">
	<Select Path="$SharedEventLog">
	  *[System[Provider[@Name='$($task.EventSource)'] and EventID=$($task.EventID)]]
	</Select>
  </Query>
</QueryList>
"@
			$Trigger.Enabled = $True

			# Define principal (run as SYSTEM with highest privileges)
			$Principal = New-ScheduledTaskPrincipal `
				-UserId "SYSTEM" `
				-LogonType ServiceAccount `
				-RunLevel Highest

			# Define settings
			$Settings = New-ScheduledTaskSettingsSet `
				-AllowStartIfOnBatteries `
				-DontStopIfGoingOnBatteries `
				-StartWhenAvailable `
				-ExecutionTimeLimit (New-TimeSpan -Hours 2)

			# Register the scheduled task
			$Task = Register-ScheduledTask `
				-TaskName $task.TaskName `
				-Action $Action `
				-Trigger $Trigger `
				-Principal $Principal `
				-Settings $Settings `
				-Description $task.TaskDescription `
				-Force

			Write-Host "Scheduled task [$($task.TaskName)] created successfully"
		} catch {
			Write-Error "Error creating scheduled task [$($task.TaskName)]: $_"
		}
	}
}
	
Stop-Transcript | Out-Null
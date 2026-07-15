<#
.SYNOPSIS
    Creates a custom scheduled task for Reenable_WiFi notifications
.DESCRIPTION
    This script:
    1. Registers a scheduled task triggered by the custom event
    2. The tasks run the applicable scripts
.NOTES
	Author: Matt Carras (mcarras8)
	Created: 06-30-2026
	
	Requires creating the custom event source first.
#>

# Configuration Variables
# Shared by all USS scripts using our custom Event Log
$SharedEventLog = "USS-EventLog"
$LogDir = "C:\USS\Logs\User\Reenable_WiFi"
$ScriptPath = "C:\USS\Scripts\User\Reenable_WiFi\Show-Toast.ps1"
# Path to the script used by our shortcut. Used for LauncherID.
$ShortcutScriptPath = "C:\USS\Scripts\User\Reenable_WiFi\Enable_WiFi.cmd"

try {
	$_scriptName = Split-Path -Leaf $PSCommandPath
} catch {
	$_scriptName = "Register-Tasks.ps1"
}
$LogPath = "$LogDir\$($_scriptName).log"

<#
 In case we need to setup multiple related tasks in the same script.
 EventSource is specific to this event log
 EventID should be 1000 unless we have more than one event per source
 
 To update tasks, change the TaskDescription. For example, add "Version: 1.1" to your task description, incrementing it as needed.

 Parameters for powershell scripts:
 Execute = "powershell.exe" `
 Argument = "-NoProfile -ExecutionPolicy Bypass -File `"$ScriptPaths`""
#>
$Tasks = @(
	@{
		EventSource = "Notify-Reenable_WiFi-Success"		
		EventID = 1000									
		TaskName = "USS-Notify-Reenable_WiFi-Success"
		TaskDescription = "Displays a toast notification when wifi is successfully re-enabled. Source: Reenable_WiFi package. Version: 1.1.1"
		Execute = "C:\Windows\System32\conhost.exe"
		Argument = "--headless powershell.exe -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$ScriptPath`" -Text `"WiFi has been successfully re-enabled.`" -Title `"WiFi Re-enabled`" -LauncherID `"$ShortcutScriptPath`""
		#WorkingDirectory = ""
	},
	
	@{
		EventSource = "Notify-Reenable_WiFi-Failure"		
		EventID = 1000									
		TaskName = "USS-Notify-Reenable_WiFi-Failure"
		TaskDescription = "Displays a toast notification when wifi fails to re-enable. Source: Reenable_WiFi package. Version: 1.1.1"
		Execute = "C:\Windows\System32\conhost.exe"
		Argument = "--headless powershell.exe -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$ScriptPath`" -Text `"An error has occurred attempting to re-enable WiFi.`" -Title `"Error`" -LauncherID `"$ShortcutScriptPath`""
		#WorkingDirectory = ""
	},
	
	@{
		EventSource = "Notify-Reenable_WiFi-Missing"		
		EventID = 1000									
		TaskName = "USS-Notify-Reenable_WiFi-Missing"
		TaskDescription = "Displays a toast notification when no disabled wifi adapters are found. Source: Reenable_WiFi package. Version: 1.1.1"
		Execute = "C:\Windows\System32\conhost.exe"
		Argument = "--headless powershell.exe -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$ScriptPath`" -Text `"No disabled WiFi adapters found.`" -Title `"Error`" -LauncherID `"$ShortcutScriptPath`""
		#WorkingDirectory = ""
	}
)

Start-Transcript $LogPath -Force

# Iterate over each task
foreach ($task in $Tasks) {
	Write-Host "Processing [$($task.TaskName)]"

	# Create the Scheduled Task
	Write-Host "Creating scheduled task [$($task.TaskName)]..."

	try {
		# Check if scheduled task already exists, or needs an update.
		$existingTask = Get-ScheduledTask -TaskName $task.TaskName -ErrorAction SilentlyContinue
		if ($existingTask -And ($existingTask.Description -ne $task.TaskDescription -Or $task.TaskDescription -eq $null)) {
			Write-Host "[$($task.TaskName)] already exists, but we have a new description. Deleting old version..."
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

		# Define principal (run as current user)
		$Principal = New-ScheduledTaskPrincipal `
			-UserId "$env:USERDOMAIN\$env:USERNAME" `
			-LogonType Interactive `
			-RunLevel Limited

		# Define settings
		$Settings = New-ScheduledTaskSettingsSet `
			-AllowStartIfOnBatteries `
			-DontStopIfGoingOnBatteries `
			-StartWhenAvailable `
			-ExecutionTimeLimit (New-TimeSpan -Hours 2) `
			-Hidden

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

Stop-Transcript | Out-Null
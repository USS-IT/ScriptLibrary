<#
	.SYNOPSIS
	Resets failed program / TS stuck installing on a remote or local computer. Also allows restarting ConfigMgr service.
	
	.DESCRIPTION
	Resets failed program / TS stuck installing on a remote or local computer. Lists all requests found with CompletionState = 'Failure'. If none are in the correct stalled state, gives choice to reset deployment policy and/or try restarting the Config Mgr service anyway.
	
	.PARAMETER ComputerName
	The name of the remote or local computer.
	
	.NOTES
	Once the request is deleted it may show up as either "Failed" or "Installed" under Available or Installation Status. Even if it says "Installed" it's probably not actually installed if the request needed to be deleted.
	
	If running locally, give the local computer name. Has additional options when running locally.
	
	Must be run as a user that has local admin privileges on the target machine (e.g., your local admin SC account).
	
	Created: 8-27-25
	Author: Matt Carras (mcarras8)
#>
param(
	[Parameter(Mandatory=$true)]
	[string]$ComputerName
)

if ([string]::IsNullOrWhitespace($ComputerName)) {
	Write-Error "Missing or invalid computer name [$ComputerName]"
} else {
	if ($ComputerName -eq ${ENV:COMPUTERNAME}) {
		Write-Host "[$ComputerName] appears to be the local machine"
		$doPingSuccessContinue = "Y"
	} elseif((Test-Connection -ComputerName $ComputerName -Count 1 -Quiet) -Or (Test-Connection -ComputerName $ComputerName -Count 1 -Quiet) -Or (Test-Connection -ComputerName $ComputerName -Count 1 -Quiet)) {
		Write-Host "[$ComputerName] appears to be online, querying..."
		$doPingSuccessContinue = "Y"
	} else {
		$doPingSuccessContinue = Read-Host "[$ComputerName] did not respond to any ping attempts, continue? (N/Y)"
	}
	
	if ($doPingSuccessContinue -eq "Y") {
		# Use WMI to check for all requests
		# CCM_ExecutionRequestEx should also include CCM_TSExecutionRequest
		$results=(gwmi -ComputerName $ComputerName -Namespace root\ccm\SoftMgmtAgent -Class CCM_ExecutionRequestEx -Filter "CompletionState = 'Failure'")
		if (($results | Measure).Count -le 0) {
			Write-Host "** No failed requests found on [$ComputerName]"
		} else {
			Write-Host "** Possible failed requests found on [$ComputerName]"
			# Output results
			$results | Select ContentID, ProgramID, @{N="IsTS"; Expression={ $_.__CLASS -eq "CCM_TSExecutionRequest"}}, @{N="ReceivedTime"; Expression={[System.Management.ManagementDateTimeConverter]::ToDateTime($_.ReceivedTime)}}, RunningState, State, CompletionState | ft
			$ignoreState = Read-Host "Reset all failed tasks regardless of completed state? (N/Y)"
			foreach($owmi in $results) {
				$cid = $owmi.ContentID
				# Open requests that are still pending with State 'Completed' are most likely stuck
				if ($owmi.State -eq 'Completed' -Or $ignoreState -eq "Y") {
					$owmi.Delete()
					Write-Host "** Deleting stalled request [$cid] on [$ComputerName]..."
					$doRestartService = $true
				}
			}
		}
		
		# Offer to reset policy, but only when run locally
		# https://learn.microsoft.com/en-us/answers/questions/123991/sccm-software-center-how-to-reset-or-cancel-an-app
		if ( $ComputerName -eq ${ENV:COMPUTERNAME} -And (Read-Host "Try resetting policy and restart the service to clear stuck deployments? (N/Y)") -eq "Y" ) {
			Invoke-CimMethod -ComputerName $ComputerName -Namespace root\ccm -ClassName SMS_Client -MethodName ResetPolicy -Arguments @{ uFlags = [uint32]1 }
			$doRestartService = $true
		} else {
			$doRestartService = (Read-Host "Try restarting ConfigMgr service and re-evaluate policy anyway? (N/Y)") -eq "Y"
		}
			
		# Restart the ccmexec service if needed
		if ($doRestartService) {
			Write-Host "** Restarting ccmexec service on [$ComputerName]..."
			$osvc = Get-Service -Name "ccmexec" -ComputerName $ComputerName
			if ($osvc) {
				Restart-Service -InputObj $osvc -Force
			}
			# Notify client to re-evaluate policy
			Write-Host "** Notifying [$ComputerName] to re-evaluate assignments"
			Invoke-WmiMethod -ComputerName $ComputerName -Namespace root\ccm -Class SMS_Client -Name TriggerSchedule -ArgumentList "{00000000-0000-0000-0000-000000000021}"
			Write-Host "** Notifying [$ComputerName] to re-evaluate machine policy"
			Invoke-WmiMethod -ComputerName $ComputerName -Namespace root\ccm -Class SMS_Client -Name TriggerSchedule -ArgumentList "{00000000-0000-0000-0000-000000000022}"
		} else {
			Write-Host "** No results match criteria, nothing to do"
		}
	}
}

Read-Host "Press enter to exit" | Out-Null

<#
Example of failed/stuck package

__GENUS                  : 2
__CLASS                  : CCM_ExecutionRequestEx
__SUPERCLASS             :
__DYNASTY                : CCM_ExecutionRequestEx
__RELPATH                : CCM_ExecutionRequestEx.RequestID="10640e55-4bbe-4085-b474-6474f83b4976"
__PROPERTY_COUNT         : 48
__DERIVATION             : {}
__SERVER                 : USS-IT-216NTH3
__NAMESPACE              : ROOT\ccm\SoftMgmtAgent
__PATH                   : \\USS-IT-216NTH3\ROOT\ccm\SoftMgmtAgent:CCM_ExecutionRequestEx.RequestID="10640e55-4bbe-4085
                           -b474-6474f83b4976"
AdvertID                 :
CompletionState          : Failure
ContentAccessRetryCount  : 0
ContentID                : CAS07337
ContentRequestGuid       : {A79219A0-C3EF-4671-8B44-E517F29718DD}
ContentType              : 0
ContentVersion           : 1
DependeePolicyExists     : True
DependencyCheckEvaluated : False
DisableMomAlerts         : False
DownloadStartedNotified  : False
ExecutionContextTempPath : C:\WINDOWS\TEMP\
IgnoreRunRerunFlags      : False
IsAdminContext           : True
LastStatusMessageID      : 0
Locations                : {}
MIFChecking              : True
MIFFileName              :
MIFPackageName           : USS-Restart Reminder (Local Script)
MIFPackagePublisher      : USS IT
MIFPackageVersion        : 1.0
MTCHandle                : {10640E55-4BBE-4085-B474-6474F83B4976}
MTCTaskPriority          : 20
NextRetryTime            : 19691231190000.000000+000
OptionalAdvertisements   : {CAS2CF48}
OwnerOfMTCTask           : True
ProcessCreationTimeHigh  : 31201101
ProcessCreationTimeLow   : 3105140909
ProcessID                : 5472
ProgramExitCode          : 0
ProgramID                : USS-Install Restart Reminder Local Script
ProgramReboot            : False
ReceivedTime             : 20250827082549.000000+000
ReferenceCount           : 0
RequestID                : 10640e55-4bbe-4085-b474-6474f83b4976
RetryCount               : 0
RetryInterval            :
RunInQuietMode           : True
RunningState             : Running
RunOnCompletion          : True
ScheduleType             :
SDKCallerId              :
State                    : Running
SuspendReboot            : False
TargetUser               : S-1-5-21-1214440339-484763869-725345543-5458093
TaskPauseReason          : 0
TSStep                   : False
UserScheduled            : False
PSComputerName           : USS-IT-216NTH3

Example of failed/stuck task sequence

__GENUS                  : 2
__CLASS                  : CCM_TSExecutionRequest
__SUPERCLASS             : CCM_ExecutionRequestEx
__DYNASTY                : CCM_ExecutionRequestEx
__RELPATH                : CCM_TSExecutionRequest.RequestID="e95a9eca-1f15-487e-bc16-92e68a6c4427"
__PROPERTY_COUNT         : 52
__DERIVATION             : {CCM_ExecutionRequestEx}
__SERVER                 : USS-IT-216NTH3
__NAMESPACE              : ROOT\ccm\SoftMgmtAgent
__PATH                   : \\USS-IT-216NTH3\ROOT\ccm\SoftMgmtAgent:CCM_TSExecutionRequest.RequestID="e95a9eca-1f15-487e
                           -bc16-92e68a6c4427"
AdvertID                 :
CompletionState          : Failure
ContentAccessRetryCount  : 0
ContentID                : CAS07587
ContentRequestGuid       : {00000000-0000-0000-0000-000000000000}
ContentType              : 0
ContentVersion           :
DependeePolicyExists     : True
DependencyCheckEvaluated : False
DisableMomAlerts         : False
DownloadStartedNotified  : False
ExecutionContextTempPath :
IgnoreRunRerunFlags      : False
IsAdminContext           : True
LastStatusMessageID      : 1073751859
Locations                : {}
MIFChecking              : False
MIFFileName              :
MIFPackageName           :
MIFPackagePublisher      :
MIFPackageVersion        :
MTCHandle                : {E95A9ECA-1F15-487E-BC16-92E68A6C4427}
MTCTaskPriority          : 20
NextRetryTime            : 19691231190000.000000+000
OptionalAdvertisements   : {CAS2CEAF}
OwnerOfMTCTask           : True
ProcessCreationTimeHigh  : 0
ProcessCreationTimeLow   : 0
ProcessID                : 0
ProgramExitCode          : 0
ProgramID                : *
ProgramReboot            : False
ReceivedTime             : 20250827084025.000000+000
ReferenceCount           : 0
RequestID                : e95a9eca-1f15-487e-bc16-92e68a6c4427
RetryCount               : 0
RetryInterval            :
RunInQuietMode           : True
RunningState             : NotifyExecution
RunOnCompletion          : True
ScheduleType             :
SDKCallerId              :
State                    : Ready
SuspendReboot            : False
TargetUser               : S-1-5-21-1214440339-484763869-725345543-5458093
TaskPauseReason          : 0
TS_ContentRequestGuid    : {dee65bfb-0707-47c3-8b96-433861bd0946, 67135b87-01cd-4d73-b94f-2c91d583d942,
                           8063c0fe-d239-4530-a001-c68e21fa3b39, 336aa8b2-0c56-44a1-85a2-93124618ae32...}
TS_MemberPackageID       : {CAS06D9F, CAS070C2, CAS070C2, CAS07586...}
TS_MemberPackageVersion  : {17, 4, 4, 2...}
TS_MemberProgramID       : {*, *, USS-Install Dell Command Update 5.4.0, USS-Dell DP-PB16250 (A02)...}
TSStep                   : False
UserScheduled            : False
PSComputerName           : USS-IT-216NTH3

#>

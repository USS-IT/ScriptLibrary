<#
	.SYNOPSIS
	Returns the status of all pending tasks/requests in ConfigMgr / SCCM client.
	
	.DESCRIPTION
	Returns the status of all pending tasks/requests in ConfigMgr / SCCM client.
	
	.NOTES
	You may also provide the IP instead of computer name. If running locally, give the computer name.
	
	Created: 8-27-25
	Author: mcarras8
#>

$comp = Read-Host "Enter computer name or IP"

if ([string]::IsNullOrWhitespace($comp)) {
	Write-Error "Missing or invalid computer name [$comp]"
} else {
	If((Test-Connection -ComputerName $comp -Count 1 -Quiet) -Or (Test-Connection -ComputerName $comp -Count 1 -Quiet) -Or (Test-Connection -ComputerName $comp -Count 1 -Quiet)) {
		Write-Host "[$comp] appears to be online, querying..."
	} else {
		Write-Warning "[$comp] did not respond to any ping attempts, querying anyway..."
	}
	
	# Use WMI to check for all requests
	# CCM_ExecutionRequestEx should also include CCM_TSExecutionRequest
	$results=(gwmi -Namespace root\ccm\SoftMgmtAgent -Class CCM_ExecutionRequestEx -ComputerName $comp)
	if (($results | Measure).Count -le 0) {
		Write-Host "** No requests found on [$comp]"
	} else {
		# Output results
		$results | Select ContentID, ProgramID, @{N="IsTS"; Expression={ $_.__CLASS -eq "CCM_TSExecutionRequest"}}, @{N="ReceivedTime"; Expression={[System.Management.ManagementDateTimeConverter]::ToDateTime($_.ReceivedTime)}}, RunningState, State, CompletionState | ft
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

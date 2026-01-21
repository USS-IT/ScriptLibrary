# MJC 3-4-22

# Wrap Get-WMIObject in a job to prevent possible hanging if WMI doesn't respond. Default: 30 seconds
function Get-WMIObjectAsJob{
	param (
		[string]$Class,
		[string[]]$ComputerName,
		[string]$Filter
	)
	$job = Get-WmiObject -Class $Class -ComputerName $ComputerName -Filter $Filter -AsJob | Wait-Job -Timeout 30
	if ($job.State -eq 'Completed') {
		return Receive-Job -Job $job
	}
	Write-Host "Get-WMIObjectAsJob: [$computerName] timed out"
}
function Get-LocalAdmins ($computerName) {
	# Properties: Caption, Domain, Name, SID
	$group = Get-WMIObjectAsJob -Class "win32_group" -ComputerName $computerName -Filter "LocalAccount=True AND SID='S-1-5-32-544'"
	If ($group) {
		# GroupComponent: Win32_Group.Domain="<COMPUTERNAME>",Name="Administrators"
		# PartComponent: \\<COMPUTERNAME>\root\cimv2:Win32_Group.Domain="WIN",Name="Domain Admins"
		$allAdmins = Get-WMIObjectAsJob -Class "win32_groupuser" -ComputerName $computerName -Filter "GroupComponent = `"Win32_Group.Domain='$($group.domain)'`,Name='$($group.name)'`"" | ForEach-Object { If ($_.PartComponent -match ".+Domain=`"([^`"]+)`",Name=`"([^`"]+)`"") { [PSCustomObject]@{ Domain = $matches[1]; Name = $matches[2]; SAMAccountName = "$($matches[1])\$($matches[2])" } } else { [PSCustomObject]@{ Domain = "<UNKNOWN>"; Name = "<UNKNOWN>"; SAMAccountName = "<UNKNOWN>" } } }
		$allAdmins
	}
}

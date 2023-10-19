# Remove compatibility flags from registry for target program and all users.
# These are the flags set when using the Compatibility tab on a program or shortcut.
# Give -Debug for more output.
# Usage: Remove-AppCompatFlags.ps1 "Path\To\Program.exe"
# MJC 9-30-22

# Target program to remove compatibility flags for.
# Ex: "${ENV:ProgramFiles(x86)}\IMR\Alchemy\ALCHEMY.EXE"
param($targetProgram)

if ([string]::IsNullOrWhitespace($targetProgram)) {
	Write-Error "Target program is required."
} else {
	Write-Debug "Target=$targetProgram"
	
	# Remove the entry from HKLM, if it exists.
	Remove-ItemProperty -Path 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\Layers' -Name $targetProgram -Force -ErrorAction SilentlyContinue

	# Loop through all domain users and remove comaptibility flags for the target program, if it exists.
	Get-ChildItem -Path Registry::HKEY_USERS | where {$_.PSChildName -match '^S\-1\-5\-21\-\d+\-\d+\-\d+\-\d+$'} | foreach {
		$result = Get-ItemProperty -Path ($_.PSPath + '\SOFTWARE\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\Layers') -Name $targetProgram -ErrorAction SilentlyContinue
		if ($result.$targetProgram -ne $null) {
			Write-Debug ('Removing from ' + $result.PSPath)
			Remove-ItemProperty -Path $result.PSPath -Name $targetProgram -Force
		}
	}
}

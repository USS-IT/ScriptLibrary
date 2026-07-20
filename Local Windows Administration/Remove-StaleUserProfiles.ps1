#Displays stale user profiles / user directories on the system and offers to delete them.

#Require RunAsAdministrator

$CutoffDate = (Get-Date).AddMonths(-1)

$userProfiles = Get-CimInstance Win32_UserProfile |
    Where-Object {
        -not $_.Special -and
        -not $_.Loaded
} | Select-Object @{	
		Name='Username'
		Expression={ Split-Path $_.LocalPath -Leaf }
	},
	LocalPath,
	LastUseTime,
	@{	Name='LastNTUserWrite'
		Expression={ 
			if (-Not [string]::IsNullOrWhitespace($_.LocalPath)) {
				$NtUserPath = Join-Path $_.LocalPath 'NTUSER.DAT'
				if((Test-Path $NtUserPath)) {
					(Get-Item $NtUserPath -Force).LastWriteTime
				}
			}
		}
	},
	@{	Name='CIMObject'
		Expression={ $_ }
	}

$userProfiles | Select Username,LocalPath,LastNTUserWrite,LastUseTime | Format-Table

if (($userProfiles | Measure).Count -gt 0) {
	foreach ($profile in $userProfiles) {
		$Username = $profile.Username
		if ($profile.CIMObject) {
			if($_.LastNTUserWrite -ge $CutoffDate -Or $_.LastUseTime -ge $CutoffDate) {
				$choice = Read-Host "$($Username) appears to be recently active. Delete their user profile anyway? (N/Y)"
			} else {
				$choice = Read-Host "Delete user profile for $($Username)? (N/Y)"
			}
			if ($choice -eq 'Y') {
				try {
					Remove-CimInstance -InputObject $profile.CIMObject
					Write-Host "Successfully deleted profile for $Username" -ForegroundColor Green
				}
				catch {
					Write-Warning "Failed to delete $Username : $_"
				}
			}
		} else {
			Write-Error "Invalid CIMObject for $Username"
		}
	}
} else {
	Write-Host "No user profiles found."
}
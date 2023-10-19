# Input CSV file must have column header "User". This can be either an email address (including aliases), UPN, or username/JHED.
# MJC 11-10-22
$defaultInputCSV = "${ENV:OneDrive}\users.csv"
$defaultOutputCSV = "${ENV:OneDrive}\results.csv"

$inputCSV = Read-Host "Input CSV filename (default: $defaultInputCSV)"
$outputCSV = Read-Host "Output CSV filename (default: $defaultOutputCSV)"
if ([string]::IsNullOrWhitespace($inputCSV)) {
	$inputCSV = $defaultInputCSV
}
if ([string]::IsNullOrWhitespace($outputCSV)) {
	$outputCSV = $defaultOutputCSV
}
$emails = Import-CSV $inputCSV
$notFoundCount = 0
Write-Host ("Processing {0} entries..." -f $emails.Count)
$emails | % { 
	$Email = $_.User
	$user=$null
	$isUsernameMatch=$false
	if (-Not $Email.Contains('@')) {
		$isUsernameMatch = $true
		$user = Get-ADUser $Email -Properties DisplayName,mail,Department,Company -ErrorAction SilentlyContinue
	} else {
		$user = Get-ADUser -LDAPFilter ("(|(UserPrincipalName=$Email)(proxyAddresses=smtp:$Email))") -Properties DisplayName,mail -ErrorAction SilentlyContinue
		# If not found, check again using only the username
		if (-Not [string]::IsNullOrEmpty($user.Name) -And $Email -match "(\w+)@" -And -Not [string]::IsNullOrWhitespace($matches[1])) {
			try {
				$user = Get-ADUser $matches[1] -Properties DisplayName,mail,Department,Company -ErrorAction SilentlyContinue
			} catch {
			}
		}
	}
	if (-Not [string]::IsNullOrEmpty($user.Name)) {
		$JHED=$user.Name
		$PrimaryEmail=$user.mail
		$DisplayName=$user.DisplayName
		$Department=$user.Department
		$Company=$user.Company
	} else {
		$Department=""
		$Company=""
		if ($isUsernameMatch) {
			$JHED=$Email
			$Email = "<USER NOT FOUND IN AD>"
			$PrimaryEmail="<USER NOT FOUND IN AD>"
			$DisplayName="<USER NOT FOUND IN AD>"
			
		} else {
			$JHED="<EMAIL/USER NOT FOUND IN AD>"
			$PrimaryEmail="<EMAIL/USER NOT FOUND IN AD>"
			$DisplayName="<EMAIL/USER NOT FOUND IN AD>"
		}
		$notFoundCount++
	}
	[PSCustomObject]@{
		User=$Email
		PrimaryEmail=$PrimaryEmail
		JHED=$JHED
		DisplayName=$DisplayName
		Department=$Department
		Company=$Company
	}
} | Export-CSV -NoTypeInformation $outputCSV
Write-Host("** {0} out of {1} users found in AD." -f ($emails.Count - $notFoundCount), $emails.Count)
Write-Host("** Results saved to [$outputCSV]")


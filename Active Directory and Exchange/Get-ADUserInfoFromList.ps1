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
$adProps = "DisplayName","mail","Department","Company","extensionAttribute2"
$emails = Import-CSV $inputCSV
$notFoundCount = 0
Write-Host ("Processing {0} entries..." -f $emails.Count)
$emails | % { 
	$Email = $_.User
	$user=$null
	$isUsernameMatch=$false
	if (-Not $Email.Contains('@')) {
		$isUsernameMatch = $true
		$user = Get-ADUser $Email -Properties $adProps -ErrorAction SilentlyContinue
	} else {
		$user = Get-ADUser -LDAPFilter ("(|(UserPrincipalName=$Email)(mail=$Email)(proxyAddresses=smtp:$Email))") -Properties $adProps -ErrorAction SilentlyContinue
		# If not found, check again using only the username
		if (-Not [string]::IsNullOrEmpty($user.Name) -And $Email -match "(\w+)@" -And -Not [string]::IsNullOrWhitespace($matches[1])) {
			try {
				$user = Get-ADUser $matches[1] -Properties $adProps -ErrorAction SilentlyContinue
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
		$Affiliation=$user.extensionAttribute2
	} else {
		$Department=""
		$Company=""
		$Affiliation=""
		if ($isUsernameMatch) {
			$JHED=$Email
			$Email = "<USER NOT FOUND>"
			$PrimaryEmail="<USER NOT FOUND>"
			$DisplayName="<USER NOT FOUND>"
			
		} else {
			$JHED="<EMAIL/USER NOT FOUND>"
			$PrimaryEmail="<EMAIL/USER NOT FOUND>"
			$DisplayName="<EMAIL/USER NOT FOUND>"
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
		Affiliation=$Affiliation
	}
} | Export-CSV -NoTypeInformation $outputCSV
Write-Host("** {0} out of {1} users found in AD." -f ($emails.Count - $notFoundCount), $emails.Count)
Write-Host("** Results saved to [$outputCSV]")


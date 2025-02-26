<#
	.SYNOPSIS
	Returns a list of users from AD given a CSV containing their Emails or JHEDs. The emails can be an alias.
	
	.DESCRIPTION
	Returns a list of users from AD given a CSV containing their Emails or JHEDs. Input CSV file must have column header "User". This can be either an email address (including aliases), UPN, or username/JHED.
	
	.NOTES
	The script will ask where the input and output CSVs should be located.

	Author: mcarras8
	Created: 11-10-22
	Last Updated: 2-26-25
#>

# Default output file locations.
$DEFAULT_INPUT_CSV = "${ENV:OneDrive}\users.csv"
$DEFAULT_OUTPUT_CSV = "${ENV:OneDrive}\results.csv"

# Default OU to restrict searches.
$USER_OU = "OU=PEOPLE,DC=win,DC=ad,DC=jhu,DC=edu"

$inputCSV = Read-Host "Input CSV filename (default: $DEFAULT_INPUT_CSV)"
$outputCSV = Read-Host "Output CSV filename (default: $DEFAULT_OUTPUT_CSV)"
if ([string]::IsNullOrWhitespace($inputCSV)) {
	$inputCSV = $DEFAULT_INPUT_CSV
}
if ([string]::IsNullOrWhitespace($outputCSV)) {
	$outputCSV = $DEFAULT_OUTPUT_CSV
}
$AD_PROPS = "DisplayName","mail","Department","Company","extensionAttribute2"
$emails = Import-CSV $inputCSV
$notFoundCount = 0
Write-Host ("Processing {0} entries..." -f $emails.Count)
$emails | % { 
	$Email = $_.User
	$user=$null
	$isUsernameMatch=$false
	$DisplayName=""
	$Department=""
	$Company=""
	$Affiliation=""
	if (-Not $Email.Contains('@')) {
		$isUsernameMatch = $true
		$user = Get-ADUser $Email -Searchbase $USER_OU -Properties $AD_PROPS -ErrorAction SilentlyContinue
	} else {
		$user = Get-ADUser -LDAPFilter ("(|(UserPrincipalName=$Email)(mail=$Email)(proxyAddresses=smtp:$Email)(proxyAddresses=SMTP:$Email))") -Searchbase $USER_OU -Properties $AD_PROPS -ErrorAction SilentlyContinue
		# If not found, check again using only the username
		if (-Not [string]::IsNullOrEmpty($user.Name) -And $Email -match "(\w+)@" -And -Not [string]::IsNullOrWhitespace($matches[1])) {
			try {
				$user = Get-ADUser $matches[1] -Searchbase $USER_OU -Properties $AD_PROPS -ErrorAction SilentlyContinue
			} catch {
			}
		}
	}
	# Check to see if we hacve more than one result.
	if ($user.Count -gt 1) {
		if ($user[0].distinguishedname -eq $user[1].distinguishedname -And -Not [string]::IsNullOrEmpty($user[0].distinguishedname)) {
			$user = $user[0]
		} else {
			Write-Warning("More than one result returned for [$Email], skipping: {0}, {1}" -f $user[0].Name, $user[1].Name)
			
			$JHED=$Email
			$PrimaryEmail="<TOO MANY RESULTS ($($user.Count))>"
			$DisplayName="<TOO MANY RESULTS ($($user.Count))>"
		}
	# Check to see if we have a valid entry.
	} elseif (-Not [string]::IsNullOrEmpty($user.Name)) {
		$JHED=$user.Name
		$PrimaryEmail=$user.mail
		$DisplayName=$user.DisplayName
		$Department=$user.Department
		$Company=$user.Company
		$Affiliation=$user.extensionAttribute2
	} else {
		if ($isUsernameMatch) {
			$JHED=$Email
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


<#
	.SYNOPSIS
	Outputs joined results from two CSV files on the given unique key.
	
	.DESCRIPTION
	Outputs joined results from two CSV files on the given unique key.
	
	.PARAMETER Left
	Required. The left CSV file to join. Values take precedence by default unless they are boolean (OR), or a datetime (either newest).
	
	.PARAMETER Right
	Required. The right CSV file to join. Values in same named columns are overriden by the Left table by default (see -Conditions).
	
	.PARAMETER On
	Primary key name from both tables to join on. Either -On or -LeftOn and -RightOn are required.
	
	.PARAMETER LeftOn
	Required with -RightOn. Primary key name from the left table to join on.
	
	.PARAMETER RightOn
	Required with -LeftOn. Primary key name from the right table to join on.
	
	.PARAMETER JoinType
	Optional. Either Inner or Outer (return both tables). Default is Outer.
	
	.PARAMETER Conditions
	Optional hash table map of precedence conditions. Key should be "ColumnName" with a nested hash table containing Precedence="Left" or "Right". If "AllowEmpty"=$true is given, it will also copy over empty values from the precedence column, if it exists.
	
	.PARAMETER LeftFilter
	Optional scriptblock filter for left CSV table as a string.
	
	.PARAMETER RightFilter
	Optional scriptblock filter for right CSV table as a string.
	
	.PARAMETER OutFile
	Optional filepath to output CSV file. Defaults to "joined_table.csv" in your OneDrive folder.

	.EXAMPLE
	. .\Join-CSVTables.ps1 -Left "snipeit_assets.csv" -Right "sccm_assets.csv" -LeftOn "name" -RightOn "Computer_Name" -LeftFilter '$_.Category -eq "PC" -And $_.status_label -notmatch "archived"' -OutFile "joined_assets.csv"
	
	.NOTES
	Author: Matthew Carras
#>
[CmdletBinding(DefaultParameterSetName='Default')]
param (
	[parameter(Mandatory=$true, Position=0)]
	[AllowEmptyCollection()]
	[array]$Left,
	
	[parameter(Mandatory=$true, Position=1)]
	[AllowEmptyCollection()]
	[array]$Right,
	
	[parameter(Mandatory=$true, ParameterSetName='Default')]
	[ValidateNotNullOrEmpty()] 
	[string]$On,
	
	[parameter(Mandatory=$true, ParameterSetName='LeftRightOn')]
	[ValidateNotNullOrEmpty()] 
	[string]$LeftOn,
	
	[parameter(Mandatory=$true, ParameterSetName='LeftRightOn')]
	[ValidateNotNullOrEmpty()]
	[string]$RightOn,
	
	[parameter(Mandatory=$false)]
	[ValidateSet("Inner","Outer")]
	[string]$JoinType="Outer",

	[parameter(Mandatory=$false)]
	[hashtable]$Conditions,
	
	[parameter(Mandatory=$false)]
	[string]$LeftFilter,
	
	[parameter(Mandatory=$false)]
	[string]$RightFilter,
	
	[parameter(Mandatory=$false)]
	[string]$OutFile="${ENV:OneDrive}\Documents\joined_table.csv"
)

# -- START FUNCTIONS --
function Join-Objects {
	<#
	.SYNOPSIS
	Joins two objects on the given unique key.
	
	.DESCRIPTION
	Joins two objects on the given unique key.

	.PARAMETER Left
	Required. The left object to join. Values take precedence by default unless they are boolean (OR), or a datetime (either newest).
	
	.PARAMETER Right
	Required. The right object to join. Values in same named columns are overriden by the Left array by default (see -Conditions).
	
	.PARAMETER On
	Primary key name from both tables to join on. Either -On or -LeftOn and -RightOn are required.
	
	.PARAMETER LeftOn
	Required with -RightOn. Primary key name from the left table to join on.
	
	.PARAMETER RightOn
	Required with -LeftOn. Primary key name from the right table to join on.
	
	.PARAMETER JoinType
	Optional. Either Inner or Outer (return both tables). Default is Outer.
	
	.PARAMETER Conditions
	Optional hash table map of precedence conditions. Key should be "ColumnName" with a nested hash table containing Precedence="Left" or "Right". If "AllowEmpty"=$true is given, it will also copy over empty values from the precedence column, if it exists.
	
	.OUTPUTS
	Joined objects.
#>
	[CmdletBinding(DefaultParameterSetName='Default')]
	param (
		[parameter(Mandatory=$true, Position=0)]
        [AllowEmptyCollection()]
		[array]$Left,
		
		[parameter(Mandatory=$true, Position=1)]
        [AllowEmptyCollection()]
		[array]$Right,
		
		[parameter(Mandatory=$true, ParameterSetName='Default')]
		[ValidateNotNullOrEmpty()] 
		[string]$On,
		
		[parameter(Mandatory=$true, ParameterSetName='LeftRightOn')]
		[ValidateNotNullOrEmpty()] 
		[string]$LeftOn,
		
		[parameter(Mandatory=$true, ParameterSetName='LeftRightOn')]
		[ValidateNotNullOrEmpty()]
		[string]$RightOn,
		
		[parameter(Mandatory=$false)]
		[ValidateSet("Inner", "Outer")]
		[string]$JoinType="Outer",
	
		[parameter(Mandatory=$false)]
		[hashtable]$Conditions
	)
	Begin {
		# Validate columns exist
		if (($Left | Measure).Count -gt 0) {
			$leftCols = $Left | Get-Member -MemberType NoteProperty | Select -ExpandProperty Name -Unique
			if (($On -ne $null -And -Not $On -in $leftCols) -Or ($LeftOn -ne $null -And -Not $LeftOn -in $leftCols)) {
				Throw "Column [$LeftOn] not found in left array"
			}
		}
		if (($Right | Measure).Count -gt 0) {
			$rightCols = $Right | Get-Member -MemberType NoteProperty | Select -ExpandProperty Name -Unique
			if (($On -ne $null -And -Not $On -in $rightCols) -Or ($RightOn -ne $null -And -Not $RightOn -in $leftCols)) {
				Throw "Column [$RightOn] not found in right array"
			}
		}
		
		# Join the two sets of columns and set the join column
		$joinedCols = ($leftCols + $rightCols) | Select -Unique
		if (-Not [string]::IsNullOrEmpty($LeftOn)) {
			if ($LeftOn -ne $RightOn -And -Not [string]::IsNullOrEmpty($RightOn)) {
				$joinedCols = $joinedCols | where {$_ -ne $RightOn}
			}
		
			$joinOn = $LeftOn
		} else {
			$joinOn = $On
		}
	}
	
	Process {
		Write-Verbose("[Join-Objects] Joining {0} objects (left) with {1} objects (right) on [{2}] with JoinType [{3}]..." -f $Left.Count,$Right.Count,$On,$JoinType)

		# If either is empty, return the other.
		if (($Right | Measure).Count -eq 0) {
			if ($JoinType -eq "Inner") {
				return $null
			} else {
				return $Left
			}
		} elseif (($Left | Measure).Count -eq 0) {
			if ($JoinType -eq "Inner") {
				return $null
			} else {
				return $Right
			}
		}
		
		# If our primary key column names are different, change the right column name to match
		if (-Not [string]::IsNullOrEmpty($LeftOn) -And -Not [string]::IsNullOrEmpty($RightOn) -And $RightOn -ne $LeftOn) {
			$Right = $Right | Select *,@{N=$LeftOn; Expression={ $_.$RightOn }} -ExcludeProperty $RightOn
		}
		
		# Construct the grouped object
		$countMap = @{}
		$results = ($Left + $Right) | where {$JoinType -eq "Outer" -Or $_.$joinOn -ne $null} | Group-Object -Property $joinOn | foreach { 
			if ($_.Count -eq 1) {
				$countMap[$_.Name] = $_.Count
				[PSCustomObject]($_.Group | Select -First 1)
			} elseif ($_.Group[0].$joinOn -ne $null -And $_.Group[1].$joinOn -ne $null) {
				$countMap[$_.Name] = $_.Count
				$o = [PSCustomObject]@{ }
				foreach ($p in $joinedCols) {
					if ($_.Group[0].$p -is [bool] -Or $_.Group[1].$p -is [bool]) {
						$val = ($_.Group[0].$p -Or $_.Group[1].$p)
					} else {
						if ($Conditions.$p.Precedence -eq "Right") {
							$lInt = 1
							$rInt = 0
						} else {
							$lInt = 0
							$rInt = 1
						}
						$val = $_.Group[$lInt].$p
						# Default conditions to use right table's values:
						# - If value is null
						# - or value is an empty string and other entry is not empty
						# - or value is a newer datetime
						if (($val -eq $null -And -Not $Conditions.$p.AllowEmpty) -Or
							($val -is [string] -And [string]::IsNullOrEmpty($val) -And -Not $Conditions.$p.AllowEmpty -And -Not [string]::IsNullOrEmpty($_.Group[$rInt].$p)) -Or 
							($val -is [DateTime] -And $_.Group[$rInt].$p -is [DateTime] -And $_.Group[$rInt].$p -gt $val)) {
							$val = $_.Group[$rInt].$p
						}
					}
					Add-Member -InputObject $o -MemberType NoteProperty -Name $p -Value $val -Force
				}
				$o
			}
		}
		
		# Default is Outer / Full Outer
		switch($JoinType) {
			"Inner" {
				$results = $results | where {$countMap[$_.$joinOn] -gt 1}
			}
		}
		
		Write-Verbose("[Join-Objects] Returning [{0}] objects..." -f ($results | Measure).Count)
		return $results
	}
	End {
	}
}
# -- END FUNCTIONS --

$leftTable = Import-CSV $Left
$rightTable = Import-CSV $Right
if ($LeftFilter) {
	$leftTable = $leftTable.where([scriptblock]::Create($LeftFilter))
}
if ($RightFilter) {
	$rightTable = $rightTable.where([scriptblock]::Create($RightFilter))
}

if ($leftTable -And $rightTable) {
	$extraParams = @{}
	if ($RightOn) {
		$extraParams["RightOn"] = $RightOn
	}
	if ($Conditions) {
		$extraParams["Conditions"] = $Conditions
	}
	$joinedTable = Join-Objects -Left $leftTable -Right $rightTable -LeftOn $LeftOn -JoinType $JoinType -Verbose @extraParams
	
	$joinedTable | Export-CSV -NoTypeInformation $OutFile
}



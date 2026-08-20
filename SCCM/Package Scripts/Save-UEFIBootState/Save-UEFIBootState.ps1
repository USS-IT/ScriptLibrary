# Captures all variables and file info related to the UEFI boot state and Secure Boot.
# Author: Matt Carras (mcarras8)

#Requires -RunAsAdministrator

param(
    [string]$OutDir = "C:\UEFI-State"
)

if (-Not (Test-Path $OutDir)) {
	New-Item -ItemType Directory -Force -Path $OutDir | Out-Null
}

Write-Host "Collecting system state into $OutDir"

# ---------------------------------------------------------------------
# System Info
# ---------------------------------------------------------------------

Get-ComputerInfo |
    Out-File "$OutDir\ComputerInfo.txt"

# ---------------------------------------------------------------------
# Secure Boot Registry State
# ---------------------------------------------------------------------

reg export HKLM\SYSTEM\CurrentControlSet\Control\SecureBoot `
    "$OutDir\SecureBoot.reg" /y

<#
reg export HKLM\SYSTEM\Setup `
    "$OutDir\Setup.reg" /y
#>

# ---------------------------------------------------------------------
# Secure Boot Variables
# ---------------------------------------------------------------------

$SecureBootDir = "$OutDir\SecureBootVariables"
New-Item -ItemType Directory -Force -Path $SecureBootDir | Out-Null

foreach ($Name in 'PK','KEK','db','dbx') {

    try {

        $Var = Get-SecureBootUEFI -Name $Name -Decoded | 
			Out-File "$SecureBootDir\$Name.decoded.txt"

        $Var = Get-SecureBootUEFI -Name $Name

        if ($Var.Bytes) {

            [IO.File]::WriteAllBytes(
                "$SecureBootDir\$Name.bin",
                $Var.Bytes
            )

            Get-FileHash "$SecureBootDir\$Name.bin" -Algorithm SHA256 |
                Export-Csv "$SecureBootDir\$Name.hash.csv" -NoTypeInformation
        }
    }
    catch {

        $_ |
            Out-File "$SecureBootDir\$Name.error.txt"
    }
}

# ---------------------------------------------------------------------
# Secure Boot SVN
# ---------------------------------------------------------------------

try {

    Get-SecureBootSVN |
        Format-List * |
        Out-File "$OutDir\SecureBootSVN.txt"
}
catch {

    $_ | Out-File "$OutDir\SecureBootSVN.error.txt"
}

# ---------------------------------------------------------------------
# BCD
# ---------------------------------------------------------------------

$bcdDir = "$OutDir\BCD"
New-Item -ItemType Directory -Force -Path $bcdDir | Out-Null

cmd /c 'bcdedit /enum all /v'      > "$bcdDir\bcd_all.txt"
cmd /c 'bcdedit /enum firmware /v' > "$bcdDir\bcd_firmware.txt"

# ---------------------------------------------------------------------
# Mount ESP
# ---------------------------------------------------------------------

$espDrive = "S:"

mountvol $espDrive /S

try {

    $EspDir = "$OutDir\ESP"
    New-Item -ItemType Directory -Force -Path $EspDir | Out-Null

    dir "$espDrive\" -Recurse -Force |
        Select-Object FullName,
                      Length,
                      CreationTime,
                      LastWriteTime |
        Export-Csv "$EspDir\FileInventory.csv" -NoTypeInformation

    Get-ChildItem "$espDrive\" -Recurse -File -Force |
    ForEach-Object {

        try {

            $Hash = Get-FileHash $_.FullName -Algorithm SHA256

            [PSCustomObject]@{
                FullName      = $_.FullName
                SHA256        = $Hash.Hash
                Length        = $_.Length
                LastWriteTime = $_.LastWriteTime
            }
        }
        catch {}
    } |
    Export-Csv "$EspDir\FileHashes.csv" -NoTypeInformation

    # Target the interesting files directly

    $Targets = @(
        "$espDrive\EFI\Microsoft\Boot\bootmgfw.efi",
        "$espDrive\EFI\Microsoft\Boot\bootmgr.efi",
        "$espDrive\EFI\Boot\bootx64.efi"
    )

    foreach ($File in $Targets) {

        if (Test-Path $File) {

            $Base = Split-Path $File -Leaf

            Get-Item $File |
                Format-List * |
                Out-File "$EspDir\$Base.properties.txt"

            Get-AuthenticodeSignature $File |
                Format-List * |
                Out-File "$EspDir\$Base.signature.txt"

            Copy-Item $File "$EspDir\$Base" -Force
        }
    }
}
finally {

    mountvol $espDrive /D
}

# ---------------------------------------------------------------------
# Windows Boot Files
# ---------------------------------------------------------------------

$BootFilesDir = "$OutDir\WindowsBootFiles"
New-Item -ItemType Directory -Force -Path $BootFilesDir | Out-Null

$WindowsBootFiles = @(
    "$env:windir\Boot\EFI\bootmgfw.efi",
    "$env:windir\Boot\EFI\bootmgr.efi",
    "$env:windir\System32\winload.efi"
)

foreach ($File in $WindowsBootFiles) {

    if (Test-Path $File) {

        $Name = Split-Path $File -Leaf

        Get-AuthenticodeSignature $File |
            Format-List * |
            Out-File "$BootFilesDir\$Name.signature.txt"

        Get-FileHash $File -Algorithm SHA256 |
            Export-Csv "$BootFilesDir\$Name.hash.csv" -NoTypeInformation
    }
}

Write-Host ""
Write-Host "Done."
Write-Host $OutDir
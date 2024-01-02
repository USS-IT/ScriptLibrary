param([int]$action)

$today=$(Get-Date)
$yesterday=$((Get-Date).AddDays(-1))
$folderPath="C:\Windows\Logs\USS-LogonHistory"
$scriptPath=$folderPath + "\logonHistory.ps1"
$todayLogPath=$folderPath + "\logonHistory - $(Get-Date -Date $today -UFormat "%m-%d-%Y") - $($env:computername).csv"
$yesterdayLogPath=$folderPath + "\logonHistory - $(Get-Date -Date $yesterday -UFormat "%m-%d-%Y") - $($env:computername).csv"

$SmtpServer='smtp.johnshopkins.edu'
$From=[mailaddress]'ussitservices@jhu.edu'
$To=[mailaddress]'ldibern1@jh.edu'

function setupCheck{
    # If the log folder path does not exist, create it!
    If(-Not (Test-Path -Path $folderPath)) {
        New-Item -Path "C:\Windows\Logs\" -Name "USS-LogonHistory" -ItemType "directory"
    }

    # If the script has not been downloaded locally, download it! It's required for the the scheduled task
    # to send the logfile in an email.
    If(-Not (Test-Path -Path $scriptPath -PathType Leaf)) {
        $scriptLink = "https://raw.githubusercontent.com/USS-IT/ScriptLibrary/LogonHistory/Startup%20or%20Login%20Scripts/LogonHistory.ps1"
        Invoke-WebRequest -URI $scriptLink -OutFile $scriptPath -UseBasicParsing
    }

    # If today's log file does not exist, create it!
    If (-Not (Test-Path -Path $todayLogPath -PathType Leaf)) {
        createTodayLogFile
    }
}

function createTodayLogFile {
    New-Item -Path $todayLogPath -ItemType "file"
    Set-Content -Path $todayLogPath -Value "date,time,type,jhed"
}

function deleteYesterdayLogFile {
    Remove-Item -Path $yesterdayLogPath
}

function handleAction{
    switch ($action)
    {
        "Logon" {
            Add-Content -Path $todayLogPath -Value "$(Get-Date -Date $today -UFormat "%m/%d/%Y"),$(Get-Date -UFormat "%R"),login,$Env:UserName"
        }
        "Log-off" {
            Add-Content -Path $todayLogPath -Value "$(Get-Date -Date $today -UFormat "%m/%d/%Y"),$(Get-Date -UFormat "%R"),logoff,$Env:UserName"
        }
        "Email"{
            Send-MailMessage -From $From -To $To -Subject "$($env:computername): $(Get-Date -Date $date -UFormat "%m/%d/%Y") Logs" -SmtpServer $SmtpServer -Attachments $yesterdayLogPath

            #now that we've sent the log history, delete yesterday's logfile.
            deleteYesterdayLogFile
        }
        Default {
        }
    }
}

main {
    setupCheck
    handleAction
}

main


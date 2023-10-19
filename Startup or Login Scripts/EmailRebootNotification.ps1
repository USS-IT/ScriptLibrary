# Sends an email notification to the given address using SMTP relay.
# Only works on domain networks to addresses in the same domain.
# MJC 5-5-23
Param (
	[mailaddress] $To,
	[mailaddress] $From=[mailaddress]'hsaitservices@jhu.edu',
	[string] $SmtpServer='smtp.johnshopkins.edu'
)

$serial = Get-WmiObject win32_bios | Select -ExpandProperty Serialnumber
$message = "This is an automated message from USS IT. Security Desk Workstation [{0}] with service tag/serial [{1}] is restarting or shutting down." -f ${ENV:COMPUTERNAME}, $serial
$subject = "Security Desk Restart Notification: {0}" -f ${ENV:COMPUTERNAME}

Send-MailMessage -From $From -To $To -Subject $subject -Body $message -SmtpServer $SmtpServer -DeliveryNotificationOption OnSuccess, OnFailure

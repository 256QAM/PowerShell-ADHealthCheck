$MailParams = @{
	'ToAddress'		= 'username@domain.com'
	'FromAddress'	= 'username@domain.com'
	'Subject' 		= 'Active Directory Health Report'
	'SMTPServer' 	= 'smtp.mailserver.com'
	'Port'			= '995'
	'SSL'			= $true
	'UseAuth'		= $true
	'UserName' 		= 'username@domain.com'
	'Password' 		= ([System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR((Get-Content ".\Resources\MailPassword" | ConvertTo-SecureString))))
}
function Send-Email {
	[CmdletBinding(DefaultParametersetName='None')]
	param(
		[Alias("T")]
		[Parameter(Mandatory=$true)]
		 [string[]]$ToAddress,
		[Alias("F")]
		[Parameter(Mandatory=$true)]
		 [string]$FromAddress,
		[Alias("S")]
		[Parameter(Mandatory=$false)]
		 [string]$Subject=' ',
		[Alias("B")]
		[Parameter(Mandatory=$false,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true,Position=0)]
		 $Body=' ',
		[Alias("Se")]
		[Parameter(Mandatory=$true)]
		 [string]$SMTPServer,
		[Alias("P")]
		[Parameter(Mandatory=$false)]
		 [int]$Port=25,
		[Alias("sS")]
		[Parameter(Mandatory=$false)]
		 [switch]$SSL,
		[Alias("sA")]
		[Parameter(ParameterSetName='Auth',Mandatory=$false)]
		 [switch]$UseAuth,
		[Alias("sAu")]
		[Parameter(ParameterSetName='Auth',Mandatory=$true)]
		 [string]$UserName,
		[Alias("sAp")]
		[Parameter(ParameterSetName='Auth',Mandatory=$true)]
		 [string]$Password
    )
	
	begin {}

	Process {
	
		$MailMessage = New-Object System.Net.Mail.MailMessage
		$SMTPClient = New-Object Net.Mail.SmtpClient($SMTPServer, $SMTPPort)
		
		$ToAddress | % {$MailMessage.To.Add($_)}
		$MailMessage.From = $FromAddress
		$MailMEssage.Subject = $Subject
		$MailMessage.Body = $Body 
		$MailMessage.isBodyHtml = $True
		
	
		$SMTPClient.EnableSsl = $SSL
		if ($UseAuth){
			$SMTPClient.Credentials = New-Object System.Net.NetworkCredential($UserName, $Password)
		}
		try {
			$Result = $SMTPClient.Send($MailMessage)
		} catch {
			$Result = $_.exception.message
		}
		Return $Result
	}
	
	end {}
}
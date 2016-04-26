Function Create-SecurePassFile {
	[CmdletBinding(DefaultParametersetName='None')]
		param(
			[Alias("O")]
			 [string]$OutputFileName = ".\Password",
			[Alias("P")]
			 [string]$PlainTextPassword
		)
	begin {
		if ((Test-Path $OutputFileName) -like $false) {
			New-Item $OutputFileName -type file
		} else {
			Clear-Content $OutputFileName
		}
	}
	Process {
		ConvertTo-SecureString -String $PlainTextPassword -AsPlainText -Force |`
		 ConvertFrom-SecureString |`
		 Set-Content ($OutputFileName)
	}
	End {
		Write-Host ("`nPassword saved to $OutputFileName`n`n")
		Return (Get-Content ($OutputFileName))
	}
}
<#   
    .SYNOPSIS   
		Checks AD Health and System/Service Integrity.
         
    .DESCRIPTION   
		This script should be able to give an automated overview of the health of an Active Directory environment
		in an HTML formatted output. It Currently only supports the automation and reporting of most standard 
		DCDiag functions, however future iterations are planned to support Exchange, Certificate Services, and 
		other AD-integreated services.
		
		The script requires an encrypted passfile to support the email functions. This can be generated by running 
		the script with the '-CreatePassFile' switch, which requires the -Password paramter to be specified. Enter 
		your password in plain text, and the script will create an encrypted copy of it within the Resources 
		directory. This password can ONLY be decrypted by the account that was used to generate it (e.g. the 
		account running this script).
		
		If you do not need the email function, the script has a '-NoEmail' switch that can bypass this function.
		
		The script logs it's output to the file path specified within. The default is: '.\AdReport.log'.
		
		E-Mail settings can be adjusted in the '.\Resources\MailParams.ps1' file. It is not reccomended or needed 
		to adjust the 'Password' paramter.
		
		This script only supports environments with Forest/Domain Functional Levels of Windows Server 2008 and 
		newer!
	
	.PARAMETER CreatePassFile
		Puts the script into encrypted password creation mode. No input required.
		
	.PARAMETER Password
		The E-Mail password to be encrypted and stored. Only needed when the '-CreatePassFile' switch is used. Input 
		should be plaintext string.
		
	.PARAMETER NoEmail
		Bypasses the email reporting module. No input required.
	
	.PARAMETER ReportPath
		Specifies the output path\file to store the generated report. Can be used to create dated reports. Input 
		should be plaintext string.
		
	.PARAMETER LogPath
		Specifies the output path\file to store the script logs. Input should be plaintext string.
		
	.PARAMETER TimeOut
		How long, in seconds, to wait for a DC to timeout before failing current test. Input should be positive 
		integer.
		
	.PARAMETER Days
		How long, in days, should the script look back for Directory Service Event Log errors. Input should be 
		positive integer.

	.NOTES
		Author: Stephen Arnold
		
		Version: 1.1 - 04/26/2016
			Added examples for those who don't like to read instructions.
		Version: 1.0 - 04/26/2016
			Added Logging functions and improved error handling.
			Started tracking version dates. Previous versions are ageless and have always existed in 
			  superposition with each other. Something about a cat in a box.
		Version 0.9
			Added Event Log error value to DC Details check.
			Cleaned up HTML (fixed the tables).
			Added passfile creation function.
		Version 0.8
			Started tracking versions...
			Skipped a few version numbers.
			Fixed the email module.
		Version 0.1
			Everything else.
  
	.EXAMPLE
		.\ADHealthCheck.ps1
		Runs the script with default values.
	
	.EXAMPLE
		.\ADHealthCheck.ps1 -CreatePassFile -Password 'ExamplePass'
		Runs the script in password creation mode, with the password 'ExamplePass' to be encrypted and saved.
	
	.EXAMPLE
		.\ADHealthCheck.ps1 -TimeOut "120"
		Runs script tests with a 120 second timeout before automatic failure.
	
	.EXAMPLE
		.\ADHealthCheck.ps1 -ReportPath 'ReportName.html' -Days 2
		Runs script with a custom report name, 'ReportName.html' and looks for Event Log errors within the last
		48 hours.
		  
	.EXAMPLE
		.\ADHealthCheck.ps1 -NoEmail
		Runs script and generates html report only. Does not email the report. 
		
#> 
[CmdletBinding(DefaultParametersetName='None')]
	param(
		[Parameter(ParameterSetName='PassGen',Mandatory=$false)]
		 [switch]$CreatePassFile,
		[Parameter(ParameterSetName='PassGen',Mandatory=$true)]
		 [string]$Password,
		[switch]$NoEmail,
		[string]$ReportPath = ".\ADReport.htm",
		[string]$LogPath = ".\ADReport.log",
		[string]$TimeOut = 60,
		[string]$Days = 1
	)
	
begin { 

	#Set some basic variables, report paths and colors
	$ReportDate = 	(Get-Date)
	$Report = 	$ReportPath
	$Log = 		$LogPath
	$NameColor = 	'LightGrey'
	$TimeOutColor = 'GoldenRod'
	$OnlineColor = 	'LawnGreen'
	$SuccessColor = 'HoneyDew'
	$FailColor = 	'Crimson'
	$TitleColor=	'SkyBlue'
	
	if ((test-path $LogPath) -like $false) {
		new-item $LogPath -type file > $null
		Add-Content $Log ((Get-Date -Format hh:mm:ss)+': 0:1-Info: Log file created at '+($LogPath))
	} else {
		Clear-Content $Log
	}

	#Import resources, E-Mail module and sending parameters
	Add-Content $Log ((Get-Date -Format hh:mm:ss)+": 0:1-Action: Script start. Importing resources.")
	. .\Resources\Create-SecurePassFile.ps1
	. .\Resources\Send-Email.ps1
	. .\Resources\MailParams.ps1
	
	if ($CreatePassFile) {
		Add-Content $Log ((Get-Date -Format hh:mm:ss)+": 0:1-Action: Running script in password creation mode")
		try {
			Create-SecurePassfile -P $Password -O '.\Resources\MailPassword'
			Add-Content $Log ((Get-Date -Format hh:mm:ss)+': 0:2-Info:   Password file created at: ".\Resources\MailPassword"')
			Clear-Host
			Write-Host "Password created. Please re-run script without the '-CreatePassFile' switch"
			Write-Host "Press any key to end..."
			$host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown") > $null
		} catch {
			Add-Content $Log ((Get-Date -Format hh:mm:ss)+': 0:0-Error:  Passfile creation at ".\Resources\MailPassword" failed: '+($_.Exception.Message))
			Clear-Host
			Write-Host "Password creation failed. Please review logs and re-run"
			Write-Host "Press any key to end..."
			$host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown") > $null
		}
		break
	}
	
	#Check if reportfile exists; create it if it does not
	if ((test-path $ReportPath) -like $false) {
		Add-Content $Log ((Get-Date -Format hh:mm:ss)+": 0:1-Action: No report file found. Attempting to create at "+($ReportPath))
		try {
			new-item $ReportPath -type file > $null
		} catch {
			Add-Content $Log ((Get-Date -Format hh:mm:ss)+': 0:0-Error:  Report file creation at '+($ReportPath)+' failed: '+($_.Exception.Message))
		}
	}
	
	#Invert $Days paramter (Can't read the future!)
	if ($Days -gt 0) {
		$DaysA = ($Days / -1)
	}

	#Checks status of the specified service
	function ServiceStatus($SvcName, $DC) {
		Add-Content $Log ((Get-Date -Format hh:mm:ss)+": 0:1-Action: Running Service test: "+($SvcName)+" for: "+($DC))
		$SvcStatus = start-job -scriptblock {get-service -ComputerName $($args[0]) -Name $($args[1]) -ErrorAction SilentlyContinue} -ArgumentList $DC,$SvcName | Wait-Job -timeout $TimeOut
		if($SvcStatus.state -like "Running") {
			Return ("<td bgcolor=$TimeOutColor align=center><b>Timeout</b></td>")
			Stop-Job $SvcStatus
		} else {
			$SvcStatus1 = Receive-job $SvcStatus
			if ($SvcStatus1.status -eq "Running") {
				$SvcName = $SvcStatus1.name 
				$SvcState = $SvcStatus1.status          
				Return ("<td bgcolor=$SuccessColor align=center><b>$SvcState</b></td>")
			} else { 
				$SvcName = $SvcStatus1.name 
				$SvcState = $SvcStatus1.status          
				Return ("<td bgcolor=$FailColor align=center><b>$SvcState</b></td>")
			} 
		}
	}
	
	#Gets DCDIAG output of specified service test
	function ServiceDiag($SvcName, $DC) {
		Add-Content $Log ((Get-Date -Format hh:mm:ss)+": 0:1-Action: Running DCDiag test: "+($SvcName)+" for: "+($DC))
		add-type -AssemblyName microsoft.visualbasic
		$cmp = "microsoft.visualbasic.strings" -as [type]
		$SysVol = start-job -scriptblock {dcdiag /test:$($args[0]) /s:$($args[1])} -ArgumentList $SvcName,$DC | Wait-Job -timeout $TimeOut
		if($SysVol.state -like "Running") {
		   Return ("<td bgcolor=$TimeOutColor align=center><b>Timeout</b></td>")
		   Stop-Job $SysVol
		} else {
			$SysVol1 = Receive-Job $SysVol
			if($cmp::instr($SysVol1, ("passed test "+$SvcName))) {
				Return ("<td bgcolor=$SuccessColor align=center><b>Passed</b></td>")
			} else {
				Return ("<td bgcolor=$FailColor align=center><b>Failed</b></td>")
			}
		}
	}
	
	#Gets number of AD error events in last n days (n = $Days)
	function ADErrorCounter($DC, $DaysA) {
		Add-Content $Log ((Get-Date -Format hh:mm:ss)+": 0:1-Action: Attempting to get event log information for: "+($DC))
		try {
			$Count = invoke-Command -ComputerName $DC -ArgumentList $DC,$DaysA -ScriptBlock { `
			 (Get-EventLog `
			  -LogName "Directory Service" `
			  -EntryType Error `
			  -ComputerName $($args[0]) `
			  -After (Get-Date).AddDays($($args[1])) `
			 ).Count
			}
		} catch {
			$Count = ("<font color=$FailColor>Error: "+($_.Exception.Message)+"</font>")
			Add-Content $Log ((Get-Date -Format hh:mm:ss)+": 0:0-Error:  EventLog query failed: "+($_.Exception.Message))
		}
		Return $Count
	}
}

process {
	
	#Clear and lingering content in reportfile
	Clear-Content $Report

	####START OF REPORT####
	Add-Content $Report "<html>" 
	Add-Content $Report "<head>" 
	Add-Content $Report "<meta http-equiv='Content-Type' content='text/html; charset=iso-8859-1'>" 
	Add-Content $Report '<title>AD Status Report</title>' 
	add-content $Report '<STYLE TYPE="text/css">' 
	add-content $Report "<!--" 
	add-content $Report "td {" 
	add-content $Report "font-family: Tahoma;" 
	add-content $Report "font-size: 11px;" 
	add-content $Report "border-top: 1px solid #999999;" 
	add-content $Report "border-right: 1px solid #999999;" 
	add-content $Report "border-bottom: 1px solid #999999;" 
	add-content $Report "border-left: 1px solid #999999;" 
	add-content $Report "padding-top: 0px;" 
	add-content $Report "padding-right: 0px;" 
	add-content $Report "padding-bottom: 0px;" 
	add-content $Report "padding-left: 0px;" 
	add-content $Report "}" 
	add-content $Report "body {" 
	add-content $Report "margin-left: 5px;" 
	add-content $Report "margin-top: 5px;" 
	add-content $Report "margin-right: 0px;" 
	add-content $Report "margin-bottom: 10px;" 
	add-content $Report "" 
	add-content $Report "table {" 
	add-content $Report "border: thin solid #000000;" 
	add-content $Report "}" 
	add-content $Report "-->" 
	add-content $Report "</style>" 
	Add-Content $Report "</head>" 
	Add-Content $Report "<body>" 
	add-content $Report "<table width='100%'>" 
	add-content $Report "<tr bgcolor='Lavender'>" 
	add-content $Report "<td colspan='7' height='75' align='center'>" 
	add-content $Report "<font face='tahoma' color='#003399' size='4'><strong>Active Directory Health Check</strong></font>" 
	add-content $Report "</td>" 
	add-content $Report "</tr>" 
	add-content $Report "</table>" 
	add-content $Report "<table width='100%'>" 
	Add-Content $Report "<tr bgcolor=$TitleColor>" 
	
	#Title names for first table columns
	Add-Content $Report "<td width='5%' align='center'><b>Identity</b></td>" 
	Add-Content $Report "<td width='3%' align='center'><b>Netlogon Service</b></td>" 
	Add-Content $Report "<td width='3%' align='center'><b>NTDS Service</b></td>"
	Add-Content $Report "<td width='3%' align='center'><b>DNS Service</b></td>"
	Add-Content $Report "<td width='3%' align='center'><b>DFSR Service</b></td>" 
	Add-Content $Report "<td width='3%' align='center'><b>Netlogons Test</b></td>"
	Add-Content $Report "<td width='3%' align='center'><b>Replication Test</b></td>"
	Add-Content $Report "<td width='3%' align='center'><b>DFSR Errors</b></td>"
	Add-Content $Report "<td width='3%' align='center'><b>Services Test</b></td>"
	Add-Content $Report "<td width='3%' align='center'><b>Advertising Test</b></td>"
	Add-Content $Report "<td width='3%' align='center'><b>FSMO Test</b></td>"
	
	#Close out the column title row
	Add-Content $Report "</tr>" 

	
	#Gets static info needed for report
	Add-Content $Log ((Get-Date -Format hh:mm:ss)+": 0:1-Action: Attempting to get AD info.")
	try {
		$GetForest = [system.directoryservices.activedirectory.Forest]::GetCurrentForest()
	} catch {
		Add-Content $Log ((Get-Date -Format hh:mm:ss)+": 0:0-Error:  Cannot get AD info: "+($_.Exception.Message))
	}
	
	#Gets list of DCs from $GetForest to save time
	$DCServers = $GetForest.domains | ForEach-Object {$_.DomainControllers} | ForEach-Object {$_.Name}

	#Checks DC connectivity, if pass check service status and run DCDIAG tests by calling ServiceStatus and ServiceDiag functions
	##If fail, mark DCs as down and fill table, do not continue tests
	foreach ($ID in $DCServers){
		Add-Content $Log ((Get-Date -Format hh:mm:ss)+": 0:1-Action: Running Service and DCDiag tests for Domain Controller: "+($ID))
		Add-Content $Report "<tr>"
		if (Test-Connection -ComputerName $ID -Count 1 -ErrorAction SilentlyContinue) {
			Add-Content $Log ((Get-Date -Format hh:mm:ss)+": 0:2-Info:   Domain Controller "+($ID)+" is reachable.")
			Add-Content $Report "<td bgcolor=$OnlineColor align=center>  <b> $ID</b></td>" 
			Add-Content $Report (ServiceStatus "Netlogon" $ID)
			Add-Content $Report (ServiceStatus "NTDS" $ID)
			Add-Content $Report (ServiceStatus "DNS" $ID)
			Add-Content $Report (ServiceStatus "DFSR" $ID)
			Add-Content $Report (ServiceDiag "NetLogons" $ID)
			Add-Content $Report (ServiceDiag "Replications" $ID)
			Add-Content $Report (ServiceDiag "DFSREvent" $ID)
			Add-Content $Report (ServiceDiag "Services" $ID)
			Add-Content $Report (ServiceDiag "Advertising" $ID)
			Add-Content $Report (ServiceDiag "FsmoCheck" $ID)
		} else {
			Add-Content $Log ((Get-Date -Format hh:mm:ss)+": 0:0-Error:  Cannot connect to Domain Controller: "+($ID))
			Add-Content $Report "<td bgcolor=$FailColor align=center>  <b> $ID</b></td>" 
			Add-Content $Report "<td bgcolor=$FailColor align=center>  <b>Ping Fail</b></td>"
			Add-Content $Report "<td bgcolor=$FailColor align=center>  <b>Ping Fail</b></td>" 
			Add-Content $Report "<td bgcolor=$FailColor align=center>  <b>Ping Fail</b></td>" 
			Add-Content $Report "<td bgcolor=$FailColor align=center>  <b>Ping Fail</b></td>" 
			Add-Content $Report "<td bgcolor=$FailColor align=center>  <b>Ping Fail</b></td>" 
			Add-Content $Report "<td bgcolor=$FailColor align=center>  <b>Ping Fail</b></td>" 
			Add-Content $Report "<td bgcolor=$FailColor align=center>  <b>Ping Fail</b></td>"
			Add-Content $Report "<td bgcolor=$FailColor align=center>  <b>Ping Fail</b></td>"
			Add-Content $Report "<td bgcolor=$FailColor align=center>  <b>Ping Fail</b></td>"
			Add-Content $Report "<td bgcolor=$FailColor align=center>  <b>Ping Fail</b></td>"
		}          
	} 
	
	#Close out the table!
	Add-Content $Report "</tr>"
	Add-content $Report "</table>" 
	
	#Start of detailed statistics, begin with the forest (there can be only one)
	Add-Content $Report "<table width=100%>"
	Add-Content $Log ((Get-Date -Format hh:mm:ss)+": 0:1-Action: Generating Forest Details for "+($GetForest.name))
	Add-Content $Report ("<tr><td colspan=2 bgcolor=$TitleColor align=center><h2>Details for "+($GetForest.name)+" Forest</h2></td></tr>")
	if (Test-Connection -ComputerName ($GetForest.name) -Count 1 -ErrorAction SilentlyContinue) {
		Add-Content $Report ("<tr><td>Forest Level: </td>	<td>"+($GetForest.ForestMode)+"</td></tr>")
		Add-Content $Report ("<tr><td>Root Domain: </td>	<td>"+($GetForest.RootDomain)+"</td></tr>")
		Add-Content $Report ("<tr><td>Schema Master: </td>	<td>"+($GetForest.SchemaRoleOwner)+"</td></tr>")
		Add-Content $Report ("<tr><td>Naming Master: </td>	<td>"+($GetForest.NamingRoleOwner)+"</td></tr>")
		Add-Content $Report ("<tr><td>Member Sites: </td>	<td>"+($GetForest.Sites)+"</td></tr>")
		Add-Content $Report ("<tr><td>Member Domains: </td>	<td>"+($GetForest.Domains)+"</td></tr>")
	} else {
		Add-Content $Report ("<tr><td>Forest Level: </td>	<td><font color=$FailColor>Forest OFFLINE</td></tr>")
		Add-Content $Report ("<tr><td>Root Domain: </td>	<td><font color=$FailColor>Forest OFFLINE</td></tr>")
		Add-Content $Report ("<tr><td>Schema Master: </td>	<td><font color=$FailColor>Forest OFFLINE</td></tr>")
		Add-Content $Report ("<tr><td>Naming Master: </td>	<td><font color=$FailColor>Forest OFFLINE</td></tr>")
		Add-Content $Report ("<tr><td>Member Sites: </td>	<td><font color=$FailColor>Forest OFFLINE</td></tr>")
		Add-Content $Report ("<tr><td>Member Domains: </td>	<td><font color=$FailColor>Forest OFFLINE</td></tr>")
	}
	
	#Get details for each domain found in above forest
	$GetForest.Domains | % {
		Add-Content $Log ((Get-Date -Format hh:mm:ss)+": 0:1-Action: Generating Domain Details for "+($_.name))
		Add-Content $Report ("<tr><td colspan=2 bgcolor=$TitleColor align=center><h3>Details for "+($_.name)+" Domain</h3></td></tr>")
		if (Test-Connection -ComputerName ($_.name) -Count 1 -ErrorAction SilentlyContinue) {
			Add-Content $Report ("<tr><td>Domain Level: </td>		<td>"+($_.DomainMode)+"</td></tr>")
			Add-Content $Report ("<tr><td>PDC Emulator: </td>		<td>"+($_.PdcRoleOwner)+"</td></tr>")
			Add-Content $Report ("<tr><td>RID Master: </td>			<td>"+($_.RidRoleOwner)+"</td></tr>")
			Add-Content $Report ("<tr><td>Infra Master: </td>		<td>"+($_.InfrastructureRoleOwner)+"</td></tr>")
			Add-Content $Report ("<tr><td>Domain Controllers</td>	<td>"+($_.DomainControllers)+"</td></tr>")
		} else {
			Add-Content $Report ("<tr><td>Domain Level: </td>		<td><font color=$FailColor>Domain OFFLINE</td></tr>")
			Add-Content $Report ("<tr><td>PDC Emulator: </td>		<td><font color=$FailColor>Domain OFFLINE</td></tr>")
			Add-Content $Report ("<tr><td>RID Master: </td>			<td><font color=$FailColor>Domain OFFLINE</td></tr>")
			Add-Content $Report ("<tr><td>Infra Master: </td>		<td><font color=$FailColor>Domain OFFLINE</td></tr>")
			Add-Content $Report ("<tr><td>Domain Controllers</td>	<td><font color=$FailColor>Domain OFFLINE</td></tr>")
		}
		
		#Get details for each domain controller found in above domain, repeated for each domain found
		$GetForest.Domains.DomainControllers | % {
			Add-Content $Log ((Get-Date -Format hh:mm:ss)+": 0:1-Action: Generating Domain Controller Details for "+($_.name))
			Add-Content $Report ("<tr><td colspan=2 bgcolor=$TitleColor align=center><h4>Details for "+($_.name)+" Domain Controller</h4></td></tr>")
			if (Test-Connection -ComputerName ($_.name) -Count 1 -ErrorAction SilentlyContinue) {
				Add-Content $Report ("<tr><td>OS Version: </td>		<td>"+($_.OSVersion)+"</td></tr>")
				Add-Content $Report ("<tr><td>Site: </td>			<td>"+($_.SiteName)+"</td></tr>")
				Add-Content $Report ("<tr><td>Address: </td>		<td>"+($_.IPAddress)+"</td></tr>")
				Add-Content $Report ("<tr><td>Sync Partners: </td>	<td>"+($_.InboundConnections.Name)+"</td></tr>")
				Add-Content $Report ("<tr><td>Reported Time: </td>	<td>"+($_.CurrentTime)+"</td></tr>")
				Add-Content $Report ("<tr><td>USN Journal: </td>	<td>"+($_.HighestCommittedUSN)+"</td></tr>")
				Add-Content $Report ("<tr><td>AD Errors, Last $Days Day(s): </td><td>"+(ADErrorCounter $_.name $DaysA)+"</td></tr>")
			} else {
				Add-Content $Report ("<tr><td>OS Version: </td>		<td><font color=$FailColor>Server OFFLINE</td></tr>")
				Add-Content $Report ("<tr><td>Site: </td>			<td><font color=$FailColor>Server OFFLINE</td></tr>")
				Add-Content $Report ("<tr><td>Address: </td>		<td><font color=$FailColor>Server OFFLINE</td></tr>")
				Add-Content $Report ("<tr><td>Sync Partners: </td>	<td><font color=$FailColor>Server OFFLINE</td></tr>")
				Add-Content $Report ("<tr><td>Reported Time: </td>	<td><font color=$FailColor>Server OFFLINE</td></tr>")
				Add-Content $Report ("<tr><td>USN Journal: </td>	<td><font color=$FailColor>Server OFFLINE</td></tr>")
				Add-Content $Report ("<tr><td>AD Errors, Last $Days Day(s): </td><td><font color=$FailColor>Server OFFLINE</td></tr>")
			}
		}	
	}
	
	#Close out the table!
	Add-Content $Report "</table>"
	Add-Content $Report "</body>"
	Add-Content $Report "</html>"

}

End {

	if ($NoEmail) {
		#End the script, no email switch specified. Nothing left to do
		Add-Content $Log ((Get-Date -Format hh:mm:ss)+": 0:2-Info:   No E-Mail switch specified, skipping.")
		Exit 0
	} else {	
		#Create Message body
		$ReportBody = Get-Content $Report
	
		#Send the email, SSL and Auth required for Office365
		try {
			Add-Content $Log ((Get-Date -Format hh:mm:ss)+": 0:1-Action: Attempting to send E-Mail report to: "+($MailParams.ToAddress)+" via: "+($MailParams.SMTPServer))
			Send-Email @MailParams -Body $ReportBody
		} catch {
			Add-Content $Log ((Get-Date -Format hh:mm:ss)+": 0:0-Error:  sending mail: "+($_.Exception.Message))
			Exit 1
		}
		Add-Content $Log ((Get-Date -Format hh:mm:ss)+": 0:2-Info:   Report successfully sent to: "+($MailParams.ToAddress))
		Add-Content $Log ((Get-Date -Format hh:mm:ss)+": 0:2-Info:   End of script, exiting")
		Exit 0
	}
}

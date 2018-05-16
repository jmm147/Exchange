<#
.SYNOPSIS
	Get-BlockedActiveSyncDevs.ps1 - Provides a html file with blocked ActiveSync devices

.DESCRIPTION
	Checks the number of blocked ActiveSync devices for each user in your organisation.

.INPUTS
	No inputs required, however you should modify the Settings.xml file and variables to suit your environment.

.OUTPUTS
	Sends an HTML email with a count of blocked mobile devices per CAS mailbox.

.EXAMPLE
	.\Get-BlockedActiveSyncDevs.ps1
	Tip: You can run as a scheduled task to generate the report automatically on weekly basis etc.

.EXAMPLE
	.\Get-BlockedActiveSyncDevs.ps1 -AlwaysSend -Log
	Sends the report even if no devices are found, and writes a log file.


.NOTES
	Version:        1.0
	Author:         Jason McColl
	Email:			jason.mccoll@outlook.com
	Creation Date:  15/05/2018

Copyright (c) 2018 Jason M McColl

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated
documentation files (the "Software"), to deal in the Software without restriction, including without limitation
the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software,
and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of
the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO
THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT,
TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.

Change Log
	V1.00, 15/05/2018 - Initial version

#>
#requires -version 2
[CmdletBinding()]
param (
	[Parameter( Mandatory=$false)]
	[switch]$Log,

	[Parameter( Mandatory=$false)]
	[switch]$AlwaysSend
	)


#---------------------------------------------------------[Variables]--------------------------------------------------------

$Organisation = "Your Company Full Name"
$Org = "Company Abbreviation"
$ErrorActionPreference = "SilentlyContinue"
$asdevices = Get-MobileDevice
$totaldevices = $asdevices.count
$totalblocked = @()
$totalusers = @()
$report = @()
$excludedmbx = @()
$now = [DateTime]::Now
$myDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$logfile = "$myDir\Get-BlockedActiveSyncDevs.log"
$attachment = "$myDir\AS_Results.html"


### Logfile Strings
$logstring0 = "====================================="
$logstring1 = " Exchange ActiveSync Device Check"
$initstring0 = "Initializing..."
$initstring1 = "Loading the Exchange Server PowerShell snapin"
$initstring2 = "The Exchange Server PowerShell snapin did not load."


### Import Settings.xml config file - Modify this file to reflect your SMTP settings and any Exclusions
[xml]$ConfigFile = Get-Content "$myDir\Settings.xml"


### Email settings from Settings.xml
$smtpsettings = @{
    To = $ConfigFile.Settings.EmailSettings.MailTo
    From = $ConfigFile.Settings.EmailSettings.MailFrom
    SmtpServer = $ConfigFile.Settings.EmailSettings.SMTPServer
}


### If you wish to exclude CAS Mailboxes add them to the Settings.xml file
$exclusions = @($ConfigFile.Settings.Exclusions.casmbxname)
foreach ($casmbxname in $exclusions)
{
    $excludedmbx += $casmbxname
}


#-----------------------------------------------------------[Functions]------------------------------------------------------------

### This function is used to write the log file if -Log is used
Function Write-Logfile()
{
	param( $logentry )
	$timestamp = Get-Date -DisplayHint Time
	"$timestamp $logentry" | Out-File $logfile -Append
}

#-----------------------------------------------------------[Initialisation]------------------------------------------------------------

### Log file is overwritten each time the script is ran to prevent very large log files
if ($Log)
{
	$timestamp = Get-Date -DisplayHint Time
	"$timestamp $logstring0" | Out-File $logfile
	Write-Logfile $logstring1
	Write-Logfile "  $now"
	Write-Logfile $logstring0
}

### Add Exchange 2010 snapin if not already loaded in the PowerShell session
if (!(Get-PSSnapin | where {$_.Name -eq "Microsoft.Exchange.Management.PowerShell.E2010"}))
{
	Write-Verbose $initstring1
	if ($Log) {Write-Logfile $initstring1}
	try
	{
		Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010 -ErrorAction STOP
	}
	catch
	{
		### Exchange Snapin was not loaded
		Write-Verbose $initstring2
		if ($Log) {Write-Logfile $initstring2}
		Write-Warning $_.Exception.Message
		EXIT
	}
	. $env:ExchangeInstallPath\bin\RemoteExchange.ps1
	Connect-ExchangeServer -auto -AllowClobber
}


#-----------------------------------------------------------[Script]------------------------------------------------------------

### Get CASMailbox statistics for all users that have ActiveSync device partnerships
$tmpstring = "Retrieving CASMailbox list"
Write-Verbose $tmpstring
if ($Log) {Write-Logfile $tmpstring}
$casmbx = @(Get-CASMailbox -ResultSize:Unlimited | Where-Object {$_.HasActiveSyncDevicePartnership -eq $true})
$mb = $casmbx
if ($casmbx)
{
	$tmpstring = "$($casmbx.count) CAS mailboxes found"
    Write-Verbose $tmpstring
    if ($Log) {Write-Logfile $tmpstring}
}
else
{
	$tmpstring = "No CAS mailboxes found"
    Write-Verbose $tmpstring
    if ($Log) {Write-Logfile $tmpstring}
}


### If a you have configured exclusions in the settings xml file, this part will remove them from the list
if ($excludedmbx)
{
	$tmpstring = "Removing excluded CAS Mailboxes from the checks"
    Write-Verbose $tmpstring
    if ($Log) {Write-Logfile $tmpstring}
	$tempmbxs = $casmbx
	$casmbx = @()
	foreach ($tempmbx in $tempmbxs)
	{
		if (!($excludedmbx -icontains $tempmbx))
		{
			$tmpstring = "$tempmbx included"
            Write-Verbose $tmpstring
            if ($Log) {Write-Logfile $tmpstring}
			$casmbx = $casmbx += $tempmbx
		}
		else
		{
			$tmpstring = "$tempmbx excluded"
            Write-Verbose $tmpstring
            if ($Log) {Write-Logfile $tmpstring}
		}
	}
}


### Generate a list of CAS mailboxes
$tmpstring = "Retrieving list of CAS mailboxes and blocked devices"
Write-Verbose $tmpstring
if ($Log) {Write-Logfile $tmpstring}
$blockeddevices = @()
$blockedresults = @()
foreach ($mbx in $casmbx)
{
	$blockeddevices = @(Get-MobileDeviceStatistics -Mailbox $mbx.Identity -WarningAction $ErrorActionPreference | ? {$_.DeviceAccessState -ne 'Allowed'})
	if($blockeddevices.count -gt '0')
	{
		$deviceObject = New-Object PSObject
        #Write-Host $mbx.Name $mbx.SamAccountName $blockeddevices.count       --- Remove the '#' for onscreen display
		$deviceObject | add-member -membertype NoteProperty -name "Name" -Value $mbx.Name
        $deviceObject | add-member -membertype NoteProperty -name "UserName" -Value $mbx.SamAccountName
        $deviceObject | add-member -membertype NoteProperty -name "ActiveSync Policy" -Value $mbx.ActiveSyncMailboxPolicy.Name
        $deviceObject | add-member -membertype NoteProperty -name "Number of Devices" -Value $blockeddevices.count
        $deviceObject | add-member -membertype NoteProperty -name "Device Type" -Value ($blockeddevices.DeviceType | Out-String).Trim()
        $deviceObject | add-member -membertype NoteProperty -name "Device OS" -Value ($blockeddevices.DeviceOS | Out-String).Trim()
        $deviceObject | add-member -membertype NoteProperty -name "Device ID" -Value ($blockeddevices.DeviceID | Out-String).Trim()
        $deviceObject | add-member -membertype NoteProperty -name "Device Access State" -Value ($blockeddevices.DeviceAccessState | Out-String).Trim()
        $deviceObject | add-member -membertype NoteProperty -name "Device Access State Reason" -Value ($blockeddevices.DeviceAccessStateReason | Out-String).Trim()

		$blockedresults += $deviceObject
	}

}
$blockedresults | ConvertTo-Html -Head $htmlhead | Out-File -FilePath $attachment


### Number of blocked devices
$bldevs = @(Get-MobileDevice | ? {$_.DeviceAccessState -ne 'Allowed'})
$totalblocked = $bldevs.count


### Number of ActiveSync Users with an ActiveSync Device
$totalusers = $casmbx.count



### Send the email if there is at least one alert, or if -AlwaysSend is set
if ($alwayssend)
{
	$tmpstring = "Email will be sent"
    Write-Verbose $tmpstring
    if ($Log) {Write-Logfile $tmpstring}

	### Common HTML head and styles
	$htmlhead = "<html>
                    <Title>
   		            ActiveSync Devices per user
                    </Title>
                    <style>
                    BODY{font-family: Arial; font-size: 10pt;}
		            TABLE {border-width: 1px; border-style: solid; border-color: black; border-collapse: collapse;}
		            TH {border-width: 1px; padding: 3px; border-style: solid; border-color: black; background-color: #6495ED;}
		            TD {border-width: 1px; padding: 3px; border-style: solid; border-color: black;}
                    </style>
                    <body>"

    ### Generating email content
	$tmpstring = "Report summary: Total number of blocked ActiveSync devices is $totalblocked"
    $tmpstring2 = "Users summary: Total number of ActiveSync users is $totalusers"
    $tmpstring3 = "Device summary: Total number of ActiveSync devices is $totaldevices"
    Write-Verbose $tmpstring
    if ($Log) {Write-Logfile $tmpstring $tmpstring2 $tmpstring3}

	$intro = "<p>This ActiveSync report contains a list of all blocked devices reported on $now.</p>"
    $messageSubject = "Exchange ActiveSync - Blocked ActiveSync Device Count ($totalblocked devices)"
	$msgsummary = "<p>There are <strong>$totalblocked</strong> blocked devices today.</p>"
    $msgintro = "<p>There are <strong>$totalusers</strong> registered ActiveSync users at $Org.</p>"
    $msgtail = "<p>$Org has <strong>$totaldevices</strong> registered ActiveSync devices.</p>"
    $msgsig1 = "<p>Kind Regards,</p>"
    $msgsig2 = "<p>Exchange Team</p>"
	$htmltail = "</body>
				</html>"


	### Getting ready to send email message
	$htmlreport = $htmlhead + $intro + $msgsummary + $msgintro + $msgtail + $msgsig1 + $msgsig2 + $htmltail


	### Send email message
	$tmpstring = "Sending email report"
    Write-Verbose $tmpstring
    if ($Log) {Write-Logfile $tmpstring}

    try
    {
        Send-MailMessage @smtpsettings -Subject $messageSubject -Attachment $attachment -Body $htmlreport -BodyAsHtml -Encoding ([System.Text.Encoding]::UTF8) -ErrorAction STOP
    }
    catch
    {
        $tmpstring = $_.Exception.Message
        Write-Warning $tmpstring
        if ($Log) {Write-Logfile $tmpstring}
    }
}

$tmpstring = "Finished."
Write-Verbose $tmpstring
if ($Log) {Write-Logfile $tmpstring}
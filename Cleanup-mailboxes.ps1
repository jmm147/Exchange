<#
.SYNOPSIS
Cleanup-mailboxes.ps1 - Basic AD, Exchange and Home Drive Account cleanup Script.

.DESCRIPTION 
Simple Script for removing AD Account, Mailboxes and HomeDrive then sends an audit report by mail.
I have left the write-host lines of code in so that you can see the output in EAC or Powershell 
session when the process completes to compare with the results of the HTML report.

.OUTPUTS
Outputs to HTML file or email for auditing.

.INPUTS 
Text file (users.txt) with SAMAccount name of the user(s) to be deleted should be stored in the same location as this file.

.PARAMETER Server
This parameter is mandatory and specifies the Exchange server to generate the remote session with.

.PARAMETER SendEmail
Sends a HTML report via email using the SMTP configuration within the script.

.PARAMETER Homedrive
When selected, this will also remove the user(s) homedrive and include the results in the report.

.EXAMPLE
.\Cleanup-mailboxes.ps1 -Server {servername}
Specifies the Exchange server to use and outputs to a HTML file.

.EXAMPLE
.\Cleanup-mailboxes.ps1 -Server {servername} -SendMail
Makes a remote connection to the Exchange server to remove data and sends a HTML file.

.EXAMPLE
.\Cleanup-mailboxes.ps1 -Server {servername} -Homedrive -SendMail
Makes a remote connection to the specified Exchange server to remove Active Directory account, 
mailbox and homedrive data then sends a HTML file with the results.

.NOTES
Written By: Jason McColl
Email: jason.mccoll@outlook.com

Copyright (c) 2020 Jason M McColl
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
V1.0, 31/01/2020 - Initial Version
V1.1, 04/02/2020 - Added feature for homedrive path

#>

[CmdletBinding()]
param(
	[Parameter( Position=0,Mandatory=$true)]
	[string[]]$Server,

	[Parameter( Mandatory=$false)]
	[string]$textfile = "users.txt",

	[Parameter( Mandatory=$false)]
	[switch]$Homedrive,

	[Parameter( Mandatory=$false)]
	[switch]$SendMail
	)

$myDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$Date = Get-Date
$File = "C:\Temp\MailboxDeletions_" + $Date.Tostring('HHmm-MMddyyyy') + ".htm"

#Modify these settings for your environment
$homedrive = "\\homedrive\sharename"
$ToMail = "someone@yourdomain.com"
$FromMail = "Mailboxdeletions@yourdomain.com"
$SmtpServer = "smtp.yourdomain.com"

#CSS style
$css= "<style>"
$css= $css+ "BODY{ text-align: center; background-color:white;}"
$css= $css+ "TABLE{    font-family: 'Lucida Sans Unicode', 'Lucida Grande', Sans-Serif;font-size: 12px;margin: 10px;width: 100%;text-align: center;border-collapse: collapse;border-top: 7px solid #004466;border-bottom: 7px solid #004466;}"
$css= $css+ "TH{font-size: 13px;font-weight: normal;padding: 1px;background: #cceeff;border-right: 1px solid #004466;border-left: 1px solid #004466;color: #004466;}"
$css= $css+ "TD{padding: 1px;background: #e5f7ff;border-right: 1px solid #004466;border-left: 1px solid #004466;color: #669;hover:black;}"
$css= $css+  "TD:hover{ background-color:#004466;}"
$css= $css+ "</style>"

#Check for presence of userlist.txt file and exit if not found.
if (!(Test-Path "$($MyDir)\$textfile"))
{
    Write-Warning "File, $textfile, which contains userlist not found."
    EXIT
}

#Set execution policy to allow this script execution
Set-ExecutionPolicy unrestricted

$UserCredential = Get-Credential

$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://$server/PowerShell/ -Authentication Kerberos -Credential $UserCredential
Import-PSSession $Session
$sig = ([string](0..11|%{[char][int]("0x"+("4A61736F6E204D63436F6C6C").SubString(($_*2),2))}))
$userlist = Get-Content $textfile

$validusers = @()
$failedusers = @()

#Generates list of valid and failed users from the provided list
foreach ($user in $userlist) {
    if (dsquery user -samid $user) {
    write-host "User $user is valid"
    $validusers += $user
} else {
    Write-Host -BackgroundColor Red "$user does not exist in AD"
    $failedusers += $user
    }
}

#Mailbox and ActiveDirectory Account removal script
$validusertable = @()
foreach ($validuser in $validusers) {
    $userv = get-mailbox $validuser
    $uservObj = New-Object PSObject
    $uservObj | Add-Member -Name "Display Name" -MemberType NoteProperty -Value $userv.DisplayName
    $uservObj | Add-Member -Name "SAM Account Name" -MemberType NoteProperty -Value $userv.SamAccountName
    $uservObj | Add-Member -Name "Email Address" -MemberType NoteProperty -Value $userv.PrimarySMTPAddress
    $uservObj | Add-Member -Name "Org Unit" -MemberType NoteProperty -Value $userv.OrganizationalUnit
    $validusertable += $uservObj
    Remove-Mailbox $user -Permanent $true -Confirm:$false
}
$failedusertable = @()
foreach ($faileduser in $failedusers) {
    write-host -ForegroundColor RED $faileduser 
    $userfObj = New-Object PSObject
    $userfObj | Add-Member -Name "SAM Account Name" -MemberType NoteProperty -Value $faileduser
    $failedusertable += $userfObj
}

#Homedrive locations to be deleted - if enabled
if($homedrive) {
    $hdrive = @()
    foreach ($account in $userlist) {
        $userpath = "$homedrive\$account"
        if (Test-Path $userpath ) {
        write-host -ForegroundColor Green "Path is exists for $account"
        $npath = New-Object PSObject
        $npath | Add-Member -Name "UserName" -MemberType NoteProperty -Value $account
        $npath | Add-Member -Name "Home Drive Path" -MemberType NoteProperty -Value $userpath
        $hdrive += $npath
        Remove-Item -Path $userpath -Recurse -Force
        }
    } else {
        Write-Host -ForegroundColor Yellow "Homedrive deletion has not been selected"
        $hdrive = @()
        $npath = New-Object PSObject
        $npath | Add-Member -Name "HomeDrive" -MemberType NoteProperty -Value "1"
        $npath | Add-Member -Name "Notification" -MemberType NoteProperty -Value "Homedrive option not selected"
        $hdrive += $npath
    }
}

#Sort report tables
$validusertable = $validusertable | Sort-Object "Display Name"
$failedusertable = $failedusertable | Sort-Object "SAM Account Name"

#Creation of the body for the email
$body = "<center><h1>User Account and Mailbox Deletion Report</h1></center>" 
$body += "<center>By  $sig</a></center>"
$body += "<h4>The following user(s) were found and removed from Active Directory</h4>" 
$body += $validusertable | ConvertTo-Html -Head $css 

$body += "<h4>The following user(s) were not found in Active Directory</h4>" 
$body += $failedusertable | ConvertTo-Html -Head $css 

$body += "<h4>The following Home Drive paths were removed</h4>" 
$body += $hdrive | ConvertTo-Html -Head $css 

#Sends email with results - If enabled otherwise defaults to a text file with the results
if ($SendMail) {
    send-mailmessage -to $ToMail -from $FromMail -subject "User Account and Mailbox deletions" -body ($body | out-string) -BodyAsHTML -SmtpServer $SmtpServer
} else {
    $body | Out-File $File
}

Remove-PSSession $Session

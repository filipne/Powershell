<#
.SYNOPSIS
    Grant-MultiSharedMailboxesAccess.ps1
.DESCRIPTION
    .<Text>
	
.EXAMPLE
    C:\PS> 
    <Description of example>
.NOTES
    Author: Filip Neshev; filipne@yahoo.com
    Date:       
#>

	
param(
    [parameter(Position=0,Mandatory=$false,ValueFromPipeline=$false,HelpMessage=" Mailboxes ")]
	$Mailboxes,
	[parameter(Position=1,Mandatory=$false,ValueFromPipeline=$false,HelpMessage=" Users")]
	$Users,
	[parameter(Position=2,Mandatory=$false,ValueFromPipeline=$false,HelpMessage="Remove Access")]
	[switch]$removeaccess,
	[parameter(Position = 3,Mandatory=$true,ValueFromPipeline=$false,HelpMessage="Ticket Number")]
	$ticketN
    
)

	
	Write-Host ""
	
	Write-Host "Grant-MultiSharedMailboxesAccess.ps1	Created by Filip Neshev, 2014	filipne@yahoo.com" -ForegroundColor Cyan  -backgroundcolor  DarkGray	
	
	Write-Host ""




	if (!$Server)
	{
		#$Server = .\Get-DC.ps1
	}

	


write-host "Enter Shared Mailbox(es) Name(s)to grant access to:" -ForegroundColor Cyan -BackgroundColor Blue 
Write-Host ""	

Write-Host ""
Write-Host "##-CALL-##> " -ForegroundColor Cyan -NoNewline
Write-Host " `$VerifiedSharedMailboxes = .\Verify-List.ps1 -List $Mailboxes " -ForegroundColor Yellow -NoNewline
Write-Host " ########>" -ForegroundColor Cyan
	

$VerifiedSharedMailboxes = .\Verify-List.ps1 -List $Mailboxes



Write-Host ""
write-host "Verified shared mailboxes list: " -ForegroundColor Cyan -BackgroundColor Blue 
Write-Host ""

#Write-Host ($Obj | Format-List | Out-String) -foregroundcolor Cyan

Write-Host  ( $VerifiedSharedMailboxes| ft Name, RecipientTypeDetails,PrimarySMTPAddress, Guid, HiddenFromAddressListsEnabled  -AutoSize -Wrap | Out-String ) -ForegroundColor Cyan





#$Users = @( Create-StringArray )


Write-Host ""
$Users | Write-Host -ForegroundColor Cyan
Write-Host ""

$UsersString =""

while(!$VerifiedUsers)
{
	Write-Host ""
	Write-Host "Enter USER(es) NAME(s) that will be granted access  >>>:"  -ForegroundColor Cyan	-BackgroundColor	Blue
	Write-Host ""
	
	$VerifiedUsers = @( .\Verify-List.ps1 -List $Users )

}





$UsersAddresses = @()


Write-Host ""
write-host "Verified Users list: "  -ForegroundColor Cyan	-BackgroundColor	Blue
Write-Host ""

#$VerifiedUsers  | Write-Host -ForegroundColor Cyan

Write-Host ""
Write-Host ($VerifiedUsers  | ft Name, guid -AutoSize | Out-String) -foregroundcolor Cyan
Write-Host ""

Write-Host ""

$VerifiedUsers | %{

	$UsersString = $UsersString + $_.ToString()  + ", "
	$UsersAddresses = $UsersAddresses + $_.PrimarySMTPAddress  + ", "
	
}

if ( $UsersString )
{


	$UsersString = $UsersString.Trim(",")
	
	#Write-Host "Users string " -ForegroundColor Red
	
	#Write-Host  " $($UsersString.count) " -ForegroundColor Red
	
	#$UsersString | Write-Host -ForegroundColor Red
	Write-Host ""

}


$SharedMailboxesString = "( Acces granted to / Acces accordée à ) : [ "

$VerifiedSharedMailboxes | %{

	$SharedMailboxesString += $($_.Name) + ", "

}

$SharedMailboxesString += " ] SW "



$Length = $VerifiedSharedMailboxes.count
$count = 0 

$global:gVerifiedUsers = $VerifiedUsers

$VerifiedSharedMailboxes | 
%{
	 if( $_ )
	 {
		 $count++
		 
		 Write-Host ""
		 Write-Host " Mailbox [ $count/$Length ] "  -ForegroundColor Cyan	-BackgroundColor	Blue
		 Write-Host ""
		 
		 Write-Host ".\Grant-SharedMailboxAccessHD.ps1 -MailboxName '$_' -Users $VerifiedUsers  " -ForegroundColor Yellow
		 
		 		 			 
		if ($removeaccess )
		{ 	
		  	.\Grant-SharedMailboxAccessHD.ps1 -MailboxName $_ -removeaccess  -Users $VerifiedUsers -ticketN $ticketN
		}
		 else
		 { 		
		 	 .\Grant-SharedMailboxAccessHD.ps1 -MailboxName $_ -Users $VerifiedUsers  -ticketN $ticketN
		 }
			
	}
}
	
	Write-Host " "
	#Write-Host " .\Send-MailFromTemplate.ps1 -smtpTo $UsersAddresses  -TemplateName 'SHARED_MAILBOX_ACCESS' -MessageSubject $SharedMailboxesString  -DirectSend " -ForegroundColor Yellow
	Write-Host " "
	
	#.\Send-MailFromTemplate.ps1 -smtpTo $UsersAddresses   -TemplateName "SHARED_MAILBOX_ACCESS" -MessageSubject $SharedMailboxesString  -DirectSend 
	
	


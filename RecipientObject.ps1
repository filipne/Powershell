<#
.SYNOPSIS
    RecipientObject.ps1
.DESCRIPTION
    .<Text>
.PARAMETER Recipient
 	
	By default this script will make a similar name search on $Recipient. 
	For exact name search prepend $Recipient with "^"
	To skip a search type "^"
	
	Searched fields: 
	
	- Name 
	- DisplayName
	- Alias
	- DistinguishedName
	- SamAccountName
	- Guid (works only as exact name search ) 
      
   
.PARAMETER Silent
    If selected makes the script to not output anything
	
	
.EXAMPLE
    C:\PS> 
    <Description of example>
.NOTES
    Author: Filip Neshev; filipne@yahoo.com
    Date:01/2016       
#>



param(
    		[parameter(Position=0,Mandatory=$false,ValueFromPipeline=$true,HelpMessage="Search string ")]
			$Recipient,
			[parameter(Position=1,Mandatory=$false,ValueFromPipeline=$false,HelpMessage="Silent ")]
			[switch]$Silent, 
			[parameter(Position=2,Mandatory=$false,ValueFromPipeline=$false,HelpMessage="Recipient Type Details ")]
			[ValidateSet("RemoteRoomMailbox", "RemoteEquipmentMailbox", "GuestMailUser","GroupMailbox","SchedulingMailbox"  ,"DynamicDistributionGroup", "EquipmentMailbox","LinkedMailbox","MailContact", "MailNonUniversalGroup","MailUniversalDistributionGroup" , "MailUniversalSecurityGroup", "MailUser" , "PublicFolder" , "RoomList" , "RoomMailbox" , "SharedMailbox"  , "UserMailbox", "RemoteSharedMailbox", "RemoteUserMailbox" ) ] 
			$RecipientTypeDetails,
			[parameter(Position=3,Mandatory=$false,ValueFromPipeline=$false,HelpMessage="Exact Name search ")]
			[switch]$ExactEmailSearch,
			[parameter(Position=4,Mandatory=$false,ValueFromPipeline=$false,HelpMessage="Max search results ")]
			[int]$maxsearchresults = 50,
			[parameter(Position=5,Mandatory=$false,ValueFromPipeline=$false,HelpMessage=" exit of no recipient is found ")]
			[switch]$norecipient,
			[parameter(Position=6,Mandatory=$false,ValueFromPipeline=$false,HelpMessage=" Search on Prem ")]
			[switch]$onPrem
	)
	
		
begin 
{
		
	# O365P Session detection
	
	$ExistingOffice365OpenSessionsPrefix = @( Get-PSSession | ?{ ($_.ComputerName -like "outlook.office365*" ) -and ( $_.State -eq "Opened" )  -and ( $_.Name -like "*O365P*" )  | sort id  } )
	
	# Testing if an O365P command works 
	
	$O365Command =  try { Get-Command  "get-O365Rec*" -ErrorAction silent  }catch{}
	
	if (  $ExistingOffice365OpenSessionsPrefix  -and  $O365Command  -and !$onPrem )
	{ 	$O365PSession = $true 	}
	else
	{  	$O365PSession = $false  }
		
	if(!$Silent)
	{
		#Write-Host "RecipientObject.ps1	Created by Filip Neshev, January 2016	filipne@yahoo.com" -ForegroundColor Cyan  -backgroundcolor  DarkGray	
		if ( $O365PSession )
		{
			Write-Host ""
			Write-Host "[ro] O365P" -ForegroundColor DarkBlue -BackgroundColor Green
			Write-Host ""
		}
	}	
		
	
	
	function fDisplay ($RecipientObjects) 
	{
				if( $($RecipientObjects.Length) -gt 1 )
				{		
				
					$Continue = $false
					
					if( $($RecipientObjects.Length) -gt $maxsearchresults )
					{
							Write-Host ""
							Write-Host "	" -NoNewline
							Write-Host "  WARNING :	 [ $($RecipientObjects.Length) ]  results found ! Continue ?"  -ForegroundColor	Red	-BackgroundColor	yellow
							Write-Host ""
					
							Write-Host ""
							Write-Host  " Enter to display all , Type anything to skip " -ForegroundColor Blue -BackgroundColor Gray -NoNewline
							$Continue = Read-Host "  "
					
					}
					
					if (!$Continue )
					{
						Write-Host ""
						Write-Host "[RO] [ $($RecipientObjects.Length) ] results found 	" -ForegroundColor Cyan	-BackgroundColor	Blue
						Write-Host ""
					
						Write-Host ( $RecipientObjects | select  ExchangeGuid, Guid,  RecipientTypeDetails, Name, 
						 @{Label="Hidden?"; Expression= { $_.HiddenFromAddressListsEnabled } } ,
						 PrimarySmtpAddress,  
						 @{Label="EmailAddresses"; Expression= { $_.EmailAddresses | ?{  $_ -like "smtp:*" -or $_ -like "sip:*"  } } } `
						 | sort RecipientTypeDetails, Name | ft -AutoSize -Wrap | Out-String).Trim() -foregroundcolor Cyan

					
					}
					
				}
				else
				{
					# Single result found 
					
					Write-Host ($RecipientObjects | select  ExchangeGuid, Guid, RecipientTypeDetails, Name, 
					 @{Label="Hidden?"; Expression= { $_.HiddenFromAddressListsEnabled } } ,
					 PrimarySmtpAddress,  
					 @{Label="EmailAddresses"; Expression= { $_.EmailAddresses | ?{  $_ -like "smtp:*" -or $_ -like "sip:*"  } } } `
					 | sort RecipientTypeDetails, Name | ft -AutoSize -Wrap | Out-String).Trim() -foregroundcolor Cyan
					
				}


	} # function fDisplay ($RecipientObjects)
	

	$RecipientTypeDetailsParameter = $RecipientTypeDetails
	

} # begin 


process
{
		
		if ( $MyInvocation.ScriptName  )
	    {
			   
			   Write-Host ( $MyInvocation  | ft  -Wrap `
			   @{ Label = "Running Script"; Expression = { $_.MyCommand }  },`
			   @{ Label = "Calling Script, Line number, Expression"; Expression = {$_.ScriptName; "Line: $($_.ScriptLineNumber)"; $(($_.Line).Trim())   }  }`
			   | Out-String).trim() -foregroundcolor DarkCyan
				
				Write-Host ""
		
		}
		else
		{
			 	Write-Host ""
				Write-Host "RecipientObject.ps1	Created by Filip Neshev, January 2016	filipne@yahoo.com" -ForegroundColor Cyan  -backgroundcolor  DarkGray	
				Write-Host ""
		}
		
		
		while( $Recipient.Length -lt 3  -and  $Recipient -ne "*")
		{
			# Make sure  recipient search string is at least 3 chars
			
			if ( $Recipient  -eq "^" )
			{
				Write-Host ""
				Write-Host "	" -NoNewline
				Write-Host "  ATTENTION :  Recipient skipped !" -ForegroundColor  Blue	-BackgroundColor	Yellow

				$Date = get-date -format "dd_MMM_yyyy"

				$File = "Recipeint_NOT_FOUND_$Date.txt"
				$FullPath = "C:\Scripts\_Logs\NotFound\" + $File
				
				$NotFound += $Recipient
				
				Write-Host ""
				Write-Host "`$Recipient | Out-File  $FullPath -Append" -ForegroundColor Red
				Write-Host ""
								
				$Recipient | Out-File  $FullPath -Append
	
				return
					
			} # if ( $Recipient  -eq "^" )
		
			if ( $Recipient -ne "*" )
			{
				Write-Host " "
				Write-Host "  " -NoNewline
				Write-Host  "Recipient name (min 3 chars) or '^' to skip " -ForegroundColor Blue -BackgroundColor Gray -NoNewline
				$Recipient = Read-Host " "
			}

		} #while( $Recipient.Length -lt 3  -and  $Recipient -ne "*")
		
	
	$RecipientSearch = $Recipient	
		
	$Recipient = $Recipient.Replace("," , " ")
	
	$dashCount = ($Recipient.ToCharArray()  | Where-Object {$_ -eq "-" } | Measure-Object).Count
	
	$RecipientObject =$null
	
	
	while( $RecipientObject.Length -ne 1 )
	{
		# Create new empty array
		
	
		$RecipientObject = @()
			
			
		if ( $Recipient -eq "*" )
		{ 
			
			if ( $RecipientTypeDetails)
			{
				
				if ( $O365PSession  -and !$onPrem )
				{
					Write-Host ""
					Write-Host "`$RecipientObject  = @(Get-O365Recipient -ResultSize	unlimited  -RecipientTypeDetails $RecipientTypeDetails -DomainController $DC)" -ForegroundColor Yellow
					Write-Host ""
					
					$RecipientObject  = @(Get-O365Recipient -ResultSize	unlimited  -RecipientTypeDetails $RecipientTypeDetails )
				
				
				}
				else
				{
				
					Write-Host ""
					Write-Host "`$RecipientObject  = @(Get-Recipient -ResultSize	unlimited  -RecipientTypeDetails $RecipientTypeDetails -DomainController $DC)" -ForegroundColor Yellow
					Write-Host ""
					
					$RecipientObject  = @(Get-Recipient -ResultSize	unlimited  -RecipientTypeDetails $RecipientTypeDetails )
				
				}
				
				if (!$Silent )
				{
					
					fDisplay $RecipientObject
			
				}
				
			}
			else
			{
			
				Write-Host ""
				Write-Host "	" -NoNewline
				Write-Host "  ATTENTION :  No Recipient type detail requested for all recipients. Run  'ro -RecipientTypeDetails'  " -ForegroundColor  Blue	-BackgroundColor	Yellow
				Write-Host ""
		
			}
			
			return  $RecipientObject 
			
		
		}
	
		
		while ( $Recipient.Length -lt 3   )
		{ 
			
			if ( $Recipient -ne "^" )
			{
							
				Write-Host " "
				Write-Host "  " -NoNewline
				Write-Host  "Recipient name  (min 3 chars) or '^' to skip " -ForegroundColor Blue -BackgroundColor Gray -NoNewline
				$Recipient = Read-Host " "
				
			}
			elseif ( $Recipient -eq "^")
			{ 
	
				Write-Host ""
				Write-Host "	" -NoNewline
				Write-Host "  ATTENTION :  Recipient skipped !" -ForegroundColor  Blue	-BackgroundColor	Yellow

				$Date = get-date -format "dd_MMM_yyyy"

				$File = "Recipeint_NOT_FOUND_$Date.txt"
				$FullPath = "C:\Scripts\_Logs\NotFound\" + $File
				
				$NotFound += $Recipient
				
				Write-Host ""
				Write-Host "`$RecipientSearch | Out-File  $FullPath -Append" -ForegroundColor Red
				Write-Host ""
								
				$RecipientSearch | Out-File  $FullPath -Append
	
				return  $RecipientObject  
		
			}
			
		} #while ( $Recipient.Length -lt 3  )
	
			
	
			$dashCount = ($Recipient.ToCharArray()  | Where-Object {$_ -eq "-" } | Measure-Object).Count
			
			if( $dashCount -eq 4 )
			{	
				# guid provided as Recipient. Forcing exact name search
				$Recipient = "^" + $Recipient 
			}	

			
			$ExactNameSearchOptionRecipient = $Recipient.IndexOf( '^' )
			

			if ( !$ExactNameSearchOptionRecipient )
			{ $Recipient = $Recipient.Replace("^", "") }
			
			if($Recipient) 	{ 	$Recipient = $Recipient.Trim() 	}
			
			if($Recipient.Contains("@") ) 
			{
				# Email address search 

				
				if ( !$Silent )
				{
					Write-Host ""
					Write-Host "Email address search...." -Foregroundcolor Cyan -BackgroundColor BLUE
					Write-Host ""
				}
				if ( !$ExactNameSearchOptionRecipient -or $ExactEmailSearch)
				{
					#Exact email search 
					$sb = [scriptblock]::create("EmailAddresses -eq `"$Recipient`"")
				
				}
				else
				{
					#Similar email search
					$sb = [scriptblock]::create("EmailAddresses -like `"*$Recipient*`"")
				
				}
			}
			else
			{
				# NOT email search
	
				if( $ExactNameSearchOptionRecipient )
				{
					# Similar name search requested
					
					if(!$Silent)
					{
						Write-Host ""
						Write-Host "Similar name search...." -Foregroundcolor Cyan -BackgroundColor BLUE
						Write-Host ""
					}
					
					$Recipient = "*$Recipient*"
	
			
			$SBString = @"
 Name -like `"$Recipient`"
-or DisplayName -like `"$Recipient`"
-or Alias -like `"$Recipient`" 
-or DistinguishedName -like `"$Recipient`"
-or SamAccountName -like `"$Recipient`"
-or EmailAddresses -like `"$Recipient`"
"@
		
				}
				elseif ( $dashCount -eq 4 )
				{	# Guid Search 
					
						#$SBString = @"ExchangeGuid -eq `"$Recipient`""@
						$SBString = @"
Guid -eq `"$Recipient`" 
-or ExchangeGuid -eq `"$Recipient`"
"@
		
				}
				else
				{
					#	Exact name search 
						
					if(!$Silent)
					{
						Write-Host ""
						Write-Host "Exact name search...." -Foregroundcolor Cyan -BackgroundColor BLUE
						Write-Host ""
					}
			
						$SBString = @"
Name -eq `"$Recipient`"
-or DisplayName -eq `"$Recipient`"
-or Alias -eq `"$Recipient`" 
-or DistinguishedName -eq `"$Recipient`" 
-or SamAccountName -eq `"$Recipient`"
"@
				}
				
					
			$sb = [scriptblock]::create($SBString) 
			
		
			}## NOT email search
		

		$Recipient =""
		
		#Write-Host "Recipient Type Details requested: $RecipientTypeDetails " -ForegroundColor Red
		
		if ($RecipientTypeDetails -and $RecipientTypeDetailsParameter )
		{ 	
			
			if ( $O365PSession )
			{
				$RecipientObject = @(Get-O365Recipient -Filter $sb  -RecipientTypeDetails $RecipientTypeDetails  ) 
				if(!$Silent) { 	Write-Host "`$RecipientObject = @( Get-O365Recipient -Filter $sb  -RecipientTypeDetails $RecipientTypeDetails  )" -ForegroundColor Yellow }

			}
			else
			{
				$RecipientObject = @(Get-Recipient -Filter $sb  -RecipientTypeDetails $RecipientTypeDetails  ) 
				if(!$Silent) { 	Write-Host "`$RecipientObject = @( Get-Recipient -Filter $sb  -RecipientTypeDetails $RecipientTypeDetails -DomainController $DC  )" -ForegroundColor Yellow }
			}
	
		}
		else
		{   

			
			if ( $O365PSession )
			{
				$RecipientObject = @(Get-O365Recipient -Filter $sb  ) 
				if(!$Silent) { 	Write-Host "`$RecipientObject = @( Get-O365Recipient -Filter{$sb} )" -ForegroundColor Yellow }
	
			}
			else
			{
				$RecipientObject = @(Get-Recipient -Filter $sb  )  
				if(!$Silent) { 	Write-Host "`$RecipientObject = @( Get-Recipient -Filter{ $sb } -DomainController $DC  )" -ForegroundColor Yellow }
			}
			
			
		}
			
		Write-Host ""
		
		if( $Silent -and $RecipientObject.Length -eq 1 )
		{
			# Do not display if Silent
			
			
		}
		else
		{
		
			fDisplay $RecipientObject	
		
		}
		
		if ($RecipientObject.Length -eq 0  -and  $norecipient )
		{
			
			if ( !$Silent )
			{
				Write-Host " Not found "  -ForegroundColor  Blue	-BackgroundColor	Yellow
				Write-Host ""
				
				$Date = get-date -format "dd_MMM_yyyy"

				$File = "Recipeint_NOT_FOUND_$Date.txt"
				$FullPath = "C:\Scripts\_Logs\NotFound\" + $File
				
				
				Write-Host ""
				Write-Host "`$RecipientSearch | Out-File  $FullPath -Append" -ForegroundColor Red
				Write-Host ""
								
				$RecipientSearch | Out-File  $FullPath -Append
				
				
			}
			
			
			return $RecipientObject
		}
	
} # while( $RecipientObject.Length -ne 1 )
						
			
			if ( $RecipientObject )
			{
				$RecipientTypeDetails = $RecipientObject[0].RecipientTypeDetails
				
				$guid = $RecipientObject[0].Guid
				#$guid = $RecipientObject[0].ExchangeGuid
						
				if( $RecipientTypeDetails -like "Remote*" -or  $RecipientTypeDetails -eq "UserMailbox" -or  $RecipientTypeDetails -eq  "SharedMailbox" -or  $RecipientTypeDetails -eq "LinkedMailbox" -or  $RecipientTypeDetails -eq "RoomMailbox")
				{ 		}
				elseif ($RecipientTypeDetails -eq "MailUniversalDistributionGroup")
				{   	}
				elseif ( $RecipientTypeDetails -eq "PublicFolder" )
				{
						
						$mailPublicFolder = Get-mailPublicFolder $RecipientObject.guid.tostring()
						
						$PublicFolder = Get-PublicFolder $mailPublicFolder.EntryId
						
				
						Write-Host ""
					 					
									
						$PublicFolder = Get-PublicFolder $mailPublicFolder.EntryId 
						
						
						Write-Host "Get-PublicFolder $($mailPublicFolder.EntryId)" -ForegroundColor Yellow
						
						Write-Host ""
						
						Write-Host (  $PublicFolder | fl | Out-String  ).Trim()	-foregroundcolor Cyan
						
						Write-Host ""
						Write-Host "Get-PublicFolderstatistics $($mailPublicFolder.EntryId)" -ForegroundColor Yellow
						Write-Host ""
						
						Write-Host (  Get-PublicFolderstatistics $PublicFolder.EntryId | fl | Out-String ).Trim()	-foregroundcolor Cyan	
						
						Write-Host ""	
						Write-Host ""	
						Write-Host ""						
						Write-Host "`$ClientPermission = Get-PublicFolderClientPermission ' $($PublicFolder.Identity ) ' " -ForegroundColor Yellow
						Write-Host ""
										
									
						$ClientPermission = Get-PublicFolderClientPermission $($PublicFolder.Identity).tostring()
						
						$ClientPermission = $ClientPermission | sort AccessRights
										
						Write-Host ($ClientPermission | ft AccessRights, User  -AutoSize | Out-String) -foregroundcolor Cyan
						
										
				}
			
			}
	
	return $RecipientObject
	
	
	} #end process
	
	
	end 
	{
		#Write-Host ""
		if ( !$Silent )
		{
			Write-Host ""
			Write-Host "  End of Script  RecipientObject.ps1 " -ForegroundColor DarkBlue	-BackgroundColor Gray
			Write-Host ""
		}
	}
	
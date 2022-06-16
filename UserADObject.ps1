<#
.SYNOPSIS
     Use UserADObject to search for and display properties of
	 1. AD (Active Directory) users
	 2. On Prem mailboxes
	 3.  Office 365 ( Remote ) mailboxes.
	 
	 
.DESCRIPTION

    Requirements:  
	
	UserADObject  uses module  ActiveDirectory which is part of Microsoft RSAT (Remote System Adminsistraion Tools)
	
	Command used from this module is Get-ADUser 
	
	To enable ActiveDirectory module for Powershell follow instructions at 
	https://4sysops.com/archives/how-to-install-the-powershell-active-directory-module/ 
	
	UserADObject  does NOT require any Exchange rights. 
	All Exchange related information is read from the Active Directory. 
	
	For UserADObject to display information related to Office 365 Mailboxes ( remote mailboxes only ) a valid and prefxed with 'O365' remote session to Office 365 must be enabled 
	----------
	
	UserADObject  supports wild card search for : 
	
	- DisplayName 
	- Email Address 
	- Proxy (secondary)  Addresses
	- Sam Account Name 
	- DistinguishedName
	- UserPrincipalName
	- Description
	
	And exact search for 
	
	- ObjectGUID

	----------
	
.LINK
https://4sysops.com/archives/how-to-install-the-powershell-active-directory-module/ 
	
.PARAMETER UserName
    The search pattern

.PARAMETER return
    If present the script will return an PS ADuser object. 

.PARAMETER silent
    If selected  the script will not display any output (unless several items to choose from ). In addition -silent will turn   -return ON.  

.PARAMETER Details
	Provides additional details based on the key_word
	key_word can be composed of any of the following in any order 
	
	- ar Access Rights . Provides mailbox access rights
	- ir Inbox Rules. Provides mailbox inbox rules 
	- lic Licesnse. Provides user O365 license information

.PARAMETER memberof
    If present the script will return user's group memebrship object. 		
	
.PARAMETER OnPrem
    If present the script will return user's remote mailbox properties ( Mailbox properties as they are seen without remote session to Office 365 ). 			

.EXAMPLE
    PS>  .\UserADObject.ps1 user_name 
    Searches for user objects that have similar Display Name  or  Email Address or  any of the Proxy (secondary)  Addresses or Sam Account Name

.EXAMPLE
	PS> .\UserADObject.ps1 guid_string
    Searches for user objects that have exactly the same GUID

.EXAMPLE
	PS> $UserObject = .\UserADObject.ps1  user_name  -return 
    Searches for user objects that have similar Display Name  or  Email Address or  any of the Proxy (secondary)  Addresses or Sam Account Name,
	displays output  and returns a custom user object in  $UserObject
	
.EXAMPLE
	PS> $UserObject = .\UserADObject.ps1  user_name  -silent
	
    Searches for user objects that have similar Display Name  or  Email Address or  any of the Proxy (secondary)  Addresses or Sam Account Name,
	does not display output once the exact user is found  and returns a custom user object in  $UserObject
		
	$UserObject = fl *  	To see all properties of the returned object 

.EXAMPLE	
	
	PS> $UserObject = .\UserADObject.ps1  user_name  -Details key_word
	
	Provides additional information depending on the key_word
	
	key_word can be composed of any of the following in any order 
	
	- ar Access Rights . Provides mailbox access rights
	- ir Inbox Rules. Provides mailbox inbox rules 
	- lic Licesnse. Provides user O365 license information
	
		
.NOTES
    Author: Filip Neshev; filipne@yahoo.com
    Date:   August 2016    
#>

param(
    [parameter(Position=0,Mandatory=$false, ValueFromPipeline=$true,HelpMessage=" User Name search string")]
	[string]$UserName,
	[parameter(Position=1,Mandatory = $false, ValueFromPipeline= $false,HelpMessage="Returns the custom PS object  ")] 
	[switch]$return= $true,
	[parameter(Position=2,Mandatory = $false, ValueFromPipeline= $false,HelpMessage="Does not show details  ")] 
	[switch]$silent,
	[parameter(Position=3,Mandatory = $false, ValueFromPipeline = $false, HelpMessage=" Details")]
	$Details,
	[parameter(Position=4,Mandatory = $false, ValueFromPipeline= $false,HelpMessage="Shows group membership  ")] 
	[switch]$memberof,
	[parameter(Position=5,Mandatory = $false, ValueFromPipeline= $false,HelpMessage="Shows licence group membership  ")] 
	[switch]$licencedetails,
	[parameter(Position=6,Mandatory = $false, ValueFromPipeline= $false,HelpMessage="Shows licence group membership  ")] 
	[switch]$licencegroup,
	[parameter(Position=7, Mandatory = $false, ValueFromPipeline= $false,HelpMessage= "Gets the On Prem version of the mailbox ")] 
	[switch]$OnPrem,
	[parameter(Position = 8 , Mandatory = $false,ValueFromPipeline= $false, HelpMessage = "Msol")	]
	[switch]$msol,
	[parameter(Position = 9 ,Mandatory = $false, ValueFromPipeline = $false, HelpMessage = "includeMailboxStats")	]
	[switch]$includeMailboxStats = $true
	)


begin 
{

	#region Functions
	
		function fHtml-ToText ( $html )
		{
		 
				# param ( [System.String] $html )
				
				#write-host "$html " -ForegroundColor Magenta 
				
				 # remove line breaks, replace with spaces
				 $html = $html -replace "(`r|`n|`t)", " "
				 # write-verbose "removed line breaks: `n`n$html`n"

				 # remove invisible content
				 @('head', 'style', 'script', 'object', 'embed', 'applet', 'noframes', 'noscript', 'noembed') | % {
				  $html = $html -replace "<$_[^>]*?>.*?</$_>", ""
				 }
				 
				 # write-verbose "removed invisible blocks: `n`n$html`n"

				#write-host "$html " -ForegroundColor Red

				 # Condense extra whitespace
				 $html = $html -replace "( )+", " "
				 # write-verbose "condensed whitespace: `n`n$html`n"

				 # Add line breaks
				 @('div','p','blockquote','h[1-9]') | % { $html = $html -replace "</?$_[^>]*?>.*?</$_>", ("`n" + '$0' )} 
				 
				 # Add line breaks for self-closing tags
				 @('div','p','blockquote','h[1-9]','br') | % { $html = $html -replace "<$_[^>]*?/>", ('$0' + "`n")} 
				 
				 # write-verbose "added line breaks: `n`n$html`n"

				 #strip tags 
				 $html = $html -replace "<[^>]*?>", ""
				 
				 $html = $html -replace "&nbsp;", ""
				 
				 # write-verbose "removed tags: `n`n$html`n"
				 
				# write-host "$html " -ForegroundColor Red
				  
				 # replace common entities
				 
				 @( 
				  @("&amp;bull;", " * "),
				  @("&amp;lsaquo;", "<"),
				  @("&amp;rsaquo;", ">"),
				  @("&amp;(rsquo|lsquo);", "'"),
				  @("&amp;(quot|ldquo|rdquo);", '"'),
				  @("&amp;trade;", "(tm)"),
				  @("&amp;frasl;", "/"),
				  @("&amp;(quot|#34|#034|#x22);", '"'),
				  @('&amp;(amp|#38|#038|#x26);', "&amp;"),
				  @("&amp;(lt|#60|#060|#x3c);", "<"),
				  @("&amp;(gt|#62|#062|#x3e);", ">"),
				  @('&amp;(copy|#169);', "(c)"),
				  @("&amp;(reg|#174);", "(r)"),
				  @("&amp;nbsp;", " "),
				   @("&amp;(.{2,6});", "")
				 ) | % { $html = $html -replace $_[0], $_[1] }
				 
				 # write-verbose "replaced entities: `n`n$html`n"

				 return $html

			}

		function fDisplayADUserRegularMailbox ( $UserName )
		{
			Write-Verbose " InvocationName: $($MyInvocation.InvocationName); Line Number [$($MyInvocation.ScriptLineNumber)]; Line: $($MyInvocation.line)"
					
			#Write-Host "fDisplayADUserRegularMailbox" -ForegroundColor Magenta 
			
			if (!$silent )
			{	
						if ($UserName.msExchRecipientTypeDetails)
						{
							
								if (!($UserName.EmailAddress ) )
								{
									Write-Host ""
									Write-Host "	" -NoNewline
									Write-Host "  System FAILURE User  $($UserName.Name) is of type $($UserName.ADRecipientTypeDetails)  but EmailAddress in AD is empty ! " -ForegroundColor Yellow -BackgroundColor Red
									Write-Host ""
								}
			
						} # if ($UserName.msExchRecipientTypeDetails)
				
						$extensionAttributes =  $UserName | select extensionAttribute* | sort 
						
						#Write-Host "439" -ForegroundColor Magenta 
						
						Write-Host ( $UserName  | fl  `
						@{Label="Teams"; Expression = {  if( !$O365 -and !$EXO ) {  "SIP N/A (no O365 session)" }else { $_.EmailAddresses |  ?{ $_.ToString() -like "sip*" }   } } } ,`
						#@{Label="RemoteRoutingAddress"; Expression = { if ( $_.RemoteRoutingAddress){ if ( $remotemailbox.RemoteRoutingAddress -like "*@msoit.mail.onmicrosoft.com")  { $_.RemoteRoutingAddress } else { Write-Host "NOT VALID RemoteRoutingAddress ! $($_.RemoteRoutingAddress) Must be @msoit.mail.onmicrosoft.com  " -ForegroundColor White -BackgroundColor Red	   } }  }    }, `
						@{Label="RemoteRoutingAddress"; Expression = { if ( $_.RemoteRoutingAddress){ if ( $_.RemoteRoutingAddress -like "*@msoit.mail.onmicrosoft.com")  { $_.RemoteRoutingAddress } else { Write-Host "NOT VALID RemoteRoutingAddress ! $($_.RemoteRoutingAddress) Must be @msoit.mail.onmicrosoft.com  " -ForegroundColor White -BackgroundColor Red	   } }  }    }, `
						EmailAddressPolicyEnabled `
						| Out-String).Trim() -foregroundcolor Cyan 
						
						Write-Host ""
						
						$EMA = $UserName.proxyAddresses | ?{ $_.ToString()  -like "SMTP*" }  | sort  
						
						Write-Host ( $EMA  | fl | Out-String).Trim() -foregroundcolor Cyan
						
						Write-Host ""
				
						Write-Host ( $UserName  | fl ExchangeGuid,`
						@{Label="Alias"; Expression = { if ($_.EmailAddress) { $OnpremAlias = ( get-recipient $_.EmailAddress).Alias ;  if ( $OnpremAlias -ne  $_.Alias ) {  "`'OnPrem Alias: $OnpremAlias`' `'O365 Alias: $($_.Alias)`'  " }else{  $($_.Alias)    }  }  } },`
						 LastLogonTime,`
						@{Label=" "; Expression = { "" } },`
						@{Label="TotalItemSize"; Expression = {  if( !$O365 -and !$EXO -and !( $_.TotalItemSize ) ) { "N/A (no O365 session)"    }else{ $_.TotalItemSize }         }},`
						@{Label="Database"; Expression = {  if( !$O365 -and !$EXO  -and !( $_.Database ) ) {  "N/A (no O365 session)" }else{ $_.Database  }         }},`
						MailboxRegion,MailboxRegionLastUpdateTime,`
						#@{Label="Extention Attributes"; Expression = {$EnabledextensionAttributes }},`
						msExchWhenMailboxCreated, Mailbox-WhenCreated, Mailbox-WhenChanged,RetentionHoldEnabled ,RetentionPolicy,`
						@{Label="LitHold"; Expression = {  $_.LitigationHoldEnabled } } 	| 	Out-String).trim()  -ForegroundColor Cyan 
					
						if ( $UserName.LitigationHoldEnabled )
						{
							Write-Host ( $UserName  | fl `
							@{Label="LitHoldOwner"; Expression = {  $_.LitigationHoldOwner } },`
							@{Label="LitHoldDate"; Expression = {  $_.LitigationHoldDate } }`
							| Out-String).trim()  -ForegroundColor Cyan 					
						}
								
						Write-Host ( $extensionAttributes  | fl | Out-String).trim()  -ForegroundColor Cyan 
						
						$global:gUserName  = $UserName 
					
						Write-Host ""
						
			
			}
	
	} # fDisplayADUserRegularMailbox 
					
		function fDisplayADUserAccount ( $UserName )
		{
			Write-Verbose " InvocationName: $($MyInvocation.InvocationName); Line Number [$($MyInvocation.ScriptLineNumber)]; Line: $($MyInvocation.line)"
				
			if ( !$UserName ) 	
			{ 	
				Write-Host ""
				Write-Host "	" -NoNewline
				Write-Host "[usrad fDisplayMsolUser]  System FAILURE fDisplayADUser: No Value for input variable UserName provided  " -ForegroundColor Yellow -BackgroundColor Red
				Write-Host ""
				
				return  
			}
				
			$ADAccountName = $UserName.Name
			
			#write-host "$ADAccountName" -ForegroundColor Magenta
			
			if ( $ADAccountName[0] -eq " " )
			{
				$ADAccountName = "*" + $ADAccountName.Trim()
				
				#Write-Host ""
				Write-Host " " -NoNewline
				Write-Host "  Leading white space found in Name:  $ADAccountName " -ForegroundColor Yellow -BackgroundColor Red
				Write-Host ""
	
			}
			
			if ( $ADAccountName[$ADAccountName.lenght-1] -eq " " )
			{
				$ADAccountName = $ADAccountName.Trim() + "*"
				
				#Write-Host ""
				Write-Host " " -NoNewline
				Write-Host "  Trailing white space found in Name:  $ADAccountName " -ForegroundColor Yellow -BackgroundColor Red
				Write-Host ""
	
			}

				#Write-Host "148 Recipient-RecipientTypeDetails $($UserName.'Recipient-RecipientTypeDetails')" -ForegroundColor Magenta 
			
				if (!$silent )
				{	
						Write-Verbose "Recipient-RecipientTypeDetails $($UserName.'Recipient-RecipientTypeDetails')"
						
						$RecipientTypeDetails = $UserName."Recipient-RecipientTypeDetails"
						
						if ( !$RecipientTypeDetails )
						{
							# OnPrem mailboxes in OUs that are not synced with O365 do not appear as MailUser on O365 so we are using the onPrem attribute 
							
							$RecipientTypeDetails = "$($UserName.ADRecipientTypeDetails) (OnPrem Only)"
						
						}
						
						#Write-Host "164 $RecipientTypeDetails" -ForegroundColor Magenta
						
						$userAccountControl  = $UserName.userAccountControl.Tostring()
						
						Write-Verbose "			RecipientTypeDetails : $RecipientTypeDetails"
						Write-Verbose "			userAccountControl :  $userAccountControl "
						
					
						if ( !$RecipientTypeDetails )
						{
							Write-Verbose  " $($MyInvocation.InvocationName); Line [$($MyInvocation.ScriptLineNumber)]: $($MyInvocation.line); Recipient-RecipientTypeDetails empty"
						}
						else
						{

						}

				
						if(!$RecipientTypeDetails  -or  $RecipientTypeDetails -like  "*Not Recipient*") 
						{ 
							Write-verbose "RecipientTypeDetails $RecipientTypeDetails" 
							 
							$RecipientTypeDetails = "NOT a Recipient" 
						
							Write-Host "$RecipientTypeDetails" -NoNewline  -ForegroundColor  Blue	-BackgroundColor	Yellow
						}
						elseif ( $RecipientTypeDetails -eq "SharedMailbox" )
						{
							Write-Host " $RecipientTypeDetails " -NoNewline -foregroundcolor  Cyan   -backgroundcolor Blue 
						}
						elseif ( $RecipientTypeDetails -eq "MailUser" )
						{
							Write-Host " $RecipientTypeDetails " -NoNewline -foregroundcolor White  -BackgroundColor DarkYellow 
						}
						elseif ( $RecipientTypeDetails -eq "RoomMailbox"  -or $RecipientTypeDetails -eq "EquipmentMailbox" )
						{
							Write-Host " $RecipientTypeDetails " -NoNewline -Foregroundcolor  BLUE  -BackgroundColor   Cyan
						}				
						else
						{
							Write-Host " $RecipientTypeDetails " -NoNewline -foregroundcolor DarkBlue  -backgroundcolor Green
						}
						
						Write-Host " " -NoNewline    -backgroundcolor  DarkBlue
						Write-Host "$($UserName.DisplayName)"  -NoNewline  -foregroundcolor White   -backgroundcolor  DarkGreen 
						Write-Host " "  -NoNewline
						
						$UserLogonName = $UserName.SamAccountName
						
						$normalizedLogonName = fNormalizeUsername $UserLogonName
						
						$accountexpired  = fAccountExpires   $UserName  $true
						
						#Write-Host " $userAccountControl " -ForegroundColor Magenta 
					
						if( $userAccountControl -ne '512' -and   $userAccountControl -ne '544' -and  $userAccountControl -ne '66048' -or $accountexpired -or $PasswordExpired  ) 
						{ 	
							Write-Host "$($UserName.SamAccountName)" -foregroundcolor White   -backgroundcolor Red  -NoNewline	
							Write-Host " "  -NoNewline	
							#Write-Host "$($UserAccountControlList.Item($UserName.userAccountControl.tostring()))"  -foregroundcolor White   -backgroundcolor Red  -NoNewline	
						}
						elseif ( $UserLogonName -ne $normalizedLogonName -or $UserLogonName.Contains(" ") -and !$Msol )
						{
							 	Write-Host "$UserLogonName" -foregroundcolor red   -backgroundcolor yellow  -NoNewline	
								Write-Host  " " -NoNewline
								Write-Host " User logon name '$UserLogonName' is not valid ! Normalized: '$normalizedLogonName'  "  -ForegroundColor	Red	-BackgroundColor	yellow  -NoNewline	
						}
						else
						{ 		
								# active account 
								
								Write-host "$UserLogonName" -ForegroundColor DarkBlue -BackgroundColor Green  -NoNewline 
								#Write-verbose  "$($UserAccountControlList.Item($UserName.userAccountControl.tostring()))"
						}
						
						if ($UserName.userAccountControl -eq "514" )
						{
							Write-Host "account disabled" -foregroundcolor White    -backgroundcolor Red  -NoNewline 
						}
												
						Write-Host  " " -NoNewline		
						
						# fPasswordStatus   $UserName 	
						
						$PasswordExpired = fPasswordStatus    $UserName 	
						
						Write-Host  " " -NoNewline
			
						$accountexpired  = fAccountExpires    $UserName
						
						Write-Host  " " -NoNewline
						 
				 
						
						#$AccountIsLockedOut = $UserName.LockedOut
						
						if( $UserName.LockedOut )
						{ 
							Write-Host "Account Is Locked Out" -foregroundcolor White    -backgroundcolor DarkYellow 
						}
						

				
						Write-Host ""
						Write-Host ""
						
						# @{Label="MSExchRecipientTypeDetails"; Expression={ $_.MSExchRecipientTypeDetails; 	$RecipientTypeDetailsList.Item(  $_.MSExchRecipientTypeDetails.ToString() )} }`
						
						#write-host "$($UserName.msExchRecipientTypeDetails)" -ForegroundColor Magenta
																		
						Write-Host ( $UserName  | fl `
						Description, Title, employeeType, telephoneNumber, Company, Country,State, Office,  Department ,`
						@{n="Notes(Info)";e={  $_.Info }} ,`
						SamAccountName, UserPrincipalName,EmailAddress,`
						@{n="Manager";e={  (Get-ADUser $_.Manager ).Name }} ,`
						HomeDirectory, CanonicalName,`
						Mail,   ipPhone, EmployeeNumber | Out-String).trim() -foregroundcolor Green -NoNewline
						
						 Write-Host ""
						 Write-Host ""
						 
					 
						 Write-Host ( $UserName  | fl @{n="Creator"; e={ ($_.Creator).Name } },`
						 whenCreated, whenChanged,`
						 @{n="LastLogOn";e={  fConvertADDate  $_.LastLogonTimestamp  }  }  | Out-String).Trim() -foregroundcolor Green -NoNewline

						
						$global:gUserName = $UserName
						
						Write-Host ""
						Write-Host ""
					
						if ($UserName.msExchRecipientTypeDetails)
						{
							
								if (!($UserName.EmailAddress ) )
								{
									Write-Host ""
									Write-Host "	" -NoNewline
									Write-Host "  System FAILURE User  $($UserName.Name) is of type $($UserName.ADRecipientTypeDetails)  but EmailAddress in AD is empty ! " -ForegroundColor Yellow -BackgroundColor Red
									Write-Host ""
								}
							
								#$Alias = $UserName.Alias
				
								$extensionAttributes =  $UserName | select extensionAttribute* | sort 
								
								
								if ( $UserName."Recipient-HiddenFromAddressListsEnabled" )
								{
									Write-Host ( $UserName  | fl Recipient-PrimarySmtpAddress | Out-String).Trim()  -foregroundcolor White -BackgroundColor DarkMagenta
								}
								else
								{
									Write-Host ( $UserName  | fl Recipient-PrimarySmtpAddress | Out-String).Trim()  -ForegroundColor DarkBlue -BackgroundColor Green
														
								}
								
						}		# if ($UserName.msExchRecipientTypeDetails)
								
								Write-Host ""
								
					}	#	if (!$silent )	
					
	#fDisplayADUserRegularMailbox  $UserName 
	
	} ########### END function fDisplayADUserAccount
		
		function fDisplayADUser ( $UserName )
		{
				Write-Verbose " InvocationName: $($MyInvocation.InvocationName); Line Number [$($MyInvocation.ScriptLineNumber)]; Line: $($MyInvocation.line)"
				
				if ( !$UserName ) 	
				{ 	
					Write-Host ""
					Write-Host "	" -NoNewline
					Write-Host "  System FAILURE fDisplayADUser: No Value for input variable UserName provided  " -ForegroundColor Yellow -BackgroundColor Red
					Write-Host ""
					
					return  
				}
				
				$ADAccountName = $UserName.Name
				
				#write-host "$ADAccountName" -ForegroundColor Magenta
				
				if ( $ADAccountName[0] -eq " " )
				{
					$ADAccountName = "*" + $ADAccountName.Trim()
					
					#Write-Host ""
					Write-Host " " -NoNewline
					Write-Host "  Leading white space found in Name:  $ADAccountName " -ForegroundColor Yellow -BackgroundColor Red
					Write-Host ""
		
				}
				
				if ( $ADAccountName[$ADAccountName.lenght-1] -eq " " )
				{
					$ADAccountName = $ADAccountName.Trim() + "*"
 					
					#Write-Host ""
					Write-Host " " -NoNewline
					Write-Host "  Trailing white space found in Name:  $ADAccountName " -ForegroundColor Yellow -BackgroundColor Red
					Write-Host ""
		
				}
				
				
				#Write-Host "148 Recipient-RecipientTypeDetails $($UserName.'Recipient-RecipientTypeDetails')" -ForegroundColor Magenta 
			
				if (!$silent )
				{	
						
						<#
						
						Write-Verbose "Recipient-RecipientTypeDetails $($UserName.'Recipient-RecipientTypeDetails')"
						
						$RecipientTypeDetails = $UserName."Recipient-RecipientTypeDetails"
						
						if ( !$RecipientTypeDetails )
						{
							# OnPrem mailboxes in OUs that are not synced with O365 do not appear as MailUser on O365 so we are using the onPrem attribute 
							
							$RecipientTypeDetails = "$($UserName.ADRecipientTypeDetails) (OnPrem Only)"
						
						}
						
						#Write-Host "164 $RecipientTypeDetails" -ForegroundColor Magenta
						
						$userAccountControl  = $UserName.userAccountControl.Tostring()
						
						Write-Verbose "			RecipientTypeDetails : $RecipientTypeDetails"
						Write-Verbose "			userAccountControl :  $userAccountControl "
						
					
						if ( !$RecipientTypeDetails )
						{
							Write-Verbose  " $($MyInvocation.InvocationName); Line [$($MyInvocation.ScriptLineNumber)]: $($MyInvocation.line); Recipient-RecipientTypeDetails empty"
						}
						else
						{

						}

				
						if(!$RecipientTypeDetails  -or  $RecipientTypeDetails -like  "*Not Recipient*") 
						{ 
							Write-verbose "RecipientTypeDetails $RecipientTypeDetails" 
							 
							$RecipientTypeDetails = "NOT a Recipient" 
						
							Write-Host "$RecipientTypeDetails" -NoNewline  -ForegroundColor  Blue	-BackgroundColor	Yellow
						}
						elseif ( $RecipientTypeDetails -eq "SharedMailbox" )
						{
							Write-Host " $RecipientTypeDetails " -NoNewline -foregroundcolor  Cyan   -backgroundcolor Blue 
						}
						elseif ( $RecipientTypeDetails -eq "MailUser" )
						{
							Write-Host " $RecipientTypeDetails " -NoNewline -foregroundcolor White  -BackgroundColor DarkYellow 
						}
						elseif ( $RecipientTypeDetails -eq "RoomMailbox"  -or $RecipientTypeDetails -eq "EquipmentMailbox" )
						{
							Write-Host " $RecipientTypeDetails " -NoNewline -Foregroundcolor  BLUE  -BackgroundColor   Cyan
						}				
						else
						{
							Write-Host " $RecipientTypeDetails " -NoNewline -foregroundcolor DarkBlue  -backgroundcolor Green
						}
						
						Write-Host " " -NoNewline    -backgroundcolor  DarkBlue
						Write-Host "$($UserName.DisplayName)"  -NoNewline  -foregroundcolor White   -backgroundcolor  DarkGreen 
						Write-Host " "  -NoNewline
						
						$UserLogonName = $UserName.SamAccountName
						
						$normalizedLogonName = fNormalizeUsername $UserLogonName
						
						$accountexpired  = fAccountExpires  $UserName  $true
						
						#Write-Host " $userAccountControl " -ForegroundColor Magenta 
					
						if( $userAccountControl -ne '512' -and   $userAccountControl -ne '544' -and  $userAccountControl -ne '66048' -or $accountexpired -or $PasswordExpired  ) 
						{ 	
							Write-Host "$($UserName.SamAccountName)" -foregroundcolor White   -backgroundcolor Red  -NoNewline	
							Write-Host " "  -NoNewline	
							#Write-Host "$($UserAccountControlList.Item($UserName.userAccountControl.tostring()))"  -foregroundcolor White   -backgroundcolor Red  -NoNewline	
						}
						elseif ( $UserLogonName -ne $normalizedLogonName -or $UserLogonName.Contains(" ") -and !$Msol )
						{
							 	Write-Host "$UserLogonName" -foregroundcolor red   -backgroundcolor yellow  -NoNewline	
								Write-Host  " " -NoNewline
								Write-Host " User logon name '$UserLogonName' is not valid ! Normalized: '$normalizedLogonName'  "  -ForegroundColor	Red	-BackgroundColor	yellow  -NoNewline	
						}
						else
						{ 		
								# active account 
								
								Write-host "$UserLogonName" -ForegroundColor DarkBlue -BackgroundColor Green  -NoNewline 
								#Write-verbose  "$($UserAccountControlList.Item($UserName.userAccountControl.tostring()))"
						}
												
						Write-Host  " " -NoNewline		
						
						# fPasswordStatus   $UserName 	
						
						$PasswordExpired = fPasswordStatus   $UserName 	
						
						Write-Host  " " -NoNewline
			
						 $accountexpired  = fAccountExpires   $UserName
						
						#$AccountIsLockedOut = $UserName.LockedOut
						
						if( $UserName.LockedOut )
						{ 
							Write-Host "Account Is Locked Out" -foregroundcolor White    -backgroundcolor DarkYellow 
						}
				
						Write-Host ""
						Write-Host ""
						
						# @{Label="MSExchRecipientTypeDetails"; Expression={ $_.MSExchRecipientTypeDetails; 	$RecipientTypeDetailsList.Item(  $_.MSExchRecipientTypeDetails.ToString() )} }`
						
						#write-host "$($UserName.msExchRecipientTypeDetails)" -ForegroundColor Magenta
																		
						Write-Host ( $UserName  | fl `
						Description, Title, telephoneNumber, Company, Country,State, Office,  Department ,`
						@{n="Notes(Info)";e={  $_.Info }} ,`
						SamAccountName, UserPrincipalName,`
						@{n="Manager";e={  (Get-ADUser $_.Manager ).Name }} ,`
						HomeDirectory, CanonicalName,`
						Mail,   ipPhone, EmployeeNumber | Out-String).trim() -foregroundcolor Green -NoNewline
						
						 Write-Host ""
						 Write-Host ""
						 
					 
						 Write-Host ( $UserName  | fl @{n="Creator"; e={ ($_.Creator).Name } },`
						 whenCreated, whenChanged,`
						 @{n="LastLogOn";e={  fConvertADDate  $_.LastLogonTimestamp  }  }  | Out-String).Trim() -foregroundcolor Green -NoNewline

						
						$global:gUserName = $UserName
						
						Write-Host ""
						Write-Host ""
						#>
						
						### Display Mailbox
					
						if ($UserName.msExchRecipientTypeDetails)
						{
							
								if (!($UserName.EmailAddress ) )
								{
									Write-Host ""
									Write-Host "	" -NoNewline
									Write-Host "  System FAILURE User  $($UserName.Name) is of type $($UserName.ADRecipientTypeDetails)  but EmailAddress in AD is empty ! " -ForegroundColor Yellow -BackgroundColor Red
									Write-Host ""
								}
							
								#$Alias = $UserName.Alias
				
								$extensionAttributes =  $UserName | select extensionAttribute* | sort 
								
								<#
								$EnabledextensionAttributes = @()
								
								for( $i=1; $i -le 15; $i++)
								{
									$extensionAttributeNumber = "extensionAttribute" + $i
									
									if ( $extensionAttributes.$extensionAttributeNumber)
									{
										#$EnabledextensionAttributes = $EnabledextensionAttributes + $extensionAttributeNumber + ": " + $extensionAttributes.$extensionAttributeNumber + "; "
										
										$EnabledextensionAttributes += $extensionAttributeNumber + ": " + $extensionAttributes.$extensionAttributeNumber
										
									}
								
								}
								#>
								
								<#
								if ( $UserName.HiddenFromAddressListsEnabled)
								{
									Write-Host ( $UserName  | fl PrimarySmtpAddress | Out-String).Trim()  -foregroundcolor White -BackgroundColor DarkMagenta
								}
								else
								{
									Write-Host ( $UserName  | fl PrimarySmtpAddress | Out-String).Trim()  -ForegroundColor DarkBlue -BackgroundColor Green
								
								}
								
								#>
								
								Write-Host ""
												
								Write-Host ( $UserName  | fl  `
								@{Label="Teams"; Expression = {  if( !$O365 -and !$EXO ) {  "SIP N/A (no O365 session)" }else { $_.EmailAddresses |  ?{ $_.ToString() -like "sip*" }   } } } ,`
								@{Label="RemoteRoutingAddress"; Expression = { if ( $_.RemoteRoutingAddress){ if ( $remotemailbox.RemoteRoutingAddress -like "*@msoit.mail.onmicrosoft.com")  { $_.RemoteRoutingAddress } else { Write-Host "NOT VALID RemoteRoutingAddress ! $($_.RemoteRoutingAddress) Must be @msoit.mail.onmicrosoft.com  " -ForegroundColor White -BackgroundColor Red	   } }  }    }, `
								EmailAddressPolicyEnabled, `
								@{Label="OnPrem EmailAddressPolicyEnabled"; Expression = {  $remotemailbox.EmailAddressPolicyEnabled } }`
								| Out-String).Trim() -foregroundcolor Cyan 
								
								Write-Host ""
								
								$EMA = $UserName.proxyAddresses | ?{ $_.ToString()  -like "SMTP*" }  | sort  
								
								Write-Host ( $EMA  | fl | Out-String).Trim() -foregroundcolor Cyan
								
								Write-Host ""
						
								Write-Host ( $UserName  | fl ExchangeGuid,`
								@{Label="Alias"; Expression = { if ($_.EmailAddress) { $OnpremAlias = ( get-recipient $_.EmailAddress).Alias ;  if ( $OnpremAlias -ne  $_.Alias ) {  "`'OnPrem Alias: $OnpremAlias`' `'O365 Alias: $($_.Alias)`'  " }else{  $($_.Alias)    }  }  } },`
								 LastLogonTime,`
								@{Label=" "; Expression = { "" } },`
								@{Label="TotalItemSize"; Expression = {  if( !$O365 -and !$EXO -and !( $_.TotalItemSize ) ) { "N/A (no O365 session)"    }else{ $_.TotalItemSize }         }},`
								ProhibitSendQuota,`
								@{Label="Database"; Expression = {  if( !$O365 -and !$EXO  -and !( $_.Database ) ) {  "N/A (no O365 session)" }else{ $_.Database  }         }},`
								MailboxRegion,MailboxRegionLastUpdateTime,`
								#@{Label="Extention Attributes"; Expression = {$EnabledextensionAttributes }},`
								msExchWhenMailboxCreated, Mailbox-WhenCreated, Mailbox-WhenChanged,RetentionHoldEnabled ,RetentionPolicy,`
								@{Label="LitHold"; Expression = {  $_.LitigationHoldEnabled } } 	| 	Out-String).trim()  -ForegroundColor Cyan 
							
								if ( $UserName.LitigationHoldEnabled )
								{
									Write-Host ( $UserName  | fl `
									@{Label="LitHoldOwner"; Expression = {  $_.LitigationHoldOwner } },`
									@{Label="LitHoldDate"; Expression = {  $_.LitigationHoldDate } }`
									| Out-String).trim()  -ForegroundColor Cyan 					
								}
										
								Write-Host ( $extensionAttributes  | fl | Out-String).trim()  -ForegroundColor Cyan 
								
								$global:gUserName  = $UserName 
							
								Write-Host ""
				
								<#
								if ( $UserName.HiddenFromAddressListsEnabled)
								{
									Write-Host ( $UserName  | fl PrimarySmtpAddress | Out-String).Trim()  -foregroundcolor White -BackgroundColor DarkMagenta
								}
								else
								{
									Write-Host ( $UserName  | fl PrimarySmtpAddress | Out-String).Trim()  -ForegroundColor DarkBlue -BackgroundColor Green
								
								}
												
								Write-Host ( $UserName  | fl  `
								@{Label="Teams"; Expression = {  if( !$O365 -and !$EXO ) {  "SIP N/A (no O365 session)" }else { $_.EmailAddresses |  ?{ $_.ToString() -like "sip*" }   } } } ,`
								@{Label="RemoteRoutingAddress"; Expression = { if ( $_.RemoteRoutingAddress){ if ( $remotemailbox.RemoteRoutingAddress -like "*@msoit.mail.onmicrosoft.com")  { $_.RemoteRoutingAddress } else { Write-Host "NOT VALID RemoteRoutingAddress ! $($_.RemoteRoutingAddress) Must be @msoit.mail.onmicrosoft.com  " -ForegroundColor White -BackgroundColor Red	   } }  }    } | Out-String).Trim() -foregroundcolor Cyan 

								Write-Host ""
								
							
								$EMA = $UserName.proxyAddresses | ?{ $_.ToString()  -like "SMTP*" }  | sort  
								
								Write-Host ( $EMA  | fl | Out-String).Trim() -foregroundcolor Cyan
						
							
								Write-Host ""
								#>
					
								if ( $UserName.msExchArchiveName)
								{
									# Dsplay Online Archive mailbox properties 
											
									if ( $UserName."Mailbox-UserPrincipalName" )
									{	
										
										#Write-Host ( $UserName  | ft  -HideTableHeaders 	ArchiveGuid | Out-String).trim()   -foregroundcolor Cyan
										#Write-Host ""
										Write-Host ( $UserName  | ft -HideTableHeaders msExchArchiveName 	| Out-String).trim() -foregroundcolor White -BackgroundColor Blue
										Write-Host ( $UserName  | fl `
										 ArchiveGuid,`
										@{Label="ArchiveMailboxStats-TotalItemSize"; Expression = {  if( !$O365 -and  !($_."ArchiveMailboxStats-TotalItemSize") ) {  "N/A (no O365 session)" }else { $_."ArchiveMailboxStats-TotalItemSize" }   } }  ,`
										ArchiveQuota ,AutoExpandingArchiveEnabled,`
										@{Label="ArchiveDatabase"; Expression = {  if( !$O365 -and !($_.ArchiveDatabase) ) {  "N/A (no O365 session) $( $_.ArchiveDatabase)  " }else { $_.ArchiveDatabase }   } }`
										| Out-String).trim() -foregroundcolor Cyan
									}
								}
								else
								{
									#Write-Host ""
									Write-Host " " -NoNewline
									Write-Host " No Online archive" -ForegroundColor  Blue	-BackgroundColor	Yellow
									Write-Host ""
								
								}
							
							
								$global:gUserName = $UserName 
								
								fAutomap  $UserName
								fMailboxForward $UserName 
								fOutOfOffice $UserName
								fAcceptMessagesOnlyFrom $UserName

								if ( $UserName.ADRecipientTypeDetails -eq "RoomMailbox" -or $UserName.ADRecipientTypeDetails -eq "RemoteRoomMailbox"  -or   $UserName.ADRecipientTypeDetails -eq "EquipmentMailbox" -and !$Silent  )
								{ 
								 	#RoomMailbox  $Mailbox
									fRoomMailbox   $UserName 
								}
					
								if ( $Details -and $Details.Contains( "ar" ))
								{	 	
									fMailboxRights $UserName
								}
				
								if ( $Details -and $Details.Contains( "ir" ))
								{	 	
									fInboxRules  $UserName
								}
								
								if ( $Details -and $Details.Contains( "lic" ))
								{	 	
									Write-Host ""
									.\Get-UserLicenseDetails.ps1  $UserName.UserPrincipalName
								}
							
					}	# if ($UserName.msExchRecipientTypeDetails)
					elseif ( !$UserName.EmailAddress )
					{
						Write-Verbose "Missing email address"
					}
				
				}# if(!$Silent)
				
				if ( $memberof)
				{
					fmemberof $UserName
				}
				
				if ( $licencegroup )
				{
					flicencegroup $UserName
				}
				
				if ( $licencedetails)
				{
					
					#Write-Host "$($UserName.UserPrincipalName)" -ForegroundColor Magenta
					
					#$global:gUserName  = $UserName
					
					.\Get-UserLicenseDetails.ps1 $UserName.UserPrincipalName
				}
				Write-Host ""
				
	} ########### END function fDisplayADUser

		function fDisplayMsolUser ( $UserName )
		{

			Write-Verbose " InvocationName: $($MyInvocation.InvocationName); Line Number [$($MyInvocation.ScriptLineNumber)]; Line: $($MyInvocation.line)"
					
			if ( !$UserName ) 	
			{ 	
				Write-Host ""
				Write-Host "	" -NoNewline
				Write-Host "[usrad fDisplayMsolUser]  System FAILURE fDisplayADUser: No Value for input variable UserName provided  " -ForegroundColor Yellow -BackgroundColor Red
				Write-Host ""
				
				return  
			}
									
			$RecipientTypeDetails = $UserName.CloudExchangeRecipientDisplayType
			
			Write-Host ( $UserName  | fl  ImmutableId,LiveId | Out-String).trim()   -foregroundcolor DarkCyan
			
			Write-Host ( $UserName  | fl CloudExchangeRecipientDisplayType,`
			@{Label="MSExchRecipientTypeDetails"; Expression={ $_.MSExchRecipientTypeDetails; 	$RecipientTypeDetailsList.Item(  $_.MSExchRecipientTypeDetails.ToString() )} }`
			| Out-String).trim()    -foregroundcolor DarkCyan
			
			Write-Host ""
			Write-Host "$($UserName.UserType)"   -foregroundcolor Cyan -BackgroundColor Blue -NoNewline
			Write-Host "	" -NoNewline 
			Write-Host "$($UserName.DisplayName)"   -foregroundcolor Cyan -BackgroundColor Blue
			Write-Host ""
			
			
			Write-Host ( $UserName | fl  objectID, UserPrincipalName, SignInName, isLicensed ,`
			@{Label="RecipientType"; Expression={ 	$RecipientTypeDetailsList.Item(  $_.MSExchRecipientTypeDetails.ToString() )} },`
			BlockCredential,  WhenCreated  | Out-String ).Trim()  -foregroundcolor Green

		#Write-Host ( $UserName | fl | Out-host 	) -foregroundcolor Green
		
		#Write-Host ""
		Write-Host ( $UserName  | fl LastPasswordChangeTimestamp, PasswordNeverExpires, LastDirSyncTime, DirSyncProvisioningErrors, SoftDeletionTimestamp | Out-String).trim()   -foregroundcolor Green
		Write-Host ""

		
		#Write-Host ( $UserName  | fl  UserPrincipalName,  BlockCredential,  `
		Write-Host ( $UserName  | fl  Title, PhoneNumber, Department, Country,State, City,StreetAddress, Office, `
		IsLicensed, IndirectLicenseErrors,`
		Licenses, LicenseReconciliationNeeded,`
		PreferredDataLocation, PreferredLanguage, AlternateEmailAddresses, ProxyAddresses,`
		UsageLocation, ValidationStatus,`
		@{Name="Error";Expression={ ( $_.errors[0].ErrorDetail.objecterrors.errorrecord.ErrorDescription )  }  } `
		| Out-String).trim()   -foregroundcolor Green
		
		
		
		#$UserName.errors[0].errordetail.objecterrors.errorrecord.ErrorDescription.tostring()
		
		# Errors 
		
		#(Get-MsolUser -UserPrincipalName $UserName.UserPrincipalName ).errors.errordetail.objecterrors.errorrecord | fl
		
		#$errors = $_.Errors.ErrorDetail.objecterrors.errorrecord.ErrorDescription
		
		$errors =  (Get-MsolUser -UserPrincipalName $UserName.UserPrincipalName ).errors
		
		$errors | %{ if ( $_.errordetail.objecterrors ) { Write-Host ( $_ | fl ErrorDetail, Resolved, ServiceInstance, Timestamp | Out-String ).trim() -foregroundcolor Magenta ; Write-Host (  $_.errordetail.objecterrors.errorrecord.ErrorDescription | fl | Out-String ).trim() -foregroundcolor Magenta } }
		
		#$errors | %{ $_ }
		
		 $global:gUserName  =  $UserName
		
		#$UserName.Licenses.AccountSku.SkuPartNumber
		
		Write-Host ""
		Write-Host "`$Ro =  Get-O365Recipient $($UserName.UserPrincipalName) -ErrorAction silent " -ForegroundColor Yellow
		
		$Ro =  Get-O365Recipient $UserName.UserPrincipalName -ErrorAction silent 
		
		if ($Ro )
		{
				Write-Host ""
				Write-Host "`$Mailbox =  Get-O365Mailbox   $($Ro.PrimarySmtpAddress) " -ForegroundColor Yellow
				Write-Host ""
				
				$Mailbox =  Get-O365Mailbox   $Ro.PrimarySmtpAddress #  -ErrorAction silent
				
				
				Write-Host ( $Mailbox  | fl ExchangeGuid, RecipientTypeDetails, Alias, PrimarySmtpAddress,`
				#@{Label="Teams"; Expression = {  if($O365) { $_.EmailAddresses |  ?{ $_.ToString() -like "sip*" } }else{ "`'SIP N/A`'"  }         }},`
				#@{Label="Alias"; Expression = { if ($_.EmailAddress) { $OnpremAlias = ( get-recipient $_.EmailAddress).Alias ;  if ( $OnpremAlias -ne  $_.Alias ) {  "`'OnPrem Alias: $OnpremAlias`' `'O365 Alias: $($_.Alias)`'  " }else{  $($_.Alias)    }  }  } },`
				@{Label="LitHold"; Expression = {  $_.LitigationHoldEnabled } },`
				@{Label="LitHoldOwner"; Expression = {  $_.LitigationHoldOwner } },`
				#@{Label="TotalItemSize"; Expression = {  if( $_.TotalItemSize ) { $_.TotalItemSize   }else{ "N/A (no O365 session)"  }         }},`
				#@{Label="TotalItemSize"; Expression = {  if( !$O365 -and !$EXO -and !( $_.TotalItemSize ) ) { "N/A (no O365 session)"    }else{ $_.TotalItemSize }         }},`
				@{Label="TotalItemSize"; Expression = {   ( Get-O365MailboxStatistics $_.PrimarySmtpAddress  ).TotalItemSize.tostring()      }},`
				#@{Label="Database"; Expression = {  if( $_.Database ) { $_.Database   }else{ "N/A (no O365 session)"  }         }},`
				@{Label="Database"; Expression = {  if( !$O365 -and !$EXO  -and !( $_.Database ) ) {  "N/A (no O365 session)" }else{ $_.Database  }         }},`
				MailboxRegion,MailboxRegionLastUpdateTime,`
				@{Label="Extention Attributes"; Expression = { $EnabledextensionAttributes }},`
				msExchWhenMailboxCreated, Mailbox-WhenCreated, Mailbox-WhenChanged | Out-String).trim()  -ForegroundColor Cyan
				
				Write-Host ""
				Write-Host "`$ArchiveMailboxstats   =  Get-O365Mailbox  -Archive  $($Ro.PrimarySmtpAddress) -ea silent " -ForegroundColor Yellow
				Write-Host ""
				
				$ArchiveMailboxstats = Get-O365MailboxStatistics -Archive  $Ro.PrimarySmtpAddress -ea silent
				
				$global:gArchiveMailboxstats = $ArchiveMailboxstats
				
								
				Write-Host ( $ArchiveMailboxstats  | fl DisplayName, TotalItemSize, ArchiveQuota,  Database | Out-String).trim()  -ForegroundColor Cyan
		
		}
		
		
		Write-Host ""
		Write-Host "  Licenses:	" -ForegroundColor Cyan	-BackgroundColor	Blue
		Write-Host ""
		
		$UserLicenseDetail = @() 
		
		$msoluser = $UserName
		
		 foreach($license in $msoluser.Licenses )
		 {
			
			#Write-Host "`$license" -ForegroundColor Magenta
			
			#$license | fl
			
			$AccountSkuId  = $license.AccountSkuId.Split(":")[1]
			
			$LicenseName =  $licenseNames.Item( $AccountSkuId ) 
			
			If($license.GroupsAssigningLicense.Count -eq 0 )
			{
                #Direct adssignment 
                $AssignmentPath = "Direct"
            }
			else
			{
                 # Assignment via a group 
				    
					foreach($groupid in $license.GroupsAssigningLicense )
					{
                        #Checking each object id, if the id is same as user's object id, there is duplication of license assignment, else capture all the group names
						
                        $AssignmentPath = "Inherited"
		
						
						If( $groupid -eq $msoluser.ObjectId ) 
						{
	                            If ($license.GroupsAssigningLicense.Count -eq 1 )
								{
	                                $AssignmentPath = "Direct"    
	                            }
	                            else 
								{
	                                $AssignmentPath += " + Direct"
	                            }
	                            break
                   		}
						
						 #Capture group names
                   		 $GroupNames += Get-MsolGroup -ObjectId $groupid | Select-Object -ExpandProperty DisplayName
					
					} # foreach($groupid in $license.GroupsAssigningLicense )
				
			}
			
				#Write-Host "330" -ForegroundColor Magenta
				
				#$userlicenseerror.group
			
			    $UserLicenseDetail += [PSCustomObject]@{
                'DisplayName' = $msoluser.DisplayName
                'UserPrincipalName' = $msoluser.UserPrincipalName
                'isLicensed' = $msoluser.isLicensed
                'LicenseCount' = $msoluser.Licenses.count
				'AccountSkuId' = $AccountSkuId
                'LicenseName' = $LicenseName
                'AssignmentPath' = $AssignmentPath
                'LicensedGroups'= $GroupNames
				'Error' = if ( $userlicenseerror.group -eq $GroupNames) {  $userlicenseerror.error }
				#'ServiceStatus' = $license.ServiceStatus
            }
            
			$GroupNames = ""
	
		} # foreach($license in $msoluser.Licenses )
	
		
		$IndirectLicenseErrors = $msoluser.IndirectLicenseErrors
		
		
		if ( $IndirectLicenseErrors )
		{
		
			Write-Host ""
			Write-Host "`$IndirectLicenseErrors = `$msoluser.IndirectLicenseErrors" -ForegroundColor Yellow
			Write-Host ""
		
			$userlicenseerrors = @()
			
			$IndirectLicenseErrors  | % {

				if ( $_.Error -ne "Other")
				{
						$userlicenseerror = $_ | select    @{  Label="WhenCreated" ;      Expression= {  $msoluser.WhenCreated   }   },`
						@{  Label="DisplayName" ;      Expression= {  $msoluser.DisplayName   }   },`
						@{  Label="UserPrincipalName" ;      Expression= {  $msoluser.UserPrincipalName   }   },`
						@{  Label="LicenseName" ;      Expression= {  $licenseNames.Item( $_.AccountSku.SkuPartNumber ) }   },`
						@{  Label="Group" ;      Expression= {  (Get-MsolGroup -ObjectId    $_.ReferencedObjectId).DisplayName }   },`
						Error
						
						#$userlicenseerror | ft 
						
						$userlicenseerrors  += $userlicenseerror
				}

	  	  }
		
			if ( $userlicenseerrors )
			{
				Write-Host ""
				Write-Host ""
				Write-Host "Indirect ( via a license group) user license errors" -ForegroundColor Cyan	-BackgroundColor	Blue
				
				Write-Host ( $userlicenseerrors  | ft  | Out-String ) -foregroundcolor Magenta 
			
			}
		
		} # if ( $IndirectLicenseErrors )
		
			Write-Host ( $UserLicenseDetail | ft  | Out-String ).trim() -foregroundcolor Cyan

		
		
		if ( $memberof )
		{
			
			.\Remote-SessionConnect.ps1 -AzureAD
			
			Write-Host ""
			Write-Host ""
			Write-Host "  Membership:	" -ForegroundColor Cyan	-BackgroundColor	Blue
			Write-Host ""
			
			Write-Host "`$GroupObjectIds = (Get-AzureADUser -ObjectId  $($UserName.ObjectId)   | Get-AzureADUserMembership).ObjectId" -ForegroundColor Yellow
			Write-Host ""
			
			$GroupObjectIds = (Get-AzureADUser -ObjectId   $UserName.ObjectId  | Get-AzureADUserMembership ).ObjectId
			
			$msolgroups = @()
			
			$GroupObjectIds | 
			%{ 
			
					Write-Host "`$msolgroup =  Get-MsolGroup -ObjectId $_ -ErrorAction silent" -ForegroundColor Yellow
					
					$msolgroup =  Get-MsolGroup -ObjectId $_ -ErrorAction silent
					
					$msolgroups += $msolgroup
			
				}
			
						
			Write-Host (  $msolgroups  |sort DisplayName  |ft ObjectId, GroupType,  DisplayName, Description, EmailAddress, ProxyAddresses | Out-String) -foregroundcolor Cyan
			#Write-Host (  $msolgroups  | ft -AutoSize | Out-String) -foregroundcolor Magenta
					
			Write-Host "Get-AzureADUser -ObjectId   $($UserName.ObjectId)  | Get-AzureADUserMembership | ft ObjectId, ObjectType,MailEnabled, SecurityEnabled,  DisplayName, Description, Mail " -ForegroundColor Yellow
			
			Write-Host (  Get-AzureADUser -ObjectId   $UserName.ObjectId  | Get-AzureADUserMembership |sort DisplayName  | ft ObjectId, ObjectType,MailEnabled, SecurityEnabled,  DisplayName, Description, Mail | Out-String) -foregroundcolor Cyan
			#Write-Host (  Get-AzureADUser -ObjectId   $UserName.ObjectId  | Get-AzureADUserMembership | ft | Out-String) -foregroundcolor Cyan
		}
		
	
	} ########### END function fDisplayMsolUser

		function fSearchADUser ($UserName, $return )
		{
						
			Write-Verbose " InvocationName: $($MyInvocation.InvocationName); Line Number [$($MyInvocation.ScriptLineNumber)]; Line: $($MyInvocation.line)"
			
			$AllSearch = @()
			
			if ( $UserName -eq "^" )  
			{ 
				Write-Host "	" -NoNewline
				Write-Host "  ATTENTION :  Skipped" -ForegroundColor  Blue	-BackgroundColor	Yellow
				Write-Host ""
				return $null
				
			}
	
			while( $AllSearch.count -ne 1 )
			{
					
					while (!$UserName -or $UserName.length -lt 3)
					{ 
						# Ensure search string is at least 3 chars long
						
						Write-Host ""
						Write-Host  "Enter User Name or '^' to Skip " -ForegroundColor Blue -BackgroundColor Gray -NoNewline
						$UserName = Read-Host " "
						
						if ( $UserName -eq "^" )  
						{ 
							Write-Host ""	
							Write-Host "	" -NoNewline
							Write-Host "  ATTENTION :  Skipped" -ForegroundColor  Blue	-BackgroundColor	Yellow
							Write-Host ""
							
							return $null
						}
					}
			
				$UserName = ($UserName.ToString()).Trim()
				

				
				$UserName = $UserName.Replace("\", "")
	            $UserName = $UserName.Replace("[", "")
	            $UserName = $UserName.Replace("]", "")
	            $UserName = $UserName.Replace(":", "")
	            $UserName = $UserName.Replace(";", "")
	            $UserName = $UserName.Replace("|", "")
	            #$UserName = $UserName.Replace("=", "")
	            #$UserName = $UserName.Replace(",", "")
	            $UserName = $UserName.Replace("+", "")
	            $UserName = $UserName.Replace("*", "")
	            $UserName = $UserName.Replace("<", "")
	            $UserName = $UserName.Replace(">", "")
				
				
	            #$UserName = $UserName.Replace("`'", "")
				
				#$UserName = $UserName.Replace(".", "")
				$UserName = $UserName.Replace("!", "")
				$UserName = $UserName.Replace("^", "")
				
				#replace multiple succeding white spaces
				
				$UserName = $UserName -replace '\s+', ' '
				
				$dashCount = ($UserName.ToCharArray()  | Where-Object {$_ -eq "-" } | Measure-Object).Count
				
				Write-Verbose "Normalized search:  $UserName"
				
				if( $dashCount -eq 4 )
				{
					 try { $guid = [guid] $UserName  }catch{}

					Write-Verbose "`$AllSearch = @(Get-ADUser -Filter `" msExchMailboxGuid -eq $guid -or ObjectGUID -eq $guid -or msExchArchiveGUID -eq $guid `"  -Properties * )"
	
					$AllSearch = @(Get-ADUser -Filter "msExchMailboxGuid -eq '$guid' -or ObjectGUID -eq '$guid'  -or msExchArchiveGUID -eq '$guid' "  -Properties * )

				}
				else
				{
					# Not a guid
					
					$SBStringSimilar = "
					Name -like `"*$UserName*`" 
					-or DisplayName -like `"*$UserName*`"
					-or EmailAddress -like `"*$UserName*`"
					-or proxyAddresses -like `"*$UserName*`"
					-or SamAccountName -like `"*$UserName*`"
					-or DistinguishedName -eq `"$UserName`"
					-or UserPrincipalName -like `"$UserName`"  
					-or Description -like `"*$UserName*`"  
					-or EmployeeNumber -eq `"$UserName`"
					"
	 		
					Write-Verbose "$SBStringSimilar"
					
					Write-Verbose "`$sb = [scriptblock]::create(`$SBStringSimilar)"

					$sb = [scriptblock]::create($SBStringSimilar)
					
					$global:gsb = $sb

					Write-Verbose "`$AllSearch = @(Get-ADUser -Filter `$sb -Properties * ) "
					
					
					$AllSearch = @(Get-ADUser -Filter $sb -Properties * )
				
				}
				
				if ( !$silent -or $($AllSearch.count) -ne 1)
				{
					# https://jackstromberg.com/2013/01/useraccountcontrol-attributeflag-values/
					
					Write-Host ""
					Write-Host ($AllSearch | sort msExchRecipientTypeDetails,userAccountControl , DisplayName  |  ft ObjectGUID, DisplayName,`
					@{Label="msExchRecipientTypeDetails"; Expression= { $RecipientTypeDetailsList.Item( ($_.msExchRecipientTypeDetails).tostring() )  }  } ,`
					userAccountControl,`
					@{Label="Disabled?"; Expression= { if ( $_.userAccountControl -eq 514 -or $_.userAccountControl -eq  546  -or $_.userAccountControl  -eq  66050   -or $_.userAccountControl  -eq  	66082  ){ "True"  }else{"False"}   }  } ,`
					EmailAddress, UserPrincipalName ,  Description   -AutoSize -Wrap  | Out-String).trim()  -foregroundcolor Cyan
					
					Write-Host ""
				
					Write-Host "[$($AllSearch.count)] results found for '$UserName' " -ForegroundColor Cyan	-BackgroundColor	Blue
					Write-Host ""
				}
			
				$UserName = ""
	
		} # while( $AllSearch.count -ne 1 )

	
		$UserName = $AllSearch
			
		
		#region  CustomADUser
			
			$CustomADUser = New-Object $AllSearch –TypeName PSObject 
			
			If (!($CustomADUser.UserPrincipalName))
			{
					Write-Host ""
					Write-Host "	" -NoNewline
					Write-Host "  System FAILURE  CustomADUser: UserPrincipalName is empty ! " -ForegroundColor Yellow -BackgroundColor Red
					Write-Host ""
					
			}

			
			#Give this object a unique typename
			$CustomADUser.PSObject.TypeNames.Insert(0,'ADUser.Information')
			
			# Overloading toString() function to return object guid 
			$CustomADUser | Add-Member  scriptmethod -Name  toString -Value { [string]$this.ObjectGUID }  -Force -PassThru  | Out-Null
			
			# Set standard members 
			$CustomADUser | Add-Member MemberSet PSStandardMembers $PSStandardMembers -Force

			Write-Verbose "$($MyInvocation.InvocationName); Line [$($MyInvocation.ScriptLineNumber)]: $($MyInvocation.line)"
			Write-Verbose "	`$passexp =  (get-aduser $($AllSearch.ObjectGUID)  -Properties msDS-UserPasswordExpiryTimeComputed).'msDS-UserPasswordExpiryTimeComputed' "
			
			$passexp =  (get-aduser $AllSearch.ObjectGUID  -Properties msDS-UserPasswordExpiryTimeComputed)."msDS-UserPasswordExpiryTimeComputed"
			
			$CustomADUser | Add-Member -NotePropertyName "PWExpiration" -NotePropertyValue $passexp -Force | Out-Null
						
			$CustomADUser | Add-Member -NotePropertyName "UserLastLogon" -NotePropertyValue  ( fConvertADDate $CustomADUser.LastLogonTimestamp)  -Force | Out-Null
			
		
			Write-Host "" -ForegroundColor Magenta
			
			$Owner = try{  ( Get-Acl "ad:\$($CustomADUser.distinguishedname)" -ErrorAction SilentlyContinue    ).Owner }catch{} 
			
			if ( $Owner)
			{
				if (  $Owner -like "*Domain Admins*")
				{
					$Creator =  $Owner.Replace("msoit\", "")
				}
				else
				{
					$distinguishedname = $CustomADUser.distinguishedname
								
					$Owner = $Owner.Trim("MSOIT\")
			
					if ( $Owner  )
					{
						$Creator =  try { get-aduser  $Owner  }catch{}
					}
					
				}

			}
			
			$CustomADUser | Add-Member -NotePropertyName "Creator" -NotePropertyValue $Creator  -Force | Out-Null
			

			if ( $CustomADUser.msExchRecipientTypeDetails)
			{
			
				$CustomADUser | Add-Member -NotePropertyName "ADRecipientTypeDetails" -NotePropertyValue ( $RecipientTypeDetailsList.Item( ($CustomADUser.msExchRecipientTypeDetails).tostring() )  )  -Force | Out-Null
				
				# get recipient properties only if msExchRecipientTypeDetails is available 
				$Recipient = fGetADRecipient   $CustomADUser.UserPrincipalName
			}
			else
			{
						
				$CustomADUser | Add-Member -NotePropertyName "ADRecipientTypeDetails" -NotePropertyValue ("Not recipient"  )  -Force | Out-Null
				
				# msExchRecipientTypeDetails  is empty hence set recipient to null
				$Recipient = $null
			
			}
	
			$global:gRecipient = $Recipient
	
			if ( !$Recipient )
			{
						# OnPrem mailboxes in OUs that are not synced with O365 do not appear as MailUser on O365 so we are using the onPrem attribute 
						
						$Recipient  = " $($UserName.ADRecipientTypeDetails)"
			}
					
			if ( $Recipient )
			{
					# Add custom  recipeint properties to CustomADUser 										
					$rootRecipientProperties = $Recipient| Get-Member -ErrorAction SilentlyContinue | ? { $_.MemberType -match "Property"} 
					
					Write-Verbose "$rootRecipientProperties "
					
					$rootRecipientProperties | % {
		
						if ( !$($CustomADUser.$($_.Name)) )
						{
							# Property does not exist in $CustomADUser object yet 
															
							$CustomADUser  | Add-Member –MemberType NoteProperty –Name  "Recipient-$($_.Name)"   –Value  $Recipient.$($_.Name) -ErrorAction SilentlyContinue -Force | Out-Null 
						}
						else
						{
							# Property already exists in $CustomADUser object 
							
							$CustomADUser  | Add-Member –MemberType NoteProperty –Name  "Recipient-$($_.Name)"   –Value  $Recipient.$($_.Name) -ErrorAction SilentlyContinue  -Force | Out-Null
						
						}
					
					}
	
					Write-Verbose " InvocationName: $($MyInvocation.InvocationName); Line Number [$($MyInvocation.ScriptLineNumber)]; Line: $($MyInvocation.line)"
					
					$result = fDisplayADUserAccount   $CustomADUser				
									
									
					if ( $CustomADUser.msExchRecipientTypeDetails )
					{	
						# IS Recipient
														
						$rtd = $RecipientTypeDetailsList.Item( ($CustomADUser.msExchRecipientTypeDetails).tostring() ) 
						
						# Write-Host "927 $rtd"  -ForegroundColor Magenta
						
						if ( $RecipientTypeDetailsList.Item( ($CustomADUser.msExchRecipientTypeDetails).tostring() ) -like "Remote*")
						{
							# Remote Mailbox 
							
							Write-verbose "776 Remote $rtd " 
							
							if ( $EXO )
							{
									Write-Verbose "get-EXOmailbox -ErrorAction silent $($AllSearch.UserPrincipalName ) "
									
									$StopWatch = [System.Diagnostics.Stopwatch]::StartNew()
																									
									$Mailbox = try { get-EXOmailbox -ErrorAction silent $AllSearch.UserPrincipalName -PropertySets all  }catch{}
									
									$StopWatch.Stop()
									
									$ElapsedMilliseconds = $($StopWatch.ElapsedMilliseconds) / 1000
									
									if ( $ElapsedMilliseconds -gt 0 )
									{
									 	Write-host "get-EXOmailbox elapsed: $ElapsedMilliseconds  $($Mailbox.Database) "  -ForegroundColor Magenta
									}
				
									if ( $includeMailboxStats )
									{
										Write-Verbose "`$Mailboxstats = get-EXOmailboxStatistics  -ErrorAction silent $($AllSearch.UserPrincipalName ) "
										$StopWatch = [System.Diagnostics.Stopwatch]::StartNew()
										$Mailboxstats = try {  get-EXOmailboxStatistics -WarningAction silent    -ErrorAction SilentlyContinue $AllSearch.UserPrincipalName -PropertySets all }catch{} 	
										$StopWatch.Stop()
										
										$ElapsedMilliseconds = $($StopWatch.ElapsedMilliseconds) / 1000
										
										if ( $ElapsedMilliseconds -gt 0	 )
										{
											 Write-host "get-EXOmailboxStatistics elapsed: $ElapsedMilliseconds $($Mailboxstats.DatabaseName)"  -ForegroundColor Magenta
										}
									
									}
									
									if ( $includeMailboxStats )
									{
									
										Write-Verbose "get-EXOmailboxStatistics -Archive  -ErrorAction silent $($AllSearch.UserPrincipalName ) " 
										$StopWatch = [System.Diagnostics.Stopwatch]::StartNew()								
										$ArchiveMailboxstats = try { get-EXOmailboxStatistics -Archive  -WarningAction silent   -ErrorAction SilentlyContinue $AllSearch.UserPrincipalName -PropertySets all }catch{} 	
										$StopWatch.Stop()
										
										$ElapsedMilliseconds = $($StopWatch.ElapsedMilliseconds) /1000
										
										if ( $ElapsedMilliseconds -gt 0 )
										{
											 Write-host "get-EXOmailboxStatistics -Archive  elapsed: $ElapsedMilliseconds $($ArchiveMailboxstats.DatabaseName)"  -ForegroundColor Magenta
										}
									 }
									 
									 Write-Host ""
								
							}
							elseif ( $O365)
							{
									Write-Verbose "get-O365mailbox -ErrorAction silent $($AllSearch.UserPrincipalName ) "
									
									$StopWatch = [System.Diagnostics.Stopwatch]::StartNew()
																									
									$Mailbox = try { get-O365mailbox -ErrorAction silent $AllSearch.UserPrincipalName  }catch{}
									
									$StopWatch.Stop()
									
									$ElapsedMilliseconds = $($StopWatch.ElapsedMilliseconds) / 1000
									
									if ( $ElapsedMilliseconds -gt 0 )
									{
									 	Write-host "get-O365mailbox elapsed: $ElapsedMilliseconds  $($Mailbox.Database) "  -ForegroundColor DarkCyan
									}
					
									
									Write-Verbose "`$Mailboxstats = get-O365mailboxStatistics  -ErrorAction silent $($AllSearch.UserPrincipalName ) "

									#Write-Host "1551" -ForegroundColor Magenta
									
								
									#$Mailboxstats = try {  get-O365mailboxStatistics -WarningAction silent    -ErrorAction SilentlyContinue $AllSearch.UserPrincipalName }catch{} 	
									
									if ( $includeMailboxStats )
									{
										$StopWatch = [System.Diagnostics.Stopwatch]::StartNew()
										$Mailboxstats = try {  get-O365mailboxStatistics -WarningAction silent    -ErrorAction SilentlyContinue $AllSearch.UserPrincipalName }catch{} 	
										$StopWatch.Stop()
									
									}
									
									$ElapsedMilliseconds = $($StopWatch.ElapsedMilliseconds) / 1000
									
									if ( $ElapsedMilliseconds -gt 3 )
									{
										 Write-host "get-O365mailboxStatistics elapsed: $ElapsedMilliseconds $($Mailboxstats.Database)"  -ForegroundColor Magenta
									}
									
									Write-Verbose "get-O365mailboxStatistics -Archive  -ErrorAction silent $($AllSearch.UserPrincipalName ) " 
									
									<# 
									$Mo = Measure-Command  { get-O365mailboxStatistics -Archive  -WarningAction silent   -ErrorAction SilentlyContinue $AllSearch.UserPrincipalName }  | select   "get-O365mailboxStatistics  -Archive", Seconds, Milliseconds 

									if ( $Mo.Seconds -gt 0 )
									{
										$Mo | Write-host  -ForegroundColor Magenta
										Write-Host ""
									}
									#>
									
									
									if ( $includeMailboxStats )
									{
										$StopWatch = [System.Diagnostics.Stopwatch]::StartNew()								
																			
										$ArchiveMailboxstats = try { get-O365mailboxStatistics -Archive  -WarningAction silent   -ErrorAction SilentlyContinue $AllSearch.UserPrincipalName }catch{} 	
										
										$StopWatch.Stop()
									
									}
									
									$ElapsedMilliseconds = $($StopWatch.ElapsedMilliseconds) /1000
									
									if ( $ElapsedMilliseconds -gt 3 )
									{
									
									 Write-host "get-O365mailboxStatistics -Archive  elapsed: $ElapsedMilliseconds $($ArchiveMailboxstats.Database)"  -ForegroundColor Magenta
									
									 }
									 
									 
									 
									 Write-Host ""
												
							} # if ( $O365)
							else
							{
									# NO O365 Sessiom
									
									Write-verbose "get-remotemailbox -ErrorAction silent $($AllSearch.UserPrincipalName ) " 
									
									$Mailbox = try { get-remotemailbox -ErrorAction silent $AllSearch.UserPrincipalName  }catch{}
							}
		
					} # if ( $RecipientTypeDetailsList.Item( ($CustomADUser.msExchRecipientTypeDetails).tostring() ) -like "Remote*")
					elseif ($CustomADUser.msExchRecipientTypeDetails -eq 128 )
					{
							# Mail-enabled User, i.e. Mailbox is only on O365
							
							#Write-Host "$($CustomADUser.msExchRecipientTypeDetails) " -ForegroundColor Magenta 
						
							if ( $O365)
							{
									Write-Verbose "get-O365mailbox -ErrorAction silent $($AllSearch.UserPrincipalName ) "
									
									#$Mo = Measure-Command { get-O365mailbox -ErrorAction silent $AllSearch.UserPrincipalName  } | select  "get-O365mailbox",  Seconds, Milliseconds 
									
									if ( $Mo.Seconds -gt 0 )
									{
										$Mo | Write-host  -ForegroundColor Magenta
										Write-Host ""
									}
									
									$StopWatch = [System.Diagnostics.Stopwatch]::StartNew()
																									
									$Mailbox = try { get-O365mailbox -ErrorAction silent $AllSearch.UserPrincipalName  }catch{}
									
									$StopWatch.Stop()
									
									$ElapsedMilliseconds = $($StopWatch.ElapsedMilliseconds) / 1000
									
									 Write-host "get-O365mailbox elapsed: $ElapsedMilliseconds  $($Mailbox.Database) "  -ForegroundColor Magenta
					
									Write-Verbose "`$Mailboxstats = get-O365mailboxStatistics  -ErrorAction silent $($AllSearch.UserPrincipalName ) "
									
									#$Mo = Measure-Command {  get-O365mailboxStatistics  -ErrorAction silent   $AllSearch.UserPrincipalName }  | select "get-O365mailboxStatistics "  ,Seconds, Milliseconds 
																	
																
									if ( $Mo.Seconds -gt 0 )
									{
										$Mo | Write-host  -ForegroundColor Magenta
										Write-Host ""
									}
									
									#Write-Host "1605" -ForegroundColor Magenta
									
									$StopWatch = [System.Diagnostics.Stopwatch]::StartNew()
								
									#$Mailboxstats = try {  get-O365mailboxStatistics -WarningAction silent    -ErrorAction SilentlyContinue $AllSearch.UserPrincipalName }catch{} 	
									
									$StopWatch.Stop()
									
									$ElapsedMilliseconds = $($StopWatch.ElapsedMilliseconds) / 1000
									
									Write-host "get-O365mailboxStatistics elapsed: $ElapsedMilliseconds $($Mailboxstats.Database)"  -ForegroundColor Magenta
							
									Write-Verbose "get-O365mailboxStatistics -Archive  -ErrorAction silent $($AllSearch.UserPrincipalName ) " 
									
									#$Mo = Measure-Command  { get-O365mailboxStatistics -Archive  -WarningAction silent   -ErrorAction SilentlyContinue $AllSearch.UserPrincipalName }  | select   "get-O365mailboxStatistics  -Archive", Seconds, Milliseconds 

									if ( $Mo.Seconds -gt 0 )
									{
										$Mo | Write-host  -ForegroundColor Magenta
										Write-Host ""
									}
									
									
									if ( $includeMailboxStats )
									{
									
										$StopWatch = [System.Diagnostics.Stopwatch]::StartNew()								
																		
										$ArchiveMailboxstats = try { get-O365mailboxStatistics -Archive  -WarningAction silent   -ErrorAction SilentlyContinue $AllSearch.UserPrincipalName }catch{} 	
										
										$StopWatch.Stop()
																			
										$ElapsedMilliseconds = $($StopWatch.ElapsedMilliseconds) /1000
										
										 Write-host "get-O365mailboxStatistics -Archive  elapsed: $ElapsedMilliseconds $($ArchiveMailboxstats.Database)"  -ForegroundColor Magenta
										 
										 Write-Host ""
									 
									 }
												
							} # if ( $O365)
					
					}
					else
					{
							# NOT a remote recipient 
							
							Write-verbose  "`$Mailbox =  get-mailbox -ErrorAction silent $($AllSearch.UserPrincipalName) " 
							$Mailbox = try { get-mailbox -ErrorAction silent $AllSearch.UserPrincipalName  }catch{}
							
							if ( $includeMailboxStats )
							{
								Write-verbose "1627 `$Mailboxstats = get-mailboxStatistics  -ErrorAction silent $($AllSearch.UserPrincipalName ) " 
								$Mailboxstats = try { get-mailboxStatistics  -ErrorAction SilentlyContinue $AllSearch.UserPrincipalName }catch{} 	
													
								Write-verbose "`$ArchiveMailboxstats = get-mailboxStatistics -Archive  -ErrorAction silent $($AllSearch.UserPrincipalName ) " 
								$ArchiveMailboxstats = try { get-mailboxStatistics -Archive  -ErrorAction SilentlyContinue $AllSearch.UserPrincipalName }catch{} 	
						
							}
					}
						
					}# if ( $CustomADUser.msExchRecipientTypeDetails )
			
					if ( $Mailbox)
					{
												
						$rootMailboxProperties = $Mailbox | Get-Member -ErrorAction SilentlyContinue | ? { $_.MemberType -match "Property"} 
						
						Write-verbose  "rootMailboxProperties  count : $($rootMailboxProperties.count)" 
					
						$rootMailboxProperties | 
						% {
						
								if ( !$($CustomADUser.$($_.Name)) )
								{
										# Mailbox property does not exist in $CustomADUser
										
										$CustomADUser  | Add-Member –MemberType NoteProperty –Name  $($_.Name)   –Value  $Mailbox.$($_.Name) -ErrorAction SilentlyContinue -Force | Out-Null 
										
								}
								else
								{
										# Mailbox property already present  in $CustomADUser. Adding it with prefix 'Mailbox-'
										
										$CustomADUser  | Add-Member –MemberType NoteProperty –Name  "Mailbox-$($_.Name)"   –Value  $Mailbox.$($_.Name) -ErrorAction SilentlyContinue -Force | Out-Null
								
								}
						
							}
					
						} # if ( $Mailbox)
						
					if ( $Mailboxstats)
					{
													
							$rootMailboxStatsProperties = $Mailboxstats | Get-Member -ErrorAction SilentlyContinue | ? { $_.MemberType -match "Property"} 
							
							Write-verbose  "Mialbox stats  properties count : $($rootMailboxStatsProperties.count)" 
							
							$rootMailboxStatsProperties | % {
							
									
									if ( !$($CustomADUser.$($_.Name)) )
									{
											# Mailbox Stats property does not exist in $CustomADUser
											
											$CustomADUser  | Add-Member –MemberType NoteProperty –Name  $($_.Name)   –Value   $Mailboxstats.$($_.Name) -ErrorAction SilentlyContinue -Force | Out-Null 
															
									}
									else
									{
											# Mailbox Stats property already present in $CustomADUser
											
											$CustomADUser  | Add-Member –MemberType NoteProperty –Name  "MailboxStats-$($_.Name)"   –Value   $Mailboxstats.$($_.Name) -ErrorAction SilentlyContinue -Force | Out-Null
			
									}
							
								}
						
							}# if ( $Mailboxstats)
								
						if ( $ArchiveMailboxstats )
						{
													
							$rootArchiveMailboxStatsProperties = $ArchiveMailboxstats  | Get-Member -ErrorAction SilentlyContinue | ? { $_.MemberType -match "Property"} 
							
							Write-verbose  "Archive Mailbox stats  properties count : $($rootArchiveMailboxStatsProperties.count)" 

							$rootArchiveMailboxStatsProperties | % {
							
									
									$CustomADUser  | Add-Member –MemberType NoteProperty –Name  "ArchiveMailboxStats-$($_.Name)"   –Value  $ArchiveMailboxstats.$($_.Name) -ErrorAction SilentlyContinue -Force | Out-Null
									
									
									#Write-Host " `$CustomADUser  | Add-Member –MemberType NoteProperty –Name  'ArchiveMailboxStats-$($_.Name)'   –Value  $($ArchiveMailboxstats.$($_.Name)) -ErrorAction SilentlyContinue -Force | Out-Null " -ForegroundColor Magenta
									
									
									<#
									
									if ( !$($CustomADUser.$($_.Name)) )
									{
											# if $CustomADUser does not have a property that have the same name as the current property in $mailbox properties    $($_.Name) 
																			
											$CustomADUser  | Add-Member –MemberType NoteProperty –Name  $($_.Name)   –Value   $ArchiveMailboxstats.$($_.Name) -ErrorAction SilentlyContinue -Force | Out-Null
																			
									}
									else
									{
										# Add duplicate properties with prepended name 
							
										$CustomADUser  | Add-Member –MemberType NoteProperty –Name  "ArchiveMailboxStats-$($_.Name)"   –Value  $ArchiveMailboxstats.$($_.Name) -ErrorAction SilentlyContinue -Force | Out-Null
									
									
									}
									#>
							
								}
						
							}	
					
												
						#$remotemailbox = get-remotemailbox $UserName.UserPrincipalName	  -ErrorAction silent
					
						$remotemailbox = try {get-remotemailbox $UserName.UserPrincipalName	  -ErrorAction silent}catch{}
						
						$CustomADUser 	| Add-Member  -NotePropertyName "RemoteRoutingAddress" -NotePropertyValue $($remotemailbox.RemoteRoutingAddress)  -Force | Out-Null
						
						#$accountdisbaled = if ($CustomADUser.userAccountControl -eq 512 ){ "False"  }else{"True"} 
						$accountdisbaled = if ($CustomADUser.userAccountControl -eq 512 ){ $false }else{ $true } 
						
						Write-Verbose "accountdisbaled $accountdisbaled"
						
						$CustomADUser 	| Add-Member  -NotePropertyName "ADAccountDisabled" -NotePropertyValue $accountdisbaled   -Force	| Out-Null
						
									
			} # if ( $Recipient )
				
		#endregion  CustomADUser
			
		$result = fDisplayADUser $CustomADUser
		
		if($return ) 	
		{ 
			return $CustomADUser  
		}

	} ########### END function fSearchADUser ($UserName, $return )

	function fGetADRecipient ($UserName)
	{		
		Write-Verbose " InvocationName: $($MyInvocation.InvocationName); Line Number [$($MyInvocation.ScriptLineNumber)]; Line: $($MyInvocation.line)"
		
		if ( !$UserName ) 	{ 	
		
				Write-Host ""
				Write-Host "	" -NoNewline
				Write-Host "  System FAILURE  fGetADRecipient: No Value for input variable UserName provided  " -ForegroundColor Yellow -BackgroundColor Red
				Write-Host ""
				
				return  
		
		}
	
		$recipient = $null 
	
		if ( $UserName)
		{
			If ( $O365 )
			{
				$recipient = try { get-O365recipient -ErrorAction silent $UserName }catch{}
			}
			else
			{
				$recipient = try { get-recipient -ErrorAction silent $UserName }catch{}
			}
				
		}	

		Write-Verbose "fGetADRecipient	Alias $($recipient.alias) "
		
		
		return $recipient 

	} ########### END function fGetADRecipient

	function fGetADMailbox ($UserName)
	{
		Write-Verbose " InvocationName: $($MyInvocation.InvocationName); Line Number [$($MyInvocation.ScriptLineNumber)]; Line: $($MyInvocation.line)"
		
		if ( !$UserName ) 	{ 	
	
			Write-Host ""
			Write-Host "	" -NoNewline
			Write-Host "  System FAILURE  fGetADMailbox: No Value for input variable UserName provided  " -ForegroundColor Yellow -BackgroundColor Red
			Write-Host ""
			
			return  
	
		}

		$mailbox = $null
		
		Write-Host " 315 get-mailbox -ErrorAction silent $UserName " -ForegroundColor Red 
		
		$mailbox = try { get-mailbox -ErrorAction silent $UserName }catch{}
		
		Write-Verbose " $mailbox " 
		
		return $mailbox

	} ########### END function fGetADMailbox

	function fNormalizeUsername ( $UserLogonName )
	{
		# removes all chars but small and cap letters as well as ' . ' and ' - '
		
		Write-Verbose " InvocationName: $($MyInvocation.InvocationName); Line Number [$($MyInvocation.ScriptLineNumber)]; Line: $($MyInvocation.line)"
		
		if ( !$UserLogonName) 	
		{ 
			Write-Host ""
			Write-Host "	" -NoNewline
			Write-Host "[usra fNormalizeUsername] System FAILURE  : No Value for input variable  UserLogonName  provided  " -ForegroundColor Yellow -BackgroundColor Red
			Write-Host ""
			return 	
		}

		#$pattern = '[^a-zA-Z0-9.-]'
		#$pattern = '[^a-zA-Z0-9.-_]'
		#$pattern = '[^a-zA-Z0-9._\-]'
		
		#$pattern = '[^a-zA-Z0-9._\-]'
		
		# https://serverfault.com/questions/604547/rules-for-active-directory-user-name-string/968960
		
		 $pattern = "[ ] : ; | = + * ? < > / \ ,"
		
		$normalized = $UserLogonName -replace $pattern, ''
		
		$normalized = $normalized.trim()
		
		#$normalized | Write-Host -ForegroundColor Red
		
		return $normalized 


	} ########### END function fNormalizeUsername

	function fPasswordStatus ($UserParameter)
	{
		
		Write-Verbose " InvocationName: $($MyInvocation.InvocationName); Line Number [$($MyInvocation.ScriptLineNumber)]; Line: $($MyInvocation.line)"
		
		if ( !$UserParameter ) 	{ 	exit  }
		
		$DaysToAlert = 10
		
		#$UserParameter | fl 
		
		$PasswordExpired = $false
		
		$global:gUserParameter = $UserParameter
		
		
		if ($UserParameter)
		{
			if ( $UserParameter.PasswordNeverExpires  )
			{
				
				if (   $UserParameter.RecipientTypeDetails -notlike "*Shared*" )
				{
					 $UserParameter | ft *Details* 
					 
					Write-Host  "Password Never Expires" -ForegroundColor Red	-BackgroundColor	yellow	-NoNewline
					Write-Host " " -NoNewline
				}
				
			}
			elseif ($UserParameter.PWExpiration)
			{
					$PasswordExpires =   Get-Date  ( [datetime]::FromFileTime($UserParameter.PWExpiration))  -format "dd/MMM/yyyy HH:mm"
					
					if ( $UserParameter.PasswordLastSet )
					{
						$PasswordLastSet = Get-Date $UserParameter.PasswordLastSet  -format "dd/MMM/yyyy HH:mm"
					}
					
					#Write-Host  "$($UserParameter.PasswordLastSet)"  -ForegroundColor DarkCyan -NoNewline
					
					Write-Host  "$PasswordLastSet"  -ForegroundColor DarkCyan -NoNewline
					
					Write-Host " " -NoNewline
					
					if ( $PasswordExpires )
					{
						$DaysToExpire = ( New-TimeSpan -end $PasswordExpires ).Days
						
						if (  $DaysToExpire -lt  0 )
						{
							#Write-Host  "Password Expired on $PasswordExpires"   -ForegroundColor Red	-BackgroundColor	yellow	 -NoNewline	
							Write-Host  "Password Expired on $PasswordExpires"    -ForegroundColor Red	-BackgroundColor	yellow		 -NoNewline	
							
							$PasswordExpired =  $true
						}
						elseif (  $DaysToExpire -lt  $DaysToAlert  )
						{
							Write-Host  "Password Expires on $PasswordExpires "  -ForegroundColor Yellow -NoNewline
						}
													
				 	}
						
			} # elseif ($UserParameter.PWExpiration)
		
		} #if ($UserParameter)
		
		#Write-Host " $PasswordExpired " -ForegroundColor Red
		
		return $PasswordExpired

	} # fPasswordStatus ($UserParameter)

	function fConvertADDate ([long] $ticks) 
	{
	    # https://msdn.microsoft.com/en-us/library/ms675098(v=vs.85).aspx
		
		Write-Verbose " InvocationName: $($MyInvocation.InvocationName); Line Number [$($MyInvocation.ScriptLineNumber)]; Line: $($MyInvocation.line)"
	 
	 
	    if ( ($ticks -eq 0) -or ($ticks -eq 9223372036854775807) ) {
	        $expires = $null 
	    }
	    else {
	        $expires = [DateTime]::FromFileTime($ticks) 
	    }
	 
	 	#Write-Host "expires $expires" -ForegroundColor Magenta
	 	$global:gexpires = $expires
	 
	    if ( $expires )
		{
			$expires =  get-date $expires -format "dd/MMM/yyyy HH:mm"
		}
		
		write-output $expires 
		

	} ########### END function fConvertADDate 

	function fAccountExpires ($UserParameter, $silent )
	{
		
			
		Write-Verbose " InvocationName: $($MyInvocation.InvocationName); Line Number [$($MyInvocation.ScriptLineNumber)]; Line: $($MyInvocation.line)"
		
		if ( !$UserParameter ) 	{ 	exit  }
		
		if ($UserParameter)
		{
			#Write-Host "800" -ForegroundColor Magenta 
			
			if ( $UserParameter.AccountExpires)
			{
				#adchange
				
				$AccountExpires  = fConvertADDate $($UserParameter.AccountExpires)
			}
			
			
			if ( $AccountExpires )
			{
			
				Write-Verbose " $($MyInvocation.InvocationName); Line [$($MyInvocation.ScriptLineNumber)]: $($MyInvocation.line); Account Expires ? $AccountExpires "
				
				$DaysToExpire = ( New-TimeSpan -end $AccountExpires ).Days
				
				$HoursToExpire = ( New-TimeSpan -end $AccountExpires ).Hours
				
				Write-Verbose "  $($MyInvocation.InvocationName); Line [$($MyInvocation.ScriptLineNumber)]: $($MyInvocation.line); Days to expire : $DaysToExpire "
				
				#write-host "HoursToExpire: $HoursToExpire $DaysToExpire" -ForegroundColor Magenta
								
				if (  $DaysToExpire -le  0 -and $HoursToExpire -lt 0 )
				{
					
					if ( !$silent)
					{
						Write-Host  "Account Expired on $AccountExpires "   -ForegroundColor Red	-BackgroundColor	yellow -NoNewline
					}
				
					$accountExpired = $true
					
				}
				else
				{
					if ( !$silent)
					{
						Write-Host  "Account Expires on $AccountExpires "  -ForegroundColor Yellow -NoNewline
					}
					
					$accountExpired = $false
					
				}
			
		 	}
		
		}
	
		return $accountExpired

	}########### END function AccountExpires ($UserParameter)) ###########

	function fMailboxForward( $UserParameter )
	{ 
		Write-Verbose " InvocationName: $($MyInvocation.InvocationName); Line Number [$($MyInvocation.ScriptLineNumber)]; Line: $($MyInvocation.line)"
				
		
		#Write-Host "fMailboxForward" -ForegroundColor Magenta
		
		if (!$UserParameter ) { return;  }
		
		#Write-Host ( $UserName  | fl *forward* | Out-String).Trim() -foregroundcolor Red

		
		if ($UserParameter.ForwardingAddress -or $UserParameter.ForwardingSmtpAddress)
		{
			Write-Host ""
			Write-Host ($UserParameter | fl  ForwardingAddress, ForwardingSmtpAddress, DeliverToMailboxAndForward    | Out-String).trim()  -ForegroundColor  White	-BackgroundColo   DarkCyan
			Write-Host ""
		}
	} ########### END function fMailboxForward( $UserParameter )

	function fMailboxRights ()
	{
		param(
	    [parameter(Position=0,Mandatory=$false,ValueFromPipeline=$false,HelpMessage="mailbox ")]
		$UserParameter,
		[parameter(Position=1,Mandatory=$false,ValueFromPipeline=$false,HelpMessage="test ")]
		[switch]$silent
		)
		
		Write-Verbose " InvocationName: $($MyInvocation.InvocationName); Line Number [$($MyInvocation.ScriptLineNumber)]; Line: $($MyInvocation.line)"
			
		if (!$UserParameter ) { 	return  }
				
				$MailboxIdentity = $UserParameter.PrimarySmtpAddress
				
				if ( $MailboxIdentity )
				{
						if ( $EXO )
						{
							# Existing EXO session
							
							Write-Host "" 
							Write-Host  "[uADo] EXO FULL ACCESS  permissions to Mailbox: '$($UserParameter.Name) ' "  -ForegroundColor Cyan	-BackgroundColor	Blue
						
								
							Write-Host "" 
							Write-Host "`$FullAccessEXO = Get-EXOMailboxPermission '$MailboxIdentity' | where{ `$_.IsInherited -eq `$False  -and `$_.AccessRights -like '*FullAccess*'  } | select User, AccessRights" -ForegroundColor Yellow
							
							$FullAccessEXO = Get-EXOMailboxPermission "$MailboxIdentity" -ErrorAction silent | where{$_.IsInherited -eq $False -and $_.AccessRights -like "*FullAccess*"} | select AccessRights, User

							Write-Host ( $FullAccessEXO | ft User, AccessRights  -AutoSize | Out-String) -foregroundcolor Cyan
							
							#$SendAsEXO  =  Get-EXORecipientPermission   $MailboxIdentity  -AccessRights SendAs   -erroraction silent 
							
							$SendAsEXO  =  Get-EXORecipientPermission   $MailboxIdentity    -erroraction silent 
							
							#Write-Host ""
							Write-Host "[uADo] EXO SEND AS   permissions to Mailbox: '$($UserParameter.Name)' "  -ForegroundColor Cyan	-BackgroundColor	Blue
							Write-Host ""
							
							Write-Host "`$SendAsEXO  = get-EXORecipientPermission   $MailboxIdentity  -AccessRights SendAs   -erroraction silent  " -ForegroundColor Yellow 
							
							Write-Host ""
							Write-Host ( $SendAsEXO | ft  Identity,Trustee,  AccessControlType, AccessRights,   IsInherited -AutoSize | Out-String).Trim() -foregroundcolor Cyan
							Write-Host ""
						
						
						}
						elseif (  $O365  )
						{
							# Existing O365 session

							Write-Host "" 
							Write-Host  "[uADo] O365 FULL ACCESS  permissions to Mailbox: '$($UserParameter.Name) ' "  -ForegroundColor Cyan	-BackgroundColor	Blue
						
								
							Write-Host "" 
							Write-Host "`$FullAccessO365 = Get-O365MailboxPermission '$MailboxIdentity' | where{ `$_.IsInherited -eq `$False  -and `$_.AccessRights -like '*FullAccess*'  } | select User, AccessRights" -ForegroundColor Yellow
							
							$FullAccessO365 = Get-O365MailboxPermission "$MailboxIdentity" -ErrorAction silent | where{$_.IsInherited -eq $False -and $_.AccessRights -like "*FullAccess*"} | select AccessRights, User

							Write-Host ( $FullAccessO365 | ft User, AccessRights  -AutoSize | Out-String) -foregroundcolor Cyan
											
							$SendAsO365  =  get-O365RecipientPermission   $MailboxIdentity  -AccessRights SendAs   -erroraction silent 
							
							#Write-Host ""
							Write-Host "[uADo] O365 SEND AS   permissions to Mailbox: '$($UserParameter.Name)' "  -ForegroundColor Cyan	-BackgroundColor	Blue
							Write-Host ""
							
							Write-Host "`$SendAsO365  = get-O365RecipientPermission   $MailboxIdentity  -AccessRights SendAs   -erroraction silent  " -ForegroundColor Yellow 
							
							Write-Host ""
							Write-Host ( $SendAsO365| ft  Identity,Trustee,  AccessControlType, AccessRights,   IsInherited -AutoSize | Out-String).Trim() -foregroundcolor Cyan
							Write-Host ""
							
							$GrantSendOnBehalfTo = $UserParameter.GrantSendOnBehalfTo | sort

							if ($GrantSendOnBehalfTo  -and !$silent ) 
							{
								Write-Host "[UO] O365 SEND ON BEHALF permissions to Mailbox: '$($UserParameter.Name)' " -ForegroundColor Cyan	-BackgroundColor	Blue
								Write-Host ""
								Write-Host "`$GrantSendOnBehalfTo = (Get-O365mailbox $MailboxIdentity ).GrantSendOnBehalfTo"  -ForegroundColor Yellow
								Write-Host "`$GrantSendOnBehalfTo | %{ Write-Host  '`$(`$_.Name)'  }" -ForegroundColor Yellow
								Write-Host ""


								$GrantSendOnBehalfTo | %{ 
								
								if( $_.Name )
								{	Write-Host  "$($_.Name)" -ForegroundColor Cyan }
								else
								{	Write-Host  $_ -ForegroundColor Cyan }
								
								}

							}
					
						}
						elseif ($RecipientTypeDetails -like "Remote*")
						{
							# No O365 session
								
							if ( !$silent )
							{
								Write-Host "" 
								Write-Host  "[uADo]   FULL ACCESS  permissions to Mailbox: '$($UserParameter.Name) ' "  -ForegroundColor Cyan	-BackgroundColor	Blue
							}	
							
							Write-Host ""
							
							$FullAccess = " Remote Mailbox & No O365 Session. To get O365 Full Access rights create O365 Session "
							
							Write-Host ( $FullAccess | ft -AutoSize | Out-String) -foregroundcolor Cyan
							
							Write-Host "" 
							Write-Host "`$SendAs = Get-QADPermission $($UserParameter.PrimarySMTPAddress.ToString()) -WarningAction silentlycontinue  | ?{ `$_.RightsDisplay -eq  'Send As' }" -ForegroundColor Yellow 
							
							
							$SendAs = Get-QADPermission $UserParameter.PrimarySMTPAddress.ToString() -WarningAction silentlycontinue  | ?{ $_.RightsDisplay -eq  "Send As" }
							
							Write-Host ""
							Write-Host "[uADo] OnPrem (no O365 Session) SEND AS permissions to Mailbox: '$Mailbox ' "  -ForegroundColor Cyan	-BackgroundColor	Blue
							Write-Host ""
							
							Write-Host ""
							Write-Host ( $SendAs  |select RightsDisplay, Account,   Source |  ft       -AutoSize | Out-String).Trim() -foregroundcolor Cyan
							Write-Host ""
						
						}
						else
						{
								# OnPrem mailbox 
								
								if ( !$silent )
								{
									Write-Host "" 
										Write-Host  "[uADo] OnPrem  FULL ACCESS  permissions to OnPrem Mailbox: '$($UserParameter.Name) ' "  -ForegroundColor Cyan	-BackgroundColor	Blue
								}
								
								
								Write-Host "" 
								Write-Host "`$FullAccess = Get-MailboxPermission '$MailboxIdentity' -ErrorAction silent  |where{ `$_.IsInherited -eq `$False -and `$_.AccessRights -like '*FullAccess*'  } | select User, AccessRights | Out-host" -ForegroundColor Yellow
								#Write-Host ""
								
								$FullAccess = Get-MailboxPermission  "$MailboxIdentity"   -ErrorAction silent | where{$_.IsInherited -eq $False -and $_.AccessRights -like "*FullAccess*"} | select User,AccessRights

								Write-Host ( $FullAccess | ft AccessRights, User  -AutoSize | Out-String) -foregroundcolor Cyan	
								
								Write-Host "`$SendAs = Get-QADPermission $($UserParameter.PrimarySMTPAddress.ToString()) -WarningAction silentlycontinue  | ?{ `$_.RightsDisplay -eq  'Send As' }" -ForegroundColor Yellow 
								Write-Host "" 
								
								$SendAs = Get-QADPermission $UserParameter.PrimarySMTPAddress.ToString() -WarningAction silentlycontinue  | ?{ $_.RightsDisplay -eq  "Send As" }
								
								Write-Host ""
								Write-Host "[uADo] OnPrem SEND AS   permissions to Mailbox: '$Mailbox ' "  -ForegroundColor Cyan	-BackgroundColor	Blue
								Write-Host ""
								
								Write-Host ""
								Write-Host ( $SendAs  |select RightsDisplay, Account,   Source |  ft       -AutoSize | Out-String).Trim() -foregroundcolor Cyan
								Write-Host ""
									
						}
					
				}# 	if ( $MailboxIdentity )
	
			
	} ########### END function fMailboxRights 
	
	function fOutOfOffice($UserParameter )
	{ 
			Write-Verbose " InvocationName: $($MyInvocation.InvocationName); Line Number [$($MyInvocation.ScriptLineNumber)]; Line: $($MyInvocation.line)"
			
			if ( !$UserParameter ) 	{ 	return "No UserParameter provided to fOutOfOffice"  }
		
			$CanonicalName = $UserParameter.CanonicalName
			

			if( $UserParameter.RecipientTypeDetails )
			{
					$State = try { Get-O365MailboxAutoReplyConfiguration  $UserParameter.UserPrincipalName  -ErrorAction SilentlyContinue}catch{}
					
					if ( !$State )
					{
						$State = Get-MailboxAutoReplyConfiguration  $CanonicalName  -ErrorAction SilentlyContinue
					}
					
					if ( $State -and   $State.AutoReplyState -ne "Disabled" )
					{
							Write-Host ""
							Write-Host "Out of Office: "  -foregroundcolor Cyan -BackgroundColor BLUE -NoNewline

							if($($State.AutoReplyState) -eq "Enabled") { Write-Host " ENABLED "		-ForegroundColor DarkBlue -BackgroundColor Green   -NoNewline }

							elseif ($($State.AutoReplyState) -eq "Scheduled") 	
							{ 
									$StartTime = Get-Date $($State.StartTime)  -format "dd/MM/yyyy HH:mm"
									$EndTime  = Get-Date $($State.EndTime)   -format "dd/MM/yyyy HH:mm"
									Write-Host " SCHEDULED  Start Time: $StartTime ; End Time: $EndTime " -foregroundcolor White -BackgroundColor DarkGreen
							}
							
							$InternalText =  fHtml-ToText $($State.InternalMessage)
							$ExternalText = fHtml-ToText $($State.ExternalMessage)

							#$global:gInternalText = $InternalText
							
							Write-Host ""	
							Write-Host ""
							Write-Host "Internal Message:  $InternalText" -ForegroundColor Green
							Write-Host ""	
							
							Write-Host "External Message: $ExternalText" -ForegroundColor Green
							
							
						
					}
					

					
					
					Write-Host ""	

			}

	} ########### END function fOutOfOffice($UserParameter, $MailboxParameter) ###########

	function fAcceptMessagesOnlyFrom($UserParameter)
	{
			Write-Verbose " InvocationName: $($MyInvocation.InvocationName); Line Number [$($MyInvocation.ScriptLineNumber)]; Line: $($MyInvocation.line)"
			
			$AcceptMessagesOnlyFromSendersOrMembers = $UserParameter.AcceptMessagesOnlyFromSendersOrMembers 
			$RequireSenderAuthenticationEnabled =   $UserParameter.RequireSenderAuthenticationEnabled
			
			
			if ( $AcceptMessagesOnlyFromSendersOrMembers  )
			{ 	Write-Host "Accepts Messages Only from :" -ForegroundColor Cyan -BackgroundColor Blue }
			
			
			if( $RequireSenderAuthenticationEnabled )
			{
				Write-Host ""
				Write-Host " " -NoNewline
				Write-Host "  WARNING:  External mailflow disabled ! " -ForegroundColor  Blue	-BackgroundColor	Yellow
			
			}
			
			
			if ($AcceptMessagesOnlyFromSendersOrMembers )
			{
				Write-Host ""
				Write-Host ($AcceptMessagesOnlyFromSendersOrMembers | sort Name | ft Name, Parent -AutoSize | Out-String) -foregroundcolor Cyan
				Write-Host ""
			}
			elseif (!$RequireSenderAuthenticationEnabled )
			{
				
			}
	

	}  ########### END function fAcceptMessagesOnlyFrom($UserParameter)

	function fMemberof( $UserParameter )
	{
		Write-Verbose " InvocationName: $($MyInvocation.InvocationName); Line Number [$($MyInvocation.ScriptLineNumber)]; Line: $($MyInvocation.line)"
		
		$Membership = @( Get-ADUser   $UserParameter.ToString() -Properties MemberOf | Select-Object -ExpandProperty MemberOf | Get-ADGroup  -Properties Description -Erroraction silent   )
		
		if ( $Membership )
		{

			$Length = $Membership.Length

			$Membership = $Membership  | select Name, GroupCategory,`
			@{  Label="Email" ; Expression= { (get-adgroup  $_.ObjectGUID  -Properties mail).mail  }}, 	Description
			
			if ( !$silent)
			{
				
				Write-Host ""
				Write-Host  " `$Membership = @( Get-ADUser  $($UserParameter.ToString()) -Properties MemberOf | Select-Object -ExpandProperty MemberOf | Get-ADGroup -Properties Description ) " -ForegroundColor Yellow
				Write-Host ""

				Write-Host "[ $Length ]  groups found " -ForegroundColor Cyan
							
				Write-Host ( $Membership  | sort GroupCategory, Name | ft -autosize | Out-String) -foregroundcolor Cyan
				
			}
		}
	
	} ###### END function fMemberof($UserParameter)
	
	function flicencegroup( $UserParameter )
	{
		Write-Verbose " InvocationName: $($MyInvocation.InvocationName); Line Number [$($MyInvocation.ScriptLineNumber)]; Line: $($MyInvocation.line)"
		
		Write-host  ""
		Write-host  "`$Membership = @( Get-ADUser  $($UserParameter.ToString()) -Properties MemberOf | Select-Object -ExpandProperty MemberOf | Get-ADGroup  -Properties Description)" -ForegroundColor Yellow 
		Write-host  ""
		
		
		$Membership = @( Get-ADUser   $UserParameter.ToString() -Properties MemberOf | Select-Object -ExpandProperty MemberOf | Get-ADGroup  -Properties Description  -ErrorAction SilentlyContinue  ) 
		
		Write-host  "`$Membership =  `$Membership | ?{ `$_.Name -like 'O365_E*' } " -ForegroundColor Yellow 
		Write-host  ""
		
		$Membership =  @( $Membership | ?{ $_.Name -like 'O365_E*' -or $_.Name -like 'O365_MG*'  } )
		
		if ( $Membership )
		{

				$Length = $Membership.count
			
				$Membership = $Membership  | select Name, `
				GroupCategory,`
				@{  Label="Email" ; Expression= { "  Email: " +  (get-adgroup  $_.ObjectGUID  -Properties mail).mail  }}, `
				@{  Label="Description" ; Expression= { "  '$($_.Description)' " } }
				

				Write-Host "[usra] [$Length] licence  group(s) found for '$($UserParameter.Name)' "  -ForegroundColor Cyan	-BackgroundColor	Blue
				
				Write-Host ""
							
				Write-Host ( $Membership  | sort GroupCategory, Name | ft -autosize -HideTableHeaders | Out-String).trim()  -ForegroundColor DarkBlue -BackgroundColor White
				#Write-Host ( $Membership  | sort GroupCategory, Name | fl | Out-String).trim()  -ForegroundColor DarkBlue -BackgroundColor White
				Write-Host ""
			
		}
		else
		{
		
			Write-Host ""
			Write-Host "	" -NoNewline
			Write-Host "  ATTENTION :  No licence groups' membership  found for  ' $($UserParameter.Name) '" -ForegroundColor  Blue	-BackgroundColor	Yellow
			Write-Host ""
		}
	
	
	} ###### END function fMemberof($UserParameter)
		
	function fInboxRules( $UserParameter)
	{ 			
		Write-Verbose " InvocationName: $($MyInvocation.InvocationName); Line Number [$($MyInvocation.ScriptLineNumber)]; Line: $($MyInvocation.line)"

		if($UserParameter)
		{
		
			$InboxRules = @(Get-InboxRule -Mailbox $UserParameter.UserPrincipalName -ErrorAction SilentlyContinue )
	
			if ( !$InboxRules )
			{
				$InboxRules = @( try {Get-O365InboxRule -Mailbox $UserParameter.UserPrincipalName -ErrorAction SilentlyContinue}catch{} )
			
			}
			
			Write-Host ""

			if ( $InboxRules )
			{
				
				$Length = $InboxRules.Length
				$rulenumber = 0 
				
				Write-Host "Inbox Rules:" -foregroundcolor Cyan -BackgroundColor BLUE 
				
				$InboxRules | %{
				
				$rulenumber++
				
				Write-Host ""
				Write-Host "Rule [$rulenumber/$Length] "  -ForegroundColor Cyan
				
				Write-Host ($_ | ft Priority, Name, Enabled, InError, StopProcessingRules, Identity, ExceptIfSubjectContainsWords -AutoSize -Wrap | Out-String) -foregroundcolor Cyan

				Write-Host ($_ | ft Name, SupportedByTask, Description -AutoSize -Wrap | Out-String) -foregroundcolor Cyan
				
				}
			}
			else
			{
				Write-host "WARNING: No Inbox rules !" -BackgroundColor Yellow  -ForegroundColor Black
				Write-Host ""
			}
		
		}


} ########### END  finboxRules($UserParameter) ###########
	
	function fAutomap ($UserParameter )
	{ 
		Write-Verbose " InvocationName: $($MyInvocation.InvocationName); Line Number [$($MyInvocation.ScriptLineNumber)]; Line: $($MyInvocation.line)"
		
		$msExchDelegateListRecipients = @()
		
		$msExchDelegateListlink  = $UserParameter.msExchDelegateListlink
		$msExchDelegateListBL  = $UserParameter.msExchDelegateListBL
				
		Write-Verbose  "msExchDelegateListBL $msExchDelegateListBL " 
		
		$global:gmsExchDelegateListBL = $msExchDelegateListBL
		
		
		$msExchDelegateListBL | %{ 	$msExchDelegateListRecipients +=  ( Get-ADUser $_ ).Name }
				
				
				<#
				$msExchDelegateListBL | %{
					
					if ( $_)
					{
						$msExchDelegateListRecipients += try { ( get-recipient $_ -erroraction silent).Name }catch{}
						$msExchDelegateListRecipients+= try { ( get-o365recipient $_ -erroraction silent).name }catch{}
					}
				}
				#>
		
	Write-Verbose  "msExchDelegateListlink  $msExchDelegateListlink " 	
		
			<#
			$msExchDelegateListlink | %{
					
					#write-host "1401 msExchDelegateListlink  $_ " -ForegroundColor Magenta 
					if ( $_)
					{
						$msExchDelegateListRecipients  +=  if ( get-recipient $_ -erroraction silent){ ( get-recipient $_ -erroraction silent).Name}
						$msExchDelegateListRecipients  += try { if (get-o365recipient $_ -erroraction silent ){  ( get-o365recipient $_ -erroraction silent).name }   }catch{}
					}
			}
			#>
			
			$global:gmsExchDelegateListRecipients = $msExchDelegateListRecipients
			
			if ( $msExchDelegateListRecipients )
			{
				Write-Host ""
				Write-Host "Automap: " -foregroundcolor Cyan -BackgroundColor BLUE -NoNewline
				Write-Host " " -NoNewline
				Write-Host ( $msExchDelegateListRecipients  -join " ; " | ft | Out-String).Trim() -foregroundcolor Cyan
				
				write-host "msExchDelegateListlink:  $msExchDelegateListlink " -foregroundcolor Cyan
				write-host "msExchDelegateListBL:   $msExchDelegateListBL " -foregroundcolor Cyan
				
				Write-Host "Set-ADUser $($UserParameter.SamAccountName) -Clear msExchDelegateListLink" -ForegroundColor Yellow 
				
			}
			
			Write-Verbose "1688"
		
		
	} ########### END  fAutomap  ($UserParameter)  ###########
	
	function fRoomMailbox ($UserParameter)
	{

		#$UserParameter  | select  ADRecipientTypeDetails | Out-Host 
				
		Write-Verbose " InvocationName: $($MyInvocation.InvocationName); Line Number [$($MyInvocation.ScriptLineNumber)]; Line: $($MyInvocation.line)"
		
		$BookingRecipients = @()
					
		Write-Host ( $UserParameter | select ResourceCapacity, ResourceCustom, ModerationEnabled, ModeratedBy  | ft -AutoSize  | Out-String).Trim() -foregroundcolo Cyan
		
		#$($UserParameter.ADRecipientTypeDetails ) | Write-Host -ForegroundColor Magenta 
					
		# https://technet.microsoft.com/en-CA/library/ms.exch.eac.EditRoomMailbox_ResourceDelegates(EXCHG.150).aspx?v=15.0.1104.0&l=0
		
		if ( $UserParameter.ADRecipientTypeDetails -like "RemoteRoom*" )
		{		
				#On O365 
				
				if ( $O365 )
				{
					Write-Verbose "`$MailboxCalendarConfiguration = Get-O365MailboxCalendarConfiguration $($UserParameter.SamAccountName) -ErrorAction silent"
					
					$MailboxCalendarConfiguration = Get-O365MailboxCalendarConfiguration $($UserParameter.SamAccountName) -ErrorAction silent
					
					Write-verbose  "`$CalendarProcessing = Get-O365CalendarProcessing -ErrorAction silent $($UserParameter.SamAccountName)" 

					$CalendarProcessing = Get-O365CalendarProcessing -ErrorAction silent $($UserParameter.SamAccountName)
					
					if ( $UserParameter.SamAccountName)
					{
						#Write-Host "`$CalEn = $($UserParameter.SamAccountName) + ':\Calendar'" -ForegroundColor Magenta 
						
						$CalEn = $UserParameter.SamAccountName + ":\Calendar"
						$CalFr  = $UserParameter.SamAccountName + ":\Calendrier"
					}
						
					$PermissionsEN = Get-O365MailboxFolderPermission $CalEn -ErrorAction SilentlyContinue  
					$PermissionsFR = Get-O365MailboxFolderPermission $CalFr -ErrorAction SilentlyContinue	
					
					if($PermissionsEN )
					{ 	$CalendarString = $CalEn}
					elseif($PermissionsFR) 
					{ 	 $CalendarString = $CalFr }
					else
					{ 	$CalendarString = "" }
						
					if ( $CalendarString )
					{
						
						#$CalendarString
						$CalendarPermissions = Get-O365MailboxFolderPermission -identity  $CalendarString  
						
						$CalendarPermissions = $CalendarPermissions | sort User
						
						$global:gCalendarPermissions = $CalendarPermissions

						Write-Host ""
						Write-Host "Add-MailboxFolderPermission -identity  $CalendarString -User ObjectID" -ForegroundColor Yellow 
						Write-Host ""
						Write-Host "Calendar Permissions  : "  -ForegroundColor Cyan -BackgroundColor Blue
						Write-Host ""
						Write-Host ($CalendarPermissions | ft User, AccessRights -AutoSize | Out-String).Trim() -foregroundcolor Cyan
						
					}
				

				}
				else
				{
					#Remote but no O365 session 
				
					$MailboxCalendarConfiguration = $null 
					$CalendarProcessing =  "NoO365"
					
				}
		}
		else
		{	
				# On Prem 
				
				#Write-Host "OnPrem" -ForegroundColor Magenta
				
				Write-Verbose "`$MailboxCalendarConfiguration = Get-MailboxCalendarConfiguration $($UserParameter.SamAccountName) -ErrorAction silent"
				
				$MailboxCalendarConfiguration = Get-MailboxCalendarConfiguration $($UserParameter.SamAccountName) -ErrorAction silent
				
				Write-verbose  "`$CalendarProcessing = Get-CalendarProcessing -ErrorAction silent $($UserParameter.SamAccountName)" 
				
				$CalendarProcessing = Get-CalendarProcessing -ErrorAction silent $($UserParameter.SamAccountName)
				
				if ( $UserParameter.SamAccountName)
				{
					#Write-Host "`$CalEn = $($UserParameter.SamAccountName) + ':\Calendar'" -ForegroundColor Magenta 
					
					$CalEn = $UserParameter.SamAccountName + ":\Calendar"
					$CalFr  = $UserParameter.SamAccountName + ":\Calendrier"
				}
					
				$PermissionsEN = Get-MailboxFolderPermission $CalEn -ErrorAction SilentlyContinue  
				$PermissionsFR = Get-MailboxFolderPermission $CalFr -ErrorAction SilentlyContinue	
				
					
				if($PermissionsEN )
				{ 	$CalendarString = $CalEn}
				elseif($PermissionsFR) 
				{ 	 $CalendarString = $CalFr }
				else
				{ 	$CalendarString = "" }
					
				if ( $CalendarString )
				{
					
					#$CalendarString
					$CalendarPermissions = Get-MailboxFolderPermission -identity  $CalendarString  
					
					$CalendarPermissions = $CalendarPermissions | sort User
					
					$global:gCalendarPermissions = $CalendarPermissions

					Write-Host ""
					Write-Host "Add-MailboxFolderPermission -identity  $CalendarString -User ObjectID" -ForegroundColor Yellow 
					Write-Host ""
					Write-Host "Calendar Permissions  : "  -ForegroundColor Cyan -BackgroundColor Blue
					Write-Host ""
					Write-Host ($CalendarPermissions | ft User, AccessRights -AutoSize | Out-String).Trim() -foregroundcolor Cyan
					
				}
						
				
				
		} # On Prem
		
		$global:gMailboxCalendarConfiguration  =  $MailboxCalendarConfiguration 
	
		$global:gCalendarProcessing  = $CalendarProcessing 
		
		if ( $CalendarProcessing -ne "NoO365" )
		{
			$ResourceDelegates = $CalendarProcessing.ResourceDelegates
			$ResourceDelegatesString = @()
		
			$ResourceDelegates | 
			%{
				if ($_) { 	$ResourceDelegatesString += $_.tostring()  + ", " }
			}
			
			Write-Host " " 
			Write-Host "Calendar processing :  " -ForegroundColor Cyan -BackgroundColor Blue
			Write-Host " " 

			
			$AutomateProcessing = $CalendarProcessing.AutomateProcessing
			$AdditionalResponse  = $CalendarProcessing.AdditionalResponse
			
			$AddAdditionalResponse = $CalendarProcessing.AddAdditionalResponse
			$AllBookInPolicy = $CalendarProcessing.AllBookInPolicy
			$BookInPolicy = @( $CalendarProcessing.BookInPolicy  )	
			
			
			
			$BookingRecipients = @()
			
			if ( $BookInPolicy )
			{
					$BookInPolicy = $BookInPolicy | sort Name 
					
					$global:gBookInPolicy = $BookInPolicy
							
					$BookInPolicy | 
					%{ 
			
							$recipient = try {Get-Recipient $_  -ErrorAction silentlycontinue }catch{} 
					
							if ( $recipient )
							{
								#$BookingRecipients += try {Get-Recipient $_  -ErrorAction silentlycontinue}catch{} 
								
								$BookingRecipients += $recipient
							}
							else
							{
								$BookingRecipients += $_
							
							}
				
					}

							$BookingRecipientscount = $BookingRecipients.count
							
							$BookingRecipients = $BookingRecipients | sort Name

			}
	
			$global:gBookingRecipients = $BookingRecipients 
			
			Write-Host "AutomateProcessing : " -NoNewline  -BackgroundColor DarkBlue -ForegroundColor Cyan
		
			if ($AutomateProcessing -eq "AutoAccept")
			{ 	Write-Host "$AutomateProcessing " -ForegroundColor DarkBlue -BackgroundColor Green }
			else
			{ 	Write-Host "$AutomateProcessing" -BackgroundColor Red -ForegroundColor White }
			
			Write-Host ""
			Write-Host "AllBookInPolicy : " -NoNewline -BackgroundColor DarkBlue -ForegroundColor Cyan
			
			if ( $AllBookInPolicy )
			{
				Write-Host "$AllBookInPolicy" -ForegroundColor DarkBlue -BackgroundColor Green
				Write-Host ""
				Write-host "Set-CalendarProcessing $($UserParameter.guid) -AllBookInPolicy:`$false" -ForegroundColor Yellow
				
			}
			else
			{	
					Write-Host "$AllBookInPolicy" -BackgroundColor Red -ForegroundColor White
					Write-Host ""
					
					Write-host "Set-CalendarProcessing $($UserParameter.guid) -AllBookInPolicy:`$true" -ForegroundColor Yellow
					
					$BookingRecipientscount = $BookingRecipients.count
					
					Write-Host ""
					Write-Host "[$BookingRecipientscount] Entities can book : " -BackgroundColor DarkBlue -ForegroundColor Cyan
					Write-Host ""
					
					Write-Host ( $BookingRecipients | ft   Name, RecipientType, guid -AutoSize  | Out-String).trim() -BackgroundColor Blue -ForegroundColor Cyan
	
					if ( $($BookingRecipients.length)  )
					{
						Write-Host ""
										
						$global:gBookingRecipients = @()
						
						#$BookingRecipients | %{ if ( $_.guid ){ $global:BookingRecipients += $_.guid.tostring() } }
						
						$BookingRecipients | %{ if ( $_.guid ){ $global:BookingRecipients += $_.primarysmtpaddress } }
						
					}
					else
					{
					
						Write-Host ""
						Write-Host "`$global:BookingRecipients = @() "   -ForegroundColor Yellow
						$global:BookingRecipients = @()
			
					}
			
					Write-Host ""
					Write-Host " " -NoNewline
					Write-Host "To grant additional users the right to book this room run :" -ForegroundColor Cyan	-BackgroundColor	Blue
					Write-Host ""
					
					#Write-Host ".\Verify-RecipientList.ps1 | %{ Try{`$BookingRecipients +=  `$_.guid.ToString() }catch{} } ;  "   -ForegroundColor Yellow
					
					Write-Host ".\Verify-RecipientList.ps1 | %{ Try{`$BookingRecipients +=  `$_.primarysmtpaddress }catch{} } ;  "   -ForegroundColor Yellow
				
					
					Write-Host "Set-CalendarProcessing $($UserParameter.guid)  -BookInPolicy `$(`$BookingRecipients | select -unique)	 "  -ForegroundColor Yellow
					
					Write-Host ""
					Write-Host "ForwardRequestsToDelegates " -NoNewline -BackgroundColor DarkBlue -ForegroundColor Cyan
		
					if ( $($CalendarProcessing.ForwardRequestsToDelegates))
					{
						Write-Host "$($CalendarProcessing.ForwardRequestsToDelegates)" -ForegroundColor DarkBlue -BackgroundColor Green
					}
					else
					{
						Write-Host "$($CalendarProcessing.ForwardRequestsToDelegates)" -BackgroundColor Red -ForegroundColor White
					
					}
	
				Write-Host "ResourceDelegates :   " -NoNewline -BackgroundColor DarkBlue -ForegroundColor Cyan 
				Write-Host  $ResourceDelegatesString -BackgroundColor DarkGreen -ForegroundColor White

			}
			
			Write-Host ""
			Write-Host "Additional response : " -NoNewline -BackgroundColor DarkBlue -ForegroundColor Cyan
			
			if ( $AddAdditionalResponse )
			{
				Write-Host " ON "  -ForegroundColor DarkBlue -BackgroundColor Green
				Write-Host ""
				Write-Host " " -NoNewline 
				Write-Host " `"  $AdditionalResponse `"  "  -ForegroundColor White -BackgroundColor DarkCyan
				Write-Host ""
			}
			else
			{ 	Write-Host "OFF "  -ForegroundColor Yellow -BackgroundColor Red }
			
			$Info =	$CalendarProcessing | select AutomateProcessing, AllowRecurringMeetings, AllowConflicts,  BookingWindowInDays, MaximumDurationInMinutes,ConflictPercentageAllowed , MaximumConflictInstances, `
			ForwardRequestsToDelegates, DeleteAttachments, DeleteComments, RemovePrivateProperty, DeleteSubject, AddOrganizerToSubject
			
			Write-Host ($Info | fl | Out-String).Trim() -foregroundcolor Cyan
			Write-Host ($MailboxCalendarConfiguration | fl WorkDays, WorkingHoursStartTime, WorkingHoursEndTime | Out-String).Trim() -foregroundcolor Cyan
			
				
		}# if ( $CalendarProcessing -ne "NoO365" )
		else
		{
			Write-Host ""
			Write-Host "	" -NoNewline
			Write-Host "  ATTENTION :  No Calendar processing (No O365 Session)" -ForegroundColor  Blue	-BackgroundColor	Yellow
			Write-Host ""
			
			$ResourceDelegates = $null 
		}
		
			
		
		<#
		if ( $BookInPolicy )
		{
			$BookInPolicy = $BookInPolicy | sort Name 
			
			$global:gBookInPolicy = $BookInPolicy
	
			
			
			$BookInPolicy | %{ 
	
					$recipient = try {Get-Recipient $_  -ErrorAction silentlycontinue }catch{} 
			
					if ( $recipient )
					{
						#$BookingRecipients += try {Get-Recipient $_  -ErrorAction silentlycontinue}catch{} 
						
						$BookingRecipients += $recipient
					}
					else
					{
						$BookingRecipients += $_
					
					}
		
			}

			$BookingRecipientscount = $BookingRecipients.count
			
			$BookingRecipients = $BookingRecipients | sort Name

		}
	
		#>
		
		<#
		if ( $AllBookInPolicy )
		{
			Write-Host "$AllBookInPolicy" -ForegroundColor DarkBlue -BackgroundColor Green
			Write-Host ""
			Write-host "Set-CalendarProcessing $($UserParameter.guid) -AllBookInPolicy:`$false" -ForegroundColor Yellow
			
		}
		else
		{	
			Write-Host "$AllBookInPolicy" -BackgroundColor Red -ForegroundColor White
			Write-Host ""
			
			Write-host "Set-CalendarProcessing $($UserParameter.guid) -AllBookInPolicy:`$true" -ForegroundColor Yellow
			
			$BookingRecipientscount = $BookingRecipients.count
			
			Write-Host ""
			Write-Host "[$BookingRecipientscount] Entities can book : " -BackgroundColor DarkBlue -ForegroundColor Cyan
			Write-Host ""
			
			Write-Host ( $BookingRecipients | ft   Name, RecipientType, guid -AutoSize  | Out-String).trim() -BackgroundColor Blue -ForegroundColor Cyan
	
			if ( $($BookingRecipients.length)  )
			{
				Write-Host ""
								
				$global:gBookingRecipients = @()
				
				#$BookingRecipients | %{ if ( $_.guid ){ $global:BookingRecipients += $_.guid.tostring() } }
				
				$BookingRecipients | %{ if ( $_.guid ){ $global:BookingRecipients += $_.primarysmtpaddress } }
				
			}
			else
			{
			
				Write-Host ""
				Write-Host "`$global:BookingRecipients = @() "   -ForegroundColor Yellow
				$global:BookingRecipients = @()
	
			}
			Write-Host ""
			Write-Host " " -NoNewline
			Write-Host "To grant additional users the right to book this room run :" -ForegroundColor Cyan	-BackgroundColor	Blue
			Write-Host ""
			
			#Write-Host ".\Verify-RecipientList.ps1 | %{ Try{`$BookingRecipients +=  `$_.guid.ToString() }catch{} } ;  "   -ForegroundColor Yellow
			
			Write-Host ".\Verify-RecipientList.ps1 | %{ Try{`$BookingRecipients +=  `$_.primarysmtpaddress }catch{} } ;  "   -ForegroundColor Yellow
			
			
			
			Write-Host "Set-CalendarProcessing $($UserParameter.guid)  -BookInPolicy `$(`$BookingRecipients | select -unique)	 "  -ForegroundColor Yellow
			
			Write-Host ""
			Write-Host "ForwardRequestsToDelegates " -NoNewline -BackgroundColor DarkBlue -ForegroundColor Cyan
		
			if ( $($CalendarProcessing.ForwardRequestsToDelegates))
			{
				Write-Host "$($CalendarProcessing.ForwardRequestsToDelegates)" -ForegroundColor DarkBlue -BackgroundColor Green
			}
			else
			{
				Write-Host "$($CalendarProcessing.ForwardRequestsToDelegates)" -BackgroundColor Red -ForegroundColor White
			
			}
	
			Write-Host "ResourceDelegates :   " -NoNewline -BackgroundColor DarkBlue -ForegroundColor Cyan 
			Write-Host  $ResourceDelegatesString -BackgroundColor DarkGreen -ForegroundColor White

		}
		
		#>
		
		
	}########### END function fRoomMailbox () ###########

	function fO365BrokenSession ()
	{
				
		$O365BrokenSessions = @( Get-PSSession | ?{ ($_.ComputerName -like "outlook.office365*" -or $_.ComputerName -like  "*online.lync.com*"  ) -and ( $_.State -ne "Opened" ) } )
			
		if ( $O365BrokenSessions )
		{
					Write-Host ""
					Write-Host "[uADo] O365P broken session processing " -ForegroundColor Magenta 
					Write-Host ""

					$O365BrokenSessions | Remove-PSSession 
					
					$Session = $O365BrokenSessions[0]
					
					$tempmodules = Get-Module | where { $_.Description.contains( $Session.ComputerName) }

					Write-Host "$tempmodules  | Remove-Module " -foregroundcolor Red
					Write-Host ""
						
					$tempmodules  | Remove-Module
					
										
					#.\Remote-SessionConnect.ps1  -silent
					
					Write-Host ""
				
			}
	
	} ########### END function O365BrokenSession  ###########
		
	function fOnPremBrokenSession ()
	{
		
		$OnPremBrokenSessions = @( Get-PSSession | ?{ ($_.ComputerName -like "*.msoit.com" -or $_.ComputerName -like  "*online.lync.com*"  ) -and ( $_.State -ne "Opened" ) } )
		
			if ( $OnPremBrokenSessions )
			{
					Write-Host ""
					Write-Host "[uADo] OnPrem broken session processing " -ForegroundColor Magenta
					Write-Host ""
					
					$OnPremBrokenSessions | Remove-PSSession 
					
					$Session = $OnPremBrokenSessions[0]

					
					$tempmodules = Get-Module | where { $_.Description.contains( $Session.ComputerName) }
					
					#Write-Host ( $tempmodules | ft -AutoSize  | Out-String ) -foregroundcolor Red
					
					Write-Host "$tempmodules  | Remove-Module " -foregroundcolor Red
					Write-Host ""
					
					$tempmodules  | Remove-Module
		
					 .\Remote-SessionConnect.ps1 -SessionType onprem -silent	
			
			}
			
	
	
	} ########### END function OnPremBrokenSession   ###########
	
	function fSearchMsolUser ( $UserParameter )
	{
		#Write-Host " .\Remote-SessionConnect.ps1 -Msol" -ForegroundColor Magenta
		 
		 .\Remote-SessionConnect.ps1 -Msol
		 .\Remote-SessionConnect.ps1 

		$results = @()

		while  ( $results.count  -ne 1 )
		{
			#Write-Host "2586" -ForegroundColor Red 
			
			if ( $UserParameter  )
			{		
				#Write-Host "2590" -ForegroundColor Red 
			
				if ( $UserParameter  -eq "*" )
				{
					Write-Host "CP1" -ForegroundColor Red 
					
					Write-Host " "
					Write-Host "	" -NoNewline
					Write-Host "  ATTENTION :  Searching for all Msol users ..." -ForegroundColor  Blue	-BackgroundColor	Yellow
					Write-Host " "
					
					Write-Host "`$results = Get-MsolUser -All" -ForegroundColor Yellow
					Write-Host " "
				
					$results = Get-MsolUser -All
				}
				elseif ( $UserParameter  -eq "^" )  
				{ 
					Write-Host ""
					Write-Host "	" -NoNewline
					Write-Host "  ATTENTION :  User Skipped" -ForegroundColor  Blue	-BackgroundColor	Yellow

					return "^"
				}
				else
				{
							$isaObjectID 	= [guid]::TryParse( $UserParameter , $([ref][guid]::Empty))
				
							if ( $isaObjectID  )
							{
								Write-Host ""
								
								Write-Host "Get-MsolUser -ObjectId  $UserParameter" -ForegroundColor Yellow 
								
								$results = Get-MsolUser -ObjectId  $UserParameter 
							 	
							}
							else
							{
								Write-Host " "
								Write-Host "Get-MsolUser -SearchString '$UserParameter'  -All" -ForegroundColor Yellow 
								Write-Host " "
								$results = Get-MsolUser -SearchString $UserParameter  -All
							}
				
				}
				
				$Ro = $null 
				
				Write-Host " " 
				Write-Host  " [$($results.count)] results found "   -ForegroundColor Cyan	-BackgroundColor	Blue
				Write-Host " " 
				
				if ( $results.count -gt 1 )
				{
						
						 $results =  $results | sort isLicensed ,MSExchRecipientTypeDetails, displayname | `
						select Objectid , UserPrincipalName, DisplayName, isLicensed,`
						@{Label="RecipientType"; Expression={   ( Get-O365Recipient $_.UserPrincipalName -ErrorAction silent ).RecipientTypeDetails   } },`
						BlockCredential,  WhenCreated
						
						Write-Host ($results | ft -AutoSize -Wrap  | Out-String ) -foregroundcolor Cyan

				}
				
				#$Ro =  Get-O365Recipient $_.UserPrincipalName -ErrorAction silent 
				
				#Write-Host (  $Ro | fl  | Out-String).trim()  -ForegroundColor Magenta
				<#
				if ($Ro )
				{
					$Mailbox =  Get-O365Mailbox   $Ro.PrimarySmtpAddress #  -ErrorAction silent
					
					Write-Host ( $Mailbox  | fl ExchangeGuid,`
					#@{Label="Teams"; Expression = {  if($O365) { $_.EmailAddresses |  ?{ $_.ToString() -like "sip*" } }else{ "`'SIP N/A`'"  }         }},`
					@{Label="Alias"; Expression = { if ($_.EmailAddress) { $OnpremAlias = ( get-recipient $_.EmailAddress).Alias ;  if ( $OnpremAlias -ne  $_.Alias ) {  "`'OnPrem Alias: $OnpremAlias`' `'O365 Alias: $($_.Alias)`'  " }else{  $($_.Alias)    }  }  } },`
					@{Label="LitHold"; Expression = {  $_.LitigationHoldEnabled } },`
					@{Label="LitHoldOwner"; Expression = {  $_.LitigationHoldOwner } },`
					LastLogonTime,`
					@{Label=" "; Expression = { "" } },`
					#@{Label="TotalItemSize"; Expression = {  if( $_.TotalItemSize ) { $_.TotalItemSize   }else{ "N/A (no O365 session)"  }         }},`
					@{Label="TotalItemSize"; Expression = {  if( !$O365 -and !$EXO -and !( $_.TotalItemSize ) ) { "N/A (no O365 session)"    }else{ $_.TotalItemSize }         }},`
					#@{Label="Database"; Expression = {  if( $_.Database ) { $_.Database   }else{ "N/A (no O365 session)"  }         }},`
					@{Label="Database"; Expression = {  if( !$O365 -and !$EXO  -and !( $_.Database ) ) {  "N/A (no O365 session)" }else{ $_.Database  }         }},`
					MailboxRegion,MailboxRegionLastUpdateTime,`
					@{Label="Extention Attributes"; Expression = {$EnabledextensionAttributes }},`
					msExchWhenMailboxCreated, Mailbox-WhenCreated, Mailbox-WhenChanged | Out-String).trim()  -ForegroundColor Magenta 
				
				}
				#>
			
				$UserParameter = $null
			
			} # if ( $UserParameter  )
			else
			{
					Write-Host " " -NoNewline 
					Write-Host  "[usra] Enter user name " -ForegroundColor Blue -BackgroundColor Gray -NoNewline
					$UserParameter  = Read-Host " "
					
			}
						
			#Write-Host " CP0 " -ForegroundColor Red
			
		}
		
		 #$results  | fl
		
		
		#Write-Host " CP0 " -ForegroundColor Red
		
		fDisplayMsolUser $results 
		
		return $results
		
	} ########### END function OnPremBrokenSession fSearchMsolUser ( $UserParameter )   ###########

	
	#endregion Functions
	
	#region	Initialization 

		if ( $licencedetails )
		{
			$licencegroup = $true
		}
			
		Write-Verbose "  $($MyInvocation.InvocationName); Line [$($MyInvocation.ScriptLineNumber)]: $($MyInvocation.line) ;  Initializing  .."
		
		$licenseNames = @{
				"POWER_BI_STANDALONE" = "POWER BI STANDALONE" 
				"ENTERPRISEPACKWITHOUTPROPLUS" = "Office 365 E3 without Microsoft 365 Apps"
				"MCOPSTNC" = "Communications Credits"
				"MCOPSTN1" = "Microsoft 365 Domestic Calling Plan"
				"MCOPSTN_5" = "Microsoft 365 Domestic Calling Plan (120 min)"
				"OFFICE365_MULTIGEO" = "Multi-Geo Capabilities in Office 365"
				"Win10_VDA_E3" = "WINDOWS 10 ENTERPRISE E3"
		         "OFFICESUBSCRIPTION_FACULTY"= "Office 365 ProPlus for Faculty" 
		         "RIGHTSMANAGEMENT_STANDARD_FACULTY"= "Azure Rights Management for faculty"       
		         "ADALLOM_S_O365"= "POWER BI STANDALONE"       
		         "CRMSTANDARD"= "Microsoft Dynamics CRM Online" 
		         "EMS"= "Enterprise Mobility Suite" 
		         "EMSPREMIUM" = "Enterprise Mobility + Security E5"
		         "O365_BUSINESS_PREMIUM"= "Office 365 BUSINESS PREMIUM" 
		         "DESKLESSPACK"= "Office 365 (Plan K1)" 
		         "DESKLESSWOFFPACK"= "Office 365 (Plan K2)" 
		         "LITEPACK"= "Office 365 (Plan P1)" 
		         "EXCHANGESTANDARD"= "Office 365 Exchange Online Only" 
		         "STANDARDPACK"= "Office 365 (Plan E1)" 
		         "STANDARDWOFFPACK"= "Office 365 (Plan E2)" 
		         "ENTERPRISEPACK"= "Office 365 E3" 
		         "ENTERPRISEPACKLRG"= "Office 365 Enterprise E3 LRG" 
		         "ENTERPRISEWITHSCAL"= "Office 365 Enterprise E4" 
		         "ENTERPRISEPREMIUM"= "Ofiice 365 Enterprise E5" 
		         "STANDARDPACK_STUDENT"= "Office 365 (Plan A1) for Students" 
		         "STANDARDWOFFPACKPACK_STUDENT"= "Office 365 (Plan A2) for Students" 
		         "ENTERPRISEPACK_STUDENT"= "Office 365 (Plan A3) for Students" 
		         "ENTERPRISEWITHSCAL_STUDENT"= "Office 365 (Plan A4) for Students" 
		         "EXCHANGESTANDARD_STUDENT"= "Exchange Online (Plan 1) for Students" 
		         "STANDARDPACK_FACULTY"= "Office 365 (Plan A1) for Faculty" 
		         "OFFICESUBSCRIPTION_STUDENT"= "Office ProPlus Student Benefit" 
		         "STANDARDWOFFPACK_FACULTY"= "Office 365 Education E1 for Faculty" 
		         "STANDARDWOFFPACK_IW_STUDENT"= "Office 365 Education for Students" 
		         "STANDARDWOFFPACK_IW_FACULTY"= "Office 365 Education for Faculty" 
		         "STANDARDWOFFPACK_STUDENT"= "Microsoft Office 365 (Plan A2) for Students" 
		         "ENTERPRISEPACK_FACULTY"= "Office 365 (Plan A3) for Faculty" 
		         "EOP_ENTERPRISE_FACULTY"= "Exchange Online Protection for Faculty" 
		         "ENTERPRISEWITHSCAL_FACULTY"= "Office 365 (Plan A4) for Faculty" 
		         "ENTERPRISEPACK_B_PILOT"= "Office 365 (Enterprise Preview)" 
		         "STANDARD_B_PILOT"= "Office 365 (Small Business Preview)" 
		         "CRMIUR"= "CRM for Partners" 
		         "AAD_PREMIUM"= "Azure Active Directory Premium" 
		         "STANDARDPACK_GOV"= "Microsoft Office 365 (Plan G1) for Government" 
		         "STANDARDWOFFPACK_GOV"= "Microsoft Office 365 (Plan G2) for Government" 
		         "ENTERPRISEPACK_GOV"= "Microsoft Office 365 (Plan G3) for Government" 
		         "ENTERPRISEWITHSCAL_GOV"= "Microsoft Office 365 (Plan G4) for Government" 
		         "DESKLESSPACK_GOV"= "Microsoft Office 365 (Plan K1) for Government" 
		         "ESKLESSWOFFPACK_GOV"= "Microsoft Office 365 (Plan K2) for Government" 
		         "EXCHANGESTANDARD_GOV"= "Microsoft Office 365 Exchange Online (Plan 1) only for Government" 
		         "EXCHANGEENTERPRISE"= "Microsoft Office 365 Exchange Online (Plan 2) only for Government" 
				  "SHAREPOINTDESKLESS_GOV"= "SharePoint Online Kiosk" 
		         "EXCHANGE_S_DESKLESS_GOV"= "Exchange Kiosk" 
		         "RMS_S_ENTERPRISE_GOV"= "Windows Azure Active Directory Rights Management" 
		         "OFFICESUBSCRIPTION_GOV"= "Microsoft 365 Apps for enterprise" 
		         "MCOSTANDARD_GOV"= "Lync Plan 2G" 
		         "SHAREPOINTWAC_GOV"= "Office Online for Government" 
		         "SHAREPOINTENTERPRISE_GOV"= "SharePoint Plan 2G" 
		         "EXCHANGE_S_ENTERPRISE_GOV"= "Exchange Plan 2G" 
		         "EXCHANGE_S_ARCHIVE_ADDON_GOV"= "Exchange Online Archiving" 
		         "EXCHANGE_L_STANDARD"= "Exchange Online (Plan 1)" 
		         "MCOLITE"= "Lync Online (Plan 1)" 
		         "SHAREPOINTLITE"= "SharePoint Online (Plan 1)" 
		         "OFFICE_PRO_PLUS_SUBSCRIPTION_SMBIZ"= "Office ProPlus" 
		         "EXCHANGE_S_STANDARD_MIDMARKET"= "Exchange Online (Plan 1)" 
		         "MCOSTANDARD_MIDMARKET"= "Lync Online (Plan 1)" 
		         "SHAREPOINTENTERPRISE_MIDMARKET"= "SharePoint Online (Plan 1)" 
		         "SHAREPOINTWAC"= "Office Online" 
		         "OFFICESUBSCRIPTION"= "Office ProPlus" 
		         "YAMMER_MIDSIZE"= "Yammer" 
		         "EXCHANGE_S_STANDARD"= "Exchange Online (Plan 2)" 
		         "MCOSTANDARD"= "Skype for Business Online (Plan 2)" 
		         "SHAREPOINTENTERPRISE"= "SharePoint Online (Plan 2)" 
		         "RMS_S_ENTERPRISE"= "Azure Active Directory Rights Management" 
		         "YAMMER_ENTERPRISE"= "Yammer Enterprise" 
		         "MCVOICECONF"= "Lync Online (Plan 3)" 
		         "EXCHANGE_S_DESKLESS"= "Exchange Online Kiosk" 
		         "SHAREPOINTDESKLESS"= "SharePoint Online Kiosk" 
		         "EXCHANGEARCHIVE"= "Exchange Online Archiving" 
		         "EXCHANGETELCO"= "Exchange Online POP" 
		         "SHAREPOINTSTORAGE"= "SharePoint Online Storage" 
		         "SHAREPOINTPARTNER"= "SharePoint Online Partner Access" 
		         "PROJECTONLINE_PLAN_1"= "Project Online (Plan 1)" 
		         "PROJECTONLINE_PLAN_2"= "Project Online with Project Pro for Office 365" 
		         "PROJECT_CLIENT_SUBSCRIPTION"= "Project Pro for Office 365" 
		         "VISIO_CLIENT_SUBSCRIPTION"= "Visio Pro for Office 365" 
		         "INTUNE_A"= "Intune for Office 365" 
		         "CRMTESTINSTANCE"= "CRM Test Instance" 
		         "ONEDRIVESTANDARD"= "OneDrive" 
		         "SQL_IS_SSIM"= "Power BI Information Services" 
		         "BI_AZURE_P1"= "Power BI Reporting and Analytics" 
				 
		         "EOP_ENTERPRISE"= "Exchange Online Protection" 
		         "PROJECT_ESSENTIALS"= "Project Lite" 
		         "PROJECTPREMIUM"= "Project Online Premium" 
		         "NBPROFESSIONALFORCRM"= "Microsoft Social Listening Professional" 
		         "MFA_PREMIUM"= "Azure Multi-Factor Authentication" 
		         "DMENTERPRISE"= "Microsoft Dynamics Marketing Online Enterprise" 
		         "DESKLESS"= "Microsoft StaffHub" 
		         "STREAM"= "Microsoft Stream Trial" 
		         "FLOW_P1"= "Microsoft Flow Plan 1" 
		         "FLOW_P2"= "Microsoft Flow Plan 2" 
		         "POWERFLOW_P1"= "Microsoft PowerApps Plan 1" 
		         "POWERFLOW_P2"= "Microsoft PowerApps Plan 2" 
		         "DYN365_ENTERPRISE_PLAN1"= "Dynamics 365 Plan 1 Enterprise Edition" 
		         "AAD_PREMIUM_P2"= "Azure Active Directory Premium P2" 
		         "POWER_BI_PRO"= "Power BI Pro" 
				 "POWER_BI_STANDARD" = "Power BI (free)" 
		         "INFOPROTECTION_P2"= "Azure Information Protection Premium P2" 
		         "WACONEDRIVESTANDARD"= "OneDrive for Business with Office Online" 
		         "ADALLOM_STANDALONE"= "Microsoft Cloud App Security" 
		         "RIGHTSMANAGEMENT"= "Azure Rights Management Premium" 
				 "STREAM_O365_E3" = "Microsoft Stream for O365 E3 SKU"
				 "FORMS_PLAN_E3" = "Microsoft Forms (Plan E3)"
				  "EXCHANGE_S_ENTERPRISE" = "Exchange Online (Plan 2)"
				  "MYANALYTICS_P2"= "Insights by MyAnalytics"
				  "MICROSOFTBOOKINGS" = "Microsoft Bookings"
				  "MICROSOFT_SEARCH" = "Microsoft Search"
				  "TEAMS1" = "Microsoft Teams"
				  "POWERAPPS_O365_P2" = "Powerapps for Office 365"
				  "SWAY"="Sway"
				  "KAIZALA_O365_P3" = "Microsoft Kaizala Pro"
				  "FLOW_O365_P2" = "Flow for Office 365"
				  "INTUNE_O365" = "Common Data Service [Intune ?] "
				  "PROJECTWORKMANAGEMENT" = "Microsoft Planner"
				  "BPOS_S_TODO_2" = "To-Do (Plan 2)"
				  "WHITEBOARD_PLAN2" = "Whiteboard (Plan 2)"
					"OFFICEMOBILE_SUBSCRIPTION" = "Office Mobile Apps for Office 365"
					"MIP_S_CLP1" = "Information Protection for Office 365 - Standard"
					"WIN10_VDA_E5" = "WINDOWS 10 ENTERPRISE E5"
					"WIN_DEF_ATP"        = "Microsoft Defender Advanced Threat Protection"
					"ATP_ENTERPRISE"        = "Office 365 Advanced Threat Protection (Plan 1)"
					"DYN365_ENTERPRISE_CUSTOMER_SERVICE"        = "DYNAMICS 365 FOR CUSTOMER SERVICE ENTERPRISE EDITION"
					"DYN365_FINANCIALS_BUSINESS_SKU"        = "DYNAMICS 365 FOR FINANCIALS BUSINESS EDITION"
					"DYN365_ENTERPRISE_SALES_CUSTOMERSERVICE"        = "DYNAMICS 365 FOR SALES AND CUSTOMER SERVICE ENTERPRISE EDITION"
					"MDATP_Server" = "Microsoft Defender Advanced Threat Protection Server"
					"PROJECTPROFESSIONAL"        = "Project Plan 3"
					"MEETING_ROOM" = "Microsoft Teams Rooms Standard"
					"SMB_APPS"      = "Business Apps (free)"
					"DYN365_ENTERPRISE_P1_IW"      = "Dynamics 365 P1 Trial for Information Workers"
	 				"DYN365_ENTERPRISE_TEAM_MEMBERS" = "Dynamics 365 For Team Members Enterprise Edition"
					"DYN365_ENTERPRISE_SALES"      = "Dynamics 365 Sales Enterprise Edition"
					"FORMS_PRO" = "Microsoft Forms Pro"
					"FLOW_FREE" = "Microsoft Power Automate Free"
					"MFA_STANDALONE" = "Microsoft Azure Multi-Factor Authentication"
					"CRMSTORAGE" = "Microsoft Dynamics CRM Online"
					"WINDOWS_STORE" = "Windows Store for Business"
					"VISIOCLIENT" = "Visio Plan 2"
					"POWERAPPS_VIRAL" = "Microsoft Power Apps Plan 2 Trial "
					"SPZA_IW"      = "Dynamics 365 Customer Voice Trial"
					"RIGHTSMANAGEMENT_ADHOC" = "Rights Management Adhoc"
					'POWERAPPS_INDIVIDUAL_USER' = 'PowerApps and Logic flows'
					"DYN365_TEAM_MEMBERS"      = "Dynamics 365 Team Members"
					'DYN365_AI_SERVICE_INSIGHTS'='Dynamics 365 Customer Service Insights Trial'
					"AX7_USER_TRIAL" 	= "Microsoft Dynamics AX7 User Trial"
					"PhoneSystem_VirtualUser" 	= "Microsoft 365 Phone System - Virtual User"
			  		"M365_E5_Suite_features" 	= "Microsoft 365 E5 Suite features"
					"M365_E5_Suite_components" 	= "Microsoft 365 E5 Suite components"
					"SKU_Dynamics_365_for_HCM_Trial" 	= "Dynamics 365 for Talent"
	         }
		
		
		$RecipientTypeDetailsList = @{
			"1"	= "UserMailbox"
			"2"		=	 "LinkedMailbox"
			"4" 		="SharedMailbox"
			"8"		  = "LegacyMailbox"	
			"16" 			 = "RoomMailbox"		
			"32" 			  = "EquipmentMailbox" 
			"64" 			 = "MailContact"		
			"128"			  = "Mail-enabled User"
			"2147483648" 	  = "RemoteUserMailbox"
			"8589934592" 	  = "RemoteRoomMailbox"
			"17179869184"	 = "RemoteEquipmentMailbox"
			"34359738368" = "RemoteSharedMailbox"
				
			}
			
		$CloudExchangeRecipientDisplayTypeList 	 = @{
			"0"	= "UserMailbox"
			"1073741824" = "Shared Mailbox"
			"-2147483642" = "RemoteUserMailbox"
		}
		
		$UserAccountControlList =  @{
		
		"1” = “SCRIPT” 
		"2” = “ACCOUNTDISABLE” 
		"8” = “HOMEDIR_REQUIRED” 
		"16” = “LOCKOUT” 
		"32” = “PASSWD_NOTREQD” 
		"64” = “PASSWD_CANT_CHANGE”
		"128” = “ENCRYPTED_TEXT_PWD_ALLOWED” 
		"256” = “TEMP_DUPLICATE_ACCOUNT” 
		"512” = “normal account” 
		"514” = “disabled account” 
		"544” = “Enabled, Password Not Required” 
		"546” = “Disabled, Password Not Required” 
		"2048” = “INTERDOMAIN_TRUST_ACCOUNT” 
		"4096” = “WORKSTATION_TRUST_ACCOUNT” 
		"8192” = “SERVER_TRUST_ACCOUNT” 
		"65536” = “DONT_EXPIRE_PASSWORD” 
		"66048” = “Enabled, Password Never Expires” 
		"66050” = “Disabled, PNE” 
		"66082” = “Disabled, Password Never Expire & Not Required” 
		"131072” = “MNS_LOGON_ACCOUNT” 
		"262144” = “SMARTCARD_REQUIRED” 
		"262656” = “Enabled, Smartcard Required” 
		"262658” = “Disabled, Smartcard Required” 
		"262690” = “Disabled, Smartcard Required, Password Not Required” 
		"328194” = “Disabled, Smartcard Required, Password Doesn’t Expire” 
		"328226” = “Disabled, Smartcard Required, Password Doesn’t Expire & Not Required” 
		"524288” = “TRUSTED_FOR_DELEGATION” 
		"532480” = “Domain controller” 
		"1048576” = “NOT_DELEGATED” 
		"2097152” = “USE_DES_KEY_ONLY” 
		"4194304” = “DONT_REQ_PREAUTH” 
		"8388608” = “PASSWORD_EXPIRED” 
		"16777216” = “TRUSTED_TO_AUTH_FOR_DELEGATION” 
		"67108864” = “PARTIAL_SECRETS_ACCOUNT” 
		
		}

		#Configure a default display set
		
		$defaultDisplaySet = 'DisplayName','EmailAddress','ObjectGUID'
		
		#Create the default property display set
		
		$defaultDisplayPropertySet = New-Object System.Management.Automation.PSPropertySet(‘DefaultDisplayPropertySet’, [  string[]  ]$defaultDisplaySet)
		$PSStandardMembers = [System.Management.Automation.PSMemberInfo[]]@($defaultDisplayPropertySet)
		
	
		fO365BrokenSession 
		fOnpremBrokenSession
		
		if ( $MyInvocation.ScriptName  )
	    {
			   Write-Host ""
			   
			   Write-Host ( $MyInvocation  | ft  -Wrap `
			   @{ Label = "Running Script"; Expression = { $_.MyCommand }  },`
			   @{ Label = "Calling Script, Line number, Expression"; Expression = {$_.ScriptName; "Line: $($_.ScriptLineNumber)"; $(($_.Line).Trim())   }  }`
			   | Out-String).trim() -foregroundcolor DarkCyan
		
		}
		else
		{		
			 	Write-Host ""
				Write-Host "UserADObject.ps1	 Created by Filip Neshev, October  2020	filipne@yahoo.com" -ForegroundColor Cyan  -backgroundcolor  DarkGray	
				Write-Host ""
				
		}
		
		$O365 = $false
		$EXO = $false 
		
		# Detect existing, prefixed with 'O365' remote session to Office 365
		
		
		#get-command "get-O365mailbox"
	
		#Write-Host "OnPrem  $OnPrem " -ForegroundColor Red
		
		$O365Command = try{ get-command "get-O365mailbox"-ErrorAction SilentlyContinue }catch{}
			
		if ( $O365Command -and !$OnPrem )
		{
			$O365 = $true
			
			Write-Host ""
			
			if ( !$silent )
			{
				
				Write-Host "[usra] O365P session" -ForegroundColor DarkBlue -BackgroundColor Green
			}
			else
			{
			 	Write-Host "[usra] O365P session" -ForegroundColor DarkBlue -BackgroundColor DarkGreen
				
			}
		}
	
		
		$psSessions = Get-PSSession | Select-Object -Property State, Name
		
		If ( ((@($psSessions) -like '@{State=Opened; Name=ExchangeOnlineInternalSession*').Count -gt 0)  -and !$OnPrem  ) 
		{
				$EXO = $true

				if ( !$silent )
				{
					Write-Host "[usra] EXO session" -ForegroundColor DarkBlue -BackgroundColor Green
				}
				else
				{
				 	Write-Host "[usra] EXO session" -ForegroundColor DarkBlue -BackgroundColor DarkGreen
					
				}
				
		}
		
		
		<#
		if( ($EXO = try{ get-command "get-EXOmailbox"-ErrorAction SilentlyContinue }catch{}) -and !$OnPrem )
		{
			$EXO = $true
			
			Write-Host ""
			
			if ( !$silent )
			{
				
				Write-Host "[usra] EXO session" -ForegroundColor DarkBlue -BackgroundColor Green
			}
			else
			{
			 	Write-Host "[usra] EXO session" -ForegroundColor DarkBlue -BackgroundColor DarkGreen
				
			}
	
		}
		#>
		
		if ( !$O365 -and !$EXO )
		{
			$O365 = $false
			$EXO = $false 
			$OnPremTxt = "[usra] OnPrem only session"
			if ( $OnPrem )
			{
				$OnPremTxt += " (requested with -onprem)"
			
			}
			
			Write-Host ""
			Write-Host  $OnPremTxt  -ForegroundColor Cyan	-BackgroundColor	Blue
		
		}
		
		#Write-Host "$O365 ; $EXO  " -ForegroundColor Red

		

	#endregion Initialization 


} # end begin 

process 
{
	fO365BrokenSession 
	fOnpremBrokenSession 
	
	if ( $msol )
	{
		 fSearchMsolUser $UserName
	}
	else
	{
	 	$CustomADUser  =   fSearchADUser $UserName $return
	 }
		
	fO365BrokenSession 
	fOnpremBrokenSession 
 
	 return $CustomADUser

}


	
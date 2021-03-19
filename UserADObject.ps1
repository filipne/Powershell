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
	[switch]$licencegroup,
	[parameter(Position=6,Mandatory = $false, ValueFromPipeline= $false,HelpMessage="Shows licence group membership  ")] 
	[switch]$licencedetails,
	[parameter(Position=7,Mandatory = $false, ValueFromPipeline= $false,HelpMessage= "Gets the On Prem version of the mailbox ")] 
	[switch]$OnPrem
	)


begin 
{

	#region Functions
	
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
				
			
				if (!$silent )
				{	
						Write-Verbose "Recipient-RecipientTypeDetails $($UserName.'Recipient-RecipientTypeDetails')"
						
						$RecipientTypeDetails = $UserName."Recipient-RecipientTypeDetails"
						
						if ( !$RecipientTypeDetails )
						{
							# OnPrem mailboxes in OUs that are not synced with O365 do not appear as MailUser on O365 so we are using the onPrem attribute 
							
							$RecipientTypeDetails = "$($UserName.ADRecipientTypeDetails) (OnPrem Only)"
						
						}
						
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
					
						if( $userAccountControl -ne '512' -and   $userAccountControl -ne '544' -and  $userAccountControl -ne '66048' -or $accountexpired  ) 
						{ 	
							Write-Host "$($UserName.SamAccountName)" -foregroundcolor White   -backgroundcolor Red  -NoNewline	
							Write-Host " "  -NoNewline	
							Write-Host "$($UserAccountControlList.Item($UserName.userAccountControl.tostring()))"  -foregroundcolor White   -backgroundcolor Red  -NoNewline	
						}
						elseif ( $UserLogonName -ne $normalizedLogonName -or $UserLogonName.Contains(" ") -and !$Msol )
						{
							 	Write-Host "' $UserLogonName '" -foregroundcolor red   -backgroundcolor yellow  -NoNewline	
								Write-Host  " " -NoNewline
								Write-Host " User logon name is not valid ! $normalizedLogonName  "  -ForegroundColor	Red	-BackgroundColor	yellow  -NoNewline	
						}
						else
						{ 		
								# active account 
								
								Write-host "$UserLogonName" -ForegroundColor DarkBlue -BackgroundColor Green  -NoNewline 
								Write-verbose  "$($UserAccountControlList.Item($UserName.userAccountControl.tostring()))"
						}
												
						Write-Host  " " -NoNewline		
						
						fPasswordStatus   $UserName 	
						
						Write-Host  " " -NoNewline
			
						 $accountexpired  = fAccountExpires  $UserName
						
						#$AccountIsLockedOut = $UserName.LockedOut
						
						if( $UserName.LockedOut )
						{ 
							Write-Host "Account Is Locked Out" -foregroundcolor White    -backgroundcolor DarkYellow 
						}
				
						Write-Host ""
						Write-Host ""
																		
						Write-Host ( $UserName  | fl `
						Description, Title, telephoneNumber, Company, Country,State, Office,  Department ,`
						@{n="Notes(Info)";e={  $_.Info }} ,`
						SamAccountName, UserPrincipalName,`
						@{n="Manager";e={  (Get-ADUser $_.Manager ).Name }} ,`
						HomeDirectory, CanonicalName,`
						Mail,   ipPhone, EmployeeNumber | Out-String).trim() -foregroundcolor Green -NoNewline
						
						 Write-Host ""
						 Write-Host ""
						 
					 
						 Write-Host ( $UserName  | ft @{n="Creator"; e={ $Creator  } },`
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
							
									$Alias = $UserName.Alias
									
							
									Write-Host ( $UserName  | ft   -hide	ExchangeGuid,`
									@{Label="Skype"; Expression = {  if($O365) { $_.EmailAddresses |  ?{ $_.ToString() -like "sip*" } }else{ "`'SIP N/A`'"  }         }},`
									@{Label="Alias"; Expression = { if ($_.EmailAddress) { $OnpremAlias = ( get-recipient $_.EmailAddress).Alias ;  if ( $OnpremAlias -ne  $_.Alias ) {  "`'OnPrem Alias: $OnpremAlias`' `'O365 Alias: $($_.Alias)`'  " }else{  "`'$($_.Alias)`'"    }  }  } },`
									@{Label="TotalItemSize"; Expression = {  if( $_.TotalItemSize ) { $_.TotalItemSize   }else{ "`'MB Size N/A`'"  }         }},`
									@{Label="Database"; Expression = {  if( $_.Database ) { $_.Database   }else{ "`'MB Database N/A`'"  }         }}`
									| Out-String).trim()  -ForegroundColor Cyan 
								
								
								$global:gUserName  = $UserName 
							
							Write-Host ""
			
							$extensionAttributes =  $UserName | select extensionAttribute*
							
							for( $i=1; $i -le 15; $i++)
							{
								$extensionAttributeNumber = "extensionAttribute" + $i
								
								if ( $extensionAttributes.$extensionAttributeNumber)
								{
									$EnabledextensionAttributes = $EnabledextensionAttributes + $extensionAttributeNumber + ": " + $extensionAttributes.$extensionAttributeNumber + "; "
								}
							
							}

							Write-Host ( $UserName  | ft msExchWhenMailboxCreated, Mailbox-WhenCreated, Mailbox-WhenChanged, LastLogonTime,`
							@{Label="LitHold"; Expression = {  $_.LitigationHoldEnabled } },`
							@{Label="Extention Attributes"; Expression = {$EnabledextensionAttributes }} | Out-String).trim() -foregroundcolor Cyan

							Write-Host ""
							
							if ( $UserName.HiddenFromAddressListsEnabled)
							{
								Write-Host ( $UserName  | fl PrimarySmtpAddress | Out-String).Trim()  -foregroundcolor White -BackgroundColor DarkMagenta
							
							}
							else
							{
								Write-Host ( $UserName  | fl PrimarySmtpAddress | Out-String).Trim()  -ForegroundColor Cyan 
							
							}
							
						
							Write-Host ( $UserName  | fl  RemoteRoutingAddress,`
							@{Label="EmailAddresses"; Expression = { $_.proxyAddresses | ?{ $_.ToString()  -like "SMTP*" }  | sort   } }`
							| Out-String).Trim() -foregroundcolor Cyan
						
							Write-Host ""
						
					
							if ( $UserName.msExchArchiveName)
							{
								# Dsplay Online Archive mailbox properties 
										
								Write-Host ( $UserName  | ft  -HideTableHeaders 	ArchiveGuid | Out-String).trim()   -foregroundcolor Cyan
								Write-Host ""
								Write-Host ( $UserName  | ft `
								msExchArchiveName, ArchiveMailboxStats-TotalItemSize, ArchiveQuota , ArchiveDatabase | Out-String).trim() -foregroundcolor Cyan
								
							}
							else
							{
								#Write-Host ""
								Write-Host " " -NoNewline
								Write-Host " No Online archive" -ForegroundColor  Blue	-BackgroundColor	Yellow
							
							}
							Write-Host ""
							
							$global:gUserName = $UserName 
							
							fAutomap  $UserName
							fMailboxForward $UserName 
							fOutOfOffice $UserName
							fAcceptMessagesOnlyFrom $UserName

							if ($UserName.ADRecipientTypeDetails -eq "RoomMailbox"  -or   $UserName.ADRecipientTypeDetails -eq "EquipmentMailbox" -and !$Silent  )
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
					
					.\Get-UserLicenseDetails.ps1 $UserName.UserPrincipalName
								
				}
				
				
				Write-Host ""
				
	} ########### END function fDisplayADUser

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
				
				$dashCount = ($UserName.ToCharArray()  | Where-Object {$_ -eq "-" } | Measure-Object).Count
				
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
					Write-Host ""
					Write-Host ($AllSearch | sort msExchRecipientTypeDetails,userAccountControl , DisplayName  |  ft ObjectGUID, DisplayName,`
					@{Label="msExchRecipientTypeDetails"; Expression= { $RecipientTypeDetailsList.Item( ($_.msExchRecipientTypeDetails).tostring() )  }  } ,`
					@{Label="Disabled?"; Expression= { if ($_.userAccountControl -eq 512 ){ "False"  }else{"True"}   }  } ,`
					EmailAddress, Description   -AutoSize -Wrap  | Out-String).trim()  -foregroundcolor Cyan
					
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
			Write-Verbose "			`$passexp =  (get-aduser $($AllSearch.ObjectGUID)  -Properties msDS-UserPasswordExpiryTimeComputed).'msDS-UserPasswordExpiryTimeComputed' "
			
			$passexp =  (get-aduser $AllSearch.ObjectGUID  -Properties msDS-UserPasswordExpiryTimeComputed)."msDS-UserPasswordExpiryTimeComputed"
			
			$CustomADUser | Add-Member -NotePropertyName "PWExpiration" -NotePropertyValue $passexp -Force | Out-Null
						
			$CustomADUser | Add-Member -NotePropertyName "UserLastLogon" -NotePropertyValue  ( fConvertADDate $CustomADUser.LastLogonTimestamp)  -Force | Out-Null
			
				
			$Creator = try{  ( Get-Acl "ad:\$($CustomADUser.distinguishedname)" -ErrorAction SilentlyContinue    ).Owner }catch{} 
			
		
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
									
					if ( $CustomADUser.msExchRecipientTypeDetails )
					{				
														
						$rtd = $RecipientTypeDetailsList.Item( ($CustomADUser.msExchRecipientTypeDetails).tostring() ) 
						
						if ( $RecipientTypeDetailsList.Item( ($CustomADUser.msExchRecipientTypeDetails).tostring() ) -like "Remote*")
						{
							Write-verbose "776 Remote $rtd " 

							if ( $O365)
							{
							
								Write-Verbose "get-O365mailbox -ErrorAction silent $($AllSearch.UserPrincipalName ) "
								
								$Mailbox = try { get-O365mailbox -ErrorAction silent $AllSearch.UserPrincipalName  }catch{}
								
								Write-Verbose "get-O365mailboxStatistics  -ErrorAction silent $($AllSearch.UserPrincipalName ) "
								
								$Mailboxstats = try {  get-O365mailboxStatistics -WarningAction silent    -ErrorAction SilentlyContinue $AllSearch.UserPrincipalName }catch{} 	
								
								Write-Verbose "get-O365mailboxStatistics -Archive  -ErrorAction silent $($AllSearch.UserPrincipalName ) " 
								
								$ArchiveMailboxstats = try { get-O365mailboxStatistics -Archive  -WarningAction silent   -ErrorAction SilentlyContinue $AllSearch.UserPrincipalName }catch{} 	
							
							}
							else
							{
									Write-verbose "get-remotemailbox -ErrorAction silent $($AllSearch.UserPrincipalName ) " 
									
									$Mailbox = try { get-remotemailbox -ErrorAction silent $AllSearch.UserPrincipalName  }catch{}
							}
		
					}
						else
						{
							Write-verbose  "get-mailbox -ErrorAction silent $($AllSearch.UserPrincipalName) " 
							$Mailbox = try { get-mailbox -ErrorAction silent $AllSearch.UserPrincipalName  }catch{}
							
							Write-verbose "get-mailboxStatistics  -ErrorAction silent $($AllSearch.UserPrincipalName ) " 
							$Mailboxstats = try { get-mailboxStatistics  -ErrorAction SilentlyContinue $AllSearch.UserPrincipalName }catch{} 	
							
							Write-verbose "`$ArchiveMailboxstats = get-mailboxStatistics -Archive  -ErrorAction silent $($AllSearch.UserPrincipalName ) " 
							$ArchiveMailboxstats = try { get-mailboxStatistics -Archive  -ErrorAction SilentlyContinue $AllSearch.UserPrincipalName }catch{} 	
						
						}
						
					}# if ( $CustomADUser.msExchRecipientTypeDetails )
			
					if ( $Mailbox)
					{
												
						$rootMailboxProperties = $Mailbox | Get-Member -ErrorAction SilentlyContinue | ? { $_.MemberType -match "Property"} 
						
						Write-verbose  "rootMailboxProperties  count : $($rootMailboxProperties.count)" 

						
						$rootMailboxProperties | % {
						
			
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
					
						}
											
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
							
								}
						
							}	
					
						$remotemailbox = get-remotemailbox $UserName.UserPrincipalName	  -ErrorAction silent
						
						$CustomADUser 	| Add-Member  -NotePropertyName "RemoteRoutingAddress" -NotePropertyValue $($remotemailbox.RemoteRoutingAddress)  -Force | Out-Null
						
						$accountdisbaled = if ($CustomADUser.userAccountControl -eq 512 ){ "False"  }else{"True"} 
						
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
			Write-Host "  System FAILURE  fNormalizeUsername : No Value for input variable  UserLogonName  provided  " -ForegroundColor Yellow -BackgroundColor Red
			Write-Host ""
			return 	
		}

		#$pattern = '[^a-zA-Z0-9.-]'
		#$pattern = '[^a-zA-Z0-9.-_]'
		#$pattern = '[^a-zA-Z0-9._\-]'
		
		$pattern = '[^a-zA-Z0-9._\-]'
		
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
					
					$PasswordExpires =   Get-Date  ( [datetime]::FromFileTime($UserParameter.PWExpiration))  -format dd/MMM/yyyy
					
					Write-Host  "$($UserParameter.PasswordLastSet)"  -ForegroundColor DarkCyan -NoNewline
					Write-Host " " -NoNewline
					
					if ( $PasswordExpires )
					{
					
												
						$DaysToExpire = ( New-TimeSpan -end $PasswordExpires ).Days
						
												
						if (  $DaysToExpire -lt  0 )
						{
							Write-Host  "Password Expired on $PasswordExpires"   -ForegroundColor Red	-BackgroundColor	yellow	 -NoNewline	
						}
						elseif (  $DaysToExpire -lt  $DaysToAlert  )
						{
							Write-Host  "Password Expires on $PasswordExpires "  -ForegroundColor Yellow -NoNewline
						}
						else
						{
							
						}
							
				 	}
						
			}
		
		}
		#>

	}### fPasswordStatus ($UserParameter)

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
				
				Write-Verbose "  $($MyInvocation.InvocationName); Line [$($MyInvocation.ScriptLineNumber)]: $($MyInvocation.line); Days to expire : $DaysToExpire "
								
				if (  $DaysToExpire -lt  0 )
				{
					
					if ( !$silent)
					{
						Write-Host  "Account Expired on $AccountExpires "   -ForegroundColor Red	-BackgroundColor	yellow		
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
				
		if (!$UserParameter ) { return;  }
		
		#Write-Host ( $UserName  | fl *forward* | Out-String).Trim() -foregroundcolor Red

		
		if ($UserParameter.ForwardingAddress -or $UserParameter.ForwardingSmtpAddress)
		{

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
						if (  $O365  )
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
							
							$FullAccess = " Remote N/A. To get Full Access rights create O365 Session "
							
							Write-Host ( $FullAccess | ft -AutoSize | Out-String) -foregroundcolor Cyan
							
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
					
							Write-Host "Out of Office: "  -foregroundcolor Cyan -BackgroundColor BLUE -NoNewline

							if($($State.AutoReplyState) -eq "Enabled") { Write-Host " ENABLED "		-ForegroundColor DarkBlue -BackgroundColor Green   -NoNewline }

							elseif ($($State.AutoReplyState) -eq "Scheduled") 	
							{ 
									$StartTime = Get-Date $($State.StartTime)  -format "dd/MM/yyyy HH:mm"
									$EndTime  = Get-Date $($State.EndTime)   -format "dd/MM/yyyy HH:mm"
									Write-Host " SCHEDULED  Start Time: $StartTime ; End Time: $EndTime " -foregroundcolor White -BackgroundColor DarkGreen -NoNewline
							}
						
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
			{ 	Write-Host "Accept Messages Only from " -ForegroundColor Cyan }
			
			
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
		
		$Membership =  @( $Membership | ?{ $_.Name -like 'O365_E*' } )
		
		if ( $Membership )
		{

				$Length = $Membership.count
			
				$Membership = $Membership  | select Name, `
				GroupCategory,`
				@{  Label="Email" ; Expression= { "  Email: " +  (get-adgroup  $_.ObjectGUID  -Properties mail).mail  }}, `
				@{  Label="Description" ; Expression= { "  '$($_.Description)' " } }
				

				Write-Host "[$Length] licence  groups found for ' $($UserParameter.Name) ' " -ForegroundColor Cyan
				
				Write-Host ""
							
				Write-Host ( $Membership  | sort GroupCategory, Name | ft -autosize -HideTableHeaders | Out-String).trim()  -ForegroundColor White -BackgroundColor DarkGreen
				
			
		}
		else
		{
		
			Write-Host ""
			Write-Host "	" -NoNewline
			Write-Host "  ATTENTION :  No licence groups' membership  found for  ' $($UserParameter.Name) '" -ForegroundColor  Blue	-BackgroundColor	Yellow
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
				$msExchDelegateListBL | %{
					
					if ( $_)
					{
						$msExchDelegateListRecipients += try { ( get-recipient $_ -erroraction silent).Name }catch{}
						$msExchDelegateListRecipients+= try { ( get-o365recipient $_ -erroraction silent).name }catch{}
					}
				}
		
	Write-Verbose  "msExchDelegateListlink  $msExchDelegateListlink " 	
		
			$msExchDelegateListlink | %{
					
					#write-host "1401 msExchDelegateListlink  $_ " -ForegroundColor Magenta 
					if ( $_)
					{
						$msExchDelegateListRecipients  +=  if ( get-recipient $_ -erroraction silent){ ( get-recipient $_ -erroraction silent).Name}
						$msExchDelegateListRecipients  += try { if (get-o365recipient $_ -erroraction silent ){  ( get-o365recipient $_ -erroraction silent).name }   }catch{}
					}
			}
			
			$global:gmsExchDelegateListRecipients = $msExchDelegateListRecipients
			
			if ( $msExchDelegateListRecipients )
			{
				Write-Host ""
				Write-Host "Automap: " -foregroundcolor Cyan -BackgroundColor BLUE -NoNewline
				Write-Host " " -NoNewline
				Write-Host ( $msExchDelegateListRecipients  -join " " | ft | Out-String).Trim() -foregroundcolor Cyan
			}
			
			Write-Verbose "1688"
		
		
	} ########### END  fAutomap  ($UserParameter)  ###########
	
	function fRoomMailbox ($UserParameter)
	{
		
		Write-Verbose " InvocationName: $($MyInvocation.InvocationName); Line Number [$($MyInvocation.ScriptLineNumber)]; Line: $($MyInvocation.line)"
		
		$BookingRecipients = @()
					
		Write-Host ( $UserParameter | select ResourceCapacity, ResourceCustom, ModerationEnabled, ModeratedBy  | ft -AutoSize  | Out-String).Trim() -foregroundcolor Cyan
		
					
		# https://technet.microsoft.com/en-CA/library/ms.exch.eac.EditRoomMailbox_ResourceDelegates(EXCHG.150).aspx?v=15.0.1104.0&l=0
		
			
		$MailboxCalendarConfiguration = Get-MailboxCalendarConfiguration $($UserParameter.SamAccountName) -ErrorAction silent
		
		if ( !$MailboxCalendarConfiguration -and $O365Session)
		{
				$MailboxCalendarConfiguration = Get-MailboxCalendarConfiguration $($UserParameter.SamAccountName) -ErrorAction silent
		}
		
		$CalendarProcessing = Get-CalendarProcessing -ErrorAction silent $($UserParameter.SamAccountName)
		
		if ( !$CalendarProcessing -and $O365Session )
		{
				$CalendarProcessing = Get-O365CalendarProcessing -ErrorAction silent $($UserParameter.SamAccountName)
		}
		
		$global:gCalendarProcessing  = $CalendarProcessing 
		
		
		$ResourceDelegates = $CalendarProcessing.ResourceDelegates
		
		$ResourceDelegatesString = @()
		
		$ResourceDelegates | %{
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

		if ( $BookInPolicy )
		{
			$BookInPolicy = $BookInPolicy | sort Name 
			
			$global:gBookInPolicy = $BookInPolicy
	
			#$BookingRecipients = @()
			
			$BookInPolicy | %{ 
	
					$recipient = try {Get-Recipient $_  -ErrorAction silentlycontinue }catch{} 
			
					if ( $recipient )
					{
						$BookingRecipients += try {Get-Recipient $_  -ErrorAction silentlycontinue}catch{} 
					}
					else
					{
						$BookingRecipients += $_
					
					}
		
			}

			$BookingRecipients = $BookingRecipients | sort Name

		}
	
	
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
			
			Write-Host ""
			Write-Host "[$($BookingRecipients.length)] Entities can book : " -BackgroundColor DarkBlue -ForegroundColor Cyan
			Write-Host ""
			
			Write-Host ( $BookingRecipients | ft   Name, RecipientType, guid -AutoSize  | Out-String).trim() -BackgroundColor Blue -ForegroundColor Cyan
	
			if ( $($BookingRecipients.length)  )
			{
				Write-Host ""
								
				$global:gBookingRecipients = @()
				
				$BookingRecipients | %{ if ( $_.guid ){ $global:BookingRecipients += $_.guid.tostring() } }
				
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
			Write-Host ".\Verify-RecipientList.ps1 | %{ Try{`$BookingRecipients +=  `$_.guid.ToString() }catch{} } ;  "   -ForegroundColor Yellow
			
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


		if ( $UserParameter.SamAccountName)
		{
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
			Write-Host "Add-MailboxFolderPermission -identity  $CalendarString" -ForegroundColor Yellow 
			Write-Host ""
			Write-Host "Calendar Permissions  : "  -ForegroundColor Cyan -BackgroundColor Blue
			Write-Host ""
			Write-Host ($CalendarPermissions | ft User, AccessRights -AutoSize | Out-String).Trim() -foregroundcolor Cyan
			
		}
	
	}########### END function fRoomMailbox () ###########



	#endregion Functions
	
	#region	Initialization 
			
		Write-Verbose "  $($MyInvocation.InvocationName); Line [$($MyInvocation.ScriptLineNumber)]: $($MyInvocation.line) ;  Initializing  .."

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
		
		Write-Host ""
		Write-Host ""
	
		
		# Detect existing, prefixed with 'O365' remote session to Office 365. 	
			
		if ( ($O365 = try{ get-command "get-O365mailbox"-ErrorAction SilentlyContinue }catch{}) -and !$OnPrem )
		{
			$O365 = $true
			Write-Host "[usra] O365P session" -ForegroundColor DarkBlue -BackgroundColor Green
		}
		else
		{
			$O365 = $false
			Write-Host "[usra] OnPrem only session" -ForegroundColor Cyan	-BackgroundColor	Blue
		}

	#endregion Initialization 


} # end begin 

process 
{

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
	
		
	 $CustomADUser  =   fSearchADUser $UserName $return
	 
	 

	 return $CustomADUser

}


	
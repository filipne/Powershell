<#
.SYNOPSIS
    Get-GroupMembers displays a list of  members of a provided AD Group
	and exports the list to a CSV file (if requested )
.DESCRIPTION
    .<Text>
.PARAMETER One
    <Parameter One explanation>
.PARAMETER Two
    <Parameter Two explanation>
	
.EXAMPLE
    C:\PS> 
    <Description of example>
.NOTES
    Author: Filip Neshev
    Date:   2011    
#>

param(
    [parameter(Position=0,Mandatory=$true,ValueFromPipeline=$true,HelpMessage="Name ")]
	$GroupName, 
    [parameter(Position=1,Mandatory=$false,ValueFromPipeline=$true,HelpMessage="Highlight")]
	$Highlight,   
	[parameter(Position=2,Mandatory=$false,ValueFromPipeline=$true,HelpMessage="CSVExport")]
	[switch]$CSVExport,
	[parameter(Position=4,Mandatory=$false,ValueFromPipeline=$true,HelpMessage="Return objects")]
	[switch]$return,
	[parameter(Position=5,Mandatory=$false,ValueFromPipeline=$true,HelpMessage="Export")]
	[switch]$Export,
	[parameter(Position=6,Mandatory=$false,ValueFromPipeline=$true,HelpMessage="Silent")]
	[switch]$Silent,
	[parameter(Position=7,Mandatory=$false,ValueFromPipeline=$true,HelpMessage="Server")]
	[switch]$usedc,
	[parameter(Position=8,Mandatory=$false,ValueFromPipeline=$true,HelpMessage="Server")]
	$ExportPath = "C:\scripts\_Logs\Groups\",
	[parameter(Position=9,Mandatory=$false,ValueFromPipeline=$true,HelpMessage="TXTExport")]
	[switch]$TXTExport
	)
    
	#Write-Host  "41 usedc $usedc" -ForegroundColor Red 
	
	#Write-Host "Get-GroupMembers.ps1	Created by Filip Neshev, 2011	filipne@yahoo.com" -ForegroundColor Cyan  -backgroundcolor  DarkGray	
	
	begin 
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
				Write-Host "Get-GroupMembers.ps1	Created by Filip Neshev, 2011	filipne@yahoo.com" -ForegroundColor Cyan  -backgroundcolor  DarkGray	
				Write-Host ""
		}
	
	
		Write-Host ""
		
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
		
		$groupTypes = @{
		"2"		=	 	"Global distribution group"
		"4" 	= 		"Domain local distribution group"
		"8"		= 		"Universal distribution group"	
		"-2147483646"  = "Global security group"		
		"-2147483644"  = "Domain local security group" 
		"-2147483640"  = "Universal security group"		
		}
		
		
		
 }   
process 
{
		$mcount = 0
		$Members = @()
		$GroupMembersguids =@()

		if($usedc)
		{
		
			Write-Host ".\GroupADObject $GroupName -usedc" -ForegroundColor Yellow
			
			Write-Host " "
			
			$GroupObject = .\GroupADObject $GroupName -usedc
			
			$Server =  .\Get-DC.ps1 
	
		}
		else
		{
			Write-Host ".\GroupADObject $GroupName " -ForegroundColor Yellow
			
			Write-Host " "
			
			if ( $Silent )
			{
			
				$GroupObject = .\GroupADObject $GroupName  -silent 
				
			}
			else
			{
				$GroupObject = .\GroupADObject $GroupName 
			}
		}
		
		$global:gGroupObject = $GroupObject
		
		
		$GName  = $GroupObject.Name
		
		Write-Host " "

		write-verbose  "113 RecipientTypeDetails  $($GroupObject.RecipientTypeDetails)" 
		
		if ( !( $GroupObject.RecipientTypeDetails))
		{
				# Security not mail enabled group
				
				Write-Verbose "	Security not mail enabled"
											
				if ($usedc )
				{
					
					Write-Host "`$AllGroupMembers = @( Get-ADGroupMember $($GroupObject.tostring())  -Server $Server | sort objectGUID ) "  -ForegroundColor Yellow
					
					$AllGroupMembers = @( Get-ADGroupMember  $GroupObject.tostring() -Server $Server   | sort objectGUID | Get-Unique )
					
				}
				else
				{
					Write-Host "`$AllGroupMembers = @( Get-ADGroupMember $($GroupObject.tostring())  | sort objectGUID ) "  -ForegroundColor Yellow
					
					$AllGroupMembers = @( Get-ADGroupMember  $GroupObject.tostring()   | sort objectGUID | Get-Unique )
										
				}
				
				Write-Host ""
				Write-Host "[$($AllGroupMembers.count)]  All Types Group member(s) in group '$GName'" -ForegroundColor Cyan	-BackgroundColor	Blue
				Write-Host ""
				
				Write-Host "`$GroupMembersUsers = @( `$AllGroupMembers  | Get-ADUser -Properties Name,userAccountControl ,  objectguid , ObjectClass , mail, Objectguid,  DisplayName , UserPrincipalName, title, department, physicalDeliveryOfficeName, msExchRecipientTypeDetails ,Description, EmployeeType   -ErrorAction silent ) "  -ForegroundColor Yellow
				
				$GroupMembersUsers  =  @( $AllGroupMembers  | Get-ADUser -Properties Name,userAccountControl ,  objectguid , ObjectClass , mail, Objectguid,  DisplayName , UserPrincipalName, title, department, physicalDeliveryOfficeName, msExchRecipientTypeDetails , Description, EmployeeType   -ErrorAction silent )
				
				#Write-Host ""
				Write-Host "`$GroupMembersGroups = @( `$AllGroupMembers | Get-ADGroup  -Properties * -ErrorAction silent ) "  -ForegroundColor Yellow
						
				$GroupMembersGroups   =  @( $AllGroupMembers | Get-ADGroup  -Properties * -ErrorAction silent )
				
				$Length = $( $GroupMembersUsers.count)
				
				Write-Host ""
				Write-Host "[$Length] User member(s) in group '$GName' " -ForegroundColor Cyan	-BackgroundColor	Blue
				Write-Host "[$($GroupMembersGroups.count)] Group member(s) in group '$GName'" -ForegroundColor Cyan	-BackgroundColor	Blue
				Write-Host "" 
			
				#$Members =  $GroupMembersUsers + $GroupMembersGroups
				$Members =  @( $GroupMembersUsers ) 
				
				$global:gMembers = $Members
			
				
				if ( !$Silent )
				{
					Write-Host "All Group members Custom View ( sort Type, Enabled, Name ) :" -ForegroundColor Cyan	-BackgroundColor	Blue
				}
					
					$Members = @( $Members | select Name, DisplayName , `
					@{Label = "Enabled"; Expression = { if ( $_.userAccountControl -eq 512 ){ "True" } else { "False" }  } },`
					@{Label=  "Type"; Expression= { $Type = $RecipientTypeDetailsList.Item( ($_.msExchRecipientTypeDetails).tostring() ); if ($Type){$Type }else{$_.msExchRecipientTypeDetails }  }  } , EmployeeType,`
					@{Label = "PrimarySMTPAddress"; Expression = { if ($_.PrimarySMTPAddress){  $_.PrimarySMTPAddress  }else{ $_.mail  } } },`
					@{Label = "Guid"; Expression = { if ($_.Guid){$_.Guid} else {$_.Objectguid }  } },`
					UserPrincipalName, title, department, physicalDeliveryOfficeName | sort  Type,Enabled, Name )
					
				#}
				
					#$global:gMembers = $Members
			
		} # Security not mail enabled group
		elseif ( $GroupObject.RecipientTypeDetails -like "*Distribution*" -or $GroupObject.RecipientTypeDetails  -like "*Mail*" )
		{
				# Distribution Group or Mail Enabled Security 
				
				Write-Verbose "	Distribution or Mail Enabled Security"
	
				if ($usedc )
				{
					
					Write-Host "`$Members = @(Get-DistributionGroupMember $($GroupObject.PrimarySMTPAddress) -ResultSize unlimited  -DomainController $Server  | sort RecipientType, Name )"  -ForegroundColor Yellow
					
					$Members =@( Get-DistributionGroupMember  -ResultSize unlimited $GroupObject.PrimarySMTPAddress -DomainController $Server   | sort RecipientType, Name )
					
				}
				else
				{
					Write-Host "`$Members =  @(Get-DistributionGroupMember $($GroupObject.PrimarySMTPAddress) -ResultSize unlimited   | sort RecipientType, Name ) "  -ForegroundColor Yellow
										
					$Members = @( Get-DistributionGroupMember  -ResultSize unlimited  $GroupObject.PrimarySMTPAddress.tostring()  | sort RecipientType, Name  )
					
				}
				Write-Host ""
				Write-Host "[$($Members.count)] members found in group '$GName' :" -ForegroundColor Cyan 	-BackgroundColor	Blue
				
				$Members = @( $Members | select Name,`
				@{Label = "Enabled"; Expression = { if ( $_.userAccountControl -eq 512 ){ "False" } else { "True" }  } },`
				@{Label = "Type"; Expression = { $_.RecipientTypeDetails } },`
				@{Label = "PrimarySMTPAddress"; Expression = { if ($_.PrimarySMTPAddress){  $_.PrimarySMTPAddress  }else{ $_.mail  } } },`
				@{Label = "Guid"; Expression = { if ($_.Guid){$_.Guid} else {$_.Objectguid }  } },EmployeeType,`
				DisplayName , UserPrincipalName, title, department, physicalDeliveryOfficeName | sort Type,Enabled, Name )
				
				#Write-Host "[$($Members.count)] members found in group '$GName' :" -ForegroundColor Cyan	-BackgroundColo DarkGreen
	
		} # Distribution Group or Mail Enabled Security 
		else
		{
				# Something different from distribution , securiy and mail enabled security
				
				Write-Verbose "	Something different from distribution , securiy and mail enabled security"
				
				Write-Host "`$Members = @( Get-ADGroup $($GroupObject.ObjectGUID) -Properties member  | Select-Object -ExpandProperty member | Get-ADObject -Properties *)" -ForegroundColor Yellow

				# Must be used to show mail contact and none mail enabled  members !
				
				if ($usedc )
				{
					$Members = @( Get-ADGroup $GroupObject.ObjectGUID -Properties member   | Select-Object -ExpandProperty member | Get-ADObject -Properties *)
				}
				else
				{
					$Members = @( Get-ADGroup $GroupObject.ObjectGUID -Properties member    | Select-Object -ExpandProperty member | Get-ADObject  -Properties *)
				}
				
				$Members = $Members | sort ObjectClass, RecipientType,  Name
		
		} # Something differenet from distribution , securiy and mail enabled security
					
		Write-Host "" 
				
	
		if($Members)
		{

			#$GName =  [string]  $GroupObject.GroupName
			
			$GName =  [string]  $GroupObject.Name
			
			$ExportGroupName = $GName.Replace(" ", "_")
			$ExportGroupName = $ExportGroupName.Replace("'", "_")
			
			$PathU = "U:\scripts\_Logs\Groups\"
			$PathC = $ExportPath
			
			if ( !$Silent )
			{
				
				$GroupMembersGroupscount = $GroupMembersGroups.count
				
				$global:goMembers = $Members 
				
				# Display group members 	
		
				if ( $GroupObject.RecipientTypeDetails)
				{
					Write-Host ( $Members |  ft DisplayName , Name,  Enabled , Type, EmployeeType, PrimarySMTPAddress,  guid   -AutoSize | Out-String).Trim() -foregroundcolor Cyan
				}
				else
				{
					Write-Host ( $Members |  ft DisplayName , Enabled , Type,  EmployeeType, PrimarySMTPAddress, UserPrincipalName, guid   -AutoSize | Out-String).Trim() -foregroundcolor Cyan
				}
				
				Write-Host ""
				Write-Host "[$($Members.count)] members found in group $($GroupObject.Name) " -ForegroundColor Cyan	-BackgroundColo Blue
		
				Write-Host "" 
		
				if ( $GroupMembersGroups )
				{
				
					#$groupTypes.Item($_.groupType)
				
					$global:gGroupMembersGroups  = $GroupMembersGroups
					
					$GroupMembersGroups  = $GroupMembersGroups  | select Name,Created,Modified,`
					@{ Label = "Type"; Expression = { if ( $_.RecipientTypeDetails ){ $_.RecipientTypeDetails } elseif ( get-recipient  $_.objectguid.tostring() -erroraction silent  ) { ( get-recipient $_.objectguid.tostring() -erroraction silent ).RecipientTypeDetails }else {  $groupTypes.Item( $_.groupType.ToString()) } } }, `
					#@{ Label = "Type"; Expression = { $groupTypes.Item( $_.groupType.ToString() ) }  }, `
					@{Label = "PrimarySMTPAddress"; Expression = { if ($_.PrimarySMTPAddress){  $_.PrimarySMTPAddress  }else{ $_.mail  } } },`
					@{Label = "Guid"; Expression = { if ($_.Guid){$_.Guid} else {$_.Objectguid }  } },`
					DisplayName, Description   | sort Type,Name
										
					Write-Host "[$GroupMembersGroupscount]  Group members in group '$GName':" -ForegroundColor Cyan	-BackgroundColor	Blue
					Write-Host "" 
					
					Write-Host ( $GroupMembersGroups  |  ft DisplayName ,  Type, PrimarySMTPAddress, guid, Description  -Wrap  -AutoSize | Out-String).Trim() -foregroundcolor Cyan
					
				}
			
				Write-Host "" 
			
				
			
				#Write-Host ( $GroupADMembers  | ft  -AutoSize Name, DisplayName, RecipientTypeDetails, PrimarySMTPAddress, @{Label="Guid"; Expression= { if ($_.Guid){$_.Guid} else {$_.Objectguid }        } },  ObjectClass  | Out-String).Trim() -foregroundcolor Cyan

			}
			
			#$global:gMembers = $Members
	
			if ( $CSVExport  )
			{
				Write-Host ""
				Write-Host "INFO:  CSV Export requested !" -BackgroundColor Blue  -ForegroundColor White
				Write-Host ""		
			
				$File = "@Get-GroupMembers@" + "$ExportGroupName.csv"
				$FullPathU = "U:\Scripts\_Exports\" + $File
				$FullPathC = "C:\Scripts\_Exports\" + $File

				Write-Host "$FullPathC"   -foregroundcolor Yellow

				write-host ""
				write-host "`$Members | Select FirstName,  LastName,  EmailAddress, Title, Office, PhoneNumber, ParentContainer   | Export-Csv -Encoding UTF8  $FullPathC -NoTypeInformation "   -ForegroundColor Yellow 
				write-host ""
				
				$Members | select * | Export-Csv -Encoding UTF8  $FullPathC -NoTypeInformation
				
				$global:gMembers = $Members

				
			}
	
			if ($TXTExport  )
			{
				Write-Host ""
				Write-Host "INFO: TXT Export requested !" -BackgroundColor Blue -ForegroundColor White
				Write-Host ""		
				
				$File = "@Get-GroupMembers@" + "$ExportGroupName" + "_GroupInfo" + ".txt"
				$FullPathU = $PathU + $File
				$FullPathC = $PathC + $File
				
				$FileNameGroupMembers = "@Get-GroupMembers@" + "$ExportGroupName" + "_GroupMembers" + ".txt"
				$FullFilePathGroupMembersU = $PathU + $FileNameGroupMembers
				$FullFilePathGroupMembersC = $PathC + $FileNameGroupMembers
				Write-Host ""

				Write-Host "`$gGroupObject | Select-Object * | out-File $FullPathC " -foregroundcolor Yellow
				Write-Host "" 

				Write-Host "`$gMembers | Select Name | out-File $FullFilePathGroupMembersC	" -foregroundcolor Yellow
				
				$GroupObject | Select-Object * | out-File $FullPathC
				
				#$Members | Select Name | out-File $FullFilePathGroupMembersC
				
				($Members | Select PrimarySmtpAddress).PrimarySmtpAddress   | out-File $FullFilePathGroupMembersC		
				
				return ($Members | Select PrimarySmtpAddress).PrimarySmtpAddress 
	
			}
				
			if ($Export)
			{
				Write-Host ""
				Write-Host "INFO: Export requested !" -BackgroundColor Blue -ForegroundColor White
				Write-Host ""
				
				$File = "@Get-GroupMembers@" + "$ExportGroupName" + "_GroupInfo" + ".xml"
				$FullPathC = $PathC + $File
				
				Write-Host " " 
				Write-Host "`$GroupObject | Export-Clixml '$FullPathC' " -ForegroundColor Yellow
				Write-Host " " 
				
				$GroupObject | Export-Clixml $FullPathC	
				
				$File = "@Get-GroupMembers@" + "$ExportGroupName" + "_GroupInfo" + ".txt"
				$FullPathU = $PathU + $File
				$FullPathC = $PathC + $File
				
				$FileNameGroupMembers = "@Get-GroupMembers@" + "$ExportGroupName" + "_GroupMembers" + ".txt"
				$FullFilePathGroupMembersU = $PathU + $FileNameGroupMembers
				$FullFilePathGroupMembersC = $PathC + $FileNameGroupMembers
				Write-Host ""

				#Write-Host "$FullPathU " -foregroundcolor Yellow
				Write-Host "`$gGroupObject | Select-Object * | out-File $FullPathC" -ForegroundColor Yellow
				Write-Host "" 
				#Write-Host "$FullFilePathGroupMembersU" -foregroundcolor Yellow
				Write-Host "( `$Members | Select PrimarySmtpAddress).PrimarySmtpAddress  | out-File $FullFilePathGroupMembersC" -ForegroundColor Yellow
				
				#$GroupObject | Select-Object * | out-File $FullPathU
				$GroupObject | Select-Object * | out-File $FullPathC
				
				#$Members | Select Name | out-File $FullFilePathGroupMembersU	
				($Members | Select PrimarySmtpAddress).PrimarySmtpAddress   | out-File $FullFilePathGroupMembersC		
		
			}
			
			<#
			$AllHighLights = @()
		    
		    foreach($Member in $Members)
		    {
			
			        $mcount++
					$DisplayName = $Member.Name
					$MembersList = $MembersList + $DisplayName + ", "
				    $match = $false
						
					foreach($HL in $Highlight)
					{
							#Write-Host "HL: $HL"
							if($HL)
							{ 
							  $HL = [string]$HL
							  $HL = $HL.Trim()
							}
												
							if( $HL -like $DisplayName )
							{

								$AllHighLights += $Member
								
								$match = $true
							}
													
						}
			
				if (!$match ) 
				{ 
					#Write-Host "$DisplayName,"	
				}			
		   	
			}
			#>
			
			Write-Host ""
			
			Write-Host ( $AllHighLights | ft  -AutoSize  Name, RecipientType, PrimarySMTPAddress | Out-String).Trim()  -ForegroundColor Green	


}
else
{
	Write-Host ""
	Write-Host "	" -NoNewline
	Write-Host "  ATTENTION : No members found in group ' $GName '   " -ForegroundColor  Blue	-BackgroundColor	Yellow
	Write-Host ""
}

#Write-Host ""
Write-Host "  End of Script Get-GroupMembers.ps1" -ForegroundColor DarkBlue	-BackgroundColor Gray
Write-Host ""

if($return) { Return $Members }

}


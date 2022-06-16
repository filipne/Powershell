<#
.SYNOPSIS
    Verify-List.ps1 reads a list of input objects and verifies them in AD
.DESCRIPTION
    
	Requirements:  
	----------
	Verify-List.ps1  uses module  ActiveDirectory which is part of Microsoft RSAT (Remote System Adminsistraion Tools)
	
	To enable ActiveDirectory module for Powerhsell follow instructions at 
	https://4sysops.com/archives/how-to-install-the-powershell-active-directory-module/ 
	
	Add-GroupADMembers  does NOT neeed any Exchange rights. All Exchange related information is read from AD. 
	----------
	
		Dependances : 
	
	- .\Read-List.ps1
	- .\GroupObject.ps1
	- .\UserObject.ps1
	
	------------
.PARAMETER List
    This is the list of items to be read. Accepts:
	
	- Comma separated values
	- Array of objects. 
	
.PARAMETER ItemType
    Item Type. Accepts : 
	- User
	- Group

.PARAMETER CSVItemField
   This is the name of the field in the CSV file from which values will be read.
   Default is 'Name'. 
   
.PARAMETER   $MinItemLength
	If a list item length is less that this number , it will not be verified thus avoding long search reasults in User and Group Object functions

	
.EXAMPLE
    C:\PS> 
    <Description of example>
.NOTES
    Author: Filip Neshev; filipne@yahoo.com
    Date: August 2016        
#>

param(
    [parameter(Position=0,Mandatory=$false,ValueFromPipeline=$true,HelpMessage="Input List ")]
	$List,
	[parameter(Position=1,Mandatory=$false,ValueFromPipeline=$true,HelpMessage=" Type of list members ") ]
	[ValidateSet("User","Group") ]
	[string]$ItemType="User",
	[parameter(Position=2,Mandatory=$false,ValueFromPipeline=$false,HelpMessage="The name of the Item field to read from a CSV file ")]
	[string]$CSVItemField="Name",
	[parameter(Position=3,Mandatory=$false,ValueFromPipeline=$false,HelpMessage="The length of the smallest item that will be read ")]
	[int]$MinItemLength=3,
   	[parameter(Position=4,Mandatory=$false,ValueFromPipeline=$false,HelpMessage="Display brief user info ")]
	[switch]$brief
)
	
Write-Host "	Starting 'Verify-List.ps1' ; Created by Filip Neshev August  2014; filipne@yahoo.com" -ForegroundColor  DarkGray	
	
	#Variables
	
	$SkipPath = "Temp\"
	$VerifiedObjects = @()
	$NotFound = @()
	
	
	
	Write-Host ""
	Write-Host " List members type is :" -ForegroundColor Cyan	-BackgroundColor	Blue -NoNewline
	Write-Host $ItemType -ForegroundColor Cyan
	Write-Host ""
	
	Write-Host ""
	Write-Host "##-CALL-##> " -ForegroundColor Cyan -NoNewline
	Write-Host " `$List = .\Read-List.ps1 -Items  $List " -ForegroundColor Yellow -NoNewline
	Write-Host " ########>" -ForegroundColor Cyan
	Write-Host ""
		
	$List = @(.\Read-List.ps1 -Items $List -CSVItemField $CSVItemField )
				
				
	Write-Host ""		
	Write-Host "[Verify-List] Imported list of Names to be verified :" -ForegroundColor Cyan	-BackgroundColor	Blue
	Write-Host ""	
	
	Write-Host ($List |  Out-String) -foregroundcolor Cyan
	
		
	$ListLength = $List.Length
	
			$count = 0
			
			$List | %{
			
			$count++
			
			Write-Host ""
			Write-Host "List Item [$count/$ListLength] " -ForegroundColor  Cyan -BackgroundColor Blue
			Write-Host ""
			
			$Item = $_.ToString()
			$ItemLength = $Item.Length
			
			if ( $ItemLength -lt $MinItemLength  -and $CSVItemField -ne "*") 
			{  
					#Item from the list is shorter than the minimun chars length and will not be verified thus avoiding user or group object functions to perform  long searches
					
					Write-Host ""
					Write-Host "	" -NoNewline
					Write-Host "  WARNING: Item ' $_ ' is skipped because it is shorter than minimun length [$MinItemLength]  " -ForegroundColor  Blue	-BackgroundColor	Yellow
					Write-Host ""
										
					#continue 
			}
			else
			{
				if ($ItemType -eq "User")
				{
					#Write-Host "`$UserObject = .\UserADObject.ps1 $($_.ToString()) " -ForegroundColor Yellow
					
					<#
					if($brief)
					{
						$UserObject = .\UserObject.ps1 $($_.ToString())   -Details usrvb
						
						
					}
					else
					{
						$UserObject = .\UserObject.ps1 $($_.ToString())   
					}
					#>
					$UserObject = .\UserADObject.ps1 $($_.ToString())
					
					if ( $UserObject -eq "^") 
					{# User NOT found 
							$Date = get-date -format "dd_MMM_yyyy"

							$File = "List_Users_NOT_FOUND_$Date.txt"
							$FullPath = $SkipPath + $File
							$NotFound += $_
							Write-Host ""
							Write-Host "`$_ | Out-File  $FullPath -Append" -ForegroundColor Yellow
							Write-Host ""
											
							$_ | Out-File  $FullPath -Append
							
					}
					else
					{ $VerifiedObjects += $UserObject }
				
				}
				elseif ($ItemType -eq "Group")
				{
								
					Write-Host "`$GroupObject =  .\GroupObject.ps1  $($_.ToString()) -return" -ForegroundColor Yellow
				
					$GroupObject = .\GroupObject.ps1 $($_.ToString())
					
					if ( $GroupObject -eq "^") 
					{# User NOT found 
							$Date = get-date -format "dd_MMM_yyyy"

							$File = "List_Groups_NOT_FOUND_$Date.txt"
							$FullPath = $SkipPath + $File
							$NotFound += $_
							Write-Host ""
							Write-Host "`$_ | Out-File  $FullPath -Append" -ForegroundColor Yellow
							Write-Host ""
											
							$_ | Out-File  $FullPath -Append
							
					}
					else
					{ $VerifiedObjects += $GroupObject }
				
				}
				else
				{
					Write-Host ""
					Write-Host "	" -NoNewline
					Write-Host "  WARNING: Unknown Item Type ' $ItemType ' " -ForegroundColor  Blue	-BackgroundColor	Yellow			
				
				}
					
			}
		
		}


	Write-Host ""		
	Write-Host "[Verify-List] Verified list of $ItemType objects :>"  -ForegroundColor Cyan	-BackgroundColor	Blue
	Write-Host ""	
	
	$VerifiedObjects =  $VerifiedObjects | sort Name
	
	Write-Host ($VerifiedObjects | ft Name, PrimarySMTPAddress , guid -AutoSize | Out-String ) -foregroundcolor Cyan
	Write-Host ""
	Write-Host "  End of Script Verify-List.ps1" -ForegroundColor DarkBlue	-BackgroundColor Gray
	Write-Host ""
		
	
	return $VerifiedObjects
	
	
	
	
	
	
<#
.SYNOPSIS
   Read-List.ps1 reads a list of names from the input 
   
.DESCRIPTION
  Read-List.ps1  can read from 
  -	The  prompt (comma separated values ) .   
  - A temp file. This allows for the names to be copied in the file as a colomn
  - A CSV file 
  
  	Dependances : 
	
	None
	
		
.PARAMETER Items
    This is a list of items to read
	
   
.PARAMETER CSVItemField
   This is the name of the field in the CSV file from which values will be read.
   Default is '*". 
   If '*' will return all column names data
	
.EXAMPLE
    PS> .\Read-List.ps1 
	ENTER
	
	PS> # Press ENTER to import from a file (CSV or List.txt)  OR
	PS> # Type name(s) (separated by comma)  >>>:
	ENTER
	
	PS> # Enter the full path to the CSV file OR
    PS> # Type '^' to import from ' Temp\List.txt ' >>>:
	
.EXAMPLE
    PS> .\Read-List.ps1 -MinItemLength 10 -$CSVItemField 'UserName'
	Reads items 10 or more chars in length. Reads only from a field with name 'UserName' from a CSV file 
	
    <Description of example>
.NOTES
    Author: Filip Neshev; filipne@yahoo.com
    Date:  August 2016     
#>


param(
   	[parameter(Position=0,Mandatory=$false,ValueFromPipeline=$true,HelpMessage="Input List ")]
	$Items,
	[parameter(Position=1,Mandatory=$false,ValueFromPipeline=$false,HelpMessage="The name of the Item field to read from a CSV file ")]
	[string]$CSVItemField="*",
	[parameter(Position=2,Mandatory=$false,ValueFromPipeline=$true,HelpMessage="Do not clean temp file")]
	[switch]$no_temp_file_clean,
	[parameter(Position=3,Mandatory=$false,ValueFromPipeline=$true,HelpMessage="Do not order")]
	[switch]$keep_order
)
	
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
					Write-Host "Read-List.ps1	Created by Filip Neshev, August 2014 filipne@yahoo.com" -ForegroundColor Cyan  -backgroundcolor  DarkGray	
					#Write-Host ""
			}
			
			
			
			$dirname = "Temp"
			if( !(Test-Path $dirname ) ) { New-Item -Name $dirname -ItemType Directory | out-Null  }
			
			
			$TXTFile = "Temp\List.txt"
			
			if( !(Test-Path $TXTFile ) ) { New-Item -Name $TXTFile -ItemType File | out-Null }
			
			$AllItems =@()
			$AllItemsTrim  =@()
				
			$AllItems  | Write-Verbose
			
	
	}

	process 
	{
	
		Write-Verbose "Stuff in Process block to perform"
	
		function fRead-List ( $Items )
		{
			#$Items | Write-Verbose 
			
			#Write-Host "fRead-List " -ForegroundColor Magenta							
										
			if(!$Items )
			{			
				# First prompt
							
				Write-Host "`r`n Enter Item(s) (separated by comma)`r`n `"^`" to paste in '$TXTFile'   `r`n <ENTER> to import from CSV or TXT " -ForegroundColor Blue -BackgroundColor Gray 
				Write-Host "`r`n [Read-List] " -NoNewline -ForegroundColor Blue -BackgroundColor Gray 
				Write-Host " " -NoNewline
				$Items = read-host -prompt " "  
				
				$Items = $Items.Replace(";", ",")
				
				if ( $Items  -eq "^" )
				{
						# "^" to import from the Notapad file
				
						Write-Host ""
						Write-Host " Requested import  from $TXTFile "   -ForegroundColor Cyan	-BackgroundColor	Blue
						Write-Host ""
						
						notepad $TXTFile 
						
						$Confirm  = $true
					
						while ($Confirm )
						{ 
							
							Write-Host ""
							Write-Host  "CTRL + C to Cancel, to Import from $TXTFile press  ENTER " -ForegroundColor Blue -BackgroundColor Gray -NoNewline
							$Confirm  = Read-Host " "
																
							#$Confirm = read-host -prompt "|#| CTRL + C to Cancel, to import from $TXTFile press >> ENTER << -->> "
						}
						
						Write-Host ""
						Write-Host "Importing content from $TXTFile "   -ForegroundColor Cyan	-BackgroundColor	Blue
						Write-Host ""
						Write-Host " `$ItemsFromFile= Get-Content $TXTFile" -ForegroundColor Yellow
						Write-Host ""
													
						$ItemsFromFile = Get-Content $TXTFile 
						
						$ItemsFromFile | Write-Verbose
						
						$AllItems += $ItemsFromFile
						
						Clear-Content $TXTFile
				
				}
				elseif(!$Items )
				{ 
					# Second prompt for CSV or List.txt import
		
					Write-Host ""
					Write-Host " Enter the full path to the CSV file" -ForegroundColor Blue -BackgroundColor Gray -NoNewline
					Write-Host " " -NoNewline
					$Path = read-host -prompt " "
					
					if($Path)
					{
							# Import from  file
							
							if ( $Path -eq "^" )
							{
								# "^" to import from the Notapad file
						
								Write-Host ""
								Write-Host " Requested import  from $TXTFile "   -ForegroundColor Cyan	-BackgroundColor	Blue
								Write-Host ""
								
								notepad $TXTFile 
								
								$Confirm  = $true
							
								while ($Confirm )
								{ 
									
									Write-Host ""
									Write-Host  "CTRL + C to Cancel, to Import from $TXTFile press  ENTER " -ForegroundColor Blue -BackgroundColor Gray -NoNewline
									$Confirm  = Read-Host " "
																		
									#$Confirm = read-host -prompt "|#| CTRL + C to Cancel, to import from $TXTFile press >> ENTER << -->> "
								}
								
								Write-Host ""
								Write-Host "Importing content from $TXTFile "   -ForegroundColor Cyan	-BackgroundColor	Blue
								Write-Host ""
								Write-Host " `$ItemsFromFile= Get-Content $TXTFile" -ForegroundColor Yellow
								Write-Host ""
															
								$ItemsFromFile = Get-Content $TXTFile 
								
								$ItemsFromFile | Write-Verbose
								
								$AllItems += $ItemsFromFile
								
								Clear-Content $TXTFile
							
						}
							elseif( Test-Path $Path )
				        	{
									# Import from a CSV file
									
									$script:Path = $Path
									
									Write-Host "" 
									
									Write-Host "`$Content = Get-Content  $Path -ErrorAction SilentlyContinue " -ForegroundColor Yellow
						
									$Content = Get-Content  $Path -ErrorAction SilentlyContinue 
									 	
									Write-Host ""
									Write-Host "[$($Content.count)] Items found in  $Path "
									
																		
									if( $Content )
									{

										$ItemsFromFile = @()
										
										Write-Host ""
										Write-Host "Importing content from CSV file "   -ForegroundColor Cyan	-BackgroundColor	Blue -NoNewline
										Write-Host "  " -NoNewline
										Write-Host " $Path"   -ForegroundColor Cyan
										Write-Host ""
										Write-Host "`$ItemsFromCSVFile  =  Import-Csv  $Path " -ForegroundColor Yellow
																		 	 
										$ItemsFromCSVFile = Import-Csv  $Path 
										
										$GroupMemberscount = $ItemsFromCSVFile.count
										
										Write-Host "`$ItemsFromCSVFileProperties =  `$ItemsFromCSVFile| Get-Member -ErrorAction SilentlyContinue | ? { `$_.MemberType -match 'Property'} | select Name, MemberType" -ForegroundColor Yellow
										
										$ItemsFromCSVFileProperties =  $ItemsFromCSVFile | Get-Member -ErrorAction SilentlyContinue | ? { $_.MemberType -match "Property"} | select Name, MemberType
										
										Write-Host ""
										Write-Host ( $ItemsFromCSVFileProperties | ft -AutoSize  | Out-String ) -foregroundcolor Cyan
										Write-Host ""
										
										$totalCount = 0
										
										if ($CSVItemField  -eq "*" )
										{
										
										
										
										}
										
																		
										$ItemsFromCSVFile | 
										% {
								
												#$Name = $($ItemsFromCSVFileProperties[1].Name)

												$totalCount++
																	
												write-host "$([string]::Format( "`r {0:d3}{1:d3}{2:s50}" ,  "[$totalCount/$GroupMemberscount] " ," $($ItemsFromCSVFileProperties[0].Name): $($_.$($ItemsFromCSVFileProperties[0].Name)) ; $($ItemsFromCSVFileProperties[1].Name): $($_.$($ItemsFromCSVFileProperties[1].Name)) ; " , "$($ItemsFromCSVFileProperties[2].Name): $($_.$($ItemsFromCSVFileProperties[2].Name))					" )  ) "   -ForegroundColor Cyan	-BackgroundColor Blue -nonewline
												
												if ($CSVItemField  -eq "*" )
												{
													$ItemsFromFile += $_
												}
												else
												{ 	$ItemsFromFile +=  $_.$CSVItemField  }
															
										
											}
									
										$script:from_csv = $true
									
										$AllItems += $ItemsFromFile
									
									}
									else
									{
											Write-Host ""
											Write-Host "  WARNING : No content read ! Exiting ..." -NoNewline -ForegroundColor	Red	-BackgroundColor	yellow
											Write-Host "  :  "  -ForegroundColor	Red	-BackgroundColor	yellow
											Write-Host ""
									}
								
						} #if($Path)
				        else
						{
							Write-Host ""
							Write-Host "	" -NoNewline
							Write-Host "WARNING: Not a  valid path! Exiting.. " -ForegroundColor  Blue	-BackgroundColor	Yellow
			       		 }
		
					}
					else
					{
						# No Path provided 
						
						Write-Host ""
						Write-Host "	" -NoNewline
						Write-Host "  ATTENTION : No Parth provided " -ForegroundColor  Blue	-BackgroundColor	Yellow
											
					
					}
				
				} # elseif(!$Items )
				else 
				{ 	
						$ItemsFromList = $Items.Split(",") 	
						
						#Write-Verbose $ItemsFromList
						
						$AllItems += $ItemsFromList
					
				}
			
				
			}
			else
			{ 	
			
				if( $Items.gettype().Name -like "String")
				{
					$AllItems = $Items.Split(",")  	
							
				}
				else
				{
					$AllItems = $Items
				}
				
			}
		
		#$AllItems  | Write-Verbose
		
		return $AllItems
					
		}# End function 
				
		# Add items from each function call from the pipe 
		
		$AllItems += fRead-List  $Items 
		
	}#process
	
	end {
	
			$AllItems  | %{ 	
			
				if ($CSVItemField  -eq "*" -and  $($script:Path))
				{
					$AllItemsTrim += $_
					
					#Write-Host "352" -ForegroundColor Magenta
				}
				elseif ($_)
				{

					#$AllItemsTrim += ($_.ToString()).Trim() 
					
					#Write-Host "' $_ '" -ForegroundColor Red
					
					$TrimedItemStringArray  = @()
					
					# from string with spaces to array 
					$ItemStringArray = ($_.ToString()).Split("")
					
					#Write-Host ( $ItemStringArray  | fl  | Out-String) -foregroundcolor Cyan
					
						
						
						$ItemStringArray | %{
							
							if ( $_ )
							{
								
								#Write-Host "' $_ '" -ForegroundColor Red
								# trimming each arary item and populating a new array
								
								$trimed =  $_.Trim(" ")
								$trimed = $trimed.Trim(",")
								
								$TrimedItemStringArray += $trimed 
							}
						}
						
					
					#$TrimedItemStringArray | Write-Host -ForegroundColor Green
					
					
					# from array of trimmed items back to string of trimmed names
					$TrimedItemStringArray= $TrimedItemStringArray -join ' '
					
					
					#$TrimedItemStringArray | Write-Host -ForegroundColor Red
									
					if ( $TrimedItemStringArray )
					{
						$AllItemsTrim += $TrimedItemStringArray
					}
					
					#Write-Host "  $($AllItemsTrim.count) " -ForegroundColor Red
					
					#>
					
					#$AllItemsTrim += ($_.ToString()).Trim() 
					
					
				}
			
			}
		
		
			if(!$keep_order)
			{
				$AllItemsTrim  = @( $AllItemsTrim  | sort )
			}
			
			$Length = $AllItemsTrim.Length
			
			if ( $script:from_csv )
			{
			
				Write-Host ""		
				Write-Host ""	
				Write-Host "(Read-List) [ $Length ] Items from field '$CSVItemField' from file ' $($script:Path) '  Read " -ForegroundColor Cyan	-BackgroundColor	Blue
				Write-Host ""
			
			}
			else
			{
				Write-Host ""		
				Write-Host "(Read-List) [ $Length ] Items Read " -ForegroundColor Cyan	-BackgroundColor	Blue
				Write-Host ""	
			}					
					
			Write-Host ($AllItemsTrim | ft | Out-String) -foregroundcolor Cyan		
			
			if ( $no_temp_file_clean )
			{
				$AllItemsTrim | Set-Content $TXTFile -Encoding Unicode 
			}
			
			Write-Host "return" -ForegroundColor Yellow
			
			return $AllItemsTrim
		
	
	}

	
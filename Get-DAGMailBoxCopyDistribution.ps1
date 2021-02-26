	<#
.SYNOPSIS

    Use the Get-DAGMailBoxCopyDistribution cmdlet to view if DAG mailbox copies are mounted on the Preference One server. 
	In addition, this cmdlet displays Content Index Status,  Copy Queue Length and  Replication Queue Length
     
.DESCRIPTION

	Get-DAGMailBoxCopyDistribution.ps1 works either in 
	
	- Database Display mode: Shows the following information for DAG Databases in a table view :
		
		AP -  Activation Preference number
		MtdOn1 - Mounted on - shows [OK] if on server preference 1 or [X] if not, 
		for the server row on which the database is mounted on  
		ContIndSt - Content Index Status
		CopyQL -Copy Queue Length
		RplyQL - Replication Queue Length
		Name - in the format DATABASE\SERVER
		Status - Database copy Status
		ErrorMessage
		
	OR
	
	- Server Display mode: Shows all Active Database copies on specified servers. 
	
.PARAMETER ServerName 
    
	If provided, DBName is ignored and the script switches to Server Display Mode.
	In this mode 
	
	Provide a Server Name Search String : Server_Name
	
	Actual wildcard  search:
	*Server_Name* - returns all servers whose names include Server_Name at any position (start, end,  in between )
	
	Default: All Mailbox Servers
	
	
.PARAMETER DBName

	Evaluated ONLY If ServerName parameter is not provided (i.e. ServerName is Empty ) !!
	Script runs in Database Display Mode. 
	
	Provide a Database Name Search String : Database_Name
	
	Actual wildcard  search: 
	*Database_Name* - returns all databases whose names include Database_Name at any position (start, end,  in between ).
	
	Default: All DAG Databases	
	
.PARAMETER	ReturnObject   
	
	If provided will make the script to return a result object that can be saved in a variable.
	If not provided, the script will only show result on screen
	
	
.EXAMPLE
    C:\PS> .\Get-DAGMailBoxCopyDistribution.ps1
    
	Searches for all DAG Databases , on all servers, in Database Display Mode
	
.EXAMPLE
    C:\PS> $ResultObject = .\Get-DAGMailBoxCopyDistribution.ps1 -ReturnObject
	
	Searches for all DAG Databases, on all servers, in Database Display Mode and returns a result object
	in variable $ResultObject 
	
.EXAMPLE
    C:\PS> .\Get-DAGMailBoxCopyDistribution.ps1 -DBName	Database_Name
	
	Searches for all DAG Databases which names are like Database_Name, on all servers, in Database Display Mode
	
.EXAMPLE
    C:\PS> .\Get-DAGMailBoxCopyDistribution.ps1 -ServerName Server_Name
	
	Searches for all DAG Databases , on servers which names are like Server_Name , in Server Display Mode	
	
	
	
	
.NOTES
    Author: Filip Neshev
    Date:   October, 2013    
#>

param(
    [parameter(Position=0,Mandatory=$false,ValueFromPipeline=$true,HelpMessage="Database Name ")][string]$DBName,
    [parameter(Position=1,Mandatory=$false,ValueFromPipeline=$true,HelpMessage="Server Name ")][string]$ServerName,
	[parameter(Position=2,Mandatory=$false,ValueFromPipeline=$true,HelpMessage="DAG Name ")][string]$DAGName="DAG",
	[parameter(Position=3,Mandatory=$false,ValueFromPipeline=$true,HelpMessage="ReturnObject")][switch]$ReturnObject
	
	
	)

Write-Host ""
Write-Host ""
Write-Host " Get-DAGMailBoxCopyDistribution.ps1	 Created by Filip Neshev, 2013	filipne@yahoo.com" -ForegroundColor Cyan  -backgroundcolor  DarkGray	
Write-Host ""

	
if( $ServerName )
{
		Write-Host ""
		Write-Host "<<< Server Display Mode. (Database search string ' $DBName ' is ignored) >>>" -ForegroundColor White -BackgroundColor Blue
		Write-Host ""
		Write-Host "For how to start in Database Mode and other info type: 
help  Get-DAGMailBoxCopyDistribution.ps1  " -Foregroundcolor Cyan -BackgroundColor BLUE  
		Write-Host ""
}
else
{
			Write-Host ""
			Write-Host "<<< Database Display Mode >>> " -ForegroundColor White -BackgroundColor Blue
			Write-Host ""
			
			Write-Host "For how to start in Server Mode and other info type: 
	help  Get-DAGMailBoxCopyDistribution.ps1  " -Foregroundcolor Cyan -BackgroundColor BLUE  
			Write-Host ""
}


$DAGNames = @(Get-DatabaseAvailabilityGroup)

Write-Host ""

Write-Host "Detected DAG(s): " -BackgroundColor DarkBlue -ForegroundColor Cyan

Write-Host "" 

Write-Host ($DAGNames | ft Name, Servers  -AutoSize | Out-String) -foregroundcolor Cyan



#$DAGNames | Write-Host -BackgroundColor DarkBlue -ForegroundColor Cyan

Write-Host "" 
Write-Host " `$DAGNames = @(Get-DatabaseAvailabilityGroup $DAGName -ErrorAction SilentlyContinue) " -ForegroundColor Yellow
Write-Host ""

$DAGNames = @(Get-DatabaseAvailabilityGroup $DAGName -ErrorAction SilentlyContinue)

Write-Host ""

while ($($DAGNames.Length) -ne 1 )
{
	$DAGName = read-host -prompt "# Enter a valid  single DAG Name  >>>"
	$DAGNames = @(Get-DatabaseAvailabilityGroup $DAGName -ErrorAction SilentlyContinue )
}



$DAGNames = [string]$DAGNames 


$MBServers = Get-MailboxServer *$ServerName* -ErrorAction SilentlyContinue  | ?{$_.DatabaseAvailabilityGroup -eq $DAGNames}

while ( !$MBServers)
{
	Write-Host ""
	Write-warning "Server Name *$ServerName* Not found !"
	Write-Host ""
	
	$ServerName = read-host -prompt "--- Type a valid Server Name or CTRL + C to Cancel; ->> "
	
	$MBServers = @(Get-MailboxServer *$ServerName* -ErrorAction SilentlyContinue | ?{$_.DatabaseAvailabilityGroup -eq $DAGNames })
}

Write-Host " "

Write-Host "Mailbox Role Servers members of $DAGNames  >>" -ForegroundColor Cyan	-BackgroundColor	Blue

Write-Host "
`$MBServers = Get-MailboxServer *$ServerName* | ?{`$_.DatabaseAvailabilityGroup -eq $DAGNames }"  -ForegroundColor Yellow
 
Write-Host ($MBServers | ft -AutoSize | Out-String) -foregroundcolor Cyan

#$MBServers | out-host

if( $ServerName )
{
		Write-Host ""
		Write-Host "Active Database Copies on  $MBServers (search string  '*$ServerName*') " -ForegroundColor Cyan	-BackgroundColor	Blue
		Write-Host ""
	
		$MBServers | % {
	
			Write-Host "Server $_ "-BackgroundColor DarkBlue -ForegroundColor Cyan
			Write-Host ""
			Write-Host "`$MailboxDatabaseCopyStatus = @(Get-MailboxDatabaseCopyStatus -Server ' $_ '  | ?{`$_.activecopy -eq `$true} )" -ForegroundColor Yellow
			Write-Host ""
			$MailboxDatabaseCopyStatus = @(Get-MailboxDatabaseCopyStatus -Server $_  | ?{$_.activecopy -eq $true} )
					
			$MailboxDatabaseCopyStatusLength = $MailboxDatabaseCopyStatus.Length 
					
			#if ( !$MailboxDatabaseCopyStatusLength ) { $MailboxDatabaseCopyStatusLength = 0	}
					
			Write-Host " [$MailboxDatabaseCopyStatusLength] Active Copies found " -ForegroundColor Cyan	-BackgroundColor	Blue
			Write-Host ""		
						
			$MailboxDatabaseCopyStatus = $MailboxDatabaseCopyStatus | sort ContentIndexState , Name 
			
			Write-Host ($MailboxDatabaseCopyStatus | sort ContentIndexState , Name | ft ActiveCopy,  Name, CopyQueueLength, ReplayQueueLength, ContentIndexState , ContentIndexErrorMessage -AutoSize -Wrap | Out-String) -foregroundcolor Cyan


		} # $MBServers | %
		
} # if( $ServerName )
else
{
			Write-Host ""
			Write-Host "Database Name Search for ' *$DBName* ' "  -ForegroundColor Cyan	-BackgroundColor	Blue

			Write-Host ""
			
			Write-Host "`$MailboxDatabases = Get-MailboxDatabase  -IncludePreExchange2013 -Identity *$DBName* -ErrorAction SilentlyContinue  | ?{ `$_.MasterServerOrAvailabilityGroup -eq $DAGNames }  | sort Name" -ForegroundColor Yellow
			
			Write-Host ""
			
			$MailboxDatabases = Get-MailboxDatabase  -IncludePreExchange2013 -Identity *$DBName* -ErrorAction SilentlyContinue  | ?{ $_.MasterServerOrAvailabilityGroup -eq $DAGNames }  | sort Name
			
			if ( $MailboxDatabases ) 
			{
				$MailboxDatabasesLength = $MailboxDatabases.Length
				if (!$MailboxDatabasesLength) { $MailboxDatabasesLength = "1" }
			}
			else
			{
				Write-Host "`$MailboxDatabases = Get-MailboxDatabase  -Identity *$DBName* -ErrorAction SilentlyContinue  | ?{ `$_.MasterServerOrAvailabilityGroup -eq $DAGNames }  | sort Name" -ForegroundColor Yellow
				$MailboxDatabases = Get-MailboxDatabase -Identity  *$DBName* -ErrorAction SilentlyContinue  | ?{ $_.MasterServerOrAvailabilityGroup -eq $DAGNames }  | sort Name
				
				
				$MailboxDatabasesLength = $MailboxDatabases.Length
				if (!$MailboxDatabasesLength) { $MailboxDatabasesLength = "1" }
			}
			
			#"MailboxDatabases $MailboxDatabases "

			$dbnumber = 0

			while ( !$MailboxDatabases )
			{
				$DBName = read-host -prompt "--- Type a valid Database Name or CTRL + C to Cancel; For all Databases press >> ENTER << --  >>> "
				Write-Host " `$MailboxDatabases = Get-MailboxDatabase $DBName -ErrorAction SilentlyContinue | sort Name	" -ForegroundColor Yellow
				$MailboxDatabases = Get-MailboxDatabase $DBName -ErrorAction SilentlyContinue | sort Name	
			}

			Write-Host  "[ $($MailboxDatabasesLength) ] Databases like ' $DBName ' found on $DAGNames >>" -ForegroundColor Cyan	-BackgroundColor	Blue

			$MailboxDatabaseCopyStatusExtendedArrayAll = @()
			$MailboxDatabaseCopyNotOnPreferenceOneArrayAll = @()
			$detected = 0
			$numberOfNotOnPr1 = 0

			$MailboxDatabases | %{	# First Loop Databases

			$DatabaseName = $_.Name
			$dbnumber ++
			$notOnPref1 = $false
			
			Write-Verbose  " DatabaseName $DatabaseName "
			
			#-Status "Scanning Mailbox Databases that are similar to ' $DBName* ' "
			
			Write-Progress  -status " " `
			-Activity    "Analyzing Database ` [$dbnumber/$MailboxDatabasesLength] : $DatabaseName ; $($_.ActivationPreference)   "`
			-CurrentOperation "[$detected] Detected Active Database copies not on DAG ' $DAGNames '  Preference ONE Server " -PercentComplete ( ($dbnumber/$MailboxDatabasesLength) * 100 )
			     
			$MailboxDatabaseCopyStatuses = Get-MailboxDatabaseCopyStatus $_.Name
				
			$ActivationPreference =  $_.ActivationPreference
			
			Write-Verbose "ActivationPreference $ActivationPreference"
			
			#$ActivationPreference
			
			# ActivationPreferenceOne style 2010
			$ActivationPreferenceOne = ($ActivationPreference | ?{ $_.Value -eq "1"  } ).Key.Name
			
			if( !$ActivationPreferenceOne ) 
			{# ActivationPreferenceOne style 2013
						
				foreach ($Pref in $ActivationPreference ) 
				{
					$Pref = $Pref.Replace("[", "")	
					$Pref = $Pref.Replace("]", "")
					$Pref = $Pref -split (",")
					$Pref[1] = $Pref[1].Trim()
					
					if( $Pref[1] -eq "1")
					{
						$ActivationPreferenceOne = $Pref[0].Trim()
				
					}
				}
				
			}
			
			
			#
			#$ActivationPreferenceOne = ($ActivationPreference | ?{ $_ -eq "1"  } ).Key.Name
			#$ActivationPreferenceOne = ($ActivationPreference | ?{ $_.Vallue -eq "1"  } )
			
			
			Write-Verbose " ActivationPreferenceOne  $ActivationPreferenceOne  "
			
			$MailboxDatabaseCopyStatusExtendedArray = @()
			$MailboxDatabaseCopyNotOnPreferenceOneArray = @()

			foreach ( $MBDBCS in $MailboxDatabaseCopyStatuses)
			{ # Second Loop true Database Copy Statuses for a given Database ( $_ )

				$MailboxServer = $MBDBCS.MailboxServer
					
				#$MBDBCSSelected  = $MBDBCS | select Name, Status, Errormessage 
				$MBDBCSSelect  = $MBDBCS | select Name, DatabaseName,  Status
				#, DatabaseName
				
				$MailboxDatabaseCopyStatusExtended = New-Object $MBDBCSSelect –TypeName PSObject 
						
				#$Errormessage = $MBDBCS.Errormessage
				
				$Status = $MBDBCS.Status
				
				$DatabaseName = $MBDBCS.DatabaseName
				#$DatabaseSize = ((Get-MailboxDatabase $DatabaseName -Status).DatabaseSize).ToMB()
				$DatabaseSize = [Microsoft.Exchange.Data.ByteQuantifiedSize]::Parse((Get-MailboxDatabase $DatabaseName -Status).DatabaseSize).ToMB()
				$DatabaseSize = $DatabaseSize / 1024 
				$DatabaseSize = "{0:N0}" -f $DatabaseSize
		
				$ContentIndexState = $MBDBCS.ContentIndexState
				$CopyQueueLength = $MBDBCS.CopyQueueLength
				$ReplayQueueLength = $MBDBCS.ReplayQueueLength
				$Errormessage = $MBDBCS.Errormessage
					
				
				foreach ($AP in  $ActivationPreference ) 
				{  # Third Loop true Activation Preferences for a given Database Status ( $MBDBCS )
				
					$ActivationPreferenceKeyName = $AP.Key.Name
					$ActivationPreferenceValue = $AP.Value
					
					if(!$ActivationPreferenceKeyName)
					{
						$AP= $AP.Replace("[", "")	
						$AP= $AP.Replace("]", "")
						$AP= $AP-split (",")
						$AP[1] = $AP[1].Trim()
	
						$ActivationPreferenceKeyName = $AP[0].Trim()
						$ActivationPreferenceValue= $AP[1].Trim()
										
					}
	
					Write-Verbose "ActivationPreferenceKeyName $ActivationPreferenceKeyName  "
					#Write-Host "$ActivationPreferenceValue $ActivationPreferenceValue  "
					
							
					if ( $ActivationPreferenceKeyName -eq $MailboxServer )
					{
						#$ActivationPreferenceValue = $AP.Value
						$MailboxDatabaseCopyStatusExtended | Add-Member –MemberType NoteProperty –Name AP –Value $ActivationPreferenceValue
						
						if ( $Status -eq 'Mounted' -and $ActivationPreferenceValue -eq 1)
						{
							$MailboxDatabaseCopyStatusExtended | Add-Member  –MemberType NoteProperty –Name MtdOn1 –Value "[OK]"
						}
						elseif  ( $Status -eq 'Mounted' -and $ActivationPreferenceValue -ne 1)
						{
							$notOnPref1 = $true
							$MailboxDatabaseCopyStatusExtended | Add-Member –MemberType NoteProperty –Name MtdOn1 –Value "[X]"
							#$MailboxDatabaseCopyNotOnPreferenceOneArray += $MailboxDatabaseCopyStatusExtended
							$detected++
						}
						else
						{
							$MailboxDatabaseCopyStatusExtended | Add-Member –MemberType NoteProperty –Name MtdOn1 –Value " "
						}
						
	
						
						$MailboxDatabaseCopyStatusExtendedArray += $MailboxDatabaseCopyStatusExtended
												
						#Write-Host " MailboxDatabaseCopyStatusExtended" 
						#$MailboxDatabaseCopyStatusExtended | Out-Host
						#$MailboxDatabaseCopyStatusExtended.Gettype()
						
						Write-Verbose " MailboxDatabaseCopyStatusExtendedArray " 
						#$MailboxDatabaseCopyStatusExtendedArray | Out-Host

						
					}

				 } # Third Loop Tru Activation Preferences for a given Database Status ( $MBDBCS )
		
		
				$MailboxDatabaseCopyStatusExtended | Add-Member –MemberType NoteProperty –Name "ContIndSt"  –Value  $ContentIndexState 
				$MailboxDatabaseCopyStatusExtended | Add-Member –MemberType NoteProperty –Name "CopyQL"   –Value  $CopyQueueLength
				$MailboxDatabaseCopyStatusExtended | Add-Member –MemberType NoteProperty –Name "RplyQL"   –Value  $ReplayQueueLength
				$MailboxDatabaseCopyStatusExtended | Add-Member –MemberType NoteProperty –Name "Size GB"   –Value  $DatabaseSize
				#$Errormessage = $MBDBCS.Errormessage
				$MailboxDatabaseCopyStatusExtended | Add-Member –MemberType NoteProperty –Name Errormessage –Value $Errormessage
				
	
			} # Second Loop true Database Copy Statuses for a given Database ( $_ )
	
		
		
		$MailboxDatabaseCopyStatusExtendedArray = $MailboxDatabaseCopyStatusExtendedArray | sort AP
		
		#Write-Host " MailboxDatabaseCopyStatusExtendedArray"
		#$MailboxDatabaseCopyStatusExtendedArray | ft -AutoSize -Wrap
		
		$MailboxDatabaseCopyStatusExtendedArrayAll = $MailboxDatabaseCopyStatusExtendedArrayAll + $MailboxDatabaseCopyStatusExtendedArray  + ""
		
		#$MailboxDatabaseCopyStatusExtendedArrayAll | ft -AutoSize -Wrap
	
		if ($notOnPref1)
		{# Add the whole status extended object
			$MailboxDatabaseCopyNotOnPreferenceOneArrayAll = $MailboxDatabaseCopyNotOnPreferenceOneArrayAll + $MailboxDatabaseCopyStatusExtendedArray + ""
			$numberOfNotOnPr1 ++
		}
	
	}#$MailboxDatabases | %

}

	if ( !$ReturnObject  )
	{
		# Show them all  on the screen 
		
		#Write-Host " CP 1"
		#$MailboxDatabaseCopyStatusExtendedArrayAll  | select  Name, Status,  AP,  MtdOn1, ContIndSt, CopyQL, RplyQL, "Size GB", Errormessage
		
		Write-Host ( $MailboxDatabaseCopyStatusExtendedArrayAll  | ft  Name, Status,  AP,  MtdOn1, ContIndSt, CopyQL, RplyQL, "Size GB", Errormessage  -AutoSize -Wrap | Out-String) -foregroundcolor Cyan

		#Write-Host ($Obj | Format-List | Out-String) -foregroundcolor Cyan

		
		#Write-Host " CP 2"
		Write-Host ""
		#Write-Host "  detected :   $detected "
		
		if ( !$ServerName )
		{
		
		
		if ( $($MailboxDatabaseCopyNotOnPreferenceOneArrayAll.Length) )
		{
			#write-host " [ $($MailboxDatabaseCopyNotOnPreferenceOneArrayAll.Length) ] " -NoNewline -BackgroundColor Red -ForegroundColor White
			write-host " [ $detected ] " -NoNewline -BackgroundColor Red -ForegroundColor White
		}
		else
		{		
			#write-host " [ $detected  ] " -NoNewline -BackgroundColor DarkGreen -ForegroundColor White
			
			write-host " [ $($MailboxDatabaseCopyNotOnPreferenceOneArrayAll.Length) ] " -NoNewline -BackgroundColor DarkGreen -ForegroundColor White
		}
		
		

		
		write-host	" Databases NOT mounted on DAG ' $DAGNames ' Preference ONE " -ForegroundColor Cyan -BackgroundColor Blue
		
		$MailboxDatabaseCopyNotOnPreferenceOneArrayAll = $MailboxDatabaseCopyNotOnPreferenceOneArrayAll | select AP, MtdOn1, Name, ContIndSt, CopyQL, RplyQL, "Size GB",     Status, Errormessage  

		Write-Host ( $MailboxDatabaseCopyNotOnPreferenceOneArrayAll | ft -AutoSize -Wrap | Out-String ) -ForegroundColor Cyan
		
		Write-Host ""
		
		}
		
	}
	else
	{
		
		#return $MailboxDatabaseCopyStatusExtendedArrayAll
		return $MailboxDatabaseCopyNotOnPreferenceOneArrayAll
	}



<#
			if ( !$ReturnObject  )
			{
				$MailboxDatabaseCopyStatus | sort Name | ft ActiveCopy,  Name, CopyQueueLength, ReplayQueueLength, ContentIndex, ErrorMessage -AutoSize -Wrap
				Write-Host ""
			}
			else
			{ 
				return $MailboxDatabaseCopyStatus 
			}
			
#>
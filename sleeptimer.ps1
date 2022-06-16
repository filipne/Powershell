param(
    [parameter(Position=0,Mandatory=$false,ValueFromPipeline=$false,HelpMessage="Sleep Time in seconds ")]
	[int]$requestedtime=5
           
	)
	
	#$timecounter = $requestedtime

#"requestedtime $requestedtime "


Write-Host ""
Write-Host "sleeptimer.ps1;	Created by Filip Neshev filipne@yahoo.com " -ForegroundColor Cyan  -backgroundcolor  DarkGray	
#Write-Host ""	
	
	
	if ( $requestedtime -lt 0 )
	{
		
		$requestedtime = 5
	
	}	
	
function  Timecounter ($timecounter )
{
	#Write-Host "$($ts.hours)"
	
	Write-Host ""
	
	$requestedtimedisplay = ""
	
	
	$Time = [System.Diagnostics.Stopwatch]::StartNew()
		

	
	
	$ts = new-timespan -Seconds $requestedtime 
	
		
	$requestedtimedisplay = "$($ts.hours):$($ts.minutes):$($ts.Seconds)"	
	
	while ($timecounter) 
	{	
	
		$CurrentTime = $Time.Elapsed

		#$ts = new-timespan -Seconds $requestedtime 
			
		#$requestedtimedisplay = "$($ts.hours) hour : $($ts.minutes) min : $($ts.Seconds) sec"
		#write-host "$([string]::Format("`rTime: {0:d2}:{1:d2}:{2:d2} / $requestedtime [sec] / $([math]::Round( $($requestedtime / 60)  , 2 )) [min]", $CurrentTime.hours, $CurrentTime.minutes,  $CurrentTime.seconds)) "  -nonewline -ForegroundColor Cyan	-BackgroundColor	Blue


		write-host "$([string]::Format("`r {0:d2}:{1:d2}:{2:d2} / $requestedtimedisplay [hr:min:sec] ",$CurrentTime.hours, $CurrentTime.minutes,  $CurrentTime.seconds ) ) "  -nonewline -ForegroundColor Cyan	-BackgroundColor	Blue


		$timecounter -= 1
		
		sleep 1
		
	}
	
	

	#Write-Host ""
	Write-Host "  Time is UP ! " -ForegroundColor Cyan	-BackgroundColor	Blue
	Write-Host ""
}

Timecounter $requestedtime
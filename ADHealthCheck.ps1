$VerbosePreference = 'Continue' # Show all verbose messages
#  Active Directory Health Report
# 
#  Author: Graeme Evans
#  Original Author: Vikas Sukhija 
#  https://gallery.technet.microsoft.com/scriptcenter/Active-Directory-Health-709336cd
# 
#  Original Date: 12/25/2014
#  Original Status: Ping,Netlogon,NTDS,DNS,DCdiag Test(Replication,sysvol,Services)
#  Original Update: Added Advertising
#
# Tests: Ping, Netlogon, NTDS, DNS, DCdiag Test, Replication, sysvol, Services, Advertising, Mcafee DAT, 
#
# Version  Date        Notes
# 1.0      10/02/2017  Original Script
# 1.1      10/02/2017  Merged Uptime, C Drive, FRSEvent, KCCCheck from other scripts
# 1.2      10/02/2017  Added McAfee Dat Date
# 1.3      10/02/2017  Replaced .Net Mail Method, Added method for sending errors with email
# 1.4      10/02/2017  Added GUID to capture session data (errors)
# 1.5      10/02/2017  Added measure-command, replaced names with ticks
# 1.6      13/02/2017  Fully indented, removed measure command
# 1.7      13/02/2017  Replaced start HTML with here-string
# 1.8      13/02/2017  AV Check tidied code
# 1.9      13/02/2017  Updated all tests to create objects
# 2.0      14/02/2017  Convert Objects to HTML Dynamically
# 2.1      17/02/2017  After stropping about set-cellcolor I used a switch{} - then stripped the manual HTML
# 2.2      17/02/2017  Added out-null to each scriptblock to tidy the live screen data
# 2.3      18/02/2017  Removed the manual tests, converted Services and DCDIAG to loops
# 2.4      20/02/2017  Added WSUS data
# 2.5      22/02/2017  Get all services in one go
# 2.6      02/03/2017  Added Comparison between A record for domain, and the list of DCs
# 2.7.3    04/06/2017  Add DC DataTable - removed DNS one
# 2.7.4    15/06/2017  Not sure why im using point releases, but this updates the DC Data table to be conditional of the if statemnet (forest, domain, list)
# 2.7.5    06/12/2017  Updated the colour coding and error handling, drives and timeouts for WSUS to timeout - not just error

#region Define Variables
    #region Script Block Timeout
    	$timeout = "60" # timeout for each job in seconds - below 60 can cause false timeouts - the WSUS one was recorded at 27.22 seconds.
    #endregion Script Block Timeout	

	#region SMTP Settings
		$smtphost = 'x' #''
		$from = "AD Health Check <NoReply@??????.co.uk>"
		$recipients = "graeme.evans@??????.com"
	#endregion SMTP Settings
	
	#region Domain Controller Scope
    # This allows control of the scope. 
    ## Single (List of) DC
    ## Single Domain
    ## Entire Forest
    
		# If set, will only scan this DC. Set to $null (or #) to follow logic
		$DCServers = $null
		
    # Restrict to DCs in this domain, if null - scope is all DCs in >forest<
		$singleDomain = 'uk.megacorp.local' 
	#endregion Domain Controller Scope
	
	#region HTML Colours
    # These needs to be HTML safe pastel colours. Red is not pastel, as this is a Critical.
	#	$testPassedColour = "PaleGreen"
	#	$testFailedColour = "LightCoral" # so this is like 3 gb free space
	#	$testWarningColour = "LightSteelBlue"
	#	$testCriticalColour = "Red" # and this one is like must action NOW
	#	$testInformationColour = "Lavender"
	#endregion HTML Colours

    #region Tests
    # MUST be the "Name" value NOT DisplayName
    $testServices = "RSCDsvc" #, "LanmanServer", "Netlogon", "NTDS", "DNS" These are tested by DCDIAG
    # Case SenSiTiVe!
    $testDCDIAG = "NetLogons","Replications","Services","Advertising","FsmoCheck","KccEvent","FrsEvent","SysVolCheck","RidManager","Topology","VerifyReferences","VerifyReplicas"
    #endregion Tests
#endregion Define Variables

#region Check settings are OK
    if ($DCServers -or $singleDomain)
    {Write-Warning "Override for Domain Controllers or Forest is set to $dcservers $singledomain"
    Read-Host "Press any key to continue, or Control-C to exit"}
#endregion 


#region Create GUID, Folder and Report file
    $ScriptPathComplete = Split-Path $MyInvocation.MyCommand.Path -Parent
    Set-Location $ScriptPathComplete
    $ScriptPathComplete = $ScriptPathComplete.TrimEnd('\')
    
    if(!$(Test-Path $ScriptPathComplete))
        {
            Write-Warning "You must specify a path to the settings for the script to run!"
            Return
        }
                
	$date = Get-Date -Format 'yyyy-MM-dd'
	
	$guid = [guid]::NewGuid().guid
	Write-Verbose "$(Get-Date): GUID Created for this session: $guid"
	
	$reportpath = "$ScriptPathComplete\$guid\" 
    $advreportname = $reportpath + 'advADReport.htm'
    $objExportname = $reportpath + 'report.xml'
    $objDCDataName = $reportpath + 'dcdata.xml'
    $objErrorsName = $reportpath + 'errors.log'
	
	if(!(test-path $reportpath)) # test for GUID folder
	{
		new-item $reportpath -type directory | Out-Null
		Write-Verbose "$(Get-Date): Folder path created $reportpath"
	}
	
	if(!(test-path $advreportname)) # test Report file exists
	{
		new-item $advreportname -type file | Out-Null
		Write-Verbose "$(Get-Date): Report HTML File Created $advreportname"
	}

#endregion Create GUID, Folder and Report file

#region Get ALL DC Servers
	if(!($DCServers)) # tests if you have an override
	{
		if($singleDomain)
		{
			# Get all DCs in the singleDomain
			Write-Warning "Scope targetted to $singleDomain only"
			$getForest = [system.directoryservices.activedirectory.Forest]::GetCurrentForest()
			$getDCData = $getForest.domains | Where-Object {$_.Name -eq $singleDomain} | ForEach-Object {$_.DomainControllers} 
            $DCServers = $getDCData | ForEach-Object {$_.Name}
			$subject = "Active Directory Health Monitor - Domain: $singleDomain"
		}
		else
		{
			# Get all DCs in the forest
			Write-Warning "Scope targetted to all DCs in the forest"
            $getForest = [system.directoryservices.activedirectory.Forest]::GetCurrentForest()
			$getDCData = $getForest.domains | ForEach-Object {$_.DomainControllers} 
            $DCServers = $getDCData | ForEach-Object {$_.Name}
			$subject = "Active Directory Health Monitor - Forest: $getForest.name" 
		}
	}
	else
	{
		Write-Warning "$(Get-Date): Scope overrided to only $DCServers"
		$subject = "Active Directory Health Monitor - Server: $DCServers"
        $getDCData = $null
	}
#endregion Get ALL DC Servers

#region create empty array
    # Null an array
    $HealthReport = $null
    # Generate an Empty Array
    $HealthReport = New-Object System.Collections.ArrayList
#endregion create empty array


# get a set of DateTime values from the pipeline
# filter out $nulls and produce the latest of them
# (c) Dmitry Sotnikov
Write-Verbose "$(Get-Date): Loading Function Measure-Latest"
function Measure-Latest {
    BEGIN { $latest = $null }
    PROCESS {
            if (($_ -ne $null) -and (($latest -eq $null) -or ($_ -gt $latest))) {
                $latest = $_ 
            }
    }
    END { $latest }
}




#region Set-CellColor HTML Helper

Function Set-CellColor
{   <#
    .SYNOPSIS
        Function that allows you to set individual cell colors in an HTML table
    .DESCRIPTION
        To be used inconjunction with ConvertTo-HTML this simple function allows you
        to set particular colors for cells in an HTML table.  You provide the criteria
        the script uses to make the determination if a cell should be a particular 
        color (property -gt 5, property -like "*Apple*", etc).
        
        You can add the function to your scripts, dot source it to load into your current
        PowerShell session or add it to your $Profile so it is always available.
        
        To dot source:
            .".\Set-CellColor.ps1"
            
    .PARAMETER Property
        Property, or column that you will be keying on.  
    .PARAMETER Color
        Name or 6-digit hex value of the color you want the cell to be
    .PARAMETER InputObject
        HTML you want the script to process.  This can be entered directly into the
        parameter or piped to the function.
    .PARAMETER Filter
        Specifies a query to determine if a cell should have its color changed.  $true
        results will make the color change while $false result will return nothing.
        
        Syntax
        <Property Name> <Operator> <Value>
        
        <Property Name>::= the same as $Property.  This must match exactly
        <Operator>::= "-eq" | "-le" | "-ge" | "-ne" | "-lt" | "-gt"| "-approx" | "-like" | "-notlike" 
            <JoinOperator> ::= "-and" | "-or"
            <NotOperator> ::= "-not"
        
        The script first attempts to convert the cell to a number, and if it fails it will
        cast it as a string.  So 40 will be a number and you can use -lt, -gt, etc.  But 40%
        would be cast as a string so you could only use -eq, -ne, -like, etc.  
    .PARAMETER Row
        Instructs the script to change the entire row to the specified color instead of the individual cell.
    .INPUTS
        HTML with table
    .OUTPUTS
        HTML
    .EXAMPLE
        get-process | convertto-html | set-cellcolor -Propety cpu -Color red -Filter "cpu -gt 1000" | out-file c:\test\get-process.html

        Assuming Set-CellColor has been dot sourced, run Get-Process and convert to HTML.  
        Then change the CPU cell to red only if the CPU field is greater than 1000.
        
    .EXAMPLE
        get-process | convertto-html | set-cellcolor cpu red -filter "cpu -gt 1000 -and cpu -lt 2000" | out-file c:\test\get-process.html
        
        Same as Example 1, but now we will only turn a cell red if CPU is greater than 100 
        but less than 2000.
        
    .EXAMPLE
        $HTML = $Data | sort server | ConvertTo-html -head $header | Set-CellColor cookedvalue red -Filter "cookedvalue -gt 1"
        PS C:\> $HTML = $HTML | Set-CellColor Server green -Filter "server -eq 'dc2'"
        PS C:\> $HTML | Set-CellColor Path Yellow -Filter "Path -like ""*memory*""" | Out-File c:\Test\colortest.html
        
        Takes a collection of objects in $Data, sorts on the property Server and converts to HTML.  From there 
        we set the "CookedValue" property to red if it's greater then 1.  We then send the HTML through Set-CellColor
        again, this time setting the Server cell to green if it's "dc2".  One more time through Set-CellColor
        turns the Path cell to Yellow if it contains the word "memory" in it.
        
    .EXAMPLE
        $HTML = $Data | sort server | ConvertTo-html -head $header | Set-CellColor cookedvalue red -Filter "cookedvalue -gt 1" -Row
        
        Now, if the cookedvalue property is greater than 1 the function will highlight the entire row red.
        
    .NOTES
        Author:             Martin Pugh
        Twitter:            @thesurlyadm1n
        Spiceworks:         Martin9700
        Blog:               www.thesurlyadmin.com
          
        Changelog:
            1.5             Added ability to set row color with -Row switch instead of the individual cell
            1.03            Added error message in case the $Property field cannot be found in the table header
            1.02            Added some additional text to help.  Added some error trapping around $Filter
                            creation.
            1.01            Added verbose output
            1.0             Initial Release
    .LINK
        http://community.spiceworks.com/scripts/show/2450-change-cell-color-in-html-table-with-powershell-set-cellcolor
    #>

    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true,Position=0)]
        [string]$Property,
        [Parameter(Mandatory=$true,Position=1)]
        [string]$Color,
        [Parameter(Mandatory=$true,ValueFromPipeline=$true)]
        [Object[]]$InputObject,
        [Parameter(Mandatory=$true)]
        [string]$Filter,
        [switch]$Row
    )
    
    Begin {
        Write-Verbose "$(Get-Date): Function Set-CellColor begins"
        If ($Filter)
        {   If ($Filter.ToUpper().IndexOf($Property.ToUpper()) -ge 0)
            {   $Filter = $Filter.ToUpper().Replace($Property.ToUpper(),"`$Value")
                Try {
                    [scriptblock]$Filter = [scriptblock]::Create($Filter)
                }
                Catch {
                    Write-Warning "$(Get-Date): ""$Filter"" caused an error, stopping script!"
                    Write-Warning $Error[0]
                    Exit
                }
            }
            Else
            {   Write-Warning "Could not locate $Property in the Filter, which is required.  Filter: $Filter"
                Exit
            }
        }
    }
    
    Process {
        ForEach ($Line in $InputObject)
        {   If ($Line.IndexOf("<tr><th") -ge 0)
            {   Write-Verbose "$(Get-Date): Processing headers..."
                $Search = $Line | Select-String -Pattern '<th ?[a-z\-:;"=]*>(.*?)<\/th>' -AllMatches
                $Index = 0
                ForEach ($Match in $Search.Matches)
                {   If ($Match.Groups[1].Value -eq $Property)
                    {   Break
                    }
                    $Index ++
                }
                If ($Index -eq $Search.Matches.Count)
                {   Write-Warning "$(Get-Date): Unable to locate property: $Property in table header"
                    Exit
                }
                Write-Verbose "$(Get-Date): $Property column found at index: $Index"
            }
            If ($Line -match "<tr( style=""background-color:.+?"")?><td")
            {   $Search = $Line | Select-String -Pattern '<td ?[a-z\-:;"=]*>(.*?)<\/td>' -AllMatches
                $Value = $Search.Matches[$Index].Groups[1].Value -as [double]
                If (-not $Value)
                {   $Value = $Search.Matches[$Index].Groups[1].Value
                }
                If (Invoke-Command $Filter)
                {   If ($Row)
                    {   Write-Verbose "$(Get-Date): Criteria met!  Changing row to $Color... - matched $($Filter)"
                        If ($Line -match "<tr style=""background-color:(.+?)"">")
                        {   $Line = $Line -replace "<tr style=""background-color:$($Matches[1])","<tr style=""background-color:$Color"
                        }
                        Else
                        {   $Line = $Line.Replace("<tr>","<tr style=""background-color:$Color"">")
                        }
                    }
                    Else
                    {   Write-Verbose "$(Get-Date): Criteria met!  Changing cell to $Color... matched $( ($Filter).ToString() )"
                        $Line = $Line.Replace($Search.Matches[$Index].Value,"<td style=""background-color:$Color"">$Value</td>")
                    }
                }
            }
            Write-Output $Line
        }
    }
    
    End {
        Write-Verbose "$(Get-Date): Function Set-CellColor completed"
    }
}

#endregion Set-CellColour HTML Helper

#region Main
Write-Verbose "$(Get-Date): Enter Main Foreach"
	foreach ($DC in $DCServers){
        $Identity = $DC
		Write-Verbose "$(Get-Date): Working with $DC"

        $HealthObject = New-Object -TypeName System.Management.Automation.PSObject
        $HealthObject | Add-Member -MemberType NoteProperty -Name 'ComputerName' -Value $DC

		#region Ping Test 
        Write-Verbose "$(Get-Date): Testing Ping for $DC"
		#if ( Test-Connection -ComputerName $DC -Count 1 -ErrorAction SilentlyContinue ) {	
		if ( Get-WmiObject -ComputerName $DC 'Win32_ComputerSystem' -ErrorAction SilentlyContinue ) {
		#if ($true) {
		    Write-Host $DC `t Ping Success -ForegroundColor Green
            $HealthObject | Add-Member -MemberType NoteProperty -Name 'Test-Ping' -Value PASS
		#endregion Ping Test
		
#region foreach services test
                Write-Verbose "$(Get-Date): Service: $testServices $DC"
				$serviceStatus = start-job -scriptblock {get-service -ComputerName $($args[0]) -Name $($args[1]) -ErrorAction SilentlyContinue} -ArgumentList $DC, $testServices
				Wait-Job $serviceStatus -timeout $timeout
				if($serviceStatus.state -like "Running") # after the timeout, its still going so give up
				{
					Write-Host $DC `t Service TimeOut -ForegroundColor Yellow
					Stop-Job $serviceStatus
                   
                    foreach ($loopService in $testServices) {
                    $HealthObject | Add-Member -MemberType NoteProperty -Name "Service-$($loopService)" -Value TIME
                    Write-Host $DC `t $loopService Service Failed -ForegroundColor Red}
                    #$HealthObject | Add-Member -MemberType NoteProperty -Name Services-All -Value TIME
				}
				else
				{
					$serviceStatus1 = Receive-job $serviceStatus
					if (!($serviceStatus1))
                        {
                           # $HealthObject | Add-Member -MemberType NoteProperty -Name $thisServiceString -Value 'FAIL'
                           # Write-Host $DC `t Service Failed -ForegroundColor Red

                            foreach ($loopService in $testServices) {
                                $HealthObject | Add-Member -MemberType NoteProperty -Name "Service-$($loopService)" -Value FAIL
                                Write-Host $DC `t $loopService Service Failed -ForegroundColor Red}
                                
                                

                        }
                        else
                        {
                           foreach ($loopService in $testServices) {
                                $thisLoopService = $serviceStatus1 | Where-Object {$_.Name -eq $loopService}
                                
                                if ($thisLoopService.Status){
                                $HealthObject | Add-Member -MemberType NoteProperty -Name "Service-$($thisLoopService.Name)" -Value $thisLoopService.status
                                Write-Host $DC `t $thisLoopService Service Passed -ForegroundColor Green}
                                else{
                                $HealthObject | Add-Member -MemberType NoteProperty -Name "Service-$($loopService)" -Value MISSING
                                Write-Host $DC `t $loopService Service Missing -ForegroundColor Magenta
                                    }
                                }
                        }
				}

#} #end foreach test services
#endregion foreach services test

#region foreach DCDIAGS test
   Foreach ($thisDCDIAG in $testDCDIAG) {
                $thisDCDIAGString = "DCDIAG-$thisDCDIAG"
                Write-Verbose "$(Get-Date): DCDIAG-$thisDCDIAG $DC"
				add-type -AssemblyName microsoft.visualbasic 
				$cmp = "microsoft.visualbasic.strings" -as [type]
				$sysvol = start-job -scriptblock {dcdiag /test:$($args[1]) /s:$($args[0])} -ArgumentList $DC, $thisDCDIAG
				wait-job $sysvol -timeout $timeout
				if($sysvol.state -like "Running")
				{
					Write-Host $DC `t $thisDCDIAGString Test TimeOut -ForegroundColor Yellow
					stop-job $sysvol
                    $HealthObject | Add-Member -MemberType NoteProperty -Name $thisDCDIAGString -Value TIME
				}
				else
				{
					$sysvol1 = Receive-job $sysvol
					if($cmp::instr($sysvol1, "passed test $thisDCDIAG"))
					{
						Write-Host $DC `t $thisDCDIAG Test passed -ForegroundColor Green
						$HealthObject | Add-Member -MemberType NoteProperty -Name $thisDCDIAGString -Value PASS
					}
					else
					{
						Write-Host $DC `t $thisDCDIAG Test Failed -ForegroundColor Red
						$HealthObject | Add-Member -MemberType NoteProperty -Name $thisDCDIAGString -Value FAIL
                        $sysvol1 | Out-File $($reportpath + $DC + "_" + $thisDCDIAGString + "_" + $date + ".txt")
					}
				}
        } 
#endregion foreach dcdiags 

			
			#region AntiVirus status
                Write-Verbose "$(Get-Date): Security-AVDAT: $DC"
				$sysvol = start-job -scriptblock {  
					try {
						$Reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine',$($args[0]))
						$RegKey= $Reg.OpenSubKey("SOFTWARE\\McAfee\\AVEngine")
						$datDate = $RegKey.GetValue("AVDATDate")
					}
					
					catch {
						$Reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine',$($args[0]))
						$RegKey= $Reg.OpenSubKey("SOFTWARE\\Wow6432Node\\McAfee\\AVEngine")
						$datDate = $RegKey.GetValue("AVDATDate")
					}
					if ($datDate) {Get-Date $datDate}else{Get-Date '1955-11-05' -Format 'dd/MM/yyyy'}
				} -ArgumentList $DC
				
				Wait-Job $sysvol -timeout $timeout
				
				if($sysvol.state -like "Running")
				{
					Write-Host $DC `t AntiVirus Test TimeOut -ForegroundColor Yellow
					#Add-Content $report "<td bgcolor= $testHTMLWarningColour align=center><B>!</B></td>"
					Stop-Job $sysvol
                    $HealthObject | Add-Member -MemberType NoteProperty -Name 'Security-AVDAT' -Value TIME
				}
				else
				{
					$sysvol1 = Receive-Job $sysvol
					if($sysvol1 -gt $((Get-Date).AddDays(-3)))
					{
						Write-Host $DC `t AV Test passed -ForegroundColor Green
						$HealthObject | Add-Member -MemberType NoteProperty -Name 'Security-AVDAT' -Value PASS
					}
					else
					{
						Write-Host $DC `t AV Test Failed -ForegroundColor Red
                        $HealthObject | Add-Member -MemberType NoteProperty -Name 'Security-AVDAT' -Value $(Get-date $sysvol1 -Format 'dd/MM/yyyy HH:mm')
					}
				}
			#endregion AntiVirus status
			
			#region Uptime
                Write-Verbose "$(Get-Date): System-Uptime: $DC"
				$wmi = 0
				$upt=0
				$Lastb =0
				$wmijob = start-job -scriptblock {Get-WmiObject -ComputerName $($args[0]) -Query "SELECT LastBootUpTime FROM Win32_OperatingSystem"} -ArgumentList $DC
				wait-job $wmijob -timeout $timeout
				if($wmijob.state -like "Running")
				{
					Write-Host $DC `t Uptime Test TimeOut -ForegroundColor Yellow
					Stop-Job $wmijob
                    $HealthObject | Add-Member -MemberType NoteProperty -Name 'System-UptimeDays' -Value TIME
				}
				else
				{
					$wmi = Receive-job $wmijob
					$now = 0
					$now = Get-Date
					$boottime = 0
					$boottime = [System.Management.ManagementDateTimeConverter]::ToDateTime($wmi.LastBootUpTime) 
					$uptime = 0
					$uptime = $now - $boottime
					$d = 0
					$d =$uptime.days
					$h = 0
					$h =$uptime.hours
					$m = 0
					$m =$uptime.Minutes
					$s = 0
					$s = $uptime.Seconds
					$upt = "$d Days $h Hours $m Min $s Sec"
					
                    if($boottime -eq 0){$boottime = '1955-11-05'}
                    $Lastb = $boottime | Get-Date -Format 'dd/MM/yyyy HH:mm'
					
					$HealthObject | Add-Member -MemberType NoteProperty -Name 'System-UptimeDays' -Value $d
                    $HealthObject | Add-Member -MemberType NoteProperty -Name 'System-BootTime' -Value $Lastb
					Write-Host $DC `t Uptime $upt Lastboot $Lastb   -ForegroundColor Green
					
					
				}
			#endregion Uptime	       
			
			#region C Drive Utilisation
                Write-Verbose "$(Get-Date): Drive-C: $DC"
				$wmi = 0
				$cdriv = 0
				$wmijob = start-job -scriptblock {Get-WmiObject win32_logicaldisk -ComputerName $($args[0]) |  Where-Object {$_.drivetype -eq 3}} -ArgumentList $DC
				wait-job $wmijob -timeout $timeout
				if($wmijob.state -like "Running")
				{
					Write-Host $DC `t Uptime Test TimeOut -ForegroundColor Yellow
					stop-job $wmijob
                    $HealthObject | Add-Member -MemberType NoteProperty -Name 'Drive-C' -Value TIME
				}
				else
				{
				$wmi = Receive-job $wmijob
					if($wmi)
					{
						$cdriv = $wmi | where{$_.DeviceID -eq "C:"}
					    $fspace = [math]::truncate($cdriv.Freespace/1Gb)
					    Write-Host $DC `t C Drive Free Space $fspace   -ForegroundColor Green
                        $HealthObject | Add-Member -MemberType NoteProperty -Name 'Drive-C' -Value $fspace
					}
					else
					{
						$fspace = 'Error'
					    Write-Host $DC `t C Drive Free Space $fspace   -ForegroundColor Red
                        $wmi | Out-File $($reportpath + $DC + "_DiskSpace_$date.txt")
                        $HealthObject | Add-Member -MemberType NoteProperty -Name 'Drive-C' -Value FAIL
                    }
				}
			#endregion C Drive Utilisation 
            
            #region NTDS Drive Utilisation
                Write-Verbose "$(Get-Date): Drive-NTDS: $DC"
                $wmi = 0
				$cdriv = 0
				$wmijob = start-job -scriptblock {
                    
                    $Reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine',$($args[0]))
                    $RegKey= $Reg.OpenSubKey("SYSTEM\\CurrentControlSet\\services\\NTDS\\Parameters")
                    $NTDSDrive = $RegKey.GetValue("DSA Database file")
                    $NTDSDrive = ($NTDSDrive -split "\\")[0]	
                    
                    Get-WmiObject win32_logicaldisk -ComputerName $($args[0]) |  Where-Object {$_.drivetype -eq 3} | Where-Object {$_.DeviceID -eq $NTDSDrive}

                    } -ArgumentList $DC
				wait-job $wmijob -timeout $timeout
				if($wmijob.state -like "Running")
				{
					Write-Host $DC `t Uptime Test TimeOut -ForegroundColor Yellow
					stop-job $wmijob
                    $HealthObject | Add-Member -MemberType NoteProperty -Name 'Drive-NTDS' -Value TIME
				}
				else
				{
				$wmi = Receive-job $wmijob
					if($wmi)
					{
					    $fspace = [math]::truncate($wmi.Freespace/1Gb)
					    Write-Host $DC `t NTDS Drive Free Space $fspace   -ForegroundColor Green
                        $HealthObject | Add-Member -MemberType NoteProperty -Name 'Drive-NTDS' -Value $fspace
					}
					else
					{
						$fspace = 'Error'
					    Write-Host $DC `t NTDS Drive Free Space $fspace   -ForegroundColor Red
                        $wmi | Out-File $($reportpath + $DC + "_DiskSpace_$date.txt")
                        $HealthObject | Add-Member -MemberType NoteProperty -Name 'Drive-NTDS' -Value FAIL
                    }
				}
			#endregion NTDS Drive Utilisation 
            
            
           #region WSUS status
                Write-Verbose "$(Get-Date): WSUS-PatchDate: $DC"
				$sysvol = Start-Job -scriptblock {  
                    (Get-HotFix -ComputerName $($args[0]) | Where-Object {$_.hotfixid -ne "file 1"} | Select @{label="InstalledOn";e={[DateTime]::Parse($_.psbase.properties["installedon"].value,$([System.Globalization.CultureInfo]::GetCultureInfo("en-US")))}} | Sort-Object InstalledOn | Select-Object -Last 1).InstalledOn
                        } -ArgumentList $DC
				
				Wait-Job $sysvol -timeout $timeout
				
				if($sysvol.state -like "Running")
				{
					Write-Host $DC `t WSUS Test TimeOut -ForegroundColor Yellow
					Stop-Job $sysvol
                    $HealthObject | Add-Member -MemberType NoteProperty -Name 'WSUS-PatchDate' -Value TIME
                    $HealthObject | Add-Member -MemberType NoteProperty -Name 'WSUS-DaysSince' -Value TIME
				}
				else
				{
                    $jobWSUSResult = Receive-job $sysvol
					if($jobWSUSResult){
                    #jobWSUSResult1 = [DateTime]::Parse($jobWSUSResult,$([System.Globalization.CultureInfo]::GetCultureInfo("en-GB")))

                    $now = 0
					$now = Get-Date
					
                    $WSUSAge = 0
                    $WSUSAge = $now - $jobWSUSResult
					
                    $WSUSDays = 0
					$WSUSDays =$WSUSAge.days
					
					$WSUSLastPatched = $jobWSUSResult | Get-Date -Format 'dd/MM/yyyy HH:mm'
					
					$HealthObject | Add-Member -MemberType NoteProperty -Name 'WSUS-DaysSince' -Value $WSUSDays
                    $HealthObject | Add-Member -MemberType NoteProperty -Name 'WSUS-PatchDate' -Value $WSUSLastPatched
					Write-Host $DC `t WSUS Days Old $WSUSDays Last Patched $WSUSLastPatched   -ForegroundColor Green}
					else { 
                    Write-Host $DC `t No data returned   -ForegroundColor Red 
                    
                    $HealthObject | Add-Member -MemberType NoteProperty -Name 'WSUS-DaysSince' -Value FAIL
                    $HealthObject | Add-Member -MemberType NoteProperty -Name 'WSUS-PatchDate' -Value FAIL
                    }
				}
#endregion WSUS status
			
           
            
            Write-Verbose "$(Get-Date): Adding `$HealthObject to `$HealthReport"
			$HealthReport.Add($HealthObject)
			Write-Verbose "$(Get-Date): End if ping -eq `$true"
		} # End if PING
		else # If the Ping fails - skip all tests
		{
			Write-Verbose "$(Get-Date): Ping failed"
            Write-Host $DC `t Ping Fail -ForegroundColor Red
			$HealthObject | Add-Member -MemberType NoteProperty -Name 'Test-Ping' -Value FAIL
            $HealthReport.Add($HealthObject)
		} # end else
     Write-Verbose "$(Get-Date): End Foreach DC"
	} # end foreach $dc 

<#region DNS A Record Check
if($singleDomain) {
$fqdn = Resolve-DnsName $singleDomain 
$res = $fqdn | foreach {Resolve-DnsName $_.IPAddress} | foreach {Resolve-DnsName $_.NameHost}
$dclist = $getForest.domains | Where-Object {$_.Name -eq $singleDomain} | ForEach-Object {$_.DomainControllers} | ForEach-Object {$_.Name} | Resolve-DnsName
$compare = Compare-Object $fqdn $dclist -Property IPAddress
$DNSTestHTML = $compare | select IPAddress, @{n="Location";e={ if ($_.SideIndicator -eq '=>') { "DC List" }  else { "DNS Record" } }} | ConvertTo-Html -Fragment
}
#endregion DNS A Record Check #>

#region Sites and Site Links Check

#$SitesandLinksTest = Get-ADObject -Filter "objectClass -eq 'siteLink'" -Searchbase (([System.DirectoryServices.DirectoryEntry] "LDAP://RootDSE").Get("configurationNamingContext")) -Property Options, Cost, ReplInterval, SiteList, Schedule, Description, WhenCreated | Select-Object Name, Description, @{Name="SiteCount";Expression={$_.SiteList.Count}}, Cost, ReplInterval, @{Name="Schedule";Expression={If($_.Schedule){If(($_.Schedule -Join "").Contains("240")){"NonDefault"}Else{"24x7"}} Else{"24x7"}}}, Options, WhenCreated | Sort Name | Converto-Html -Fragment

#endregion Site and Services Links Check

#region DC Data
            $dcDataTable = $getDCData | Select Name, IPAddress, SiteName, CurrentTime, @{n='USN';e={$_.HighestCommittedUsn}}, OSVersion, @{n='Roles';e={($_.Roles) -join ", "}}
            $dcDataTableHTML = $dcDataTable | ConvertTo-Html -Fragment
            $dcDataTable | Export-Clixml $objDCDataName
            #export this at some point - thats why it is an obj and HTML
#endregion DC Data


Write-Verbose "$(Get-Date): End Main"	
#endregion Main

#$HealthReport = $HealthReport | Sort-Object -Property Test-Ping

Write-Verbose "$(Get-Date): Create HTML Table Headers"
$HealthHeaders = $HealthReport | Get-Member | Where-Object {$_.MemberType -eq 'NoteProperty'} | Select-Object Name


$temp1=@()
foreach ($HealthObject in $HealthReport) {
        if ($HealthObject.'Test-Ping' -eq 'FAIL') 
            {
                foreach ($Header in $HealthHeaders)
                    {
                        $HealthObject | Add-Member -MemberType NoteProperty -Name $header.name -Value FAIL
                    }
            } 
        $temp1 += $HealthObject
        $healthreport.add($temp1)
        $temp1 = $null
        }

$HealthReport = $HealthReport | Sort-Object 'Test-Ping' -Descending
  
$healthreport | Export-Clixml $objExportname

#region Advanced HTML Output
Write-Verbose "$(Get-Date): Start Advanced HTML Output"

Write-Verbose "$(Get-Date): Convert `$HealthReport to HTML Fragment"
$AdvancedHTML = $HealthReport | ConvertTo-Html -Fragment

#region Colour Code the Cells
#region HealthReport Colour
foreach ($header in $HealthHeaders)
    { 
        Write-Verbose "$(Get-Date): Switching $($Header.Name) for cell colours"
        switch -Wildcard ($header.name)
        {
                Security-AVDAT {$AdvancedHTML = $AdvancedHTML | Set-CellColor -Property Security-AVDAT -Color LightCoral -Filter "Security-AVDAT -like ""*/*"" "}
				
				Drive-* {
                $AdvancedHTML = $AdvancedHTML | Set-CellColor -Property $Header.Name -Color Red -Filter "$Header.Name -le 5"
                $AdvancedHTML = $AdvancedHTML | Set-CellColor -Property $Header.Name -Color Khaki -Filter "$Header.Name -gt 5"
                $AdvancedHTML = $AdvancedHTML | Set-CellColor -Property $Header.Name -Color PaleGreen -Filter "$Header.Name -gt 10"
                $AdvancedHTML = $AdvancedHTML | Set-CellColor -Property $Header.Name -Color PaleGreen -Filter "$Header.Name -eq 'Running'"
                $AdvancedHTML = $AdvancedHTML | Set-CellColor -Property $Header.Name -Color LightCoral -Filter "$Header.Name -eq 'Stopped'"
                $AdvancedHTML = $AdvancedHTML | Set-CellColor -Property $Header.Name -Color Khaki -Filter "$Header.Name -eq 'Paused'"
                $AdvancedHTML = $AdvancedHTML | Set-CellColor -Property $Header.Name -Color LightSteelBlue -Filter "$Header.Name -eq 'TIME'"
                $AdvancedHTML = $AdvancedHTML | Set-CellColor -Property $Header.Name -Color LightCoral -Filter "$Header.Name -eq 'FAIL'"
                $AdvancedHTML = $AdvancedHTML | Set-CellColor -Property $Header.Name -Color Orange -Filter "$Header.Name -eq 'MISSING'"}

                System-UptimeDays {
                $AdvancedHTML = $AdvancedHTML | Set-CellColor -Property System-UptimeDays -Color PaleGreen -Filter "System-UptimeDays -ge 0"
                $AdvancedHTML = $AdvancedHTML | Set-CellColor -Property System-UptimeDays -Color Khaki -Filter "System-UptimeDays -gt 30"
                $AdvancedHTML = $AdvancedHTML | Set-CellColor -Property System-UptimeDays -Color LightCoral -Filter "System-UptimeDays -gt 35"
                $AdvancedHTML = $AdvancedHTML | Set-CellColor -Property System-UptimeDays -Color Red -Filter "System-UptimeDays -gt 60"}
        
                WSUS-DaysSince {
                $AdvancedHTML = $AdvancedHTML | Set-CellColor -Property WSUS-DaysSince -Color PaleGreen -Filter "WSUS-DaysSince -ge 0"
                $AdvancedHTML = $AdvancedHTML | Set-CellColor -Property WSUS-DaysSince -Color Khaki -Filter "WSUS-DaysSince -gt 35"
                $AdvancedHTML = $AdvancedHTML | Set-CellColor -Property WSUS-DaysSince -Color LightCoral -Filter "WSUS-DaysSince -gt 40"
                $AdvancedHTML = $AdvancedHTML | Set-CellColor -Property WSUS-DaysSince -Color Red -Filter "WSUS-DaysSince -gt 60"
                $AdvancedHTML = $AdvancedHTML | Set-CellColor -Property WSUS-DaysSince -Color PaleGreen -Filter "WSUS-DaysSince -eq 'Running'"
                $AdvancedHTML = $AdvancedHTML | Set-CellColor -Property WSUS-DaysSince -Color LightCoral -Filter "WSUS-DaysSince -eq 'Stopped'"
                $AdvancedHTML = $AdvancedHTML | Set-CellColor -Property WSUS-DaysSince -Color Khaki -Filter "WSUS-DaysSince -eq 'Paused'"
                $AdvancedHTML = $AdvancedHTML | Set-CellColor -Property WSUS-DaysSince -Color LightSteelBlue -Filter "WSUS-DaysSince -eq 'TIME'"
                $AdvancedHTML = $AdvancedHTML | Set-CellColor -Property WSUS-DaysSince -Color LightCoral -Filter "WSUS-DaysSince -eq 'FAIL'"
                $AdvancedHTML = $AdvancedHTML | Set-CellColor -Property WSUS-DaysSince -Color Orange -Filter "WSUS-DaysSince -eq 'MISSING'"}

                WSUS-PatchDate {$AdvancedHTML = $AdvancedHTML | Set-CellColor -Property 'WSUS-PatchDate' -Color Lavender -Filter "WSUS-PatchDate -like ""*/*"" "}

                # ComputerName and BootTime to Lavender
                Computername {$AdvancedHTML = $AdvancedHTML | Set-CellColor -Property ComputerName -Color Lavender -Filter "Computername -like ""*.*"" "}
                System-BootTime {$AdvancedHTML = $AdvancedHTML | Set-CellColor -Property System-BootTime -Color Lavender -Filter "System-BootTime -like ""*"" "}
            
                Service-* {
                # Main - Pass Fail Warn Time
                $AdvancedHTML = $AdvancedHTML | Set-CellColor -Property $header.Name -Color PaleGreen -Filter "$header.Name -eq 'Running'"
                $AdvancedHTML = $AdvancedHTML | Set-CellColor -Property $header.Name -Color LightCoral -Filter "$header.Name -eq 'Stopped'"
                $AdvancedHTML = $AdvancedHTML | Set-CellColor -Property $header.Name -Color Khaki -Filter "$header.Name -eq 'Paused'"
                $AdvancedHTML = $AdvancedHTML | Set-CellColor -Property $header.Name -Color LightSteelBlue -Filter "$header.Name -eq 'TIME'"
                $AdvancedHTML = $AdvancedHTML | Set-CellColor -Property $header.Name -Color LightCoral -Filter "$header.Name -eq 'FAIL'"
                $AdvancedHTML = $AdvancedHTML | Set-CellColor -Property $header.Name -Color Orange -Filter "$header.Name -eq 'MISSING'"
                    }

                Test-Ping {$AdvancedHTML = $AdvancedHTML | Set-CellColor -Property 'Test-Ping' -Color LightCoral -Filter "Test-Ping -eq 'FAIL'"}

                default {
                # Main - Pass Fail Warn Time
                $AdvancedHTML = $AdvancedHTML | Set-CellColor -Property $header.Name -Color PaleGreen -Filter "$header.Name -eq 'PASS'"
                $AdvancedHTML = $AdvancedHTML | Set-CellColor -Property $header.Name -Color LightCoral -Filter "$header.Name -eq 'FAIL'"
                $AdvancedHTML = $AdvancedHTML | Set-CellColor -Property $header.Name -Color Khaki -Filter "$header.Name -eq 'WARN'"
                $AdvancedHTML = $AdvancedHTML | Set-CellColor -Property $header.Name -Color LightSteelBlue -Filter "$header.Name -eq 'TIME'"
                    }
        } #end switch
    }
    #endregion HealthReport
    if ($SitesandLinksTest){
    $SitesandLinksTest = $SitesandLinksTest | Set-CellColor -Property $header.Name -Color PaleGreen -Filter "$header.Name -eq 'PASS'"}

    #endregion Colour Code

    # Swap out the Pass Warn Fail words with cool HTML icons.
    Write-Verbose "$(Get-Date): Swapping words with HTML icons"
    $AdvancedHTML = $AdvancedHTML -Replace '>PASS<', '>&#x2713;<'
    $AdvancedHTML = $AdvancedHTML -Replace '>FAIL<', '>&#x2717;<'
    $AdvancedHTML = $AdvancedHTML -Replace '>WARN<', '>!<'
    $AdvancedHTML = $AdvancedHTML -Replace '>TIME<', '>&#x23F0;<'
    $AdvancedHTML = $AdvancedHTML -Replace '>MISSING<', '>??<'

    $AdvancedHTML = $AdvancedHTML -Replace '>Running<', '>&#x2713;<'
    $AdvancedHTML = $AdvancedHTML -Replace '>Stopped<', '>&#x2717;<'
    # tick &#x2713; cross &#x2717; clock &#x23F0; hourglass &#x231B;
   $AdvancedreportStart = @" 
<html> 
	<head>
		<meta http-equiv='Content-Type' content='text/html; charset=iso-8859-1'>
		<title>AD Status Report</title>
		<STYLE TYPE="text/css">
			<!--
				td {
				    font-family: "Segoe UI", Tahoma;
				    font-size: 11px;
				    border-top: 0px solid #999999;
				    border-right: 0px solid #999999;
				    border-bottom: 0px solid #999999;
				    border-left: 0px solid #999999;
				    padding-top: 0px;
				    padding-right: 0px;
				    padding-bottom: 0px;
				    padding-left: 0px;
                    text-align: center;
				}
				th {
				    font-family: "Segoe UI", Tahoma;
				    font-size: 11px;
				    border-top: 0px solid #999999;
				    border-right: 0px solid #999999;
				    border-bottom: 0px solid #999999;
				    border-left: 0px solid #999999;
				    padding-top: 0px;
				    padding-right: 0px;
				    padding-bottom: 0px;
				    padding-left: 0px;
					background-color:Lavender;
                    text-align: center;
				}
				body {
				    margin-left: 5px;
				    margin-top: 5px;
				    margin-right: 0px;
				    margin-bottom: 10px;
				    }
				table {
				    border: thin solid #999999;
				}
			-->
		</style>
	</head>
	<body>
		<table width='100%'>
			<tr bgcolor='Lavender'>
				<td colspan='7' height='25' align='center'>
					<font face='Segoe UI' color='#003399' size='4'><strong>$($subject) $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</strong></font>
				</td>
			</tr>
		</table>
"@

Write-Verbose "$(Get-Date): Exporting to $advreportname"

#region composing html
Add-content $advreportname $AdvancedreportStart # the headers
Add-content $advreportname $AdvancedHTML # the processed coolness
if($DNSTestHTML) {  Add-Content $advreportname "<br /> <br /> The following table displays IPs that are not identical between the list of DCs from AD and the DNS A record for the domain.<br />"
                    Add-Content $advreportname $DNSTestHTML}
if($dcDataTableHTML) {  Add-Content $advreportname "<br /> <br />"
                        Add-Content $advreportname $dcDataTableHTML} #DC Data Table
Add-Content $advreportname "Report generated at $date from $env:COMPUTERNAME"
Add-Content $advreportname "</body>" # close
Add-Content $advreportname "</html>" # close
#endregion composing html

#endregion Advanced HTML Output

#region send errors to text file
$Error | Out-File $objErrorsName
#endregion send errors to text file

#region Send Email
    Write-Verbose "$(Get-Date): Starting to send email"
    $body = Get-Content $advreportname | Out-String
	$attachments = Get-ChildItem .\$guid\* | Select FullName
    if ($attachments) 
        {Write-Verbose "$(Get-Date): Adding attachments, Sending mail to $recipients"
		$mailAttachments = @()
		$attachments | foreach {$mailAttachments += $_.fullname}
		Send-MailMessage -Subject $subject -To $recipients -From $from -Body $body -BodyAsHtml -SmtpServer $smtphost -Attachments $mailAttachments}
    else
        {Write-Verbose "$(Get-Date): No attachments, Sending mail to $recipients"; Send-MailMessage -Subject $subject -To $recipients -From $from -Body $body -BodyAsHtml -SmtpServer $smtphost}
#endregion Send Email

#end			

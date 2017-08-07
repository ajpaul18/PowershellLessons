set-psdebug -strict


#powershell ".\Get-OutlookLogs.ps1" > C:\test.log


$DebugPreference = 'Continue'
<#
1.  Collect logs using the following tools from an impacted computer with Outlook:
a.       Turn on Outlook Troubleshooting Logging.
b.       Psping to outlook.office365.com
c.       Tracert to outlook.office365.com
d.       Collect network traffic using Netmon
e.       Collect HTTPS logs using Flidder Trace (more details here)

#>

function SetLoggingKey([bool]$bLogSetting)
{  
    $RegistryPath = 'HKCU:\SOFTWARE\Microsoft\Office\16.0\Outlook\Options\Mail'
    $Name = 'EnableLogging'

    $Value = 0
    if ($bLogSetting -eq $true)
    {
        $Value = 1
    }
    elseif ($bLogSetting -eq $false) 
    {
        $Value = 0
    }
    
    Write-Debug "Setting Outlook troubleshooting logging to [$Value]"
    #if the path doesn't exist, create it
    If((Test-Path $RegistryPath) -eq $false)
    {
        New-Item -Path $RegistryPath -Force | Out-Null
    }
    New-ItemProperty -Path $RegistryPath -Name $Name -Value $Value -PropertyType DWORD -Force | Out-Null

}


##### Script Starts Here ######  

#script "constants"
$logFile = "$PSScriptRoot\filename.cap"
$ServerNames = @("outlook.office365.com")
$strSectionDivision = "-" * 40

#todo: fix the outlook debug stuff
#Write-Host 'Setting Outlook troubleshooting logging to on'
#SetLoggingKey $true
 
#Write-Debug "RUNNING PSPING"
#"RUNNING PSPING" | out-File $logFile

Write-Debug "Using target servers [$ServerNames]"
"Using target servers [$ServerNames]" | out-File $logFile -append


foreach ($Server in $ServerNames)
{ 
    $arrCommands = `
        @(
            "cmd.exe /c ping outlook.office365.com",
            "cmd.exe /c tracert outlook.office365.com",
            "cmd.exe /c nmcap /network * /capture /file $PSScriptRoot\filename.cap:50M"            
        )
    #Test-NetConnection outlook.office365.com -TraceRoute
    #don't use the above one - it takes longer to run than tracert

    foreach ($strCmd in $arrCommands)
    {
        Write-Debug "Running [$strCmd]"
        $arrCmdResult = Invoke-Expression -Command $strCmd
        $strCmdResult = $arrCmdResult -join "`n"

        Write-Debug "writing to file"
        "$strSectionDivision`n" | out-File $logFile -append        
        "Results of [$strCmd]`n" | out-File $logFile -append
        "$strSectionDivision`n" | out-File $logFile -append
        "$strCmdResult`n" | out-File $logFile -append
    }
}


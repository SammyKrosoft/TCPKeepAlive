<#
.SYNOPSIS
    Quick description of this script

.DESCRIPTION
    Longer description of what this script does

.PARAMETER FirstNumber
    This parameter does blablabla

.PARAMETER CheckVersion
    This parameter will just dump the script current version.

.INPUTS
    None. You cannot pipe objects to that script.

.OUTPUTS
    None for now

.EXAMPLE
.\Do-Something.ps1
This will launch the script and do someting

.EXAMPLE
.\Do-Something.ps1 -CheckVersion
This will dump the script name and current version like :
SCRIPT NAME : Do-Something.ps1
VERSION : v1.0

.NOTES
None

.LINK
    https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_comment_based_help?view=powershell-6

.LINK
    https://github.com/SammyKrosoft
#>
[CmdLetBinding(DefaultParameterSetName = "NormalRun")]
Param(
    [Parameter(Mandatory = $False, Position = 1, ParameterSetName = "NormalRun")][string[]]$Servers,
    [Parameter(Mandatory = $false, Position = 1, ParameterSetName = "CheckOnly")][switch]$CheckVersion
)

<# ------- SCRIPT_HEADER (Only Get-Help comments and Param() above this point) ------- #>
#Initializing a $Stopwatch variable to use to measure script execution
$stopwatch = [system.diagnostics.stopwatch]::StartNew()
#Using Write-Debug and playing with $DebugPreference -> "Continue" will output whatever you put on Write-Debug "Your text/values"
# and "SilentlyContinue" will output nothing on Write-Debug "Your text/values"
$DebugPreference = "Continue"
# Set Error Action to your needs
$ErrorActionPreference = "SilentlyContinue"
#Script Version
$ScriptVersion = "0.1"
<# Version changes
v0.1 : first script version
v0.1 -> v0.5 : 
#>
$ScriptName = $MyInvocation.MyCommand.Name
If ($CheckVersion) {Write-Host "SCRIPT NAME     : $ScriptName `nSCRIPT VERSION  : $ScriptVersion";exit}
# Log or report file definition
# NOTE: use $PSScriptRoot in Powershell 3.0 and later or use $scriptPath = split-path -parent $MyInvocation.MyCommand.Definition in Powershell 2.0
$scriptPath = split-path -parent $MyInvocation.MyCommand.Definition
$OutputReport = "$ScriptPath\$($ScriptName)_$(get-date -f yyyy-MM-dd-hh-mm-ss).csv"
# Other Option for Log or report file definition (use one of these)
$ScriptLog = "$ScriptPath\$($ScriptName)-$(Get-Date -Format 'dd-MMMM-yyyy-hh-mm-ss-tt').txt"
<# ---------------------------- /SCRIPT_HEADER ---------------------------- #>
<# -------------------------- DECLARATIONS -------------------------- #>
$ADSite = $null
New-Alias -Name "Log" Write-Host

<# /DECLARATIONS #>
<# -------------------------- FUNCTIONS -------------------------- #>

Function Title1 ($title, $TotalLength = 100, $Back = "Yellow", $Fore = "Black") {
    $TitleLength = $Title.Length
    [string]$StarsBeforeAndAfter = ""
    $RemainingLength = $TotalLength - $TitleLength
    If ($($RemainingLength % 2) -ne 0) {
        $Title = $Title + " "
    }
    $Counter = 0
    For ($i=1;$i -le $(($RemainingLength)/2);$i++) {
        $StarsBeforeAndAfter += "*"
        $counter++
    }
    
    $Title = $StarsBeforeAndAfter + $Title + $StarsBeforeAndAfter
    Write-Host $Title -BackgroundColor $Back -foregroundcolor $Fore
    Write-Host
    
}
Function LogRed ($Message){
    Write-Host $message -ForegroundColor Red
}

Function LogGreen ($message){
    Write-Host $message -ForegroundColor Green
}

Function LogYellow ($message){
    Write-Host $message -ForegroundColor Yellow
}

Function LogBlue ($message){
    Write-Host $message -ForegroundColor Blue
}
<# /FUNCTIONS #>
<# -------------------------- EXECUTIONS -------------------------- #>
# Getting local site name
$ADSite = [System.DirectoryServices.ActiveDirectory.ActiveDirectorySite]::GetComputerSite().Name

If ($Servers -eq $null) {
    
    $Message = "No servers specified with the -Servers Server1, Server2, ... property, Local site is $ADSite ... about to load the Exchange cmdlets if not loaded, you MUST have Exchange Management Tools installed locally, or call the script from an Exchange-enabled PowerShell host, do you want to continue ?"
    $Yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes","help";
    $No = New-Object System.Management.Automation.Host.ChoiceDescription "&No","help";
    $choices = [System.Management.Automation.Host.ChoiceDescription[]]($Yes,$no);
    $answer = $host.UI.PromptForChoice($caption,$message,$choices,1)
    
    switch ($answer){
        0 {LogGreen "Continuing Script..."; Start-Sleep -Seconds 3}
        1 {LogRed "Exiting Script..."; exit}
    }

    Title1 "Preparing the environment"
    $CheckSnapin = (Get-PSSnapin | Where {$_.Name -eq "Microsoft.Exchange.Management.PowerShell.E2010"} | Select Name)
    if($CheckSnapin -like "*Exchange.Management.PowerShell*"){
        LogGreen "Exchange Snap-in already loaded, continuing...." -ForegroundColor Green
    }
    Else{
        LogYellow "Loading Exchange Snap-in Please Wait..."
        Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010 -ErrorAction SilentlyContinue
    }
    Title1 "Getting Exchange Servers list"
    $Servers = Get-ExchangeServer | Where-Object {$_.Site -match $ADSite}
    $SErvers | out-host

    #Connect to Each server that it finds from above and open the KeepAliveTime registry key if it exists and record the value.
    Title1 "RCP KeepAlive registry Key settings"
    foreach ($Server in $Servers){
        $EXCHServer = $Server.name
        LogRed "Treating server $EXCHServer"
        $OpenReg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine',$EXCHServer)
        $RegKeyPath = 'SYSTEM\CurrentControlSet\Services\Tcpip\Parameters'
        $RegKey = $OpenReg.OpenSubKey($RegKeyPath)
        $TCPKeepAlive = $RegKey.GetValue('KeepAliveTime')
        $Exists = if($TCPKeepAlive){$true} else {$false}
        
        #Dump the scripts findings into an object.
        $Report = [PSCustomObject]@{
            "Server Name" = $EXCHServer;
            "Key Present" = $Exists;
            "TCP Keep Alive Time" = $TCPKeepAlive}
        
        #Display report on screen
        $Report | out-host
        #Write the output to a report file
        $Report | Export-Csv ($OutputReport) -Append -NoTypeInformation
        notepad $OutputReport
    } 

    }

<# /EXECUTIONS #>
<# -------------------------- CLEANUP VARIABLES -------------------------- #>

<# /CLEANUP VARIABLES#>
<# ---------------------------- SCRIPT_FOOTER ---------------------------- #>
#Stopping StopWatch and report total elapsed time (TotalSeconds, TotalMilliseconds, TotalMinutes, etc...
$stopwatch.Stop()
$msg = "`n`nThe script took $([math]::round($($StopWatch.Elapsed.TotalSeconds),2)) seconds to execute..."
Write-Host $msg
$msg = $null
$StopWatch = $null
<# ---------------- /SCRIPT_FOOTER (NOTHING BEYOND THIS POINT) ----------- #>
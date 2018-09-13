#GetADSites

#region Variables init
$ExchangeAndSitesCollection = @()

#endregion

#region FUNCTIONS

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
#endregion

#region AD queries
cls
Title1 "Active Directory quick discovery"

LogGreen "Getting forest details"
$CurrentForest = [system.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest()
LogRed "Forest name $CurrentForest"
LogYellow "Forest details:"
$CurrentForest

LogGreen "Getting AD sites details"
$ADSitesDetails = [System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest().Sites
$CountADSites = ($ADSitesDetails).count
LogRed "We have $CountADSites sites..."
$ADSitesDetails | Ft Name

#endregion

#region Getting Exchange servers per site
Title1 "Getting Exchange servers for each site"

Foreach ($ADSite in $ADSitesDetails) {
    $CompleteSiteName = $CurrentForest.Name + "/Configuration/Sites/" + $($ADSite.Name)
    $ExchangeServers = Get-ExchangeServer | ? {$_.Site -like $CompleteSiteName}
    $ExchangeServersCount = ($ExchangeServers).count
    LogBlue "Site $ADSite"
    LogYellow "There are $ExchangeServersCount Exchange servers in $ADSite"

    $ExchAndSites = [PSCustomObject]@{
        Site = $ADSite
        ExchangeServers = $ExchangeServers

    }

    $ExchangeAndSitesCollection += $ExchAndSites
}

$ExchangeAndSitesCollection | out-host

#endregion

#region Getting number of mailboxes per AD site (datacenter)
Title1 "Getting number of mailboxes per AD site"

Foreach ($Item in $ExchangeAndSitesCollection){
    If (($Item.ExchangeServers).count -eq 0) {
        LogRed "No ExchangeServer in Site $($Item.Site) ... moving on to next site"
    } Else {
        $AllDatabases = @()
        LogBlue "Working on Site $($Item.Site) that has $(($item.ExchangeServers).count) exchange servers"
        $ExchangeServersForSite = $Item.ExchangeServers
        LogGreen "Getting all databases of all servers from site $($Item.Site)"
        Foreach ($Server in $ExchangeServersForSite) {
            $Databases = Get-MailboxDatabase -server $Server
            LogBlue "Found $($Databases.count) databases for server $Server ..."
            $AllDatabases += $Databases
        }
        LogRed "Total databases found for site $($Item.Site) : $($AllDatabases.count)"

        LogRed "Counting mailboxes for all these databases for $($Item.Site) ..."
        $MailboxesCount = 0
        $Counter = 0
        Foreach ($Database in $AllDatabases) {
            $Counter++
            Write-Progress -Activity "Browing databases" -Status "Browsing Database $($Database.Name) ..." -PercentComplete $($Counter/($AllDatabases.Count))
            $MailboxesCount += (Get-Mailbox -Database $Database | Select Name).count
        }

        LogBlue "There are $MailboxesCount mailboxes in site $($Item.Site)"

    }
}

#endregion
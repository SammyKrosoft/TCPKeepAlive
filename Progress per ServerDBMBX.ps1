cls
$E2016 = get-exchangeserver | ? {$_.AdminDisplayversion -like "*15.1*"} 

$ServerCounter = 0
Foreach ($Server in $E2016) {
$ServerCounter++
Write-Progress -id 1 -Activity "Servers" -Status "Server : $($Server.Name)" -PercentComplete $($ServerCounter/$($E2016.Count)*100)
    $Databases = $Server| get-mailboxdatabase
    $DatabaseCounter = 0
    Foreach ($DB in $Databases){
        $DatabaseCounter++
        Write-Progress -id 2 -ParentId 1 -Activity "Databases" -Status "Database : $($DB.Name)" -PercentComplete $($DatabaseCounter/$($Databases.Count)*100)
        $Mailboxes = $DB | Get-Mailbox -ResultSize Unlimited
        Write-Host "Number of mailboxes : $($Mailboxes.count)"
        $MAilboxCounter = 0
        Foreach($MBX in $Mailboxes){
            $MAilboxCounter++
            Write-Progress -id 3 -ParentId 2 -Activity "Mailboxes" -Status "Mailbox: $($MBX.DisplayName)" -PercentComplete $($MailboxCounter/$($Mailboxes.count)*100)
            $MBX | Get-MailboxStatistics | fl
        }
    }
}

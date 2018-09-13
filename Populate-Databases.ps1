cls

$AllDatabases = Get-mailboxdatabase

$Counter = 0
1..50 | Foreach {
    $Counter++
    $RandomDatabase = $AllDatabases[$(Get-Random -Maximum $($AllDatabases.count))]
    $MailboxSet = "NewGuy$_"
    Write-Progress -Activity "Creating mailboxes..." -Status "Processing mailbox $MailboxSet" -PercentComplete $($Counter/50*100)
    Write-Host "Creating $MAilboxSet on database $RandomDatabase"
    Net User $MailboxSet P@ssw0rd1 /ADD /Domain
    Enable-Mailbox $MailboxSet -Database $RandomDatabase

}
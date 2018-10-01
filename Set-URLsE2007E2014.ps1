"IRBEXCH3","NCEMIRB0037","NCEMIRB0038" | Resolve-DNSName -Type A
<########### OUTPUT ###########
Name Type TTL Section IPAddress 
---- ---- --- ------- --------- 
IRBEXCH3.IRB-CISR.IRBNET.GC.CA A 3600 Answer 159.177.130.42 
NCEMIRB0037.IRB-CISR.IRBNET.GC.CA A 1200 Answer 10.120.0.52 
NCEMIRB0038.IRB-CISR.IRBNET.GC.CA A 1200 Question 10.120.0.53 
#>

Resolve-DnsName mail.irb-cisr.gc.ca # Round Robin DNS to E2013 servers => Ok
<#
Name Type TTL Section IPAddress 
---- ---- --- ------- --------- 
mail.irb-cisr.gc.ca A 600 Answer 10.120.0.53 
mail.irb-cisr.gc.ca A 600 Answer 10.120.0.52 
#>
　
　
Resolve-DNSName legacy.irb-cisr.gc.ca #=> points to single E2007 server => Ok
<########### OUTPUT ###########
Source Destination IPV4Address IPV6Address Bytes Time(ms)
------ ----------- ----------- ----------- ----- --------
NCEMIRB0038 legacy.irb-cisr.gc.ca 159.177.130.42 32 0 
#>

cls;Resolve-DnsName autodiscover.irb-cisr.gc.ca # => round robin DNS to E2013 servers => Ok
<########### OUTPUT ###########
Name Type TTL Section IPAddress 
---- ---- --- ------- --------- 
autodiscover.irb-cisr.gc.ca A 600 Answer 10.120.0.52 
autodiscover.irb-cisr.gc.ca A 600 Answer 10.120.0.53
#>

#Plan:
#1 - update URLs on Exchange 2007 as follows:
#Set OWA, EWS, ActiveSync, OAB to : legacy.irb-cisr.gc.ca


#On Exchange 2007 do:
$Exchange2007Server = "IRBEXCH3"
$OWA2007 = Get-OwaVirtualDirectory -Server $Exchange2007Server | ? {$_.Name -like "*OWA*"}
$OWA2007 | Set-OwaVirtualDirectory -InternalUrl https://legacy.irb-cisr.gc.ca/OWA -ExternalURL https://legacy.irb-cisr.gc.ca/OWA

$EWS2007 = Get-WebServicesVirtualDirectory -Server $Exchange2007Server
$EWS2007 | Set-WebServicesVirtualDirectory -InternalUrl https://legacy.irb-cisr.gc.ca/OWA -ExternalURL https://legacy.irb-cisr.gc.ca/OWA

$EAS2007 = Get-ActiveSyncVirtualDirectory -Server $Exchange2007Server
$EAS2007 | Set-ActiveSyncVirtualDirectory -InternalUrl https://legacy.irb-cisr.gc.ca/OWA -ExternalURL https://legacy.irb-cisr.gc.ca/OWA

$OAB2007 = Get-OabVirtualDirectory -Server $Exchange2007Server
$OAB2007 | Set-OabVirtualDirectory -InternalUrl https://legacy.irb-cisr.gc.ca/OWA -ExternalURL https://legacy.irb-cisr.gc.ca/OWA

Set-ClientAccessServer $Exchange2007Server -AutoDiscoverServiceInternalUri https://autodiscover.irb-cisr.gc.ca/Autodiscover/Autodiscover.xml


#2 - update URLs on Exchange 2013 as follows:
Set OWA, ECP, EWS, ActiveSync, OAB AND Outlook Anywhere to : mail.irb-cisr.gc.ca

$E2013Servers = Get-ExchangeServer | ? {$_.AdminDisplayVersion -like "*15.0*"}

Foreach ($E2013 in $E2013Servers) {
    $OWA2013 = Get-OwaVirtualDirectory -Server $E2013 | ? {$_.Name -like "*OWA*"}
    $OWA2013 | Set-OwaVirtualDirectory -InternalUrl https://mail.irb-cisr.gc.ca/OWA -ExternalURL https://mail.irb-cisr.gc.ca/OWA

    $ECP2013 = Get-ECPVirtualDirectory -Server $E2013 | ? {$_.Name -like "*OWA*"}
    $ECP2013 | Set-ECPVirtualDirectory -InternalUrl https://mail.irb-cisr.gc.ca/OWA -ExternalURL https://mail.irb-cisr.gc.ca/OWA

    $EWS2013 = Get-WebServicesVirtualDirectory -Server $E2013
    $EWS2013 | Set-WebServicesVirtualDirectory -InternalUrl https://mail.irb-cisr.gc.ca/OWA -ExternalURL https://mail.irb-cisr.gc.ca/OWA

    $EAS2013 = Get-ActiveSyncVirtualDirectory -Server $E2013
    $EAS2013 | Set-ActiveSyncVirtualDirectory -InternalUrl https://mail.irb-cisr.gc.caOWA -ExternalURL https://mail.irb-cisr.gc.ca/OWA

    $OAB2013 = Get-OabVirtualDirectory -Server $E2013
    $OAB2013 | Set-OabVirtualDirectory -InternalUrl https://mail.irb-cisr.gc.ca/OWA -ExternalURL https://mail.irb-cisr.gc.ca/OWA

    Set-ClientAccessServer $E2013 -AutoDiscoverServiceInternalUri https://autodiscover.irb-cisr.gc.ca/Autodiscover/Autodiscover.xml
}

$E2013Servers | Get-OutlookAnywhere -ADPropertiesOnly | Set-Outlookanywhere -ExternalClientAuthenticationMethod Negotiate -ExternalClientsRequireSSL $true -ExternalHostName mail.irb-cisr.gc.ca -IISAuthenticationMethods 'Basic', 'Ntlm', 'Negotiate' -InternalClientAuthenticationMethod Negotiate -InternalClientsRequireSSL $true -InternalHostName mail.irb-cisr.gc.ca

$E2013Servers | Get-MapiVirtualDirectory -ADPropertiesOnly | Set-MapiVirtualDirectory -InternalUrl https://mail.irb-cisr.gc.ca/mapi -ExternalUrl https://mail.irb-cisr.gc.ca/mapi -IISAuthenticationMethods Ntlm, OAuth, Negotiate 



#3 - update ALL servers (2007 and 2013) AutodiscoverURI to : autodiscover.irb-cisr.gc.ca
　
Function HereStringToArray ($HereString) {
    Return $HereString -split "`n" | %{$_.trim()}
}

# BACKWARDS:

$CSV = @"
"ServerName","EAS-vDirNAme","EAS-InternalURL","EAS-ExternalURL","OAB-vDirNAme","OAB-InternalURL","OAB-ExernalURL","OWA-vDirNAme","OWA-InternalURL","OWA-ExernalURL","ECP-vDirNAme","ECP-InternalURL","ECP-ExernalURL","AutoDisc-vDirNAme","AutoDisc-URI","EWS-vDirNAme","EWS-InternalURL","EWS-ExernalURL","OutlookAnywhere-InternalHostName(NoneForE2010)","OutlookAnywhere-ExternalHostNAme(E2010+)"
"IRBEXCH3","Microsoft-Server-ActiveSync (Default Web Site)","https://irbexch3.irb-cisr.irbnet.gc.ca/Microsoft-Server-ActiveSync","https://irbsmtp.irb-cisr.gc.ca/Microsoft-Server-ActiveSync","OAB (Default Web Site)","http://irbexch3.irb-cisr.irbnet.gc.ca/OAB","http://irbsmtp.irb-cisr.gc.ca/OAB",,,,,,,"IRBEXCH3","https://irbexch3.irb-cisr.irbnet.gc.ca/Autodiscover/Autodiscover.xml","EWS (Default Web Site)","https://irbexch3.irb-cisr.irbnet.gc.ca/EWS/Exchange.asmx","https://irbexch3.irb-cisr.irbnet.gc.ca/ews/Exchange.asmx",,
"NCEMIRB0037","Microsoft-Server-ActiveSync (Default Web Site)","https://ncemirb0037.irb-cisr.irbnet.gc.ca/Microsoft-Server-ActiveSync",,"OAB (Default Web Site)","https://ncemirb0037.irb-cisr.irbnet.gc.ca/OAB",,,,,"ecp (Default Web Site)","https://ncemirb0037.irb-cisr.irbnet.gc.ca/ecp",,"NCEMIRB0037",,"EWS (Default Web Site)","https://ncemirb0037.irb-cisr.irbnet.gc.ca/EWS/Exchange.asmx",,"ncemirb0037.irb-cisr.irbnet.gc.ca",
"NCEMIRB0038","Microsoft-Server-ActiveSync (Default Web Site)","https://ncemirb0038.irb-cisr.irbnet.gc.ca/Microsoft-Server-ActiveSync",,"OAB (Default Web Site)","https://ncemirb0038.irb-cisr.irbnet.gc.ca/OAB",,,,,"ecp (Default Web Site)","https://ncemirb0038.irb-cisr.irbnet.gc.ca/ecp",,"NCEMIRB0038","https://ncemirb0038.irb-cisr.irbnet.gc.ca/autodiscover/autodiscover.xml","EWS (Default Web Site)","https://ncemirb0038.irb-cisr.irbnet.gc.ca/EWS/Exchange.asmx",,"ncemirb0038.irb-cisr.irbnet.gc.ca",
"@

$OldValues = HereStringToArray -HereString $CSV

$OldValues | ConvertTo-Csv -Delimiter ","

For ($i=0;$i -le $($OldValues.Count);$i++){
    Write-Host "Server: $($OldValues[$i].ServerName)" -b yellow -f blue
    $OldValues[$i]
}

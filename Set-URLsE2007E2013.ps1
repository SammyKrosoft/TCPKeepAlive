"IRBEXCH3","NCEMIRB0037","NCEMIRB0038" | Resolve-DNSName -Type A
<########### OUTPUT ###########
Name Type TTL Section IPAddress 
---- ---- --- ------- --------- 
IRBEXCH3.IRB-CISR.IRBNET.GC.CA A 3600 Answer 159.177.130.42 
NCEMIRB0037.IRB-CISR.IRBNET.GC.CA A 1200 Answer 10.120.0.52 
NCEMIRB0038.IRB-CISR.IRBNET.GC.CA A 1200 Question 10.120.0.53 
#>
　
Resolve-DnsName mail.irb-cisr.gc.ca
<#
Name Type TTL Section IPAddress 
---- ---- --- ------- --------- 
mail.irb-cisr.gc.ca A 600 Answer 10.120.0.53 
mail.irb-cisr.gc.ca A 600 Answer 10.120.0.52 
#>
　
　
Resolve-DNSName legacy.irb-cisr.gc.ca
<########### OUTPUT ###########
Source Destination IPV4Address IPV6Address Bytes Time(ms)
------ ----------- ----------- ----------- ----- --------
NCEMIRB0038 legacy.irb-cisr.gc.ca 159.177.130.42 32 0 
#>
cls;Resolve-DnsName autodiscover.irb-cisr.gc.ca
<########### OUTPUT ###########
Name Type TTL Section IPAddress 
---- ---- --- ------- --------- 
autodiscover.irb-cisr.gc.ca A 600 Answer 10.120.0.52 
autodiscover.irb-cisr.gc.ca A 600 Answer 10.120.0.53
#>
<# CONCLUSION 1 : 
Mail.irb points to Exchange 2007
Legacy.irb points to Exchange 2007
autodiscover.irb points to Exchange 2013
#>
　
　
cls;Get-ExchangeCertificate -Thumbprint 907B11C81002A41C8EAF0250BEB69A2AD1C5DA3B | select -ExpandProperty CertificateDomains|ft domain
<#
Domain 
------ 
irb.gc.ca 
irbsmtp.irb-cisr.gc.ca 
autodiscover.irb-cisr.gc.ca 
autodiscover.cisr-irb.gc.ca 
autodiscover.irb.gc.ca 
autodiscover.cisr.gc.ca 
mail.irb-cisr.gc.ca 
appmail.irb-cisr.irbnet.gc.ca 
ncemirb0037.irb-cisr.irbnet.gc.ca 
ncemirb0038.irb-cisr.irbnet.gc.ca 
ncemirb0039.irb-cisr.irbnet.gc.ca 
ncemirb0040.irb-cisr.irbnet.gc.ca 
activesync.irb-cisr.irbnet.gc.ca 
legacy.irb-cisr.gc.ca 
mail.cisr-irb.gc.ca 
Irbexch3.irb-cisr.irbnet.gc.ca 
legacy.irb-cisr.irbnet.gc.ca 
activesync.irb-cisr.gc.ca 
irb-cisr.gc.ca 
www.irb.gc.ca
#>
<#E2007 Names:
Domain
------
irbexch3.irb-cisr.irbnet.gc.ca
irbsmtp.irb-cisr.gc.ca
autodiscover.irb-cisr.gc.ca
autodiscover.cisr-irb.gc.ca
autodiscover.irb.gc.ca
autodiscover.cisr.gc.ca
IRBEXCH3
IRBEXCH3.IRB-CISR.IRBNET.GC.CA
#>
<# CONCLUSION 2 : 
Mail.irb points to Exchange 2007
Legacy.irb points to Exchange 2007
autodiscover.irb points to Exchange 2013
autodiscover.irb is in both Exchange 2007 and Exchange 2013 certificates
#>
　
　
Get-ExchangeCertificate -Server IRBEXCH3
　
cls;Get-ClientAccessServer | fl name,AutodiscoverServiceinternalURI
<#
Name : IRBEXCH3
AutoDiscoverServiceInternalUri : https://irbexch3.irb-cisr.irbnet.gc.ca/Autodiscover/Autodiscover.xml
Name : NCEMIRB0037
AutoDiscoverServiceInternalUri : 
Name : NCEMIRB0038
AutoDiscoverServiceInternalUri : https://ncemirb0038.irb-cisr.irbnet.gc.ca/autodiscover/autodiscover.xml
#>
　
cls; get-exchangeserver | fl name,*Role*, admindisplay*
<#
Name : irbexch1
ExchangeLegacyServerRole : 0
ServerRole : Mailbox
AdminDisplayVersion : Version 8.3 (Build 83.6)
Name : IRBEXCH3
ExchangeLegacyServerRole : 0
ServerRole : ClientAccess, HubTransport
AdminDisplayVersion : Version 8.3 (Build 83.6)
Name : NCEMIRB0037
ExchangeLegacyServerRole : 0
ServerRole : Mailbox, ClientAccess
AdminDisplayVersion : Version 15.0 (Build 1395.4)
Name : NCEMIRB0038
ExchangeLegacyServerRole : 0
ServerRole : Mailbox, ClientAccess
AdminDisplayVersion : Version 15.0 (Build 1395.4)
#>
cls
Write-host "EWS config:***************" -BackgroundColor yellow -ForegroundColor Red
Get-WebServicesVirtualDirectory -ADPropertiesOnly | fl Server, name,InternalURL, ExternalURL
Write-host "OWA config:***************" -BackgroundColor yellow -ForegroundColor Red
Get-OwaVirtualDirectory -ADPropertiesOnly | ? {$_.Name -like "*OWA*"} | fl Server, name, InternalURL, ExternalURL
Write-Host "ECP Config:***************" -BackgroundColor yellow -ForegroundColor Red
Get-ECPVirtualDirectory -ADPropertiesOnly | ? {$_.Name -like "*ECP*"} | fl Server, name, InternalURL, ExternalURL
Write-Host "OAB config:***************" -BackgroundColor yellow -ForegroundColor Red
Get-OABVirtualDirectory -ADPropertiesOnly | fl Server, name, InternalURL, ExternalURL
Write-Host "Activesync config:********" -BackgroundColor yellow -ForegroundColor Red
Get-ActiveSyncVirtualDirectory -ADPropertiesOnly | fl Server, name, InternalURL, ExternalURL
Write-Host "AUTODISCOVER:*************" -BackgroundColor yellow -ForegroundColor Red
Get-ClientAccessServer | fl Name,AutodiscoverServiceInternalURI


Exchange 2007 : Set your URLs to legacy.

<#
EWS config:***************
　
Server : IRBEXCH3
Name : EWS (Default Web Site)
InternalUrl : https://irbexch3.irb-cisr.irbnet.gc.ca/EWS/Exchange.asmx
ExternalUrl : https://irbexch3.irb-cisr.irbnet.gc.ca/ews/Exchange.asmx
Server : NCEMIRB0037
Name : EWS (Default Web Site)
InternalUrl : https://ncemirb0037.irb-cisr.irbnet.gc.ca/EWS/Exchange.asmx
ExternalUrl : 
Server : NCEMIRB0038
Name : EWS (Default Web Site)
InternalUrl : https://ncemirb0038.irb-cisr.irbnet.gc.ca/EWS/Exchange.asmx
ExternalUrl : 
　
　
OWA config:***************
　
Server : IRBEXCH3
Name : owa (Default Web Site)
InternalUrl : https://irbexch3.irb-cisr.irbnet.gc.ca/owa
ExternalUrl : https://irbsmtp.irb-cisr.gc.ca/owa
Server : NCEMIRB0037
Name : owa (Default Web Site)
InternalUrl : https://ncemirb0037.irb-cisr.irbnet.gc.ca/owa
ExternalUrl : 
Server : NCEMIRB0038
Name : owa (Default Web Site)
InternalUrl : https://ncemirb0038.irb-cisr.irbnet.gc.ca/owa
ExternalUrl : 
　
　
ECP Config:***************
　
Server : NCEMIRB0038
Name : ecp (Default Web Site)
InternalUrl : https://ncemirb0038.irb-cisr.irbnet.gc.ca/ecp
ExternalUrl : 
Server : NCEMIRB0037
Name : ecp (Default Web Site)
InternalUrl : https://ncemirb0037.irb-cisr.irbnet.gc.ca/ecp
ExternalUrl : 
　
　
OAB config:***************
　
Server : IRBEXCH3
Name : OAB (Default Web Site)
InternalUrl : http://irbexch3.irb-cisr.irbnet.gc.ca/OAB
ExternalUrl : http://irbsmtp.irb-cisr.gc.ca/OAB
Server : NCEMIRB0037
Name : OAB (Default Web Site)
InternalUrl : https://ncemirb0037.irb-cisr.irbnet.gc.ca/OAB
ExternalUrl : 
Server : NCEMIRB0038
Name : OAB (Default Web Site)
InternalUrl : https://ncemirb0038.irb-cisr.irbnet.gc.ca/OAB
ExternalUrl : 
　
　
Activesync config:********
　
Server : IRBEXCH3
Name : Microsoft-Server-ActiveSync (Default Web Site)
InternalUrl : https://irbexch3.irb-cisr.irbnet.gc.ca/Microsoft-Server-ActiveSync
ExternalUrl : https://irbsmtp.irb-cisr.gc.ca/Microsoft-Server-ActiveSync
Server : NCEMIRB0037
Name : Microsoft-Server-ActiveSync (Default Web Site)
InternalUrl : https://ncemirb0037.irb-cisr.irbnet.gc.ca/Microsoft-Server-ActiveSync
ExternalUrl : 
Server : NCEMIRB0038
Name : Microsoft-Server-ActiveSync (Default Web Site)
InternalUrl : https://ncemirb0038.irb-cisr.irbnet.gc.ca/Microsoft-Server-ActiveSync
ExternalUrl : 
　
　
AUTODISCOVER:*************
　
Name : IRBEXCH3
AutoDiscoverServiceInternalUri : https://irbexch3.irb-cisr.irbnet.gc.ca/Autodiscover/Autodiscover.xml
Name : NCEMIRB0037
AutoDiscoverServiceInternalUri : 
Name : NCEMIRB0038
AutoDiscoverServiceInternalUri : https://ncemirb0038.irb-cisr.irbnet.gc.ca/autodiscover/autodiscover.xml
#>
<#
Plan:
#1 - update URLs on Exchange 2007 as follows:
Set OWA, ECP, EWS, ActiveSync, OAB to : legacy.irb-cisr.gc.ca
　
#2 - update URLs on Exchange 2013 as follows:
Set OWA, ECP, EWS, ActiveSync, OAB AND Outlook Anywhere to : mail.irb-cisr.gc.ca
#3 - update ALL servers (2007 and 2013) AutodiscoverURI to : autodiscover.irb-cisr.gc.ca
　
　
#>
Get-OutlookAnywhere | ft servername, internalhostname, externalhostname
<############# OUTPUT
ServerName InternalHostname ExternalHostname 
---------- ---------------- ---------------- 
NCEMIRB0037 ncemirb0037.irb-cisr.irbnet.gc.ca 
NCEMIRB0038 ncemirb0038.irb-cisr.irbnet.gc.ca 
#>

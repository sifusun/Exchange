##Install Windows Roles and Features
Install-WindowsFeature AS-HTTP-Activation, Desktop-Experience, NET-Framework-45-Features, RPC-over-HTTP-proxy, RSAT-Clustering, RSAT-Clustering-CmdInterface, RSAT-Clustering-Mgmt, RSAT-Clustering-PowerShell, Web-Mgmt-Console, WAS-Process-Model, Web-Asp-Net45, Web-Basic-Auth, Web-Client-Auth, Web-Digest-Auth, Web-Dir-Browsing, Web-Dyn-Compression, Web-Http-Errors, Web-Http-Logging, Web-Http-Redirect, Web-Http-Tracing, Web-ISAPI-Ext, Web-ISAPI-Filter, Web-Lgcy-Mgmt-Console, Web-Metabase, Web-Mgmt-Console, Web-Mgmt-Service, Web-Net-Ext45, Web-Request-Monitor, Web-Server, Web-Stat-Compression, Web-Static-Content, Web-Windows-Auth, Web-WMI, Windows-Identity-Foundation, RSAT-ADDS

##Download and Install .NET Framework 4.8
https://download.visualstudio.microsoft.com/download/pr/2d6bb6b2-226a-4baa-bdec-798822606ff1/8494001c276a4b96804cde7829c04d7f/ndp48-x86-x64-allos-enu.exe

##Verify the .Net Framework version
Get-ChildItem 'HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP' -Recurse | Get-ItemProperty -Name Version,Release -EA 0 | Where { $_.PSChildName -match '^(?!S)\p{L}'} | Select PSChildName, Version, Release

##Download and Install Visual C++ Redistributable Package for Visual Studio 2012
https://download.microsoft.com/download/1/6/B/16B06F60-3B20-4FF2-B699-5E9B7962F9AE/VSU_4/vcredist_x64.exe

##Download and Install Visual C++ Redistributable Package for Visual Studio 2013
https://download.microsoft.com/download/2/E/6/2E61CFA4-993B-4DD4-91DA-3737CD5CD6E3/vcredist_x64.exe

##Download and install Microsoft Unified Communications Managed API 4.0, Core Runtime 64-bit
https://download.microsoft.com/download/2/C/4/2C47A5C1-A1F3-4843-B9FE-84C0032C61EC/UcmaRuntimeSetup.exe

###Prepare Active Directory
##Verify the Active Directory and Domain Controller information
Get-ADForest
Get-ADDomainController | Select Name,OperatingSystem

##Download the latest Cumulative Update for Exchange Server 2013
##Cumulative Update 23
https://download.microsoft.com/download/7/F/D/7FDCC96C-26C0-4D49-B5DB-5A8B36935903/Exchange2013-x64-cu23.exe

##Extend the Active Directory schema
.\Setup.exe /IAcceptExchangeServerLicenseTerms_DiagnosticDataON /PrepareSchema

##The First time Prepare the active directory
#.\Setup.exe /PrepareAD /OrganizationName:"<organization name>" /IAcceptExchangeServerLicenseTerms

##Prepare the active directory
.\Setup.exe /IAcceptExchangeServerLicenseTerms_DiagnosticDataON /PrepareAD

##Prepare active directory domains
.\Setup.exe /IAcceptExchangeServerLicenseTerms_DiagnosticDataON /PrepareAllDomains

##Install Exchange Server 2013 Mailbox and Client Access Server Roles
.\Setup.exe /IAcceptExchangeServerLicenseTerms_DiagnosticDataON /mode:Install /r:CA,MB

##Import the SSL certificate
#Import-ExchangeCertificate -Server <Exchange Server name> -FileName "<FilePath or UNC>\<file name>.pfx" -Password (Get-Credential).password
Import-ExchangeCertificate -Server EX01-2013 -FileName "\\EX01-2013\C$\Post-Install\certificate\wildcard_gooddealmart_ca.pfx" -Password (Get-Credential).password

##Enable exchange services with the certificate
#Enable-ExchangeCertificate -Thumbprint <Thumbprint> -Services <Service name>,<Service name>... [-Server < Exchange Server Name >]
Enable-ExchangeCertificate -Thumbprint 0F8621FD7020DE4DC46CA88A58E239B4EC2598CA -Services SMTP, IIS -Server Ex01-2013

##Verify that you have successfully assigned a certificate to the Exchange services.
Get-ExchangeCertificate | Format-List FriendlyName,Subject,CertificateDomains,Thumbprint,Services

##Check the Outlook Anywhere namespace from Exchange 2010
Get-OutlookAnywhere | Select Server,Internalhostname,Externalhostname | Fl

##Check OWA Virtual Directory namespace from Exchange 2010
Get-OWAVirtualDirectory -ADPropertiesOnly | Select Server,Internalurl,Externalurl | Fl

##Check ECP Virtual Directory namespace from Exchange 2010
Get-ECPVirtualDirectory -ADPropertiesOnly | Select Server,Internalurl,Externalurl | Fl

##check OAB Virtual Directory namespace from Exchange 2010
Get-OABVirtualDirectory -ADPropertiesOnly | Select Server,Internalurl,Externalurl | Fl

##Check WebServices Virtual Directory from Exchange 2010
Get-WebServicesVirtualDirectory -ADPropertiesOnly | Select Server,Internalurl,Externalurl | Fl

##Check ActiveSync Virtual Directory from Exchange 2010
Get-ActiveSyncVirtualDirectory -ADPropertiesOnly | Select Server,Internalurl,Externalurl | Fl

##Check ClientAccessServer Virtual Directory from Exchange 2010
Get-ClientAccessServer | Select Identity,AutodiscoverServiceInternaluri

##Check Client Access Array settings from Exchange 2010
Get-MailboxDatabase |fl RPCClientAccessServer

###Configuring the Autodiscover Services Connection Point (SCP) for Exchange 2013 Server
##Check the existing Autodiscover SCP settings from Exchange 2013
#Get-ClientAccessServer -Identity <Exchange Server Name> | Select Name,AutodiscoverServiceInternalUri | Fl
Get-ClientAccessServer -Identity "EX01-2013" | Select Name,AutodiscoverServiceInternalUri | Fl

##Update Autidiscover SCP
#Set-ClientAccessServer -Identity <Exchange Server Name> -AutoDiscoverServiceInternalUri  <Uri>
Set-ClientAccessServer -Identity "EX01-2013" -AutoDiscoverServiceInternalUri "https://mail.gooddealmart.ca/Autodiscover/Autodiscover.xml"

###Configuring the Client Access Namespaces for Exchange 2013 Server
#The External and Internal Name is likely behind a load balancer / VIP
#Make sure to configure that first
#These steps need to be performed on each server
#Adjust the DNS Name according and Server name Accordingly
$ExternalHostname = “mail.gooddealmart.ca”
$InternalHostname = “mail.gooddealmart.ca”
$Servername = “EX01-2013”
Get-OWAVirtualDirectory -Server $Servername | Set-OWAVirtualDirectory -ExternalUrl https://$ExternalHostname/owa -InternalUrl https://$InternalHostname/owa
Get-ECPVirtualDirectory -Server $Servername | Set-ECPVirtualDirectory -ExternalUrl https://$ExternalHostname/ecp -InternalUrl https://$InternalHostname/ecp
Get-ActiveSyncVirtualDirectory -Server $Servername | Set-ActiveSyncVirtualDirectory -ExternalUrl https://$ExternalHostname/Microsoft-Server-ActiveSync -InternalUrl https://$InternalHostname/Microsoft-Server-ActiveSync
Get-WebServicesVirtualDirectory -Server $Servername | Set-WebServicesVirtualDirectory -ExternalUrl https://$ExternalHostname/EWS/Exchange.asmx -InternalUrl https://$InternalHostname/EWS/Exchange.asmx
Get-OABVirtualDirectory -Server $Servername | Set-OABVirtualDirectory -ExternalUrl https://$ExternalHostname/OAB -InternalUrl https://$InternalHostname/OAB
Get-MapiVirtualDirectory -Server $Servername | Set-MapiVirtualDirectory -ExternalUrl https://$ExternalHostname/mapi -InternalUrl https://$InternalHostname/Mapi
Get-OutlookAnywhere -Server $Servername | Set-OutlookAnywhere -ExternalHostname $ExternalHostname -InternalHostname $InternalHostname -ExternalClientsRequireSsl $true -InternalClientsRequireSsl $false -DefaultAuthenticationMethod Basic,NTLM

###Creating a new Offline Address Book for Exchange 2013 Server
##Create a new Offline Address Book
New-OfflineAddressBook -Name "OAB2013" -AddressLists "\Default Global Address List"

##Set the Offline Address Book as default
Set-OfflineAddressBook -Identity "OAB2013" -IsDefault $true

##Check Offline Address Book Status
Get-OfflineAddressBook


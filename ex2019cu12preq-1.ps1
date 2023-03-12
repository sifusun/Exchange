﻿##Install Windows Roles and Features for Exchange 2019 Mailbox server
Install-WindowsFeature Server-Media-Foundation, NET-Framework-45-Features, RPC-over-HTTP-proxy, RSAT-Clustering, RSAT-Clustering-CmdInterface, RSAT-Clustering-Mgmt, RSAT-Clustering-PowerShell, WAS-Process-Model, Web-Asp-Net45, Web-Basic-Auth, Web-Client-Auth, Web-Digest-Auth, Web-Dir-Browsing, Web-Dyn-Compression, Web-Http-Errors, Web-Http-Logging, Web-Http-Redirect, Web-Http-Tracing, Web-ISAPI-Ext, Web-ISAPI-Filter, Web-Lgcy-Mgmt-Console, Web-Metabase, Web-Mgmt-Console, Web-Mgmt-Service, Web-Net-Ext45, Web-Request-Monitor, Web-Server, Web-Stat-Compression, Web-Static-Content, Web-Windows-Auth, Web-WMI, Windows-Identity-Foundation, RSAT-ADDS

##Download and Install .NET Framework 4.8
https://download.visualstudio.microsoft.com/download/pr/2d6bb6b2-226a-4baa-bdec-798822606ff1/8494001c276a4b96804cde7829c04d7f/ndp48-x86-x64-allos-enu.exe

##Verify the .Net Framework version
Get-ChildItem 'HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP' -Recurse | Get-ItemProperty -Name Version,Release -EA 0 | Where { $_.PSChildName -match '^(?!S)\p{L}'} | Select PSChildName, Version, Release

##Download and Install Visual C++ Redistributable Package for Visual Studio 2012
https://download.microsoft.com/download/1/6/B/16B06F60-3B20-4FF2-B699-5E9B7962F9AE/VSU_4/vcredist_x64.exe

##Download and Install Visual C++ Redistributable Package for Visual Studio 2013
https://download.microsoft.com/download/2/E/6/2E61CFA4-993B-4DD4-91DA-3737CD5CD6E3/vcredist_x64.exe

##Download and Install  IIS URL Rewrite Module
https://download.microsoft.com/download/1/2/8/128E2E22-C1B9-44A4-BE2A-5859ED1D4592/rewrite_amd64_en-US.msi

##Download and install Microsoft Unified Communications Managed API 4.0, Core Runtime 64-bit
https://download.microsoft.com/download/2/C/4/2C47A5C1-A1F3-4843-B9FE-84C0032C61EC/UcmaRuntimeSetup.exe

###Prepare Active Directory
##Verify the Active Directory and Domain Controller information
Get-ADForest
Get-ADDomainController | Select Name,OperatingSystem

##Download the latest Cumulative Update for Exchange Server 2019
##Cumulative Update 23
https://download.microsoft.com/download/b/c/7/bc766694-8398-4258-8e1e-ce4ddb9b3f7d/ExchangeServer2019-x64-CU12.ISO

##Mount the Exchange Server 2019 iso image file
##Extend the Active Directory schema
.\Setup.exe /IAcceptExchangeServerLicenseTerms_DiagnosticDataON /PrepareSchema

##The First time Prepare the active directory
#.\Setup.exe /IAcceptExchangeServerLicenseTerms_DiagnosticDataON /PrepareAD /OrganizationName:"<organization name>"
##Prepare the active directory
.\Setup.exe /IAcceptExchangeServerLicenseTerms_DiagnosticDataON /PrepareAD

##Prepare active directory domains
.\Setup.exe /IAcceptExchangeServerLicenseTerms_DiagnosticDataON /PrepareAllDomains

##Install Exchange Server 2019 Mailbox and Client Access Server Roles
.\Setup.exe /IAcceptExchangeServerLicenseTerms_DiagnosticDataON /mode:Install /r:MB

###Install SSL Certificate
##Check the Thumbprint of the certificate
Get-ExchangeCertificate | Select CertificateDomains, Thumbprint

##Export the certificate
#$bincert = Export-ExchangeCertificate -Thumbprint <Thumbprint> -BinaryEncoded -Password (Get-Credential).password
$bincert = Export-ExchangeCertificate -Thumbprint 0F8621FD7020DE4DC46CA88A58E239B4EC2598CA -BinaryEncoded -Password (Get-Credential).password

##Save the certificate file
#[System.IO.File]::WriteAllBytes('<path>\<file name>.pfx', $bincert.FileData)
[System.IO.File]::WriteAllBytes('C:\Post-Install\certificate\wildcard_gooddealmart_cat.pfx', $bincert.FileData)


##Import the SSL certificate and save the Thumbprint 
#Import-ExchangeCertificate -Server <server name> -FileData ([System.IO.File]::ReadAllBytes('<path>\<certificate file name>')) -Password (Get-Credential).password
Import-ExchangeCertificate -Server EX01-2019 -FileData ([System.IO.File]::ReadAllBytes('C:\Post-Install\certificate\wildcard_gooddealmart_ca.pfx')) -Password (Get-Credential).password

##Enable exchange services with the certificate
#Enable-ExchangeCertificate -Thumbprint <Thumbprint> -Services <Service name>,<Service name>... [-Server < Exchange Server Name >]
Enable-ExchangeCertificate -Thumbprint 0F8621FD7020DE4DC46CA88A58E239B4EC2598CA -Services SMTP, IIS -Server Ex01-2019

##Verify that you have successfully assigned a certificate to the Exchange services.
Get-ExchangeCertificate | Format-List FriendlyName,Subject,CertificateDomains,Thumbprint,Services

####Configure Exchange Server 2019 Coexistence Exchange Server 2016
###Configuring the Autodiscover Services Connection Point (SCP) for Exchange 2019 Server
##Check the existing Autodiscover SCP settings
#Get-ClientAccessService -Identity <Exchange Server Name> | Select Name,AutodiscoverServiceInternalUri | Fl
Get-ClientAccessService -Identity "EX01-2019" | Select Name,AutodiscoverServiceInternalUri | Fl

##Update Autidiscover SCP
#Set-ClientAccessService -Identity <Exchange Server Name> -AutoDiscoverServiceInternalUri <Uri>
Set-ClientAccessService -Identity "EX01-2019" -AutoDiscoverServiceInternalUri "https://mail.gooddealmart.ca/Autodiscover/Autodiscover.xml"

###Configuring the Client Access Namespaces for Exchange 2019 Server
#The External and Internal Name is likely behind a load balancer / VIP
#Make sure to configure that first
#These steps need to be performed on each server
#Adjust the DNS Name according and Server name Accordingly
$ExternalHostname = “mail.gooddealmart.ca”
$InternalHostname = “mail.gooddealmart.ca”
$Servername = “EX01-2019”
Get-OWAVirtualDirectory -Server $Servername | Set-OWAVirtualDirectory -ExternalUrl https://$ExternalHostname/owa -InternalUrl https://$InternalHostname/owa
Get-ECPVirtualDirectory -Server $Servername | Set-ECPVirtualDirectory -ExternalUrl https://$ExternalHostname/ecp -InternalUrl https://$InternalHostname/ecp
Get-ActiveSyncVirtualDirectory -Server $Servername | Set-ActiveSyncVirtualDirectory -ExternalUrl https://$ExternalHostname/Microsoft-Server-ActiveSync -InternalUrl https://$InternalHostname/Microsoft-Server-ActiveSync
Get-WebServicesVirtualDirectory -Server $Servername | Set-WebServicesVirtualDirectory -ExternalUrl https://$ExternalHostname/EWS/Exchange.asmx -InternalUrl https://$InternalHostname/EWS/Exchange.asmx
Get-OABVirtualDirectory -Server $Servername | Set-OABVirtualDirectory -ExternalUrl https://$ExternalHostname/OAB -InternalUrl https://$InternalHostname/OAB
Get-MapiVirtualDirectory -Server $Servername | Set-MapiVirtualDirectory -ExternalUrl https://$ExternalHostname/mapi -InternalUrl https://$InternalHostname/Mapi
Get-OutlookAnywhere -Server $Servername | Set-OutlookAnywhere -ExternalHostname $ExternalHostname -InternalHostname $InternalHostname -ExternalClientsRequireSsl $true -InternalClientsRequireSsl $false -DefaultAuthenticationMethod NTLM,NTLM

###Creating a new Offline Address Book for Exchange 2019 Server
##Create a new Offline Address Book
New-OfflineAddressBook -Name "OAB2019" -AddressLists "\Default Global Address List"

##Set the Offline Address Book as default
Set-OfflineAddressBook -Identity "OAB2019" -IsDefault $true

##Check Offline Address Book Status
Get-OfflineAddressBook

###Migrate Mailboxes to Exchange 2019
###Configuring the Default Mailbox Database for Exchange Server 2019
##Get all mailbox database information
Get-MailboxDatabase -IncludePreExchange2013

##Change the Default Mailbox Database name
#Set-MailboxDatabase “<Default Mailbox Database Name>” -Name "<New box Database Name>"
Set-MailboxDatabase “Mailbox Database 1909041848” -Name DB01-2019

##Check the Mailbox Database path
#Get-MailboxDatabase <Mail Database Nmae> | Fl *path*
Get-MailboxDatabase DB01-2019 | Fl *path*

##Move the Mailbox Database to the right path
#Move-DatabasePath -Identit <Mailbox Database Name> -EdbFilePath <EdbFilePath> -LogFolderPath <NonRootLocalLongFullPath>  
Move-DatabasePath -Identity DB01-2019 -EdbFilePath D:\DB01-2019\DB01-2019_DB\DB01-2019.edb -LogFolderPath D:\DB01-2019\DB01-2019_LOGS

##Create a Mailbox Database for Exchange 2016
#New-MailboxDatabase -Name <Mailbox Database Name> -Server <Exchange Server Name> -EdbFilePath <EdbFilePath> -LogFolderPath <NonRootLocalLongFullPath>
New-MailboxDatabase -Name DB02-2019 -Server EX01-2019 -EdbFilePath D:\DB02-2019\DB02-2019_DB\DB02-2019.edb -LogFolderPath D:\DB02-2019\DB02-2019_LOGS

#Mount Mailbox Database
#Mount-Database -Identity <Mailbox Database Name>
Mount-Database -Identity DB02-2019

###Configuring Mailbox Database Quota
##Check Mailbox Database Quota settings
Get-MailboxDatabase -IncludePreExchange2013 | Select Name,IssueWarningQuota,ProhibitSendQuota,ProhibitSendReceiveQuota

##Change Mailbox Database Quota settings
#Get-MailboxDatabase -Server <Exchange Server Name> | Set-MailboxDatabase –IssueWarningQuota <Value> –ProhibitSendQuota <Value> -ProhibitSendReceiveQuota <Value>
Get-MailboxDatabase -Server Ex01-2019 | Set-MailboxDatabase –IssueWarningQuota 15GB –ProhibitSendQuota 16GB -ProhibitSendReceiveQuota 20GB
Get-MailboxDatabase -IncludePreExchange2013 | Select Name,IssueWarningQuota,ProhibitSendQuota,ProhibitSendReceiveQuota

###Configuring Offline Address Book
##check Mailbox Database Offline Address Book settings
Get-MailboxDatabase -IncludePreExchange2013 | Select Name,offline*

#Create Mailbox Database Offline Address Book
#New-OfflineAddressBook -Name <String> -AddressLists "Default Global Address List"  -GlobalWebDistributionEnabled $true
New-OfflineAddressBook -Name "OAB2019" -AddressLists "Default Global Address List"  -GlobalWebDistributionEnabled $true

##Check OAB Status
Get-OfflineAddressBook

##Change Mailbox Database Offline Address Book settings
Get-MailboxDatabase -Server EX01-2019 | Set-MailboxDatabase -OfflineAddressBook “OAB2019”
Get-MailboxDatabase -IncludePreExchange2013 | Select Name,offline*

###Migrate Arbitration Mailboxes
##Check Arbitration Mailboxes
Get-Mailbox -Arbitration | Select Name,Database

##Move Arbitration Mailboxes
Get-Mailbox -Arbitration |New-MoveRequest -TargetDatabase DB01-2019

##Check Arbitration Mailboxes move Statistics
Get-MoveRequest | Get-MoveRequestStatistics

##Check Arbitration Mailboxes
Get-Mailbox -Arbitration | Select Name,Database

##Remove completed move requests
Get-MoveRequest -MoveStatus Completed | Remove-MoveRequest

###Migrate Audit Log Mailbox
##Check the audit log mailbox
Set-ADServerSettings -ViewEntireForest $true
Get-Mailbox -AuditLog | Format-Table Name, ServerName, Database, AdminDisplayVersion

##Move the audit log mailbox
Get-Mailbox -AuditLog | New-MoveRequest -TargetDatabase DB01-2019

##Check the move Statistics
Get-MoveRequest | Get-MoveRequestStatistics

##Check the audit log mailbox
Get-Mailbox -AuditLog | Format-Table Name, ServerName, Database, AdminDisplayVersion

##Remove completed move requests
Get-MoveRequest -MoveStatus Completed | Remove-MoveRequest

###Migrate Discovery Search Mailbox
##Check the discovery search mailbox
Get-Mailbox -RecipientTypeDetails DiscoveryMailbox | Format-Table Name, Database

##Move the discovery search mailbox
Get-Mailbox -RecipientTypeDetails DiscoveryMailbox | New-MoveRequest -TargetDatabase DB01-2019

##Check the move Statistics
Get-MoveRequest | Get-MoveRequestStatistics

##Check the discovery search mailbox
Get-Mailbox -RecipientTypeDetails DiscoveryMailbox | Format-Table Name, Database

##Remove completed move requests
Get-MoveRequest -MoveStatus Completed | Remove-MoveRequest

###Create Database Availability Group
##Create a new DAG
New-DatabaseAvailabilityGroup -Name DAG-2019 -WitnessServer DC01-2022.gooddealmart.ca -WitnessDirectory C:\DAG-2019_Witness -FileSystem ReFS

##add Exchange servers to the DAG
Add-DatabaseAvailabilityGroupServer -Identity DAG-2019 -MailboxServer EX01-2019
Add-DatabaseAvailabilityGroupServer -Identity DAG-2019 -MailboxServer EX02-2019

##Adding Mailbox database copies
Add-MailboxDatabaseCopy -Identity DB01-2019 -MailboxServer EX02-2019 -ActivationPreference 2

###Migrate the public folders from Exchange 2010 to 2019
##Ensure the arbitration mailbox migrated
Get-Mailbox -Arbitration -Identity Migration.*

##Download the Migration Scripts from Microsoft Website
https://download.microsoft.com/download/0/1/E/01EF0592-9338-4D22-827B-46834BF70581/Create-PublicFolderMailboxesForMigration.ps1
https://download.microsoft.com/download/0/1/E/01EF0592-9338-4D22-827B-46834BF70581/PublicFolderToMailboxMapGenerator.strings.psd1
https://download.microsoft.com/download/0/1/E/01EF0592-9338-4D22-827B-46834BF70581/Export-PublicFolderStatistics.ps1
https://download.microsoft.com/download/0/1/E/01EF0592-9338-4D22-827B-46834BF70581/Export-PublicFolderStatistics.strings.psd1
https://download.microsoft.com/download/0/1/E/01EF0592-9338-4D22-827B-46834BF70581/PublicFolderToMailboxMapGenerator.ps1
https://download.microsoft.com/download/0/1/E/01EF0592-9338-4D22-827B-46834BF70581/CreatePublicFolderMailboxesForMigration.strings.psd1

##Preparation on the Exchange 2010 server
##Take a snapshot of source public folder structure
Get-PublicFolder -Recurse | Export-CliXML C:\Post-Install\PFMigration\Legacy_PFStructure.xml

##Take a snapshot of all public folder statistics
Get-PublicFolderStatistics | Export-CliXML C:\Post-Install\PFMigration\Legacy_PFStatistics.xml

##Take a snapshot of the permissions
Get-PublicFolder -Recurse | Get-PublicFolderClientPermission | Select-Object Identity,User -ExpandProperty AccessRights | Export-CliXML C:\Post-Install\PFMigration\Legacy_PFPerms.xml

##Renaming any public folders with a backslash in the name
Get-PublicFolderStatistics -ResultSize Unlimited | Where {($_.Name -like "*\*") -or ($_.Name -like "*/*") } | Format-List Name, Identity

##Rename any public folders that are returned, change "/" and "\" to '|"
Set-PublicFolder -Identity <public folder> -Name <new public folder name>

##Check the public folder migration status
Get-OrganizationConfig | Format-List PublicFoldersLockedforMigration, PublicFolderMigrationComplete

##Ensure the status of the PublicFoldersLockedforMigration or PublicFolderMigrationComplete properties is $false
Set-OrganizationConfig -PublicFoldersLockedforMigration:$false -PublicFolderMigrationComplete:$false

###Preparation on the Exchange 2019 server
##Discover any existing serial migration requests
Get-PublicFolderMigrationRequest | Get-PublicFolderMigrationRequestStatistics -IncludeReport | Format-List

##Remove any existing public folder serial migration requests.
Get-PublicFolderMigrationRequest | Remove-PublicFolderMigrationRequest

##Discover any existing batch migration requests
Get-MigrationBatch | where {$_.MigrationType.ToString() -eq "PublicFolder"}

##Remove any existing batch migration requests
$batch = Get-MigrationBatch | where {$_.MigrationType.ToString() -eq "PublicFolder"}
$batch | Remove-MigrationBatch -Confirm:$false

##Ensure no public folder mailboxes exist on the Exchange 2019 server
Get-Mailbox -PublicFolder

##To see if any public folders exist
Get-PublicFolder

##Remove Public Folders
Get-Mailbox -PublicFolder | Where{$_.IsRootPublicFolderMailbox -eq $false} | Remove-Mailbox -PublicFolder -Force -Confirm:$false
Get-Mailbox -PublicFolder | Remove-Mailbox -PublicFolder -Force -Confirm:$false

###Create .csv files and Migrate public folders
##Create the folder name-to-folder size mapping CSV file
.\Export-PublicFolderStatistics.ps1 \\MBX01-2010\PFMigration\PFSizeMap.csv EX01-2010.gooddealmart.ca

##Create the public folder-to-mailbox mapping file
.\PublicFolderToMailboxMapGenerator.ps1 20000000000 \\MBX01-2010\PFMigration\PFSizeMap.csv \\MBX01-2010\PFMigration\PFMailboxMap.csv

##Create the target public folder mailboxes at Exchange Server 2019
.\Create-PublicFolderMailboxesForMigration.ps1 -FolderMappingCsv PFMailboxMap.csv -EstimatedNumberOfConcurrentUsers:200

##create the migration batch
New-MigrationBatch -Name PFMigration -SourcePublicFolderDatabase (Get-PublicFolderDatabase -Server EX01-2016) -CSVData (Get-Content C:\Post-Install\PFMigration\PFMailboxMAP.csv -Encoding Byte) -NotificationEmails csun@gooddealmart.ca

##Start the migration
Start-MigrationBatch PFMigration

##Lock the legacy public folders for finalization at Exchange server 2010
Set-OrganizationConfig -PublicFoldersLockedForMigration:$true

#Change the Exchange 2016 deployment type to Remote at Exchange server 2019
Set-OrganizationConfig -PublicFoldersEnabled Remote

##Complete the public folder migration
Complete-MigrationBatch PFMigration

##Assign some pilot mailboxes to use the migrated public folder mailbox as the default public folder mailbox
Set-Mailbox -Identity GDMUSER2 -DefaultPublicFolderMailbox PFMailbox1

##Unlock the public folders
Get-Mailbox -PublicFolder | Set-Mailbox -PublicFolder -IsExcludedFromServingHierarchy $false

##Indicate that the public folder migration is complete at Exchange 2010 server
Set-OrganizationConfig -PublicFolderMigrationComplete:$true

##Enable the public folder locally at Exchange 2019 Server
Set-OrganizationConfig -PublicFoldersEnabled Local

##Compare files contents with previous files and verify success
Get-PublicFolder -Recurse | Export-CliXML C:\Post-Install\PFMigration\EX2019_PFStructure.xml
Get-PublicFolderStatistics -ResultSize Unlimited | Export-CliXML C:\Post-Install\PFMigration\Ex2019_PFStatistics.xml
Get-PublicFolder -Recurse | Get-PublicFolderClientPermission | Select-Object Identity,User -ExpandProperty AccessRights | Export-CliXML  C:\Post-Install\PFMigration\Ex2019_PFPerms.xml

####Decommission Exchange 2016 servers
###Uninstall Exchange 2016 Mailbox role
##Ensure all mailboxes are moved
Get-MailboxDatabase -Server <Exchange 2016 server name> | Get-Mailbox | select Name,Database

##Remove the mailbox database
Remove-MailboxDatabase -Identity <mailbox database name>

##Remove the mailbox database copy
Remove-MailboxDatabaseCopy -Identity <mailbox database name>\<server name>

##Remove the server from the database availability group
Remove-DatabaseAvailabilityGroupServer -Identity <DAG name> -MailboxServer <server name>

##Remove the database availability group
Remove-DatabaseAvailabilityGroup -Identity <DAG name>

##Remove public folders
Get-PublicFolder -Server <server name> "\NON_IPM_SUBTREE" -Recurse -ResultSize:Unlimited | Remove-PublicFolder -Server <server name> -Recurse -ErrorAction:SilentlyContinue
Get-PublicFolder -Server <server name> "\" -Recurse -ResultSize:Unlimited | Remove-PublicFolder -Server <server name> -Recurse -ErrorAction:SilentlyContinue

##Move all replicas to another server
.\MoveAllReplicas.ps1 -Server <source server name> -NewServer <destination server name>

##Remove the public folder database
Remove-PublicFolderDatabase -Identity <public folder database name>

##Get the offline address book
Get-OfflineAddressBook

##Remove the Exchange 2010 offline address book
Remove-OfflineAddressBook -identity <offline address book name>
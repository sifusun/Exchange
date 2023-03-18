################################################################
#                                                              #       
#      Hot Patches for Exchange Servers with DAG               #
#                By Cary Sun - MVP                             #
#Follow along with the step by step guide for the installation #
#                                                              #
################################################################

##Disable the Veeam Backup and Replica Jobs
##Put the server in maintenance mode in the monitoring systems (SCOM, VeeamONE etc.)
##Log in to Exchange Server1
##Download the Exchange 2016 Cumulative Update 23 for the Exchange server
https://download.microsoft.com/download/8/d/2/8d2d01b4-5bbb-4726-87da-0e331bc2b76f/ExchangeServer2016-x64-CU23.ISO

##Download the Microsoft .NET Framework 4.8 version as CU requirements
https://download.visualstudio.microsoft.com/download/pr/2d6bb6b2-226a-4baa-bdec-798822606ff1/8494001c276a4b96804cde7829c04d7f/ndp48-x86-x64-allos-enu.exe

##Run PowerShell as Administrator and add Exchange Snapin to PowerShell
Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn

###Put the Exchange Server1 in maintenance mode, drain the Hub Transport Service. It won't accept any message
#Set-ServerComponentState -Identity "<Exchange Server1>" -Component HubTransport -State Draining -Requester Maintenance
Set-ServerComponentState -Identity "EX01" -Component HubTransport -State Draining -Requester Maintenance

##Redirect the queued messages to the Exchange Server2.
#Redirect-Message -Server "<Exchange Server1>" -Target "<Exchange Server2 FQDN>" -Confirm
Redirect-Message -Server "EX01" -Target "EX02.gooddealmart.ca" -Confirm

##Run the below command to pause the cluster node and suspend Exchange Server1 as a Cluseter Node.
#Suspend-ClusterNode "<Exchange Server1>"
Suspend-ClusterNode "EX01"

##Run the below command to disable database copy automatic activation and also move active database copies to another DAG members.
##It will take several minutes for the moving active database copies
#Set-MailboxServer "<Exchange Server1>" -DatabaseCopyActivationDisabledAndMoveNow $true
Set-MailboxServer "EX01" -DatabaseCopyActivationDisabledAndMoveNow $true

##Make a note of the database copy automatic activation policy on the Exchange Server1. 
##You need to set it back to after the end of maintenance. The default setting is Unrestricted.
#Get-MailboxServer "<Exchange Server1>" | Select DatabaseCopyAutoActivationPolicy
Get-MailboxServer "EX01" | Select DatabaseCopyAutoActivationPolicy

##Run the below command to set the Exchange Server1 to Blocked and prevent any of the databases become Active.
#Set-MailboxServer "<Exchange Server1>" -DatabaseCopyAutoActivationPolicy Blocked
Set-MailboxServer "EX01" -DatabaseCopyAutoActivationPolicy Blocked

##Run the below to check the database copies status to ensure they don't mount on the Exchange Server1.This command should return no results.
##It may take a while for the Active databases to move. If any database copies are still active on the server and other DAG members host copies of the database. Perform a manual switchover.
##Get-MailboxDatabaseCopyStatus -Server "<Exchange Server1>" | Where {$_.Status -eq "Mounted"}
##Get-MailboxDatabaseCopyStatus -Server "<Exchange Server1>" | Where {$_.Status -eq "Mounted"} | ft -AutoSiz
Get-MailboxDatabaseCopyStatus -Server "EX01" | Where {$_.Status -eq "Mounted"}
Get-MailboxDatabaseCopyStatus -Server "EX01" | Where {$_.Status -eq "Mounted"} | ft -AutoSiz

##Run the below command to check the transport queue. Queues should be empty.
##Any emails still pending in the queues will delay delivery until the Exchange Server1 takes out from maintenance mode.
Get-Queue

##Put the Exchange Server1 into maintenance mode
#Set-ServerComponentState "<Exchange Server1>" -Component ServerWideOffline -State Inactive -Requester Maintenance
Set-ServerComponentState "EX01" -Component ServerWideOffline -State Inactive -Requester Maintenance

##Check the load balancer (or firewall) and ensure the health checks take the Exchange Server1 out of the pool or marked it as offline/inactive. 
##Typically there would be SMTP and HTTPS virtual services. This will force any future connections to Exchange Server1.
##Run the below command to ensure all components show Inactive except for Monitoring and RecoveryActionsEnabled.
Get-ServerComponentState "EX01" | Select Component, State

#Rebooting the Exchange Server1

##Installing the Exchange Cumulative Update.
##out the CU ISO image file
#Run the below command to prepare AD Schema
.\Setup.exe /IAcceptExchangeServerLicenseTerms_DiagnosticDataON /PrepareSchema

#Run the below commnad to prepare AD
.\Setup.exe /IAcceptExchangeServerLicenseTerms_DiagnosticDataON /PrepareAD

#Run the below command to prepare Domains
.\Setup.exe /IAcceptExchangeServerLicenseTerms_DiagnosticDataON /PrepareAllDomains

#Run the below command to install Exchange Cumulative Update
#.\Setup.exe /IAcceptExchangeServerLicenseTerms_DiagnosticDataON /Mode:Upgrade /DomainController:<DC server FQDN>
.\Setup.exe /IAcceptExchangeServerLicenseTerms_DiagnosticDataON /Mode:Upgrade /DomainController:dc01.gooddealmart.ca

#Reboot Exchange Server1

#Run Windows Update to install the Windows patches and least security update for the Exchange Cumulative Update

##Take Exchange Server1 out of maintenance mode
#Set-ServerComponentState "<Exchange Server1>" -Component ServerWideOffline -State Active -Requester Maintenance
Set-ServerComponentState "EX01" -Component ServerWideOffline -State Active -Requester Maintenance

#Resume-ClusterNode -Name "<Exchange Server1>"
Resume-ClusterNode -Name "EX01"

#Set-MailboxServer "<Exchange Server1>" -DatabaseCopyAutoActivationPolicy Unrestricted
Set-MailboxServer "EX01" -DatabaseCopyAutoActivationPolicy Unrestricted

#Set-MailboxServer "<Exchange Server1>" -DatabaseCopyActivationDisabledAndMoveNow $false
Set-MailboxServer "EX01" -DatabaseCopyActivationDisabledAndMoveNow $false

#Set-ServerComponentState "<Exchange Server1>" -Component HubTransport -State Active -Requester Maintenance
Set-ServerComponentState "EX01" -Component HubTransport -State Active -Requester Maintenance

##Rebalance Database Availability Groups
cd $exscripts
#.\RedistributeActiveDatabases.ps1 -DagName "<DAG name>" -BalanceDbsByActivationPreference -SkipMoveSuppressionChecks
.\RedistributeActiveDatabases.ps1 -DagName "PRODDAG01-2016" -BalanceDbsByActivationPreference -SkipMoveSuppressionChecks

##Verify if the Exchange Server1 is back up and running
#Get-ClusterNode "<Exchange Server1>"
Get-ClusterNode "EX01"

##Check that the cluster node has the state up on all the Exchange Servers
#Test-ServiceHealth "<Exchange Server1>"
Test-ServiceHealth "EX01"

##Check that the required services are running on all the Exchange Servers
Get-ExchangeServer | Test-ServiceHealth

##Test the MAPI Connectivity
#Test-MAPIConnectivity -Server "<Exchange Server1>"
Test-MAPIConnectivity -Server "EX01"

##Test the MAPI Connectivity on all the Exchange Servers
Get-ExchangeServer | Test-MAPIConnectivity

##Get the result of the DAG Copy Status Health
#Get-MailboxDatabaseCopyStatus -Server "<Exchange Server1>" | Sort Name | Select Name, Status, Contentindexstate
Get-MailboxDatabaseCopyStatus -Server "EX01" | Sort Name | Select Name, Status, Contentindexstate

##Get the result of the DAG Copy Status Health on all the Exchange Servers
Get-MailboxDatabaseCopyStatus * | Sort Name | Select Name, Status, Contentindexstate

##Check the Replication Health
#Test-ReplicationHealth -Server "<Exchange Server1>"
Test-ReplicationHealth -Server "EX01"

##Check the Replication Health on all the Exchange Servers
Get-DatabaseAvailabilityGroup | Select -ExpandProperty:Servers | Test-ReplicationHealth | Sort Name

##Verify the Database Activation Policy is set to Unrestricted
#Get-MailboxServer "<Exchange Server1>" | Select Name, DatabaseCopyAutoActivationPolicy
Get-MailboxServer "EX01" | Select Name, DatabaseCopyAutoActivationPolicy

##Verify the Database Activation Policy is set to Unrestricted on all the Exchange Servers
Get-MailboxServer | Select Name, DatabaseCopyAutoActivationPolicy

##Verify that the load balancer health checks have taken the server in the pool or marked it as online/active. 
##Typically there would be SMTP and HTTPS virtual services. This will enable connections to Exchange Server1.

##Repeat steps for all exchange servers except prepare schema, prepare AD and prepare all domains.

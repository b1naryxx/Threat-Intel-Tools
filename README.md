# Threat-Intel-Tools
Collection of tools to gather or process information for threat intelligence activities which I have developed so far. Below is a brief description and usage of each tool.

## Quarantine-Download.ps1
This is a powershell tool to download quarantined emails from Office365 using Get-QuarantineMessage and Export-QuarantineMessage cmdlets from ExchangeOnlineManagement module, and save them in the specified folder for further analysis. 

**NOTE**: To install ExchangeOnlineManagmente module use the following command:
```
Install-Module -Name ExchangeOnlineManagement
```
Then create a remote session using:
```
Connect-ExchangeOnline -UserPrincipalName user@tenant.com -ShowProgress $true
```

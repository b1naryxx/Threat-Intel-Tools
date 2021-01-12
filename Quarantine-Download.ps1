######################################################################################################
#                                                                                                    #
# Name:        Quarantine-Download.ps1                                                               #
#                                                                                                    #
# Version:     1.0                                                                                   #
#                                                                                                    #
# Description: Searches the previous X days for quarantined messages date range and download the eml #
#                                                                                                    #
# Limitations: Search query is limited to 1,000,000 entries.                                         #
#                                                                                                    #
# Requires:    Remote PowerShell Connection to Exchange Online                                       #
#                                                                                                    #
# Author:      Tomas Suescun                                                                         #
#                                                                                                    #
# Usage:       .\Quarantine-Download.ps1 -Days 30 -OutputDir C:\QuarantineFiles\                     #
#                                                                                                    #
#                                                                                                    #
# Disclaimer:  This script is provided AS IS without any support. Please test in a lab environment   #
#              prior to production use.                                                              #
#                                                                                                    #
######################################################################################################

#Use this script after creating a new Exchange Online Management session:
#    Import-Module ExchangeOnlineManagement
#    Connect-ExchangeOnline -UserPrincipalName username@tenant.com -ShowProgress $true

# To-do: - Add asyncronous jobs to speed the process
#        - Handle different encodings 
#        - Add params to handle custom dates to workaround 1 millon cap

<#
    .PARAMETER  Days
        Number of days back to search.

    .PARAMETER  TypeQuarantine
        Filter by quarantine type. Can be one or multiple of the following values separeted by commas: Bulk,HighConfPhish,Malware,Phish,Spam,SPOMalware,TransportRule. 
        If none is specified, all will be selected.

    .PARAMETER  OutputDir
        Full path of the output directory to store the eml files.

    .PARAMETER  Direction
        Filter by direction of the emails, can be Inbound or Outbound. If not defined both are selected.
#>

Param(
    [Parameter(Mandatory=$True)]
        [int]$Days,
    [Parameter(Mandatory=$False)]
        $TypeQuarantine,
    [Parameter(Mandatory=$True)]
        [string]$OutputDir,
    [Parameter(Mandatory=$False)]
        $Direction
    )


[DateTime]$DateEnd = Get-Date -format g
[DateTime]$DateStart = $DateEnd.AddDays($Days * -1)

$FoundCount = 0

For($i = 1; $i -le 1000; $i++)  # Maximum allowed pages is 1000
{
    $Command = 'Get-QuarantineMessage -PageSize 1000 -Page $i -StartReceivedDate $DateStart -EndReceivedDate $DateEnd '

    if ($PSBoundParameters.ContainsKey('TypeQuarantine') = $True){
        $Command += '-QuarantineTypes $TypeQuarantine '
    } 

    if ($PSBoundParameters.ContainsKey('Direction') = $True){
        $Command += '-Direction $Direction '
    } 

    $Messages = Invoke-Expression $Command

    If($Messages.count -gt 0)
    {
        Foreach ($Message in $Messages) #Download messages using Export-QuarantineMessage
        {
            #Progress information
            $Status = $Messages[-1].ReceivedTime.ToString("MM/dd/yyyy HH:mm") + " - " + $Messages[0].ReceivedTime.ToString("MM/dd/yyyy HH:mm") + "  [" + ("{0:N0}" -f ($i*1000)) + " Searched | " + $FoundCount + " Donwloaded]"
            Write-Progress -activity "Checking Messages (Up to 1 Million)..." -status $Status

            #Export message to file
            $e = Export-QuarantineMessage -Identity $Message.Identity 
            if ($e.BodyEncoding -eq "Base64")
            {
                $bytes = [Convert]::FromBase64String($e.eml)
                [IO.File]::WriteAllBytes($OutputDir+"\"+$FoundCount+".eml", $bytes)
            }
            else 
            {
                Write-Host "Message with Identity: "+$e.Identity+" found with a different bodyEncoding, unable to handle it"
            }
            $FoundCount += 1
        }

    }
    Else
    {
        Break
    }
}  

Write-Host $FoundCount "Entries Found & Logged In"
<#
.SYNOPSIS
  Get Microsoft 365 Message Center updates and post to Teams using webhooks
.DESCRIPTION
  Script to check Microsoft 365 Message Center updates for your tenant, configured to check last 24 hours (can be adapted as required). Run as a scheduled task, Azure automation etc.

  Create a webhook in Teams and copy the URI to the variable section below.

  Replace the variables with your own where feasible
  
  Credits to Evotec for the module used (also take a look at the documentation in the links below):
  https://github.com/EvotecIT/PSWinDocumentation.O365HealthService
  https://evotec.xyz/preparing-azure-app-registrations-permissions-for-office-365-health-service/
  https://evotec.xyz/powershell-way-to-get-all-information-about-office-365-service-health

  Pre-req: As admin:
  Install-Module PSWinDocumentation.O365HealthService

  Disclaimer: This script is offered "as-is" with no warranty. 
  While the script is tested and working in my environment, it is recommended that you test the script
  in a test environment before using in your production environment.
 
.NOTES
  Version:        1.0
  Author:         Einar Asting (einar@asting.net)
  Creation Date:  Oct 7nd 2019
  Purpose/Change: Initial version
.LINK
  https://github.com/einast/PS_M365_scripts
#>

Import-Module PSWinDocumentation.O365HealthService -Force

# Variables
    $ApplicationID = 'application ID'
    $ApplicationKey = 'application key'
    $TenantDomain = 'your FQDN' # Alternatively use DirectoryID if tenant domain fails
    $URI = 'Teams webhook URI'
    $Hours = '24'
    $Now = Get-Date
    $color = '0377fc' # Blue

# Get data
$O365 = Get-Office365Health -ApplicationID $ApplicationID -ApplicationKey $ApplicationKey -TenantDomain $TenantDomain -ErrorAction Stop

# Parse data
  ForEach ($msg in $O365.MessageCenterInformationExtended){
                                                
    If (($Now - [datetime]$msg.LastUpdatedTime).TotalHours -le $Hours) {
               
                $Payload = ConvertTo-Json -Depth 4 @{
                text = 'Message Center updates past ' +  $($Hours) + ' hours - ID: ' + $($msg.ID)
                themeColor = $color
                sections = @(
                            @{
                        title = $msg.Title
                        facts = @(
                            @{
                            name = 'Service'
                            value = $msg.AffectedService
                            },
                            @{
                            name = 'Severity'
                            value = $msg.Severity
                            },
                             @{
                            name = 'Classification'
                            value = $msg.Classification
                            },
                             @{
                            name = 'Action Type'
                            value = $msg.ActionType
                            },
                            @{
                            name = 'Description'
                            value = $msg.Message
                            },
                            @{
                            name = 'Last updated'
                            value = $msg.LastUpdatedTime
                            }
                        )
                    }
                )
            }
            #Convert to UTF8
            $Payload = ([System.Text.Encoding]::UTF8.GetBytes($Payload))
            Invoke-Webrequest -URI $URI -Method POST -Body $Payload
        }
    }
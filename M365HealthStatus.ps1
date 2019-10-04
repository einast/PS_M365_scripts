<#
.SYNOPSIS
  Get Microsoft 365 Service health status and post to Teams using webhooks
.DESCRIPTION
  Script to check Microsoft 365 Health status, configured to check last 15 minutes (can be adapted as required). Run as a scheduled task, Azure automation etc.

  Create a webhook in Teams and copy the URI to the variable section below.

  The output will be color coded (can be adapted as required) according to Classification of the entry:
  
  Red = Incident
  Yellow = Advisory
  Green = Resolved (Messages with a value in "End date")

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
  Version:        1.1
  Author:         Einar Asting (einar@asting.net)
  Creation Date:  Oct 2nd 2019
  Purpose/Change: Modified version
.LINK
  https://github.com/einast/PS_M365_scripts
#>

Import-Module PSWinDocumentation.O365HealthService -Force

# Variables
    $ApplicationID = 'application ID'
    $ApplicationKey = 'application key'
    $TenantDomain = 'your FQDN' # Alternatively use DirectoryID if tenant domain fails
    $URI = 'Teams webhook URI'
    $Now = Get-Date
    $Minutes = '15'

# Poll for data
$O365 = Get-Office365Health -ApplicationID $ApplicationID -ApplicationKey $ApplicationKey -TenantDomain $TenantDomain -ErrorAction Stop

# Poll and parse data
  ForEach ($inc in $O365.Incidents){
                    
                    #Set the color line of the card according to the Classification of the event, or if it has ended
                    if ($inc.Classification -eq "Incident")
                    {
                    $color = "ff0000" # Red
                    }
                    else
                        {
                        if ($inc.EndTime -ne $null)
                            {
                            $color = "00cc00" # Green
                            }
                            else
                                {
                                $color = "ffff00" # Yellow
                                }
                            }
                                 
    If (($Now - [datetime]$inc.LastUpdatedTime).TotalMinutes -le $Minutes) {
               
                $Payload = ConvertTo-Json -Depth 4 @{
                text = 'Service status updates past ' +  $($Minutes) + ' minutes - INC ID: ' + $($inc.ID)
                themeColor = $color
                sections = @(
                            @{
                        title = $inc.Title
                        facts = @(
                            @{
                            name = 'Service'
                            value = $inc.Service
                            },
                            @{
                            name = 'Severity'
                            value = $inc.Severity
                            },
                             @{
                            name = 'Classification'
                            value = $inc.Classification
                            },
                            @{
                            name = 'Description'
                            value = $inc.ImpactDescription
                            },
                            @{
                            name = 'Last Updated'
                            value = $inc.LastUpdatedTime
                            },
                            @{
                            name = 'Incident End Time'
                            value = $inc.EndTime
                            }                           
                        )
                    }
                )
            }
            #Convert to UTF8 and post any new events to Teams
            $Payload = ([System.Text.Encoding]::UTF8.GetBytes($Payload))
            Invoke-Webrequest -URI $URI -Method POST -Body $Payload
            }
       }
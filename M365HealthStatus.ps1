<#
.SYNOPSIS
  Get Microsoft 365 Service health status and post to Teams using webhooks
.DESCRIPTION
  Script to check Microsoft 365 Health status, configured to check last 24 hours (can be adapted as required). Run as a scheduled task.
  
  Credits: 
  https://evotec.xyz/preparing-azure-app-registrations-permissions-for-office-365-health-service/
  https://evotec.xyz/powershell-way-to-get-all-information-about-office-365-service-health
  Module used:
  https://github.com/EvotecIT/PSWinDocumentation.O365HealthService

  1. Install module
  2. Create a webhook in Teams and copy the URI to the variable section below.
  3. Replace the variables with your own
  4. Create an Azure app and generate application ID and key
  5. Run as required (e.g scheduled task)

  The output will be color coded (can be adapted as required) according to Status of the classification - Red = Incident, Yellow = Advisory

  Disclaimer: This script is offered "as-is" with no warranty. 
  While the script is tested and working in my environment, it is recommended that you test the script
  in a test environment before using in your production environment.
 
.NOTES
  Version:        1.0
  Author:         Einar Asting (einar@asting.net)
  Creation Date:  Oct 2nd 2019
  Purpose/Change: Initial version
.LINK
  https://github.com/einast/PS_M365_scripts
#>

Import-Module PSWinDocumentation.O365HealthService -Force

# Variables
    $ApplicationID = 'app id'
    $ApplicationKey = 'app key'
    $TenantDomain = 'tenant domain (NOT .onmicrosoft.com)' # Alternatively DirectoryID if tenant domain fails
    $URI = 'uri to Teams webhook'
    $Now = Get-Date
    $Hours = '24'

# Poll for data
$O365 = Get-Office365Health -ApplicationID $ApplicationID -ApplicationKey $ApplicationKey -TenantDomain $TenantDomain -ErrorAction Stop

# Parse data
  ForEach ($inc in $O365.IncidentsExtended){
               #Set the color line of the card according to the Status of the incident
                if ($inc.Classification -eq "Incident")
                    {
                    $color = "ff0000"
                    }
                    else
                        {
                        $color = "ffff00"
                                        }
                                 
    If (($Now - [datetime]$inc.LastUpdatedTime).TotalHours -le $Hours) {
               
                $Payload = ConvertTo-Json -Depth 4 @{
                #title = 'Microsoft 365 Service Status'
                text = 'Service status updates past ' +  $($Hours) + ' hours - INC ID: ' + $($inc.ID)
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
                            value = $inc.Message
                            },
                            @{
                            name = 'Last updated'
                            value = $inc.LastUpdatedTime
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
     
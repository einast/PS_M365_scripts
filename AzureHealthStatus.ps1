<#
.SYNOPSIS
  Azure Resource Health
.DESCRIPTION
  Script to check your current Azure tenant service health, configured to check last 12 hours (can be adapted as required). Run as a scheduled task, Azure automation runbook etc.
  Create a webhook in Teams and copy the URI to the variable section below.
  
  Replace the variables with your own where feasible
  
  Disclaimer: This script is offered "as-is" with no warranty. 
  While the script is tested and working in my environment, it is recommended that you test the script
  in a test environment before using in your production environment.
 
.NOTES
  Version:        1.0
  Author:         Einar Asting (einar@asting.net)
  Creation Date:  Oct 29nd 2019
  Purpose/Change: First version
.LINK
  https://github.com/einast/PS_M365_scripts
#>

# User defined variables
$ApplicationID = 'application ID'
$ApplicationKey = 'application key'
$tenantid = 'your Azure tenant ID'
$subscriptionid = 'your subscription ID'
$URI = 'Teams webhook URI'
$Hours = '12'
$Now =  (Get-Date).ToUniversalTime()

# Generate access token
$TokenEndpoint = {https://login.windows.net/{0}/oauth2/token} -f $tenantid
$ARMResource = "https://management.azure.com/"

$Body = @{
            'resource' = $ARMResource
            'client_id' = $ApplicationID
            'grant_type' = 'client_credentials'
            'client_secret' = $ApplicationKey
            }

$params = @{
    ContentType = 'application/x-www-form-urlencoded'
    Headers = @{'accept'='application/json'}
    Body = $Body
    Method = 'Post'
    URI = $TokenEndpoint
}

$token = Invoke-RestMethod @params
$token | select *,
@{L='Expires';E={[timezone]::CurrentTimeZone.ToLocalTime(([datetime]'1/1/1970').AddSeconds($_.expires_on))}} | fl *


# Request health status
$url = "https://management.azure.com/subscriptions/$subscriptionid/providers/Microsoft.ResourceHealth/events?api-version=2018-07-01"
$accesstoken = $token.access_token
$header = @{
    'Authorization' = 'Bearer ' + $accesstoken
}

$Events = (Invoke-RestMethod –Uri $url –Headers $header –Method GET)
$Events = $Events.value.properties

# Parse data
ForEach ($msg in $Events){
        
# Add updates posted last x hours
If (([datetime]$Now - [datetime]$msg.lastUpdateTime).TotalHours -le $Hours){
                 
                # Convert to JSON, replace invalid characters
                $Message = $msg.description -replace '"', ''
                $Message = ConvertTo-Json $Message

                $services = ($msg.impact | Select -ExpandProperty impactedservice) -join ',<br>'
                
                #Set the color line of the card according to the Status of the environment
                 if ($msg.status -eq "Resolved")
                    {
                    $color = "00cc00"
                    }
                    else
                       {
                       $color = "ff0000"
                       }  
          
$Payload =  @"
{
    "@context": "https://schema.org/extensions",
    "@type": "MessageCard",
    "potentialAction": [
            {
            "@type": "OpenUri",
            "name": "More info",
            "targets": [
                {
                    "os": "default",
                    "uri": "https://portal.azure.com/#blade/Microsoft_Azure_Health/ServiceIssuesBlade"
                }
            ]
        },
     ],
    "sections": [
        {
            "facts": [
                {
                    "name": "Status:",
                    "value": "$($msg.status)"
                },
                {
                    "name": "Event type:",
                    "value": "$($msg.eventType)"
                },
                {
                    "name": "Affected services:",
                    "value": "$($services)"
                },
                {
                    "name": "Last updated:",
                    "value": "$([datetime]$msg.lastUpdateTime)"
                }
                
            ],
            "text": $($Message)
        }
    ],
    "summary": "$($msg.header)",
    "themeColor": "$($color)",
    "title": "$($msg.title)"
}
"@

# If any new posts, add to Teams
Invoke-RestMethod -uri $uri -Method Post -body $Payload -ContentType 'application/json; charset=utf-8'

        }
     }
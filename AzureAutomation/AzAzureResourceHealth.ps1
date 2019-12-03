<#
.SYNOPSIS
  Azure Resource Health and post to Teams using webhooks
.DESCRIPTION
  Script to check Azure Resource Health for your tenant, configured to check last 1 hour (can be adapted as required). Run as an Azure automation runbook.
    
    Create Azure Automation Variable Assets for the below, and copy the names in to the script's variable section:
    --------------------------------------------------------------------------------------------------------------
    $AzVariableAzureHealthURI = 'Your Azure Health Teams URI variable name'
    $AzVariableAzureTenantID = 'Your Azure Tenant ID variable name'
    $AzVariableAzureSubscriptionID = 'Your Azure Subscription ID variable name'
    $AzVariableAzureApplicationID = 'Your Azure Application ID (credential) variable name'
    $AzVariableAzureApplicationKey = 'Your Azure Application Key (credential) variable name (can be encrypted)'
  
  The output will be color coded (can be adapted as required) according to Status of the Feature:
  
    Red = Active incident
    Green = Resolved incident
  
  Replace the variables with your own where feasible
  
  Disclaimer: This script is offered "as-is" with no warranty. 
  While the script is tested and working in my environment, it is recommended that you test the script
  in a test environment before using in your production environment.
 
.NOTES
  Version:        2.0
  Author:         Einar Asting (einar@thingsinthe.cloud)
  Creation Date:  Nov 22nd 2019
  Purpose/Change: Rewrote for Azure Automation
.LINK
  https://github.com/einast/PS_M365_scripts/AzureAutomation
#>

# User defined variables
# ----------------------

$AzVariableAzureHealthURI = 'Your Azure Health Teams URI variable name'
$AzVariableAzureTenantID = 'Your Azure Tenant ID variable name'
$AzVariableAzureSubscriptionID = 'Your Azure Subscription ID variable name'
$AzVariableAzureApplicationID = 'Your Azure Application ID (credential) variable name'
$AzVariableAzureApplicationKey = 'Your Azure Application Key (credential) variable name'
$Hours = '1'

# ----------------------


# Read the values from Azure Automation variables and add to script variables
$URI = Get-AutomationVariable -Name $AzVariableAzureHealthURI
$tenantid = Get-AutomationVariable -Name $AzVariableAzureTenantID
$subscriptionid = Get-AutomationVariable -Name $AzVariableAzureSubscriptionID
$ApplicationID = Get-AutomationVariable -Name $AzVariableAzureApplicationID
$ApplicationKey = Get-AutomationVariable -Name $AzVariableAzureApplicationKey

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
$Now =  (Get-Date).ToUniversalTime()

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

<#
.SYNOPSIS
  [Azure Automation version] Get Office 365 Endpoints updates (Worldwide) and post to Teams using webhooks
.DESCRIPTION
  Script to check Office 365 Endpoints updates (Worldwide) RSS feed, configured to check last 24 hours (can be adapted as required).
  
  Adapted for Azure Automation
  
  Create a webhook in Teams and copy the URI to the variable section below.
    
  Replace the variables with your own where feasible
  
  Disclaimer: This script is offered "as-is" with no warranty. 
  While the script is tested and working in my environment, it is recommended that you test the script
  in a test environment before using in your production environment.
 
.NOTES
  Version:        1.0
  Author:         Einar Asting (einar@thingsinthe.cloud)
  Creation Date:  Jan 29th 2020
  Purpose/Change: Initial version
.LINK
  https://github.com/einast/PS_M365_scripts
#>

# User defined variables (Azure Automation specific)
$AzTeamsURI = 'AAVTeamsURIO365EndpointUpdates' # Your Teams channel webhook URI
$AzO365EndpointURL = 'AAVO365EndpointURL' # I define this endpoint URL: https://docs.microsoft.com/en-us/office365/enterprise/urls-and-ip-address-ranges
$AzO365EndpointRSS = 'AAVO365EndpointRSS' # I define the RSS feed for the worldwide endpoints: https://endpoints.office.com/version/worldwide?allversions=true&format=rss&clientrequestid=b10c5ed1-bad1-445f-b386-b919946339a7
$Hours = '24'

# Request data
$Now = Get-Date
$updates = (Invoke-RestMethod -Uri (Get-AutomationVariable $AzO365EndpointRSS) -Headers $headerParams -Method Get)
$O365EndpointURL = Get-AutomationVariable $AzO365EndpointURL 

# Parse data
ForEach ($update in $updates) {

    # Add updates posted last $Hours                
    If (($Now - [datetime]$update.pubDate).TotalHours -le $Hours) {
                
        # Convert MessageText to JSON beforehand, if not the payload will fail.
        $Message = ConvertTo-Json $update.description

     # Generate payload(s)          
    $Payload = @"
{
    "@context": "https://schema.org/extensions",
    "@type": "MessageCard",
    "potentialAction": [
        {
            "@type": "OpenUri",
            "name": "Changes",
            "targets": [
                {
                    "os": "default",
                    "uri": "$($update.link)"
                }
            ]
        },
        {  
                "@type": "OpenUri",
                "name": "O365 Endpoints",
                "targets": [
                    {
                        "os": "default",
                        "uri": "$($O365EndpointURL)"
                    }
                ]
            }
        ],
        "sections": [
            {
                "facts": [
                    {
                        "name": "Published date:",
                        "value": "$($update.pubDate)"
                    }
                ],
                "text": "$($update.description)"
            }
        ],
        "summary": "O365 Endpoint changes - $($update.title)",
        "themeColor": "",
        "title": "O365 Endpoint changes - $($update.title)"
    }
"@
    # If any new posts, add to Teams
    Invoke-RestMethod -uri (Get-AutomationVariable $AzTeamsURI) -Method Post -body $Payload -ContentType 'application/json; charset=utf-8'
    }
}

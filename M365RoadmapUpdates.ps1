<#
.SYNOPSIS
  Get Microsoft 365 Roadmap update and post to Teams using webhooks
.DESCRIPTION
  Script to check Microsoft 365 Roadmap RSS feed, configured to check last 24 hours (can be adapted as required). Run as a scheduled task, Azure automation runbook etc.
  Create a webhook in Teams and copy the URI to the variable section below.
  
  The output will be color coded (can be adapted as required) according to Status of the Feature:
  
    Red = In development
    Yellow = Rolling out
    Green = Launched
  
  Replace the variables with your own where feasible
  
  Disclaimer: This script is offered "as-is" with no warranty. 
  While the script is tested and working in my environment, it is recommended that you test the script
  in a test environment before using in your production environment.
 
.NOTES
  Version:        2.0
  Author:         Einar Asting (einar@asting.net)
  Creation Date:  Oct 19nd 2019
  Purpose/Change: Rewrote code, new cards with buttons, removed temp file etc
.LINK
  https://github.com/einast/PS_M365_scripts
#>

# User defined variables
$URI = 'Teams webhook URI'
$Roadmap = 'https://www.microsoft.com/en-us/microsoft-365/RoadmapFeatureRSS'
$Hours = '24'
$Now = Get-Date

# Request data
$messages = Invoke-RestMethod -Uri $Roadmap

# Parse data
ForEach ($msg in $messages) {

    # Add updates posted last 24 hours                
    If (($Now - [datetime]$msg.pubDate).TotalHours -le $Hours) {
                
        # Count, join and prepare category for use in the card
        $categoryno = $msg.category.Count
        $category = $msg.category[1..$categoryno] -join ", "
                
        # Convert MessageText to JSON beforehand, if not the payload will fail.
        $Message = ConvertTo-Json $msg.description

        #Set the color line of the card according to the Status of the environment
        if ($msg.category.Contains("In development")) {
            $color = "ff0000"
        }
        
        elseif ($msg.category.Contains("Rolling out")) {
            $color = "ffff00"
        }
        else {
            $color = "00cc00"
        }
    
    # Generate payload(s)          
    $Payload = @"
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
                    "uri": "$($msg.Link)"
                }
            ]
        },
     ],
    "sections": [
        {
            "facts": [
                {
                    "name": "Status:",
                    "value": "$($msg.category[0])"
                },
                {
                    "name": "Category:",
                    "value": "$($category)"
                }
            ],
            "text": $($message)
        }
    ],
    "summary": "$($msg.Title)",
    "themeColor": "$($color)",
    "title": "Feature ID: $($msg.guid.'#text') - $($msg.Title)"
}
"@
    # If any new posts, add to Teams
    Invoke-RestMethod -uri $uri -Method Post -body $Payload -ContentType 'application/json; charset=utf-8'
    }
}

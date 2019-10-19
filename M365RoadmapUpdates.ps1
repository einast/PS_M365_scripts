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
$ApplicationID = 'App ID'
$ApplicationKey = 'App key'
$URI = 'Teams webhook URI'
$Roadmap = 'https://www.microsoft.com/en-us/microsoft-365/RoadmapFeatureRSS'
$Hours = '24'
$Now = Get-Date

# Request data
$messages = (Invoke-RestMethod -Uri $Roadmap -Headers $headerParams -Method Get)

# Parse data
ForEach ($msg in $messages){

        # Add updates posted last 24 hours                
        If (($Now - [datetime]$msg.pubDate).TotalHours -le $Hours) {

                # Convert MessageText to JSON beforehand, if not the payload will fail.
                $Message = ConvertTo-Json $msg.description

                #Set the color line of the card according to the Status of the environment
                if ($msg.category[0] -eq "In development")
                    {
                    $color = "ff0000"
                    }
                    else
                        {
                            if ($msg.category[0] -eq "Rolling out")
                                {
                                    $color = "ffff00"
                                    }
                                    else
                                        {
                                            $color = "00cc00"
                                            }
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
                    "name": "Environment:",
                    "value": "$($msg.category[1])"
                }
            ],
            "text": $($message)
        }
    ],
    "summary": "$($msg.Title)",
    "themeColor": "$($color)",
    "title": "$($msg.Title)"
}
"@
# If any new posts, add to Teams
Invoke-RestMethod -uri $uri -Method Post -body $Payload -ContentType 'application/json; charset=utf-8'
        }
     }
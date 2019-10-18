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
  
  Example doc for registering Azure application for credentials and permissions:
  https://evotec.xyz/preparing-azure-app-registrations-permissions-for-office-365-health-service/

  Disclaimer: This script is offered "as-is" with no warranty. 
  While the script is tested and working in my environment, it is recommended that you test the script
  in a test environment before using in your production environment.
 
.NOTES
  Version:        2.0
  Author:         Einar Asting (einar@asting.net)
  Creation Date:  Oct 17th 2019
  Purpose/Change: Rewrote card appearance
.LINK
  https://github.com/einast/PS_M365_scripts
#>

# User defined variables
$ApplicationID = 'application ID'
$ApplicationKey = 'application key'
$TenantDomain = 'your FQDN' # Alternatively use DirectoryID if tenant domain fails
$URI = 'Teams webhook URI'
$Now = Get-Date
$Minutes = '15'

# Request data
$body = @{
    grant_type="client_credentials";
    resource="https://manage.office.com";
    client_id=$ApplicationID;
    client_secret=$ApplicationKey;
    earliest_time="-$($Minutes)m@s"}

$oauth = Invoke-RestMethod -Method Post -Uri "https://login.microsoftonline.com/$($tenantdomain)/oauth2/token?api-version=1.0" -Body $body
$headerParams = @{'Authorization'="$($oauth.token_type) $($oauth.access_token)"}
$messages = (Invoke-RestMethod -Uri "https://manage.office.com/api/v1.0/$($tenantdomain)/ServiceComms/Messages" -Headers $headerParams -Method Get)
$incidents = $messages.Value | Where-Object {$_.MessageType -eq 'Incident'}

# Parse data
ForEach ($inc in $incidents){
                  
                # Set the color line of the card according to the Classification of the event, or if it has ended
                if ($inc.Classification -eq "Incident" -and $inc.EndTime -eq $null)
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
                             
# Add updates posted last $Minutes
If (($Now - [datetime]$inc.LastUpdatedTime).TotalMinutes -le $Minutes) {

# Pick latest message in the message index and convert the text to JSON before generating payload (if not it will fail).
$Message = $inc.Messages.MessageText[$inc.Messages.Count-1] | ConvertTo-Json # -Compress | %{[regex]::Unescape($_)}
  
# Generate payload(s)
$Payload =  @"
{
    "@context": "https://schema.org/extensions",
    "@type": "MessageCard",
    "potentialAction": [
            {
            "@type": "OpenUri",
            "name": "Post INC document",
            "targets": [
                {
                    "os": "default",
                    "uri": "$($inc.PostIncidentDocumentUrl)"
                }
            ]
        },           
    ],
    "sections": [
        {
            "facts": [
                {
                    "name": "Service:",
                    "value": "$($inc.WorkloadDisplayName)"
                },
                {
                    "name": "Status:",
                    "value": "$($inc.Status)"
                },
                {
                    "name": "Severity:",
                    "value": "$($inc.Severity)"
                },
                {
                    "name": "Classification:",
                    "value": "$($inc.Classification)"
                }
            ],
            "text": $($Message)
        }
    ],
    "summary": "$($Inc.Title)",
    "themeColor": "$($color)",
    "title": "$($Inc.Id) - $($Inc.Title)"
}
"@

# If any new posts, add to Teams
Invoke-RestMethod -uri $uri -Method Post -body $Payload -ContentType 'application/json; charset=utf-8'
  }
}

$uri = 'https://outlook.office.com/webhook/dd9359ec-ffab-473a-b7aa-3a7a465d4902@50909e35-f818-4485-9c84-5b4313d0caec/IncomingWebhook/8aa9f61d8cc24b37a5213175da718986/d5158fb0-4d85-4213-8177-355adea2f438'
$ApplicationID = '2320f718-bc95-4c0b-a7a4-c294ca1b5d8f'
$ApplicationKey = 'EFHYl746vbcB8L@pgXRV=KDpHMs?_[y?'
$TenantDomain = 'asting.net' # Alternatively use DirectoryID if tenant domain fails
$Now = Get-Date
$Hours = '24'    
$color = '0377fc'

<#
.SYNOPSIS
  Get Microsoft 365 Message Center updates and post to Teams using webhooks
.DESCRIPTION
  Script to check Microsoft 365 Message Center updates for your tenant, configured to check last 24 hours (can be adapted as required). Run as a scheduled task, Azure automation etc.
  Create a webhook in Teams and copy the URI to the variable section below.
  Replace the variables with your own where feasible
  
  Example doc for registering Azure application for credentials and permissions:
  https://evotec.xyz/preparing-azure-app-registrations-permissions-for-office-365-health-service/
  
  Disclaimer: This script is offered "as-is" with no warranty. 
  While the script is tested and working in my environment, it is recommended that you test the script
  in a test environment before using in your production environment.
 
.NOTES
  Version:        2.0
  Author:         Einar Asting (einar@asting.net)
  Creation Date:  Oct 17h 2019
  Purpose/Change: Updated card with buttons etc
.LINK
  https://github.com/einast/PS_M365_scripts
#>

# User defined variables
$ApplicationID = 'App ID'
$ApplicationKey = 'App key'
$TenantDomain = 'tenant domain' # Alternatively use DirectoryID if tenant domain fails
$URI = 'Teams webhook URI'
$Now = Get-Date
$Hours = '24'    
$color = '0377fc'

# Request data
    $body = @{
        grant_type="client_credentials";
        resource="https://manage.office.com";
        client_id=$ApplicationID;
        client_secret=$ApplicationKey;
        earliest_time="-$($Hours)h@s"}

    $oauth = Invoke-RestMethod -Method Post -Uri "https://login.microsoftonline.com/$($tenantdomain)/oauth2/token?api-version=1.0" -Body $body
    $headerParams = @{'Authorization'="$($oauth.token_type) $($oauth.access_token)"}
    $messages = (Invoke-RestMethod -Uri "https://manage.office.com/api/v1.0/$($tenantdomain)/ServiceComms/Messages" -Headers $headerParams -Method Get)
    $incidents = $messages.Value | Where-Object {$_.MessageType -eq 'MessageCenter'}

# Parse data
ForEach ($inc in $incidents){
                             
	# Add updates posted last 24 hours
	If (($Now - [datetime]$inc.LastUpdatedTime).TotalHours -le $Hours) {

	# Convert MessageText to JSON beforehand, if not the payload will fail.
	$Message = ConvertTo-Json $inc.Messages.MessageText

# Generate payload(s)
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
                    "uri": "$($inc.ExternalLink)"
                }
            ]

        },
        {
            "@type": "OpenUri",
            "name": "Blog link",
            "targets": [
                {
                    "os": "default",
                    "uri": "$($inc.BlogLink)"
                }
            ]

        },
        {
            "@type": "OpenUri",
            "name": "Help link",
            "targets": [
                {
                    "os": "default",
                    "uri": "$($inc.HelpLink)"
                }
            ]

        }        
    ],
    "sections": [
        {
            "facts": [
                {
                    "name": "Service:",
                    "value": "$($inc.AffectedWorkloadDisplayNames)"
                },
                {
                    "name": "Action Type:",
                    "value": "$($inc.ActionType)"
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
    "themeColor": "0377fc",
    "title": "$($Inc.Id) - $($Inc.Title)"
}
"@

# If any new posts, add to Teams
Invoke-RestMethod -uri $uri -Method Post -body $Payload -ContentType 'application/json'
  }
}
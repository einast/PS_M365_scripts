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
  Version:        2.1
  Author:         Einar Asting (einar@asting.net)
  Creation Date:  Jan 6th 2022
  Purpose/Change: Updated to Graph API
.LINK
  https://github.com/einast/PS_M365_scripts
#>

# User defined variables
$ApplicationID = 'application ID'
$ApplicationKey = 'application key'
$TenantDomain = 'your FQDN' # Alternatively use DirectoryID if tenant domain fails
$URI = 'Teams webhook URI'
$Now = Get-Date
$Hours = '24'    
$color = '0377fc'

# Request data
    $body = @{
        grant_type="client_credentials";
        resource="https://graph.microsoft.com";
        client_id=$ApplicationID;
        client_secret=$ApplicationKey;
        earliest_time="-$($Hours)h@s"}

    $oauth = Invoke-RestMethod -Method Post -Uri "https://login.microsoftonline.com/$($tenantdomain)/oauth2/token?api-version=1.0" -Body $body
    $headerParams = @{'Authorization'="$($oauth.token_type) $($oauth.access_token)"}
    $messages = (Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/admin/serviceAnnouncement/messages" -Headers $headerParams -Method Get)
    $incidents = $messages.Value #| Where-Object {$_.MessageType -eq 'MessageCenter'}

# Parse data
ForEach ($inc in $incidents){
                             
	# Add updates posted last 24 hours
	If (($Now - [datetime]$inc.lastModifiedDateTime).TotalHours -le $Hours) {

	# Clean up output (replace brackets with HTML), convert MessageText to JSON beforehand, if not the payload will fail.
    	$Message = $inc.body.content -replace [regex]::Escape("["), "<br><b>" -replace [regex]::Escape("]"), "</b><br><br>" 
    	$Message = ConvertTo-Json $Message

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
                    "value": "$($inc.services)"
                },
                {
                    "name": "Category:",
                    "value": "$($inc.category)"
                },
                {
                    "name": "Severity:",
                    "value": "$($inc.severity)"
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

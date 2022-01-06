<#
.SYNOPSIS
  Get Microsoft 365 Message Center updates and post to Teams using webhooks
.DESCRIPTION
  Script to check Microsoft 365 Message Center updates for your tenant, configured to check last 24 hours (can be adapted as required). Run as a scheduled task, Azure automation etc.
  
  Adjust the variable section to your environment.
  
  You need to create an Azure application, an example doc for registering Azure application for credentials and permissions:
  https://evotec.xyz/preparing-azure-app-registrations-permissions-for-office-365-health-service/
     
  # User defined variables
  # ----------------------
  $AzVariableApplicationIDName = 'Create an Azure application, and copy the Application ID to an Azure Automation variable asset'
  $AzVariableApplicationKeyName = 'From the same Azure application, copy the Application Key to an Azure Automation variable asset (I recommend to encrypt it)'
  $AzVariableTenantDomainName = 'Create an Azure Automation variable asset containing your tenant domain FQDN'
  $AzVariableTeamsURIName = 'You need to create a Teams webhook, and copy the URI to an Azure Automation variable asset'
    
  Disclaimer: This script is offered "as-is" with no warranty. 
  While the script is tested and working in my environment, it is recommended that you test the script
  in a test environment before using in your production environment.
 
.NOTES
  Version:        1.0
  Author:         Einar Asting (einar@thingsinthe.cloud)
  Creation Date:  Jan 6th 2022
  Purpose/Change: Replaced API with Graph API
.LINK
  https://github.com/einast/PS_M365_scripts/AzureAutomation
#>

# User defined variables
# ----------------------
$AzVariableApplicationIDName = 'AzApplicationID'
$AzVariableApplicationKeyName = 'AzApplicationKey'
$AzVariableTenantDomainName = 'AzTenantDomain'
$AzVariableTeamsURIName = 'M365MessageCenterURI'
$Hours = '24'
$Color = '0377fc'
# ---------------------


# Get current date and time
$Now = Get-Date

# Converting Azure Automation variable assets to script variables
$Tenantdomain = Get-AutomationVariable -Name $AzVariableTenantDomainName
$ApplicationID = Get-AutomationVariable -Name $AzVariableApplicationIDName
$ApplicationKey = Get-AutomationVariable -Name $AzVariableApplicationKeyName
$URI = Get-AutomationVariable -Name $AzVariableTeamsURIName

# Request data
    $body = @{
        grant_type="client_credentials";
        resource="https://graph.microsoft.com";
        client_id=$ApplicationID;
        client_secret=$ApplicationKey;
        earliest_time="-$($Hours)h@s"}

    $oauth = Invoke-RestMethod -Method Post -Uri "https://login.microsoftonline.com/$($Tenantdomain)/oauth2/token?api-version=1.0" -Body $body
    $headerParams = @{'Authorization'="$($oauth.token_type) $($oauth.access_token)"}
    $messages = (Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/admin/serviceAnnouncement/messages" -Headers $headerParams -Method Get)
    $incidents = $messages.Value #| Where-Object {$_.MessageType -eq 'MessageCenter'}    

# Parse data
ForEach ($inc in $incidents){
                             
	# Add updates posted last $Hours
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
    "themeColor": "$($Color)",
    "title": "$($Inc.Id) - $($Inc.Title)"
}
"@

# If any new posts, add to Teams
Invoke-RestMethod -uri $uri -Method Post -body $Payload -ContentType 'application/json; charset=utf-8'
  }
}

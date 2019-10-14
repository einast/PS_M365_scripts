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
  Version:        1.5
  Author:         Einar Asting (einar@asting.net)
  Creation Date:  Oct 14th 2019
  Purpose/Change: Replaced module with Invoke-RestMethod
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
                             
If (($Now - [datetime]$inc.LastUpdatedTime).TotalHours -le $Hours) {

    $Payload = ConvertTo-Json -Depth 4 @{
        Text = "Office 365 Message Center updates past " + $Hours + " hours - MC ID: " + $inc.ID
        themeColor = $color
        Summary = $Inc.Title
        sections = @(
            @{
            facts = @(
            @{
                name = "Title"
                value = $inc.Title
                },
                @{
                name = "Affected Service"
                value = $inc.AffectedWorkloadDisplayNames
                },
                @{
                name = "Classification"
                value = $inc.Classification
                },
                @{
                name = "Action Type"
                value = $inc.ActionType
                },
                @{
                name = "Information"
                value = $inc.Messages.MessageText
                },
                @{
                name = 'Last Updated (UTC)'
                value = $inc.LastUpdatedTime
                },
                @{
                name = "Link"
                value = "<a href="+$($inc.ExternalLink)+">"+$($inc.ExternalLink)+"</a>"
                }
                )
            }
          )
        }
      }
    }
Invoke-RestMethod -ContentType "Application/Json" -Method Post -Body $Payload -Uri $URI -Verbose
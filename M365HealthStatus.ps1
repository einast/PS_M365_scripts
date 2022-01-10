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
  Version:        2.4
  Author:         Einar Asting (einar@thingsinthe.cloud)
  Creation Date:  January 4th 2022
  Purpose/Change: Fixed failing API by replacing with Graph API
.LINK
  https://github.com/einast/PS_M365_scripts
#>

# User defined variables
$ApplicationID = 'application ID'
$ApplicationKey = 'application key'
$TenantDomain = 'your FQDN' # Alternatively use DirectoryID if tenant domain fails
$URI = 'Teams webhook URI'
$Minutes = '15'

# Service(s) to monitor
# Leave the one(s) you DON'T want to check empty (with '' ), add a value in the ones you WANT to check (I added 'yes' for readability

$ExchangeOnline = 'yes'
$MicrosoftForms = ''
$MicrosoftIntune = ''
$MicrosoftKaizala = ''
$SkypeforBusiness = ''
$MicrosoftDefenderATP = ''
$MicrosoftFlow = ''
$FlowinMicrosoft365 = ''
$MicrosoftTeams = 'yes'
$MobileDeviceManagementforOffice365 = ''
$OfficeClientApplications = ''
$Officefortheweb = ''
$OneDriveforBusiness = 'yes'
$IdentityService = ''
$Office365Portal = 'yes'
$OfficeSubscription = ''
$Planner = ''
$PowerApps = ''
$PowerAppsinMicrosoft365 = ''
$PowerBI = ''
$AzureInformationProtection = ''
$SharePointOnline = 'yes'
$MicrosoftStaffHub = ''
$YammerEnterprise = ''
$Microsoft365Suite = ''

# Classification(s) to monitor
# Leave the one(s) you DON'T want to check empty (with '' ), add a value in the ones you WANT to check (I added 'yes' for readability)

$Incident = 'yes'
$Advisory = ''

# Build the Services array            
$ServicesArray = @()            
            
# If Services variables are present, add with 'eq' comparison            
if($ExchangeOnline){$ServicesArray += '$_.service -eq "Exchange Online"'}            
if($MicrosoftForms){$ServicesArray += '$_.service -eq "Microsoft Forms"'}
if($MicrosoftIntune){$ServicesArray += '$_.service -eq "Microsoft Intune"'}
if($MicrosoftKaizala){$ServicesArray += '$_.service -eq "Microsoft Kaizala"'} 
if($SkypeforBusiness){$ServicesArray += '$_.service -eq "Skype for Business"'}
if($MicrosoftDefenderATP){$ServicesArray += '$_.service -eq "Microsoft Defender ATP"'}
if($MicrosoftFlow){$ServicesArray += '$_.service -eq "Microsoft Flow"'}
if($FlowinMicrosoft365){$ServicesArray += '$_.service -eq "Flow in Microsoft 365"'}
if($MicrosoftTeams){$ServicesArray += '$_.service -eq "Microsoft Teams"'}
if($MobileDeviceManagementforOffice365){$ServicesArray += '$_.service -eq "Mobile Device Management for Office 365"'}
if($OfficeClientApplications){$ServicesArray += '$_.service -eq "Office Client Applications"'}
if($Officefortheweb){$ServicesArray += '$_.service -eq "Office for the web"'}
if($OneDriveforBusiness){$ServicesArray += '$_.service -eq "OneDrive for Business"'}
if($IdentityService){$ServicesArray += '$_.service -eq "Identity Service"'}
if($Office365Portal){$ServicesArray += '$_.service -eq "Office 365 Portal"'}
if($OfficeSubscription){$ServicesArray += '$_.service -eq "Office Subscription"'}
if($Planner){$ServicesArray += '$_.service -eq "Planner"'}
if($PowerApps){$ServicesArray += '$_.service -eq "PowerApps"'}
if($PowerAppsinMicrosoft365){$ServicesArray += '$_.service -eq "PowerApps in Microsoft 365"'}
if($PowerBI){$ServicesArray += '$_.service -eq "Power BI"'}
if($AzureInformationProtection){$ServicesArray += '$_.service -eq "Azure Information Protection"'}
if($SharepointOnline){$ServicesArray += '$_.service -eq "Sharepoint Online"'}
if($MicrosoftStaffHub){$ServicesArray += '$_.service -eq "Microsoft StaffHub"'}
if($YammerEnterprise){$ServicesArray += '$_.service -eq "Yammer Enterprise"'}
if($Microsoft365Suite){$ServicesArray += '$_.service -eq "Microsoft 365 Suite"'}

# Build the Services where array into a string and joining each statement with -or     
$ServicesString = $ServicesArray -Join " -or "

# Build the Classification array            
$ClassificationArray = @()            
            
# If Classification variables are present, add with 'eq' comparison            
if($Incident){$ClassificationArray += '$_.Classification -eq "incident"'}            
if($Advisory){$ClassificationArray += '$_.Classification -eq "advisory"'}            

# Build the Classification where array into a string and joining each statement with -or            
$ClassificationString = $ClassificationArray -Join " -or "

# Request data
$body = @{
    grant_type="client_credentials";
    resource="https://graph.microsoft.com";
    client_id=$ApplicationID;
    client_secret=$ApplicationKey;
    earliest_time="-$($Minutes)m@s"}

$oauth = Invoke-RestMethod -Method Post -Uri "https://login.microsoftonline.com/$($tenantdomain)/oauth2/token?api-version=1.0" -Body $body
$headerParams = @{'Authorization'="$($oauth.token_type) $($oauth.access_token)"}
$messages = (Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/admin/serviceAnnouncement/issues" -Headers $headerParams -Method Get)
$incidents = $messages.Value | Where-Object ([scriptblock]::Create($ClassificationString)) | Where-Object ([scriptblock]::Create($ServicesString))

$Now = Get-Date

# Parse data
ForEach ($inc in $incidents){
                
                # Add updates posted last $Minutes
                If (($Now - [datetime]$inc.lastModifiedDateTime).TotalMinutes -le $Minutes) {
                
                # Set the color line of the card according to the Classification of the event, or if it has ended
                if ($inc.Classification -eq "incident" -and $inc.EndTime -eq $null)
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

# Pick message in the message index and convert the text to JSON before generating payload (if not it will fail).
$Message = $inc.posts.description.content[$inc.Messages.Count-1] | ConvertTo-Json 

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
                    "value": "$($inc.service)"
                },
                {
                    "name": "Status:",
                    "value": "$($inc.Status)"
                },
                {
                    "name": "Classification:",
                    "value": "$($inc.classification)"
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

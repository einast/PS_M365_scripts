<#
.SYNOPSIS
  Get Microsoft 365 Service health status and post to Teams using webhooks
.DESCRIPTION
  Script to check Microsoft 365 Health status, configured to check last 15 minutes (can be adapted as required). Run in Azure automation.
  
  Create a webhook in Teams and copy the URI to the variable section below.
  
  The output will be color coded (can be adapted as required) according to Classification of the entry:
  
  Red = Incident
  Yellow = Advisory
  Green = Resolved (Messages with a value in "End date")
  Replace the variables with your own where feasible
  
  Add a value to the Services and Classifications you would like to be notified about (I used 'yes' for readability), leave the others empty.
  
  Pushover notification is possible, enabled by default. Comment out the entire section if you don't want to user
  Pushover.

  Example doc for registering Azure application for credentials and permissions:
  https://evotec.xyz/preparing-azure-app-registrations-permissions-for-office-365-health-service/
  Disclaimer: This script is offered "as-is" with no warranty. 
  While the script is tested and working in my environment, it is recommended that you test the script
  in a test environment before using in your production environment.
 
.NOTES
  Version:        2.6
  Author:         Einar Asting (einar@thingsinthe.cloud)
  Creation Date:  January 6th 2022
  Purpose/Change: Updated script to use Graph API
.LINK
  https://github.com/einast/PS_M365_scripts/blob/master/AzureAutomation/AzO365ServiceHealth.ps1
#>

# Adjust the values below as needed

# ------------------------------------- USER DEFINED VARIABLES -------------------------------------

$ApplicationID = Get-AutomationVariable -Name 'AzO365ServiceHealthApplicationID'
$ApplicationKey = Get-AutomationVariable -Name 'AzO365ServiceHealthApplicationKey'
$TenantDomain = Get-AutomationVariable -Name 'AzO365TenantDomain'
$URI = Get-AutomationVariable -Name 'AzO365TeamsURI'
$Minutes = '15'

# Pushover notifications in case Teams is down.
# Due to limitations and readability, the script will only send the title of the incident/advisory to Pushover. 
# COMMENT OUT THIS SECTION IF YOU DON'T WANT TO USE PUSHOVER!

$Pushover = 'yes' # Comment out if you don't want to use Pushover. I use 'yes' for readability.
$PushoverToken = Get-AutomationVariable -Name 'AzO365PushoverToken' # Your API token. Comment out if you don't want to use Pushover
$PushoverUser = Get-AutomationVariable -Name 'AzO365PushoverUser' # User/Group token. Comment out if you don't want to use Pushover
$PushoverURI = 'https://api.pushover.net/1/messages.json' # DO NOT CHANGE! Default Pushover URI. Comment out if you don't want to use Pushover

# Service(s) to monitor
# Leave the one(s) you DON'T want to check empty (with '' ), add a value in the ones you WANT to check (I added 'yes' for readability)

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
$Advisory = 'yes'

# ------------------------------------- END OF USER DEFINED VARIABLES -------------------------------------

# Build the Services array            
$ServicesArray = @()            
            
# If Services variables are present, add with 'eq' comparison            
if($ExchangeOnline){$ServicesArray += '$_.WorkloadDisplayName -eq "Exchange Online"'}            
if($MicrosoftForms){$ServicesArray += '$_.WorkloadDisplayName -eq "Microsoft Forms"'}
if($MicrosoftIntune){$ServicesArray += '$_.WorkloadDisplayName -eq "Microsoft Intune"'}
if($MicrosoftKaizala){$ServicesArray += '$_.WorkloadDisplayName -eq "Microsoft Kaizala"'} 
if($SkypeforBusiness){$ServicesArray += '$_.WorkloadDisplayName -eq "Skype for Business"'}
if($MicrosoftDefenderATP){$ServicesArray += '$_.WorkloadDisplayName -eq "Microsoft Defender ATP"'}
if($MicrosoftFlow){$ServicesArray += '$_.WorkloadDisplayName -eq "Microsoft Flow"'}
if($FlowinMicrosoft365){$ServicesArray += '$_.WorkloadDisplayName -eq "Flow in Microsoft 365"'}
if($MicrosoftTeams){$ServicesArray += '$_.WorkloadDisplayName -eq "Microsoft Teams"'}
if($MobileDeviceManagementforOffice365){$ServicesArray += '$_.WorkloadDisplayName -eq "Mobile Device Management for Office 365"'}
if($OfficeClientApplications){$ServicesArray += '$_.WorkloadDisplayName -eq "Office Client Applications"'}
if($Officefortheweb){$ServicesArray += '$_.WorkloadDisplayName -eq "Office for the web"'}
if($OneDriveforBusiness){$ServicesArray += '$_.WorkloadDisplayName -eq "OneDrive for Business"'}
if($IdentityService){$ServicesArray += '$_.WorkloadDisplayName -eq "Identity Service"'}
if($Office365Portal){$ServicesArray += '$_.WorkloadDisplayName -eq "Office 365 Portal"'}
if($OfficeSubscription){$ServicesArray += '$_.WorkloadDisplayName -eq "Office Subscription"'}
if($Planner){$ServicesArray += '$_.WorkloadDisplayName -eq "Planner"'}
if($PowerApps){$ServicesArray += '$_.WorkloadDisplayName -eq "PowerApps"'}
if($PowerAppsinMicrosoft365){$ServicesArray += '$_.WorkloadDisplayName -eq "PowerApps in Microsoft 365"'}
if($PowerBI){$ServicesArray += '$_.WorkloadDisplayName -eq "Power BI"'}
if($AzureInformationProtection){$ServicesArray += '$_.WorkloadDisplayName -eq "Azure Information Protection"'}
if($SharepointOnline){$ServicesArray += '$_.WorkloadDisplayName -eq "Sharepoint Online"'}
if($MicrosoftStaffHub){$ServicesArray += '$_.WorkloadDisplayName -eq "Microsoft StaffHub"'}
if($YammerEnterprise){$ServicesArray += '$_.WorkloadDisplayName -eq "Yammer Enterprise"'}
if($Microsoft365Suite){$ServicesArray += '$_.WorkloadDisplayName -eq "Microsoft 365 Suite"'}

# Build the Services where array into a string and joining each statement with -or     
$ServicesString = $ServicesArray -Join " -or "

# Build the Classification array            
$ClassificationArray = @()            
            
# If Classification variables are present, add with 'eq' comparison            
if($Incident){$ClassificationArray += '$_.Classification -eq "Incident"'}            
if($Advisory){$ClassificationArray += '$_.Classification -eq "Advisory"'}            

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
If (($Now - [datetime]$inc.lastModifiedDateTime).TotalMinutes -le $Minutes) {

# Pick latest message in the message index and convert the text to JSON before generating payload (if not it will fail).
$Message = $incidents.posts.description.content[$inc.Messages.Count-1] | ConvertTo-Json

  
# Generate payload(s)
$Payload =  @"
{
    "@context": "https://schema.org/extensions",
    "@type": "MessageCard",
    "potentialAction": [
            {
            "@type": "OpenUri",
            "name": "Post INC report",
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

# If Pushover alerts are enabled, send to Pushover
if($Pushover){
$POparameters = @{
  token = $PushoverToken
  user = $PushoverUser
  message = "$($Inc.Id) - $($Inc.Title) `n`n$($incidents.posts.description.content[$inc.Messages.Count-1])"
}
$POparameters | Invoke-RestMethod -Uri $PushoverUri -Method Post
}
  }
}

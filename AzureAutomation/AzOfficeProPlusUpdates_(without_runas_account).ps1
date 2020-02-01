<#
.SYNOPSIS
  Get Microsoft Office ProPlus channel updates and post to Teams using webhooks and Azure Automation variable assets
.DESCRIPTION
  Script to check Microsoft Office ProPlus channel updates, configured to check last 12 hours (can be adapted as required).
  
  Adapted for Azure automation.
  
  Create one or more webhook in Teams (if you want to split the updates into separate channels) and copy the URI(s) to the user variable section below.
  The output is color coded (can be adapted as required). Default is green.
  Script for usage without Runas account in Azure Automation. You need to manually create the variable assets before running the script:
  $AzAutomationMonthlyVariable
  $AzAutomationSACTVariable
  $AzAutomationSACVariable
  
  # Monthly channel 
  $AzAutomationURIMonthlyVariable = The Azure Automation variable asset you stored your Monthly Teams URI in (https://) Comment out to not check
  $AzAutomationMonthlyVariable = Name of Azure Automation variable asset containing your last successful Monthly payload.
  
  # SACT channel
  $AzAutomationURISactVariable = The Azure Automation variable asset you stored your SACT Teams URI in (https://) Comment out to not check
  $AzAutomationSACTVariable = Name of Azure Automation variable asset containing your last successful SACT payload.
  
  # SAC channel
  $AzAutomationURISacVariable = The Azure Automation variable asset you stored your SAC Teams URI in (https://) Comment out to not check
  $AzAutomationSACVariable = Name of Azure Automation variable asset containing your last successful SAC payload.
  
  # Generic variables
  $Hours = Last number of hours to check for updates. Align with schedule. Default set to 12 hours
  $Color = Set to green as default
  
  # Handling oversize payloads (MS Teams supports between 18-40KB), tested and set to 18KB as default, feel free to test and adjust your own value
  $MaxPayloadSize = '18000' is the default value, as I had no success with a larger number

  # Custom message appended to truncated payload, adjust as you like
  $Trunktext = "DUE TO TEAMS LIMITATIONS, CHANGELOG IS TRUNCATED. CLICK THE 'MORE INFO' LINK FOR ALL DETAILS!" # Feel free to customize, I added some 
  calculations in the script that add this text to the new payload, while keeping it under the limit set above

  Disclaimer: This script is offered "as-is" with no warranty. 
  While the script is tested and working in my environment, it is recommended that you test the script
  in a test environment before using in your production environment.
 
.NOTES
  Version:        1.8
  Author:         Einar Asting (einar@thingsinthe.cloud)
  Creation Date:  Jan 20th 2020
  Purpose/Change: Added logic to check payload size due to Microsoft limits
.LINK
  https://github.com/einast/PS_M365_scripts
#>

# User defined variables
# ----------------------
# If you want to check Monthly Channel, Semi-Annual Channel Targeted (SACT) and/or Semi-Annual Channel, add your Teams URI in the variables fields. 
# Comment out the ones you don't want to check.

# Monthly channel 
$AzAutomationURIMonthlyVariable = 'AzMonthlyURI' # Comment out to _not_ check this channel
$AzAutomationPayloadMonthlyVariable = 'MonthlyPayloadAZ' # Will be created by the script if not existing

# SACT channel
$AzAutomationURISactVariable = 'AzSactURI' # Comment out to _not_ check this channel
$AzAutomationPayloadSACTVariable = 'SACTPayloadAZ' # Will be created by the script if not existing

# SAC channel
$AzAutomationURISacVariable = 'AzSacURI' # Comment out to _not_ check this channel
$AzAutomationPayloadSACVariable = 'SACPayloadAZ' # Will be created by the script if not existing

# Generic variables
$Hours = '12' # Set the time window to check for updates, align with your schedules
$Color = '00ff00' # Green

# Handling oversize payloads (MS Teams supports between 18-40KB), tested and set to 18KB as default, feel free to test and adjust your own value
$MaxPayloadSize = '18000'

# Custom message appended to truncated payload, adjust as you like
$Trunktext = "DUE TO TEAMS LIMITATIONS, CHANGELOG IS TRUNCATED. CLICK THE 'MORE INFO' LINK FOR ALL DETAILS!"

# ---------------------

# Setting other script variables
$Now = Get-Date -Format 'yyyy-MM-dd HH:mm'
$Year = Get-Date -Format yyyy
$Monthly = 'https://docs.microsoft.com/en-us/officeupdates/monthly-channel-' +$Year
$SAC = 'https://docs.microsoft.com/en-us/officeupdates/semi-annual-channel-' +$Year
$SACT = 'https://docs.microsoft.com/en-us/officeupdates/semi-annual-channel-targeted-' +$Year
$Trunkprefix = "...<br><br><b>"
$TrunkAppend = $Trunkprefix + $Trunktext

# Looking for new updates
# ---------------------

# Monthly channel
# ---------------

# Check if channel is set for checking, if so, do stuff
If ($AzAutomationURIMonthlyVariable) {
$MonthlyURI = Get-AutomationVariable -Name $AzAutomationURIMonthlyVariable

#Get data
$Monthlyweb = Invoke-RestMethod -Uri $Monthly

# Find article's last updated time
$monthlydatepattern = '\d{4}-\d{2}-\d{2} \d{2}:\d{2} [AP]M'
$monthlyLastUpdated = $monthlyweb | select-string  -Pattern $monthlydatepattern -AllMatches | % { $_.Matches } | % { $_.Value }

# Convert match into Date/Time
$monthlyDate = Get-Date $monthlyLastUpdated

# Check if updates are newer than time variable, if so, do stuff
                             
	# Calculate time difference
	If (([datetime]$Now - [datetime]$monthlyDate).TotalHours -le $Hours) {

    # Picking out title
    $monthlytitlepattern = '(?<=\<h2 id="v.*?\>)(.*?)(?=<\/h2\>)'
    $monthlytitle = $Monthlyweb | select-string  -Pattern $monthlytitlepattern -AllMatches | % { $_.Matches } | % { $_.Value } | Select-Object -First 1

    # Tailor the "More info" button by adding suffix to link to right section of the webpage
    $monthlylinkpattern = '(?<=\<h2.*?\")(.*?)(?=\"\>)'
    $monthlylink = $Monthlyweb | Select-String -Pattern $monthlylinkpattern -AllMatches | % { $_.Matches } | % { $_.Value } | Select-Object -Index 1

    # Select latest updates
    $monthlycontentpattern = '(\<h2 id="v.+?\>)(.|\n)*?(?=(\<h2 id="v.+?\>|<div class.+?\>))'
    $monthlyupdate = $Monthlyweb | select-string  -Pattern $monthlycontentpattern -AllMatches | % { $_.Matches } | % { $_.Value } | Select-Object -First 1
    $monthlycontent = $monthlyupdate | ConvertTo-Json

#Generate payload
          
$MonthlyPayload =  @"
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
                    "uri": "https://docs.microsoft.com/en-us/officeupdates/monthly-channel-$($Year)#$($monthlylink)"
                }
            ]
        },
     ],
    "sections": [
        {
            "facts": [
                {
                    "name": "Version updated:",
                    "value": "$($monthlyDate)"
                }
                
            ],
            "text": $monthlycontent
        }
    ],
    "summary": "O365 ProPlus Monthly",
    "themeColor": "$($color)",
    "title": "Monthly Channel release: $($monthlytitle)"
}
"@

# If any new posts, add to Teams. If new content matches content of previous payload, do not post
$monthlyPayloadAZ = Get-AutomationVariable $AzAutomationPayloadMonthlyVariable

# First check if Payload is over 18K (Microsoft Teams limit)

if ($MonthlyPayload.Length -gt $MaxPayloadSize) {  

    # Find payload + append text total length
    $MonthlyTotalLength = $monthlyPayload.Length + $TrunkAppend.Length

    # Find the overshooting value
    $MonthlyPayloadOverSize = ($MaxPayloadSize - $MonthlyTotalLength)

    # At what point in the original Payload content do we have to split the JSON
    $MonthlyPayloadSplitValue = $MonthlyContent.Length - (-$MonthlyPayloadOverSize)

    # Split the JSON into a new Payload
    $MonthlyNewSplitContent = $monthlyupdate.Substring(0,$MonthlyPayloadSplitValue)
    
    # Create new truncated payload
    $MonthlyNewContent = $MonthlyNewSplitContent + $TrunkAppend 
    $MonthlyNewPayloadContent = ConvertTo-Json $MonthlyNewContent
    
#Generate truncated payload
          
$MonthlyNewPayload =  @"
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
                    "uri": "https://docs.microsoft.com/en-us/officeupdates/monthly-channel-$($Year)#$($saclink)"
                }
            ]
        },
     ],
    "sections": [
        {
            "facts": [
                {
                    "name": "Version updated:",
                    "value": "$($monthlyDate)"
                }
                
            ],
            "text": $MonthlyNewPayloadContent
        }
    ],
    "summary": "O365 ProPlus Monthly",
    "themeColor": "$($color)",
    "title": "Monthly Channel release: $($monthlytitle)"
}
"@

If ($monthlycontent -ne $MonthlyPayloadAZ) {
Invoke-RestMethod -uri $MonthlyURI -Method Post -body $MonthlyNewPayload -ContentType 'application/json; charset=utf-8'
Set-AutomationVariable -Name $AzAutomationPayloadMonthlyVariable -Value ($monthlycontent -as [string])
      }
    Else {
    }
    
    }
Else {
    If ($monthlycontent -ne $MonthlyPayloadAZ) {
    Invoke-RestMethod -uri $MonthlyURI -Method Post -body $MonthlyPayload -ContentType 'application/json; charset=utf-8'
    Set-AutomationVariable -Name $AzAutomationPayloadMonthlyVariable -Value ($monthlycontent -as [string])
          }
    Else {
    }
    
    }
        }
    }
    Else {
    }


# Semi-Annual channel (targeted) (SACT)
# -------------------------------------

# Check if channel is set for checking, if so, do stuff
If ($AzAutomationURISactVariable) {
$sactURI = Get-AutomationVariable -Name $AzAutomationURISactVariable

#Get data
$sactweb = Invoke-RestMethod -Uri $SACT

# Find article's last updated time
$sactdatepattern = '\d{4}-\d{2}-\d{2} \d{2}:\d{2} [AP]M'
$sactLastUpdated = $sactweb | select-string  -Pattern $sactdatepattern -AllMatches | % { $_.Matches } | % { $_.Value }

# Convert match into Date/Time
$sactDate = Get-Date $sactLastUpdated

# Check if updates are newer than time variable, if so, do stuff
                             
	# Calculate time difference
	If (([datetime]$Now - [datetime]$sactDate).TotalHours -le $Hours) {

    # Picking out title
    $sacttitlepattern = '(?<=\<h2 id="v.*?\>)(.*?)(?=<\/h2\>)'
    $sacttitle = $sactweb | select-string  -Pattern $sacttitlepattern -AllMatches | % { $_.Matches } | % { $_.Value } | Select-Object -First 1

    # Tailor the "More info" button by adding suffix to link to right section of the webpage
    $sactlinkpattern = '(?<=\<h2.*?\")(.*?)(?=\"\>)'
    $sactlink = $sactweb | Select-String -Pattern $sactlinkpattern -AllMatches | % { $_.Matches } | % { $_.Value } | Select-Object -Index 1

    # Select latest updates
    $sactcontentpattern = '(\<h2 id="v.+?\>)(.|\n)*?(?=(\<h2 id="v.+?\>|<div class.+?\>))'
    $sactupdate = $sactweb | select-string  -Pattern $sactcontentpattern -AllMatches | % { $_.Matches } | % { $_.Value } | Select-Object -First 1
    $sactcontent = $sactupdate | ConvertTo-Json

#Generate payload
          
$sactPayload =  @"
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
                    "uri": "https://docs.microsoft.com/en-us/officeupdates/semi-annual-channel-targeted-$($Year)#$($sactlink)"
                }
            ]
        },
     ],
    "sections": [
        {
            "facts": [
                {
                    "name": "Version updated:",
                    "value": "$($sactDate)"
                }
                
            ],
            "text": $sactcontent
        }
    ],
    "summary": "O365 ProPlus Semi-Annual (targeted)",
    "themeColor": "$($color)",
    "title": "Semi-Annual Channel (targeted) release: $($sacttitle)"
}
"@

# If any new posts, add to Teams. If new content matches content of previous payload, do not post

$sactPayloadAZ = Get-AutomationVariable $AzAutomationPayloadSACTVariable

# First check if Payload is over 18K (Microsoft Teams limit)

if ($sactPayload.Length -gt $MaxPayloadSize) {  

    # Find payload + append text total length
    $sactTotalLength = $sactPayload.Length + $TrunkAppend.Length

    # Find the overshooting value
    $sactPayloadOverSize = ($MaxPayloadSize - $sactTotalLength)

    # At what point in the original Payload content do we have to split the JSON
    $sactPayloadSplitValue = $sactContent.Length - (-$sactPayloadOverSize)

    # Split the JSON into a new Payload
    $sactNewSplitContent = $sactupdate.Substring(0,$sactPayloadSplitValue)
    
    # Create new truncated payload
    $sactNewContent = $sactNewSplitContent + $TrunkAppend 
    $sactNewPayloadContent = ConvertTo-Json $sactNewContent
    
#Generate truncated payload
          
$sactNewPayload =  @"
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
                    "uri": "https://docs.microsoft.com/en-us/officeupdates/semi-annual-channel-targeted-$($Year)#$($saclink)"
                }
            ]
        },
     ],
    "sections": [
        {
            "facts": [
                {
                    "name": "Version updated:",
                    "value": "$($sactDate)"
                }
                
            ],
            "text": $sactNewPayloadContent
        }
    ],
    "summary": "O365 ProPlus Semi-Annual (targeted)",
    "themeColor": "$($color)",
    "title": "Semi-Annual Channel (targeted) release: $($sacttitle)"
}
"@

If ($sactcontent -ne $sactPayloadAZ) {
Invoke-RestMethod -uri $sactURI -Method Post -body $sactNewPayload -ContentType 'application/json; charset=utf-8'
Set-AutomationVariable -Name $AzAutomationPayloadsactVariable -Value ($sactcontent -as [string])
      }
    Else {
    }
    
    }
Else {
    If ($sactcontent -ne $sactPayloadAZ) {
    Invoke-RestMethod -uri $sactURI -Method Post -body $sactPayload -ContentType 'application/json; charset=utf-8'
    Set-AutomationVariable -Name $AzAutomationPayloadsactVariable -Value ($sactcontent -as [string])
          }
    Else {
    }
    
    }
        }
    }
    Else {
    }


# Semi-Annual channel (SAC)
# -------------------------

# Check if channel is set for checking, if so, do stuff
If ($AzAutomationURISactVariable) {
$sacURI = Get-AutomationVariable -Name $AzAutomationURISacVariable

#Get data
$SACweb = Invoke-RestMethod -Uri $SAC

# Find article's last updated time
$sacdatepattern = '\d{4}-\d{2}-\d{2} \d{2}:\d{2} [AP]M'
$sacLastUpdated = $SACweb | select-string  -Pattern $sacdatepattern -AllMatches | % { $_.Matches } | % { $_.Value }

# Convert match into Date/Time
$SACDate = Get-Date $sacLastUpdated

# Check if updates are newer than time variable, if so, do stuff
                             
	# Calculate time difference
	If (([datetime]$Now - [datetime]$SACDate).TotalHours -le $Hours) {

    # Picking out title
    $sactitlepattern = '(?<=\<h2 id="v.*?\>)(.*?)(?=<\/h2\>)'
    $sactitle = $SACweb | select-string  -Pattern $sactitlepattern -AllMatches | % { $_.Matches } | % { $_.Value } | Select-Object -First 1

    # Tailor the "More info" button by adding suffix to link to right section of the webpage
    $saclinkpattern = '(?<=\<h2.*?\")(.*?)(?=\"\>)'
    $saclink = $SACweb | Select-String -Pattern $saclinkpattern -AllMatches | % { $_.Matches } | % { $_.Value } | Select-Object -Index 1

    # Select latest updates
    $saccontentpattern = '(\<h2 id="v.+?\>)(.|\n)*?(?=(\<h2 id="v.+?\>|<div class.+?\>))'
    $sacupdate = $SACweb | select-string  -Pattern $saccontentpattern -AllMatches | % { $_.Matches } | % { $_.Value } | Select-Object -First 1
    $saccontent = $sacupdate | ConvertTo-Json
#Generate payload
          
$SACPayload =  @"
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
                    "uri": "https://docs.microsoft.com/en-us/officeupdates/semi-annual-channel-$($Year)#$($saclink)"
                }
            ]
        },
     ],
    "sections": [
        {
            "facts": [
                {
                    "name": "Version updated:",
                    "value": "$($sacDate)"
                }
                
            ],
            "text": $saccontent
        }
    ],
    "summary": "O365 ProPlus Semi-Annual",
    "themeColor": "$($color)",
    "title": "Semi-Annual Channel release: $($sactitle)"
}
"@

# If any new posts, add to Teams. If new content matches content of previous payload, do not post

$SACPayloadAZ = Get-AutomationVariable -Name $AzAutomationPayloadSACVariable

# First check if Payload is over 18K (Microsoft Teams limit)

if ($sacPayload.Length -gt $MaxPayloadSize) {  

    # Find payload + append text total length
    $sacTotalLength = $sacPayload.Length + $TrunkAppend.Length

    # Find the overshooting value
    $sacPayloadOverSize = ($MaxPayloadSize - $sacTotalLength)

    # At what point in the original Payload content do we have to split the JSON
    $sacPayloadSplitValue = $sacContent.Length - (-$sacPayloadOverSize)

    # Split the JSON into a new Payload
    $sacNewSplitContent = $sacupdate.Substring(0,$sacPayloadSplitValue)
    
    # Create new truncated payload
    $sacNewContent = $sacNewSplitContent + $TrunkAppend 
    $sacNewPayloadContent = ConvertTo-Json $sacNewContent
    
#Generate truncated payload
          
$sacNewPayload =  @"
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
                    "uri": "https://docs.microsoft.com/en-us/officeupdates/semi-annual-channel-$($Year)#$($saclink)"
                }
            ]
        },
     ],
    "sections": [
        {
            "facts": [
                {
                    "name": "Version updated:",
                    "value": "$($sacDate)"
                }
                
            ],
            "text": $sacNewPayloadContent
        }
    ],
    "summary": "O365 ProPlus Semi-Annual",
    "themeColor": "$($color)",
    "title": "Semi-Annual Channel release: $($sactitle)"
}
"@

If ($saccontent -ne $sacPayloadAZ) {
Invoke-RestMethod -uri $sacURI -Method Post -body $sacNewPayload -ContentType 'application/json; charset=utf-8'
Set-AutomationVariable -Name $AzAutomationPayloadsacVariable -Value ($saccontent -as [string])
      }
    Else {
    }
    
    }
Else {
    If ($saccontent -ne $sacPayloadAZ) {
    Invoke-RestMethod -uri $sacURI -Method Post -body $sacPayload -ContentType 'application/json; charset=utf-8'
    Set-AutomationVariable -Name $AzAutomationPayloadsacVariable -Value ($saccontent -as [string])
          }
    Else {
    }
    
    }
        }
    }
    Else {
    }

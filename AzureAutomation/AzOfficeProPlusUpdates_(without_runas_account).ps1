<#
.SYNOPSIS
  Get Microsoft Office ProPlus channel updates and post to Teams using webhooks and Azure Automation variable assets
.DESCRIPTION
  Script to check Microsoft Office ProPlus channel updates, configured to check last 12 hours (can be adapted as required). Run as a scheduled task, Azure automation etc.
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
  
  Disclaimer: This script is offered "as-is" with no warranty. 
  While the script is tested and working in my environment, it is recommended that you test the script
  in a test environment before using in your production environment.
 
.NOTES
  Version:        1.5
  Author:         Einar Asting (einar@thingsinthe.cloud)
  Creation Date:  Dec 3rd 2019
  Purpose/Change: Simplified, rewrote to avoid runas account, using different Azure cmdlets.
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

# ---------------------

# Setting other script variables
$Now = Get-Date -Format 'yyyy-MM-dd HH:mm'
$Year = Get-Date -Format yyyy
$Monthly = 'https://docs.microsoft.com/en-us/officeupdates/monthly-channel-' +$Year
$SAC = 'https://docs.microsoft.com/en-us/officeupdates/semi-annual-channel-' +$Year
$SACT = 'https://docs.microsoft.com/en-us/officeupdates/semi-annual-channel-targeted-' +$Year

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
    $monthlytitlepattern = '(?<=\<h2.*?\>)(.*?)(?=<\/h2\>)'
    $monthlytitle = $Monthlyweb | select-string  -Pattern $monthlytitlepattern -AllMatches | % { $_.Matches } | % { $_.Value } | Select-Object -First 1

    # Tailor the "More info" button by adding suffix to link to right section of the webpage
    $monthlylinkpattern = '(?<=\<h2.*?\")(.*?)(?=\"\>)'
    $monthlylink = $Monthlyweb | Select-String -Pattern $monthlylinkpattern -AllMatches | % { $_.Matches } | % { $_.Value } | Select-Object -First 1

    # Select latest updates
    $monthlycontentpattern = '(\<h2.+?\>)((.|\n)+?(?=<h2.+?\>))'
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

$MonthlyPayloadAZ = Get-AutomationVariable $AzAutomationPayloadMonthlyVariable

If ($monthlycontent -ne $MonthlyPayloadAZ) {
    Invoke-RestMethod -uri $MonthlyURI -Method Post -body $MonthlyPayload -ContentType 'application/json; charset=utf-8'
    Set-AutomationVariable -Name $AzAutomationPayloadMonthlyVariable -Value ($monthlycontent -as [string])
    }
    Else {
    }
}
Else {
     }
}


# Semi-Annual channel (targeted) (SACT)
# -------------------------------------

# Check if channel is set for checking, if so, do stuff
If ($AzAutomationURISactVariable) {
$sactURI = Get-AutomationVariable -Name $AzAutomationURISactVariable

# Looking for Azure automation variable for SACT channel, if it's not existing, create it
If (!(Get-AzureRmAutomationVariable -AutomationAccountName $AzAutomationAccountName -Name $AzAutomationPayloadSACTVariable -ResourceGroupName $AzResourceGroup -ErrorAction SilentlyContinue)) {
    New-AzureRMAutomationVariable -AutomationAccountName $AzAutomationAccountName -Name $AzAutomationPayloadSACTVariable -ResourceGroupName $AzResourceGroup -Value $null -Encrypted $false
    }
    Else {    
    }

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
    $sacttitlepattern = '(?<=\<h2.*?\>)(.*?)(?=<\/h2\>)'
    $sacttitle = $sactweb | select-string  -Pattern $sacttitlepattern -AllMatches | % { $_.Matches } | % { $_.Value } | Select-Object -First 1

    # Tailor the "More info" button by adding suffix to link to right section of the webpage
    $sactlinkpattern = '(?<=\<h2.*?\")(.*?)(?=\"\>)'
    $sactlink = $sactweb | Select-String -Pattern $sactlinkpattern -AllMatches | % { $_.Matches } | % { $_.Value } | Select-Object -First 1

    # Select latest updates
    $sactcontentpattern = '(\<h2.+?\>)((.|\n)+?(?=<h2.+?\>))'
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

If ($sactcontent -ne $sactPayloadAZ) {
    Invoke-RestMethod -uri $sactURI.Value -Method Post -body $sactPayload -ContentType 'application/json; charset=utf-8'
    Set-AutomationVariable -Name $AzAutomationPayloadSACTVariable -Value ($sactcontent -as [string])
    }
    Else {
    }
}
Else {
     }
   }


# Semi-Annual channel (SAC)
# -------------------------

# Check if channel is set for checking, if so, do stuff
If ($AzAutomationURISactVariable) {
$sactURI = Get-AutomationVariable -Name $AzAutomationURISacVariable

# Looking for Azure automation variable for SAC channel, if it's not existing, create it
If (!(Get-AzureRmAutomationVariable -AutomationAccountName $AzAutomationAccountName -Name $AzAutomationPayloadSACVariable -ResourceGroupName $AzResourceGroup -ErrorAction SilentlyContinue)) {
    New-AzureRmAutomationVariable -AutomationAccountName $AzAutomationAccountName -Name $AzAutomationPayloadSACVariable -ResourceGroupName $AzResourceGroup -Value $null -Encrypted $False
    }
    Else {    
    }

# Setting script variables from Azure Automation variable assets
$sacURI = Get-AzureRmAutomationVariable -AutomationAccountName $AzAutomationAccountName -Name $AzAutomationURISacVariable -ResourceGroupName $AzResourceGroup

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
    $sactitlepattern = '(?<=\<h2.*?\>)(.*?)(?=<\/h2\>)'
    $sactitle = $SACweb | select-string  -Pattern $sactitlepattern -AllMatches | % { $_.Matches } | % { $_.Value } | Select-Object -First 1

    # Tailor the "More info" button by adding suffix to link to right section of the webpage
    $saclinkpattern = '(?<=\<h2.*?\")(.*?)(?=\"\>)'
    $saclink = $SACweb | Select-String -Pattern $saclinkpattern -AllMatches | % { $_.Matches } | % { $_.Value } | Select-Object -First 1

    # Select latest updates
    $saccontentpattern = '(\<h2.+?\>)((.|\n)+?(?=<h2.+?\>))'
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

If ($saccontent -ne $SACPayloadAZ.Value) {
    Invoke-RestMethod -uri $SACURI.Value -Method Post -body $SACPayload -ContentType 'application/json; charset=utf-8'
    Set-AutomationVariable -Name $AzAutomationPayloadSACVariable -Value ($saccontent -as [string])    
    }
    Else {
    }
}
Else {
     }
}
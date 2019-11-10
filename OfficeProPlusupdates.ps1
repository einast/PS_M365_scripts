<#
.SYNOPSIS
  Get Microsoft Office ProPlus channel updates and post to Teams using webhooks
.DESCRIPTION
  Script to check Microsoft Office ProPlus channel updates, configured to check last 12 hours (can be adapted as required). Run as a scheduled task, Azure automation etc.
  Create one or more webhook in Teams (if you want to split the updates into separate channels) and copy the URI(s) to the user variable section below.
  The output is color coded (can be adapted as required). Default is green.
  
  Disclaimer: This script is offered "as-is" with no warranty. 
  While the script is tested and working in my environment, it is recommended that you test the script
  in a test environment before using in your production environment.
 
.NOTES
  Version:        1.0
  Author:         Einar Asting (einar@thingsinthe.cloud)
  Creation Date:  Nov 9th 2019
  Purpose/Change: Initial version
.LINK
  https://github.com/einast/PS_M365_scripts
#>

# User defined variables
# ----------------------
# If you want to check Monthly Channel, Semi-Annual Channel Targeted (SACT) and/or Semi-Annual Channel, add your Teams URI in the variables fields. 
# Leave the ones you don't want to check blank.
 
$MonthlyURI = ''
$SACTURI = ''
$SACURI = ''
$Hours = '12'
$Color = '00ff00' # Green

# Setting other script variables
$Now = Get-Date -Format 'yyyy-MM-dd HH:mm'
$Year = Get-Date -Format yyyy
$Monthly = 'https://docs.microsoft.com/en-us/officeupdates/monthly-channel-' +$Year
$SAC = 'https://docs.microsoft.com/en-us/officeupdates/semi-annual-channel-' +$Year
$SACT = 'https://docs.microsoft.com/en-us/officeupdates/semi-annual-channel-targeted-' +$Year

# Monthly channel
# ---------------

If ($MonthlyURI) {
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

    # Select latest updates
    $monthlycontentpattern = '(\<h2.+?\>)((.|\n)+?(?=<h2.+?\>))'
    $monthlyupdate = $Monthlyweb | select-string  -Pattern $contentpattern -AllMatches | % { $_.Matches } | % { $_.Value } | Select-Object -First 1
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
                    "uri": "https://docs.microsoft.com/en-us/officeupdates/monthly-channel-$($Year)"
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

# If any new posts, add to Teams
Invoke-RestMethod -uri $MonthlyURI -Method Post -body $MonthlyPayload -ContentType 'application/json; charset=utf-8'
}
Else {
     }
}

# Semi-annual channel targeted (SACT)
# -----------------------------------

If ($SACTURI) {
#Get data
$SACTweb = Invoke-RestMethod -Uri $SACT

# Find article's last updated time
$sactdatepattern = '\d{4}-\d{2}-\d{2} \d{2}:\d{2} [AP]M'
$sactLastUpdated = $SACTweb | select-string  -Pattern $sactdatepattern -AllMatches | % { $_.Matches } | % { $_.Value }

# Convert match into Date/Time
$SACTDate = Get-Date $sactLastUpdated

# Check if updates are newer than time variable, if so, do stuff
                             
	# Calculate time difference
	If (([datetime]$Now - [datetime]$SACTDate).TotalHours -le $Hours) {

    # Picking out title
    $sacttitlepattern = '(?<=\<h2.*?\>)(.*?)(?=<\/h2\>)'
    $sacttitle = $SACTweb | select-string  -Pattern $sacttitlepattern -AllMatches | % { $_.Matches } | % { $_.Value } | Select-Object -First 1

    # Select latest updates
    $sactcontentpattern = '(\<h2.+?\>)((.|\n)+?(?=<h2.+?\>))'
    $sactupdate = $SACTweb | select-string  -Pattern $sactcontentpattern -AllMatches | % { $_.Matches } | % { $_.Value } | Select-Object -First 1
    $sactcontent = $sactupdate | ConvertTo-Json

#Generate payload
          
$SACTPayload =  @"
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
                    "uri": "https://docs.microsoft.com/en-us/officeupdates/semi-annual-channel-targeted-$($Year)"
                }
            ]
        },
     ],
    "sections": [
        {
            "facts": [
                {
                    "name": "Version updated:",
                    "value": "$($SACTDate)"
                }
                
            ],
            "text": $sactcontent
        }
    ],
    "summary": "O365 ProPlus SACT",
    "themeColor": "$($color)",
    "title": "Semi-Annual Channel (targeted) release: $($sacttitle)"
}
"@

# If any new posts, add to Teams
Invoke-RestMethod -uri $SACTURI -Method Post -body $SACTPayload -ContentType 'application/json; charset=utf-8'
}
Else {
     }
}

# Semi-annual channel (SAC)
# -------------------------

If ($SACURI) {
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
                    "uri": "https://docs.microsoft.com/en-us/officeupdates/semi-annual-channel-$($Year)"
                }
            ]
        },
     ],
    "sections": [
        {
            "facts": [
                {
                    "name": "Version updated:",
                    "value": "$($SACDate)"
                }
                
            ],
            "text": $saccontent
        }
    ],
    "summary": "O365 ProPlus SAC",
    "themeColor": "$($color)",
    "title": "Semi-Annual Channel release: $($sactitle)"
}
"@

# If any new posts, add to Teams
Invoke-RestMethod -uri $SACURI -Method Post -body $SACPayload -ContentType 'application/json; charset=utf-8'
}
Else {
     }
}
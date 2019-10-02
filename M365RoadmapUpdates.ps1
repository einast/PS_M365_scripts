<#
.SYNOPSIS
  Get Microsoft 365 Roadmap update and post to Teams using webhooks
.DESCRIPTION
  Script to check Microsoft 365 Roadmap RSS feed, configured to check last 24 hours (can be adapted as required). Run as a scheduled task.
  Create a webhook in Teams and copy the URI to the variable section below.
  The output will be color coded (can be adapted as required) according to Status of the Feature - Red = In development, Yellow = Rolling out, Green = Launched
  Replace the variables with your own where feasible
  
  Disclaimer: This script is offered "as-is" with no warranty. 
  While the script is tested and working in my environment, it is recommended that you test the script
  in a test environment before using in your production environment.
 
.NOTES
  Version:        1.0
  Author:         Einar Asting (einar@asting.net)
  Creation Date:  Oct 2nd 2019
  Purpose/Change: Initial version
.LINK
  https://github.com/einast/PS_M365_scripts
#>

# Variables

    # Your Teams Webhook URI
    $URI = 'URI to Teams webhook'

    # M365 Roadmap RSS URL
    $ParseFrom = 'https://www.microsoft.com/en-us/microsoft-365/RoadmapFeatureRSS'

    # Temp file location
    $TempFile = 'C:\Folder\test.xml'

    # Last x hours to check for updates
    $Hours = '24'

    # Read current date and time
    $Now = Get-Date

# Get content
Invoke-WebRequest -Uri $ParseFrom -OutFile $TempFile -ErrorAction Stop

# Parse
[xml]$Content = Get-Content $TempFile

  $Feed = $Content.rss.channel

  ForEach ($msg in $Feed.Item){

                #Set the color line of the card according to the Status of the environment
                if ($msg.category[0] -eq "In development")
                    {
                    $color = "ff0000"
                    }
                    else
                        {
                            if ($msg.category[0] -eq "Rolling out")
                                {
                                    $color = "ffff00"
                                    }
                                    else
                                        {
                                            $color = "00cc00"
                                            }
                                 }
    If (($Now - [datetime]$msg.pubDate).TotalHours -le $Hours) {
          
                $Payload = ConvertTo-Json -Depth 4 @{
                text = 'Roadmap updates past ' +  $($Hours) + ' hours - Feature ID: ' + $($msg.guid.'#text')
                themeColor = $color
                sections = @(
                            @{
                        title = $msg.title
                        facts = @(
                            @{
                            name = 'Environment'
                            value = $msg.category[1]
                            },
                            @{
                            name = 'Status'
                            value = $msg.category[0]
                            },
                            @{
                            name = 'Description'
                            value = $msg.description
                            },
                            @{
                            name = 'Published date'
                            value = $msg.pubDate
                            }
                            @{
                            name = 'Link'
                            value = '<a href=' + $msg.link + '>'  + $msg.link + '</a>'
                            }
                        )
                    }
                )
            }
            #Convert to UTF8
            $Payload = ([System.Text.Encoding]::UTF8.GetBytes($Payload))
            Invoke-webrequest -URI $URI -Method POST -Body $Payload
        }
     }

# Clean up
Remove-Item $TempFile -Force
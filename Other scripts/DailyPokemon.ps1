<#
.SYNOPSIS
 Get random daily Pokemon, write to JSON file on a webhost. Post the result to a DAKboard
.DESCRIPTION
  Script to get random daily Pokemon, save as JSON on a webhost.
  Stats are pulled from Pokeapi.co and image from pokeres.bastionbot.org, and uploaded to a webhost via FTP
  The JSON will be parsed on a DAKboard (https://dakboard.com/) at home to present a new Pokemon every day.

  The script use the PSFTP module (Azure Automation module: https://www.powershellgallery.com/packages/PSFTP/1.7.0.4
  alternatively on-premise: https://gallery.technet.microsoft.com/scriptcenter/PowerShell-FTP-Client-db6fe0cb).
  
  This script runs every 24 hours at midnight in an Azure Automation runbook, but can be adapted to run on-premise if needed.

  Disclaimer: This script is offered "as-is" with no warranty.
  While the script is tested and working in my environment, it is recommended that you test the script
  in a test environment before using in your production environment.
  
.NOTES
  Version:        1.0
  Author:         Einar Asting (einar@thingsinthe.cloud)
  Creation Date:  Mar 7th 2020
  Purpose/Change: Initial version
.LINK
  https://github.com/einast/PS_M365_scripts
#>

# Setting variables
# -------------------------------------------------------------------

$jsonfile = Join-Path $env:TEMP "pokemon.json" # Temp location for storing the JSON file

#FTP server
$FtpServer = Get-AutomationVariable -Name AzFTPServer
$FtpUser = Get-AutomationVariable -Name AzFTPUser
$FtpPassword = Get-AutomationVariable -Name AzFTPPassword
$userPassword = ConvertTo-SecureString -String $FtpPassword -AsPlainText -Force
$FtpCredentials = New-Object System.Management.Automation.PSCredential ($FtpUser, $userPassword)

# -------------------------------------------------------------------

# Getting Pokemon stats and image
$Id = Get-Random -Minimum 1 -Maximum 807
$pokemonApiUrl = "https://pokeapi.co/api/v2/pokemon-species/$Id"
$pokemonApiStats = "https://pokeapi.co/api/v2/pokemon/$Id"
$pokemonImage = "https://pokeres.bastionbot.org/images/pokemon/$Id.png"
$pokemon = Invoke-RestMethod -Uri $pokemonApiUrl
$name = $pokemon.name.substring(0,1).toupper()+$pokemon.name.substring(1).tolower()
$pokemonstats = Invoke-RestMethod -Uri $pokemonApiStats

# Replace the with your own text
$html = @"
{
"name":"$($name)",
"hp":"$($pokemonstats.stats.base_stat[5])",
"attack":"$($pokemonstats.stats.base_stat[4])",
"defense":"$($pokemonstats.stats.base_stat[3])",
"atkspe":"$($pokemonstats.stats.base_stat[2])",
"defspe":"$($pokemonstats.stats.base_stat[1])",
"speed":"$($pokemonstats.stats.base_stat[0])",
"image":"$($pokemonImage)"
}
"@

# Write a new HTML fil (have to be ASCII or else the JSON parsing fails at my webhost)
$html | Out-File -Encoding ASCII $jsonfile

# Upload to webhost using PSFTP
Set-FTPConnection -Credentials $FtpCredentials -Server $FtpServer -UsePassive 
Add-FTPItem -Path "/public_html/" -LocalPath $jsonfile -Overwrite
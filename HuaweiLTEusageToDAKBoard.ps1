<#
.SYNOPSIS
 Get LTE/4G bandwidth usage from Huawei B818 router, write to JSON file on a webhost. Post the result to DAKboard
.DESCRIPTION
  Script to get LTE/4G bandwidth usage from Huawei B818 router, save as JSON on a webhost.
  The JSON will be parsed on a DAKboard at home to monitor the bandwidth cap enforced by the ISP

  The script use the PSFTP module (https://gallery.technet.microsoft.com/scriptcenter/PowerShell-FTP-Client-db6fe0cb).
   
  Disclaimer: This script is offered "as-is" with no warranty.
  While the script is tested and working in my environment, it is recommended that you test the script
  in a test environment before using in your production environment.
  
.NOTES
  Version:        1.0
  Author:         Einar Asting (einar@asting.net)
  Creation Date:  Jan 7th 2020
  Purpose/Change: Initial version
.LINK
  https://github.com/
#>  

Import-Module PSFTP
# Setting variables
# -------------------------------------------------------------------
# NB! Also review the here-string starting with $json (line 61) for customizing your JSON file!


$jsonfile = 'C:\whatever\folder' # Temp local folder for storing the JSON file
$routerip = '192.168.0.1' # Your Huawei router IP
$icon = 'https://www.link.to/icon.png' # If you want a custom icon shown on the DAKboard, add the URL here

#FTP server
$FtpServer = 'ftp://ftp.server.you.use.com'
$FtpUser = 'ftpserverpassword'
$FTPPasswd = 'ftpserverpassword'
$FTPRemotePath = '/remotefolderonftp/'


# For setting Norwegian month, either comment out or replace with your own locale
$LocaleNO = New-Object System.Globalization.CultureInfo("nb-NO")
$month = (Get-Date).tostring("MMMM",$LocaleNO)

# -------------------------------------------------------------------

# Request data usage from Huawei B818 router using the builtin API URL
$response = Invoke-WebRequest -Uri http://$routerip/api/monitoring/month_statistics -Method Get

# Use regex to pick out the usage data, add incoming/outgoing for total data usage
$pattern1 = '(?<=\<CurrentMonthDownload\>)(.*?)(?=<\/Current)'
$usage1 = $response.Content | select-string  -Pattern $pattern1 -AllMatches | % { $_.Matches } | % { $_.Value } | Select-Object -First 1 | measure -Sum

$pattern2 = '(?<=\<CurrentMonthUpload\>)(.*?)(?=<\/Current)'
$usage2 = $response.Content | select-string  -Pattern $pattern2 -AllMatches | % { $_.Matches } | % { $_.Value } | Select-Object -First 1 | measure -Sum

$number = ($usage1.Sum + $usage2.Sum)
$sum = (($number)/1GB)
$data = '{0:f2}' -f $sum

# Replace the with your own text
$json = @"
{
"month":"LTE usage for $($month)",
"amount":"$($data) / 1024 GB",
"image":"$($icon)"
}
"@

# Remove the last version of the local JSON file if existing
if (Test-Path $jsonfile){Remove-Item $jsonfile -Force}

# Write new JSON file (keep encoding as ASCII, if not my webhost writes a malfunctioning JSON file (000webhost.com)
$json | Out-File -Encoding ASCII $jsonfile

# Upload JSON using PSFTP
$FtpPassword = ConvertTo-SecureString $FTPPasswd -AsPlainText -Force
$FtpCredentials = New-Object System.Management.Automation.PSCredential ($FtpUser, $FtpPassword)
Set-FTPConnection -Credentials $FtpCredentials -Server $FtpServer -UsePassive 
Add-FTPItem -Path $FTPRemotePath -LocalPath $jsonfile -Overwrite

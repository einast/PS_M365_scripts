<# 
.SYNOPSIS
  Offboard Microsoft 365 users
.DESCRIPTION
  This script was created to offboard a large amount of users simultaneously 
  using a CSV file for a specific customer case.
  
  The script will:
  -Block new sign-ins
  -Disconnect any active sessions
  -Cancel meetings organized by the user
  -Remove user from distribution list
  -Convert the user mailbox to a shared mailbox
  -Set auto-reply for the mailbox
  -Check if the CSV has defined an external email for the user, if so, create a contact object
  -Remove all assigned Microsoft 365 licenses

  The script will output the result to the console, comment/remove the Write-Host lines if required.

  Usage: Populate a CSV file with 3 columns, for example (remember to keep the header names identical to the ones below):

  Name,internal.email,external.email
  Roy Rogers,roy.rogers@acmeinc.com,royr@gmail.com
  Ola Nordmann,ola.nordmann@acmeinc.com,olanord@hotmail.com

  Requirements: Powershell modules AzureAD and ExchangeOnlineManagement

  The script will connect to Azure AD and Microsoft 365 and parse the CSV file. For each line in the CSV, the script will block sign-ins, disconnect any active sessions, cancel meetings organized
  by the user, remove the user from any Distribution Groups, convert the mailbox to a shared mailbox, create a new mail contact, and finally remove any assigned Microsoft 365 licenses.

  Disclaimer: This script is offered "as-is" with no warranty. 
  While the script is tested and working in my environment, it is recommended that you test the script
  in a test environment before using in your production environment.
.NOTES
  Version:        0.9
  Author:         Einar Asting (einar@thingsinthe.cloud)
  Creation Date:  Nov 3rd 2020
  Purpose/Change: Minor edits
.LINK
  https://github.com/einast/PS_M365_scripts
#>

# Set the path to your CSV file
$CSVfile = 'C:\Temp\test3.csv'

# Required to be able to run commands remotely against the cloud
Set-ExecutionPolicy RemoteSigned

# Import necessary modules
Import-Module -Name AzureAD
Import-Module -Name ExchangeOnlineManagement

# Connect to Azure AD and Exchange Online (two login boxes will appear)
Connect-AzureAD
Connect-ExchangeOnline

# Parse users from CSV file (adapt path as needed)
$users = Import-csv $CSVfile

# Loop through each user
foreach ($user in $users) {

    # Get-AzureADUser
    $AADuser = Get-AzureADUser | Where-Object {$_.UserPrincipalName -like $user.'internal.email'}

    # Block sign-ins
    Set-AzureADUser -ObjectId $AADuser.ObjectId -AccountEnabled $false
    Write-Host "$($user.'internal.email') Microsoft 365 sign-in blocked"

    # Disconnect any active sessions
    Revoke-AzureADUserAllRefreshToken -ObjectId $AADuser.ObjectId
    Write-Host "$($user.'internal.email') active Microsoft 365 sessions disconnected"

    # Cancel meetings organized by this user
    Remove-CalendarEvents -Identity $user.'internal.email' -CancelOrganizedMeetings -QueryWindowInDays 120 -confirm:$False
    Write-Host "$($user.'internal.email') meetings cancelled"

    # Remove user from DistributionGroups
    $DistributionGroups= Get-DistributionGroup | where { (Get-DistributionGroupMember $_.Name | foreach {$_.PrimarySmtpAddress}) -contains $user.'internal.email'}

        foreach( $dg in $DistributionGroups)
	        {
	        Remove-DistributionGroupMember $dg.name -Member $user.'internal.email' -Confirm:$false
	        }
    Write-Host "$($user.'internal.email') removed from Distribution groups"

    # Convert mailbox to shared mailbox
    Set-Mailbox $user.'internal.email' -Type “Shared” 

    # Set mailbox auto-reply message
    Set-MailboxAutoReplyConfiguration -Identity $user.'internal.email' -AutoReplyState Enabled -InternalMessage "The user no longer work here, please use $($user.'external.email')" -ExternalMessage "The user no longer work here, please use $($user.'external.email')"
    Write-Host "$($user.'internal.email') automatic reply is set"

    # Check if user has populated external.email field, if so, create a new mail contact
    if ($user.'external.email') {
        New-MailContact -Name $user.'Name' -ExternalEmailAddress $user.'external.email' | Out-Null
        Write-Host "$($user.'internal.email') external contact object created"
    }
    else {
        Write-Host "$($user.'internal.email') external contact email not found, skipping..."
    }

    # Remove assigned Microsoft 365 licenses
    $licenses = Get-AzureADUser -ObjectId $AADuser.ObjectId | select AssignedLicenses
    $licenses.AssignedLicenses.skuID | foreach { 
                    $body = @{
                                addLicenses = @()
                                removeLicenses= @($_)
                            }
        Set-AzureADUserLicense -ObjectId $AADuser.ObjectId -AssignedLicenses $body }
        Write-Host "$($user.'internal.email') Microsoft 365 licenses removed`n"
}
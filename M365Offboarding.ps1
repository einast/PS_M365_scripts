<# 
.SYNOPSIS
  Offboard Microsoft 365 users
.DESCRIPTION
  This script was created to offboard a large amount of users simultaneously 
  using a CSV file for a customer.

  PLEASE NOTE: It should be reviewed and adapted as needed, as it served a specific use case. But it might
  serve as a starting point for your own cases.
  
  The script will:
  -Block new sign-ins
  -Cancel meetings organized by the user
  -Remove user from distribution list
  -Convert the user mailbox to a shared mailbox
  -Prefix the shared mailbox for easier readability (like Quit - Roy Rogers)
  -Set SMTP addresses, ADAPT AS NEEDED
  -Set auto-reply for the mailbox
  -Check if the CSV has defined an external email for the user, if so, create a contact object
  -Remove all assigned Microsoft 365 licenses

  The script will output the result to the console, comment/remove the Write-Host lines if required.

  Usage: Populate a CSV file with 2 columns, for example (remember to keep the header names identical to the ones below):
  Leave externalemail column blank for users (or all) if you don't want to set up contact objects.

  internal.email,external.email
  roy.rogers@acmeinc.com,royr@gmaail.com
  ola.nordmann@acmeinc.com,
  john.johnson@acmeinc,jj@gmaail.com

  Requirements: Powershell modules MSOnline and ExchangeOnlineManagement

  The script will connect to Azure AD and Microsoft 365 and parse the CSV file. For each line in the CSV, the script will block sign-ins, disconnect any active sessions, cancel meetings organized
  by the user, remove the user from any Distribution Groups, convert the mailbox to a shared mailbox, create a new mail contact, and finally remove any assigned Microsoft 365 licenses.

  Disclaimer: This script is offered "as-is" with no warranty. 
  While the script is tested and working in my environment, it is recommended that you test the script
  in a test environment before using in your production environment.
.NOTES
  Version:        0.9
  Author:         Einar Asting (einar@thingsinthe.cloud)
  Creation Date:  Dec 8th 2020
  Purpose/Change: Changes after testing further
.LINK
  https://github.com/einast/PS_M365_scripts
#>

####################################################### User defined variables #######################################################

$CSVfile = 'C:\Temp\test3.csv'                # Set the path to your CSV file
$AutoReplyMessage = "Hi and thanks for your email,<br><br>The person you are trying to reach is no longer working here. For further information, please contact our customer care at help@acme.com<br><br>Best regards<br>
<b>Acme Inc</b>"                              # Auto-reply message, use HTML tags for richer text
$CustomerDomain = '@acmeinc.com'              # For setting new SMTP addresses
$SharedMbxPrefix = 'Quit - '                  # What to prefix the shared mailbox name with, for easier readability         

# Optional for removing licenses by removing users from dynamic license group
$DynamicLicenseGroup = '' # If your tenant use dynamic group for allocating licenses, leave blank if not

######################################################################################################################################

# Required to be able to run commands remotely against the cloud
Set-ExecutionPolicy RemoteSigned

# Import necessary modules
Import-Module -Name ExchangeOnlineManagement
Import-Module -Name MSOnline

# Connect to Exchange Online and MSOnline (two login boxes will appear)
#Connect-AzureAD
Connect-ExchangeOnline
Connect-MsolService

# Parse users from CSV file (adapt path as needed)
$users = Import-csv $CSVfile


# Loop through each user
foreach ($user in $users) {

  # Check if user has a mailbox, if not exit the script
  $mbx = Get-Mailbox -Identity $user.internalemail -ErrorAction SilentlyContinue
  
  if ($mbx -ne $null) {
  
      # Get User information
      $AADUser = (Get-MsolUser -UserPrincipalName $user.internalemail).UserPrincipalName
      $AADUserFullName = (Get-MsolUser -UserPrincipalName $user.internalemail).DisplayName
      $AADUserObjectId = (Get-MsolUser -UserPrincipalName $user.internalemail).ObjectId
  
      # Block sign-ins
      Set-MsolUser -UserPrincipalName $user.internalemail -BlockCredential $true
      Write-Host "$($user.internalemail) Microsoft 365 sign-in blocked"
  
      # Disconnect any active sessions
      #Revoke-AzureADUserAllRefreshToken -ObjectId $AADUser2.objectId
      #Write-Host "$($AADuser.UserPrincipalName) active Microsoft 365 sessions disconnected"
  
      # Cancel meetings organized by this user
      Remove-CalendarEvents -Identity $user.internalemail -CancelOrganizedMeetings -QueryWindowInDays 120 -confirm:$False
      Write-Host "$($user.internalemail) meetings cancelled"
  
      # Remove user from DistributionGroups
      $DistributionGroups= Get-DistributionGroup | where { (Get-DistributionGroupMember $_.Name | foreach {$_.PrimarySmtpAddress}) -contains $user.internalemail}
  
          foreach( $dg in $DistributionGroups)
            {
            Remove-DistributionGroupMember $dg.name -Member $user.internalemail -Confirm:$false
            }
      Write-Host "$($user.internalemail) removed from Distribution groups"
  
      # Preparing to convert user mailbox to shared mailbox
    
          # Selecting initials to be used when setting aliases
          # Clean out variables
          $firstname = ''
          $lastname = ''
          
          # Pick out first- and lastname, convert Norwegian characters if found
          $Firstname = (Get-MsolUser -UserPrincipalName $user.internalemail).Firstname -replace 'æ','ae' -replace 'ø','o' -replace 'å','a'
          $Lastname = (Get-MsolUser -UserPrincipalName $user.internalemail).Lastname -replace 'æ','ae' -replace 'ø','o' -replace 'å','a'
          
          # Pick first letter of firstname and two first letters of lastname
          $Firstname_short = $FirstName.substring(0,1)
          $Lastname_short = $Lastname.substring(0,2)
          
          # Selecting all given and surnames, replacing space with . for setting smtp aliases
          $AliasFirstName = $Firstname.split(" ") -join "."
          $AliasLastName = $Lastname.split(" ") -join "."
        
      # Check if user has populated external.email field, if so:
      # 	1. create a new mail contact
      #	  2. set up forwarding on the shared mailbox  
  
      if ($user.'external.email') {
          New-MailContact -Name $AADuserFullName -ExternalEmailAddress $user.'external.email' | Out-Null
      #Set-Mailbox -Identity $AADuser -ForwardingSMTPAddress $user.'external.email'
      
          Write-Host "$($user.internalemail) external contact object created"
      }
      else {
      Write-Host "$($user.internalemail) external contact email not found, skipping..."
      }
  
      # Set auto-reply
      Set-MailboxAutoReplyConfiguration -Identity $AADuser -AutoReplyState Enabled -InternalMessage $AutoReplyMessage -ExternalMessage $AutoReplyMessage
      Write-Host "$($user.internalemail) auto-reply enabled"
  
      # Convert mailbox to shared and set new primary SMTP and alias
      Set-Mailbox -Identity $AADuser -EmailAddresses "SMTP:$($AliasFirstName + '.')$($AliasLastName)$($CustomerDomain)","$($firstname_short)$($lastname_short)$($CustomerDomain)" -DisplayName "$($SharedMbxPrefix) - $($AADUserFullName)"
      Set-Mailbox -Identity $AADuser -Type "Shared"
      Start-Sleep -Seconds 10
      Write-Host "$($user.internalemail) user mailbox converted to shared mailbox"
  
  if (!$DynamicLicenseGroup)
  {
  # This part removes licenses. If you use dynamic groups for licensing, we need to remove the user from that group, if not, manual removal will fail.
  # Remove user assigned Microsoft 365 licenses
  $licenses = Get-AzureADUser -ObjectId $AADuser.ObjectId | select AssignedLicenses
  $licenses.AssignedLicenses.skuID | foreach { 
                  $body = @{
                              addLicenses = @()
                              removeLicenses= @($_)
                          }
      Set-AzureADUserLicense -ObjectId $AADuser.ObjectId -AssignedLicenses $body }
  }
  else {
  # Remove licenses by removing user from dynamic license group
  Remove-MsolGroupMember -GroupObjectId $DynamicLicenseGroup -GroupMemberType User -GroupmemberObjectId $AADUserObjectId
  Write-Host "$($user.internalemail) Office365 licenses removed"
  }
  Write-Host "$($user.'internal.email') Microsoft 365 licenses removed`n"
}
  Else {
  Write-Host "$($user.internalemail) does not have a mailbox. Please review manually" -ForegroundColor White -BackgroundColor Red
      }
  }
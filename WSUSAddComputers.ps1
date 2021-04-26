<#
.SYNOPSIS
  WSUS add computers to computer groups, handle Defender updates outside WSUS
.DESCRIPTION
  WSUS script for setting parameters in registry and handle Microsoft Defender updates outside of WSUS

  This script was created to add multiple computers from different domains and workgroups to a WSUS
  server. There are other ways to acheive this, like GPOs, but in this case, we needed something
  simple that works across the entire environment.

  Windows Defender updates handling is added. To avoid any WSUS delays, these should be downloaded
  directly from Microsoft. This script adds a scheduled task to do this every 6 hours.

  In this script, two deployment groups are set up, one for pilot and one for broad deployment. The
  script will prompt for which group you would like to add the computer to.
  Also some other options are set, adapt as required.
  
  Disclaimer: This script is offered "as-is" with no warranty. 
  While the script is tested and working in my environment, it is recommended that you test the script
  in a test environment before using in your production environment.
 
.NOTES
  Version:        1.0
  Author:         Einar Asting (einar@thingsinthe.cloud)
  Creation Date:  Apr 26th 2021
  Purpose/Change: First version
.LINK
  https://github.com/einast/PS_M365_scripts
#>

# User defined parameters
$pilotdeploy = "1. Pilot deployment"
$broaddeploy = "2. Broad deployment"
$wsusserver = "http://wsus.server:8530"
$wsusstatusserver = "http://wsus.server:8530"
$auoptions = "3" 
	<# AU options:
	1: Keep my computer up to date is disabled in Automatic Updates.
	2: Notify of download and installation.
	3: Automatically download and notify of installation.
	4: Automatically download and scheduled installation.
	#>

$title = "Add computer to WSUS server"
$message = "Which WSUS group do you want to add this computer to?"

$pilot = New-Object System.Management.Automation.Host.ChoiceDescription "&Pilot deployment", `
    "Adds this computer to the pilot deployment group in WSUS, will get updates first."
$broad = New-Object System.Management.Automation.Host.ChoiceDescription "&Broad deployment.", `
    "Adds this computer to the broad deployment group in WSUS, for more critical computers. Updates will come after pilot"
$options = [System.Management.Automation.Host.ChoiceDescription[]]($pilot, $broad)

$selectedGroup = $Host.ui.PromptForChoice($title, $message, $options, 0) 

# Add/modify registry keys to point to WSUS server and set update options based on which WSUS group the computer should belong to.
switch($selectedGroup)
{
    0 {
    Set-ItemProperty  -path "HKLM:\\Software\Policies\Microsoft\Windows\WindowsUpdate" -name "DisableOSUpgrade" -value "1"
    Set-ItemProperty  -path "HKLM:\\Software\Policies\Microsoft\Windows\WindowsUpdate" -name "TargetGroup" -value $pilotdeploy
    Set-ItemProperty  -path "HKLM:\\Software\Policies\Microsoft\Windows\WindowsUpdate" -name "TargetGroupEnabled" -value "1"
    Set-ItemProperty  -path "HKLM:\\Software\Policies\Microsoft\Windows\WindowsUpdate" -name "WUServer" -value $wsusserver
    Set-ItemProperty  -path "HKLM:\\Software\Policies\Microsoft\Windows\WindowsUpdate" -name "WUStatusServer" -value $wsusstatusserver
    Set-ItemProperty  -path "HKLM:\\Software\Policies\Microsoft\Windows\WindowsUpdate" -name "SetProxyBehaviorForUpdateDetection" -value "0"

    Set-ItemProperty  -path "HKLM:\\Software\Policies\Microsoft\Windows\WindowsUpdate\AU" -name "UseWUServer" -value "1"
    Set-ItemProperty  -path "HKLM:\\Software\Policies\Microsoft\Windows\WindowsUpdate\AU" -name "NoAutoUpdate" -value "0"
    Set-ItemProperty  -path "HKLM:\\Software\Policies\Microsoft\Windows\WindowsUpdate\AU" -name "NoAUShutdownOption" -value "1"
    Set-ItemProperty  -path "HKLM:\\Software\Policies\Microsoft\Windows\WindowsUpdate\AU" -name "NoAutoRebootWithLoggedOnUsers" -value "1"
    Set-ItemProperty  -path "HKLM:\\Software\Policies\Microsoft\Windows\WindowsUpdate\AU" -name "AUOptions" -value $auoptions
    Set-ItemProperty  -path "HKLM:\\Software\Policies\Microsoft\Windows\WindowsUpdate\AU" -name "ScheduledInstallDay" -value "0"
    Set-ItemProperty  -path "HKLM:\\Software\Policies\Microsoft\Windows\WindowsUpdate\AU" -name "ScheduledInstallTime" -value "3"
    }

    1 {
    Set-ItemProperty  -path "HKLM:\\Software\Policies\Microsoft\Windows\WindowsUpdate" -name "DisableOSUpgrade" -value "1"
    Set-ItemProperty  -path "HKLM:\\Software\Policies\Microsoft\Windows\WindowsUpdate" -name "TargetGroup" -value $broaddeploy
    Set-ItemProperty  -path "HKLM:\\Software\Policies\Microsoft\Windows\WindowsUpdate" -name "TargetGroupEnabled" -value "1"
    Set-ItemProperty  -path "HKLM:\\Software\Policies\Microsoft\Windows\WindowsUpdate" -name "WUServer" -value $wsusserver
    Set-ItemProperty  -path "HKLM:\\Software\Policies\Microsoft\Windows\WindowsUpdate" -name "WUStatusServer" -value $wsusstatusserver
    Set-ItemProperty  -path "HKLM:\\Software\Policies\Microsoft\Windows\WindowsUpdate" -name "SetProxyBehaviorForUpdateDetection" -value "0"

    Set-ItemProperty  -path "HKLM:\\Software\Policies\Microsoft\Windows\WindowsUpdate\AU" -name "UseWUServer" -value "1"
    Set-ItemProperty  -path "HKLM:\\Software\Policies\Microsoft\Windows\WindowsUpdate\AU" -name "NoAutoUpdate" -value "0"
    Set-ItemProperty  -path "HKLM:\\Software\Policies\Microsoft\Windows\WindowsUpdate\AU" -name "NoAUShutdownOption" -value "1"
    Set-ItemProperty  -path "HKLM:\\Software\Policies\Microsoft\Windows\WindowsUpdate\AU" -name "NoAutoRebootWithLoggedOnUsers" -value "1"
    Set-ItemProperty  -path "HKLM:\\Software\Policies\Microsoft\Windows\WindowsUpdate\AU" -name "AUOptions" -value $auoptions
    Set-ItemProperty  -path "HKLM:\\Software\Policies\Microsoft\Windows\WindowsUpdate\AU" -name "ScheduledInstallDay" -value "0"
    Set-ItemProperty  -path "HKLM:\\Software\Policies\Microsoft\Windows\WindowsUpdate\AU" -name "ScheduledInstallTime" -value "3"
    }
}

# Adding Microsoft Defender updates outside WSUS for faster refresh
    # Check for existing path, create if not existing
    $path = "C:\Scripts"
    If(!(test-path $path))
    {
          New-Item -ItemType Directory -Force -Path $path
    }

# Create script for updating Microsoft Defender in C:\Scripts
New-Item C:\Scripts\DefenderUpdate.ps1
Set-Content C:\Scripts\DefenderUpdate.ps1 '# Update Microsoft Defender signature files from Microsoft Update
Update-MpSignature -UpdateSource MicrosoftUpdateServer'

# Create a new scheduled task to check for Windows Defender updates
    # Create a new task action
    $taskAction = New-ScheduledTaskAction `
        -Execute 'powershell.exe' `
        -Argument '-File C:\Scripts\DefenderUpdates.ps1'
    $taskAction

    # Create a new trigger (Every 6 hours)
    $taskTrigger = New-ScheduledTaskTrigger -Once -At (Get-Date) -RepetitionInterval (New-TimeSpan -Hours 6)

    # Register the new PowerShell scheduled task
    # The name of your scheduled task.
    $taskName = "DefenderUpdate"

    # Describe the scheduled task.
    $description = "Update Microsoft Defender every 6 hours"

    # Register the scheduled task
    Register-ScheduledTask `
        -TaskName $taskName `
        -Action $taskAction `
        -Trigger $taskTrigger `
        -Description $description

# Try to force the computer to check in to the WSUS server (credit to https://pleasework.robbievance.net/howto-force-really-wsus-clients-to-check-in-on-demand/)
$updateSession = new-object -com "Microsoft.Update.Session"; $updates=$updateSession.CreateupdateSearcher().Search($criteria).Updates
cmd.exe --% /c wuauclt /reportnow

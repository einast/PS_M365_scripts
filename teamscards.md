### Scripts for posting updates from Microsoft Message Center (both incidents and announcements), and updates from the Office 365 roadmap, to Teams via webhooks ###

  | Script                                                       | Description                                                  |
  | ------------------------------------------------------------ | ------------------------------------------------------------ |
  | [M365 Health status](https://github.com/einast/PS_M365_scripts/blob/master/M365HealthStatus.ps1) | Display M365 health events relevant for your tenant          |
  | [M365 Message Center updates](https://github.com/einast/PS_M365_scripts/blob/master/M365MessageCenterUpdates.v2.ps1) | Display relevant messages for your tenant from Message Center |
  | [M365 Roadmap updates](https://github.com/einast/PS_M365_scripts/blob/master/M365RoadmapUpdates.ps1) | Display roadmap updates using RSS feed                            |

#### Pre-requisites ####

- Health status and Message Center scripts require app registration in Azure. There is a guide [here](https://evotec.xyz/preparing-azure-app-registrations-permissions-for-office-365-health-service/) that describes in detail how to set it up (the roadmap script does not need this).

- You also need to set up a webhook in your Teams channel of choice, and copy the URI in to the scripts.

- Adapt the user variables section in each script to work with you environment.

- For running the scripts, I configure them with Azure Automation runbooks set on schedules.

#### Screenshots ####

Health status:
![Screenshot](https://github.com/einast/PS_M365_scripts/blob/master/O365ServiceHealth3.PNG)

Message Center updates:
![Screenshot](https://github.com/einast/PS_M365_scripts/blob/master/M365MessageCenter2.PNG)

Roadmap updates:
![Screenshot](https://github.com/einast/PS_M365_scripts/blob/master/TeamsRoadmapWebHook3.PNG)

[![alt text][1.1]][1]

[1.1]: https://github.com/einast/PS_M365_scripts/blob/master/sc%2Blinkedin-131965017554733397_48.png

[1]: https://www.linkedin.com/in/easting/

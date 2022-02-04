# Introduction
We're working a lot at our home office these days. Several people already found inventive solutions to make working in the home office more comfortable. One of these ways is to automate activities in your home automatation system based on your status on Microsoft Teams.

Microsoft provides the status of your account that is used in Teams via the Graph API. To access the Graph API, your organization needs to grant consent for the organization so everybody can read their Teams status. Since my organization didn't want to grant consent, I needed to find a workaround, which I found in monitoring the Teams client logfile for certain changes.

This script makes use of two items that you will need to create in OpenHAB up front:
* Teams_Status
* Teams_Activity

Teams_Status displays that availability status of your Teams client based on the icon overlay in the taskbar on Windows. Teams_Activity shows if you are in a call or not based on the App updates deamon, which is paused as soon as you join a call.

# Important
This solution has been updated to work with OpenHAB. It will work with any home automation platform that provides an API, but you probably need to change the PowerShell code.

# Requirements
* Create the two Teams Items (Status and Activity) in OpenHAB

* Download the files from this repository and save them to C:\Scripts\TeamsStatus
* Edit the Settings.ps1 file and:
  * Replace `<Insert username>` and `<Insert password>` with valid OpenHAB Cloud Connector Credentials
  * Replace `<Insert Windows UserName>` with the username that is logged in to Teams and you want to monitor
  * Update `$OHUrl` if you are not using the standard OpenHAB Cloud Connector Service
  * Replace `<Insert status name>` and `<Insert activity name>` with the item names you created in OpenHAB
  * Adjust the language settings to your preferences
* Start a elevated PowerShell prompt, browse to C:\Scripts\TeamsStatus and run the following command:
```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned
Unblock-File .\Settings.ps1
Unblock-File .\Get-TeamsStatus.ps1
Start-Process -FilePath .\nssm.exe -ArgumentList 'install "Microsoft Teams Status Monitor" "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe" "-command "& { . C:\Scripts\TeamsStatus\Get-TeamsStatus.ps1 }"" ' -NoNewWindow -Wait
Start-Service -Name "Microsoft Teams Status Monitor"
```

After completing the steps below, start your Teams client and verify if the status and activity is updated as expected.

<#
.NOTES
    Name: Get-TeamsStatus.ps1
    Author: Danny de Vries
    Requires: PowerShell v2 or higher
    Version History: https://github.com/EBOOZ/TeamsStatus/commits/main
.SYNOPSIS
    Sets the status of the Microsoft Teams client to OpenHAB.
.DESCRIPTION
    This script is monitoring the Teams client logfile for certain changes. It
    makes use of two sensors that are created in OpenHAB up front.
    The status entity (sensor.teams_status by default) displays that availability 
    status of your Teams client based on the icon overlay in the taskbar on Windows. 
    The activity entity (sensor.teams_activity by default) shows if you
    are in a call or not based on the App updates deamon, which is paused as soon as 
    you join a call.
.PARAMETER SetStatus
    Run the script with the SetStatus-parameter to set the status of Microsoft Teams
    directly from the commandline.
.EXAMPLE
    .\Get-TeamsStatus.ps1 -SetStatus "Offline"
#>
# Configuring parameter for interactive run
Param($SetStatus)

# Import Settings PowerShell script
. ($PSScriptRoot + "\Settings.ps1")

$auth = $OHUser + ':' + $OHPassword
$Encoded = [System.Text.Encoding]::UTF8.GetBytes($auth)
$authorizationInfo = [System.Convert]::ToBase64String($Encoded)
$headers = @{"Authorization"="Basic $($authorizationInfo)"}

# Run the script when a parameter is used and stop when done
If($null -ne $SetStatus){
    Write-Host ("Setting Microsoft Teams status to "+$SetStatus+":")
	$params = (Get-Culture).TextInfo.ToTitleCase($SetStatus)
        
	Invoke-RestMethod -Uri "$OHUrl/rest/items/$entityStatus" -Method POST -Headers $headers -Body $params -ContentType "text/plain" 

	break
}

If ($env:APPDATA -match "C:\\Users") {
	$path = "$env:APPDATA\Microsoft\Teams\logs.txt"
}
Else {
	$path = "C:\Users\$UserName\AppData\Roaming\Microsoft\Teams\logs.txt"
}

# Start monitoring the Teams logfile when no parameter is used to run the script
Get-Content -Path $path -Tail 1000 -ReadCount 0 -Encoding Utf8 -Wait | % {
    
    # Get Teams Logfile and last icon overlay status
    $TeamsStatus = $_ | Select-String -Pattern `
        'Setting the taskbar overlay icon -',`
        'StatusIndicatorStateService: Added' | Select-Object -Last 1

    # Get Teams Logfile and last app update deamon status
    $TeamsActivity = $_ | Select-String -Pattern `
        'Resuming daemon App updates',`
        'Pausing daemon App updates',`
        'SfB:TeamsNoCall',`
        'SfB:TeamsPendingCall',`
        'SfB:TeamsActiveCall',`
        'name: desktop_call_state_change_send, isOngoing' | Select-Object -Last 1

    # Get Teams application process
    $TeamsProcess = Get-Process -Name Teams -ErrorAction SilentlyContinue

    # Check if Teams is running and start monitoring the log if it is
    If ($null -ne $TeamsProcess) {
        If($TeamsStatus -eq $null){ }
        ElseIf ($TeamsStatus -like "*Setting the taskbar overlay icon - $lgAvailable*" -or `
            $TeamsStatus -like "*StatusIndicatorStateService: Added Available*" -or `
            $TeamsStatus -like "*StatusIndicatorStateService: Added NewActivity (current state: Available -> NewActivity*") {
            $Status = $lgAvailable
            Write-Host $Status
        }
        ElseIf ($TeamsStatus -like "*Setting the taskbar overlay icon - $lgBusy*" -or `
                $TeamsStatus -like "*StatusIndicatorStateService: Added Busy*" -or `
                $TeamsStatus -like "*Setting the taskbar overlay icon - $lgOnThePhone*" -or `
                $TeamsStatus -like "*StatusIndicatorStateService: Added OnThePhone*" -or `
                $TeamsStatus -like "*StatusIndicatorStateService: Added NewActivity (current state: Busy -> NewActivity*") {
            $Status = $lgBusy
            Write-Host $Status
        }
        ElseIf ($TeamsStatus -like "*Setting the taskbar overlay icon - $lgAway*" -or `
                $TeamsStatus -like "*StatusIndicatorStateService: Added Away*" -or `
                $TeamsStatus -like "*StatusIndicatorStateService: Added NewActivity (current state: Away -> NewActivity*") {
            $Status = $lgAway
            Write-Host $Status
        }
        ElseIf ($TeamsStatus -like "*Setting the taskbar overlay icon - $lgBeRightBack*" -or `
                $TeamsStatus -like "*StatusIndicatorStateService: Added BeRightBack*" -or `
                $TeamsStatus -like "*StatusIndicatorStateService: Added NewActivity (current state: BeRightBack -> NewActivity*") {
            $Status = $lgBeRightBack
            Write-Host $Status
        }
        ElseIf ($TeamsStatus -like "*Setting the taskbar overlay icon - $lgDoNotDisturb *" -or `
                $TeamsStatus -like "*StatusIndicatorStateService: Added DoNotDisturb*" -or `
                $TeamsStatus -like "*StatusIndicatorStateService: Added NewActivity (current state: DoNotDisturb -> NewActivity*") {
            $Status = $lgDoNotDisturb
            Write-Host $Status
        }
        ElseIf ($TeamsStatus -like "*Setting the taskbar overlay icon - $lgFocusing*" -or `
                $TeamsStatus -like "*StatusIndicatorStateService: Added Focusing*" -or `
                $TeamsStatus -like "*StatusIndicatorStateService: Added NewActivity (current state: Focusing -> NewActivity*") {
            $Status = $lgFocusing
            Write-Host $Status
        }
        ElseIf ($TeamsStatus -like "*Setting the taskbar overlay icon - $lgPresenting*" -or `
                $TeamsStatus -like "*StatusIndicatorStateService: Added Presenting*" -or `
                $TeamsStatus -like "*StatusIndicatorStateService: Added NewActivity (current state: Presenting -> NewActivity*") {
            $Status = $lgPresenting
            Write-Host $Status
        }
        ElseIf ($TeamsStatus -like "*Setting the taskbar overlay icon - $lgInAMeeting*" -or `
                $TeamsStatus -like "*StatusIndicatorStateService: Added InAMeeting*" -or `
                $TeamsStatus -like "*StatusIndicatorStateService: Added NewActivity (current state: InAMeeting -> NewActivity*") {
            $Status = $lgInAMeeting
            Write-Host $Status
        }
        ElseIf ($TeamsStatus -like "*Setting the taskbar overlay icon - $lgOffline*" -or `
                $TeamsStatus -like "*StatusIndicatorStateService: Added Offline*") {
            $Status = $lgOffline
            Write-Host $Status
        }

        If($TeamsActivity -eq $null){ }
        ElseIf ($TeamsActivity -like "*Resuming daemon App updates*" -or `
            $TeamsActivity -like "*SfB:TeamsNoCall*" -or `
            $TeamsActivity -like "*name: desktop_call_state_change_send, isOngoing: false*") {
            $Activity = $lgNotInACall
            Write-Host $Activity
        }
        ElseIf ($TeamsActivity -like "*Pausing daemon App updates*" -or `
            $TeamsActivity -like "*SfB:TeamsActiveCall*" -or `
            $TeamsActivity -like "*name: desktop_call_state_change_send, isOngoing: true*") {
            $Activity = $lgInACall
            Write-Host $Activity
        }
    }
    # Set status to Offline when the Teams application is not running
    Else {
            $Status = $lgOffline
            $Activity = $lgNotInACall
            Write-Host $Status
            Write-Host $Activity
    }

    # Call OpenHAB API to set the status and activity sensors
    If ($CurrentStatus -ne $Status -and $Status -ne $null) {
        $CurrentStatus = $Status

        $params = (Get-Culture).TextInfo.ToTitleCase($Status)
        
        Invoke-RestMethod -Uri "$OHUrl/rest/items/$entityStatus" -Method POST -Headers $headers -Body $params -ContentType "text/plain" 
    }

    If ($CurrentActivity -ne $Activity) {
        $CurrentActivity = $Activity

        $params = (Get-Culture).TextInfo.ToTitleCase($Activity)
        
        Invoke-RestMethod -Uri "$OHUrl/rest/items/$entityActivity" -Method POST -Headers $headers -Body $params -ContentType "text/plain" 
    
	}
        
}

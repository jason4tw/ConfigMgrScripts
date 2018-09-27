<#
	.SYNOPSIS
		Lists all Automatic Deployment Rules to formatted console output.
	
	.DESCRIPTION
        Lists all Automatic Deployment Rules to formatted console output.
        
Version 0.5
26 September 2018
Jason Sandys

#>

[CmdletBinding()]
param
(

)

Import-Module (Join-Path (Split-Path $env:SMS_ADMIN_UI_PATH –parent) ConfigurationManager.psd1)
$siteCode = Get-PSDrive -PSProvider CMSITE

Push-Location $siteCode":"

$allAdrs = Get-CMAutoDeploymentRule

foreach($adr in $allAdrs)
{
    Write-Output " * $($adr.Name)"

    if($adr.Schedule -ne "")
    {
        $sched = Convert-CMSchedule -ScheduleString $($adr.Schedule)
        $startTime = [DateTime]$sched.StartTime

        Write-Output "   Evaluation Schedule $($sched.SmsProviderObjectPath.ToString().Substring(12))"
        Write-Output "   - Type: $($sched.SmsProviderObjectPath.ToString().Substring(12))"
        Write-Output "   - Time: $($startTime.ToString('T'))"

        if($sched.SmsProviderObjectPath -eq "SMS_ST_RecurMonthlyByWeekday")
        {
            if($sched.WeekOrder -eq 1)
            {
                $week = "First"
            }
            elseif($sched.WeekOrder -eq 2)
            {
                $week = "Second"
            }
            elseif($sched.WeekOrder -eq 3)
            {
                $week = "Third"
            }
            elseif($sched.WeekOrder -eq 4)
            {
                $week = "Fourth"
            }
            elseif($sched.WeekOrder -eq 5)
            {
                $week = "Fifth"
            }

            Write-Output "   - Day: $week $((Get-Culture).DateTimeFormat.GetDayName($sched.Day))"
            Write-Output "   - Recur every $($sched.ForNumberOfMonths) month(s)"
        }
        elseif($sched.SmsProviderObjectPath -eq "SMS_ST_RecurInterval")
        {
            if($sched.HourSpan -gt 0)
            {
                Write-Output "   - Recur every $($sched.HourSpan) Hours"
            }
            elseif($sched.DaySpan -gt 0)
            {
                Write-Output "   - Recur every $($sched.DaySpan) Days"
            }
            elseif($sched.MinuteSpan -gt 0)
            {
                Write-Output "   - Recur every $($sched.MinuteSpan) Minutes"
            }
        }
        elseif($sched.SmsProviderObjectPath -eq "SMS_ST_RecurWeekly")
        {
            Write-Output "   - Day: $((Get-Culture).DateTimeFormat.GetDayName($sched.Day))"
            Write-Output "   - Recur every $($sched.ForNumberOfWeeks) week(s)"            
        }
    }
    else
    {
        Write-Output "   Evaluation disabled"

    }
    Write-Output "--------------------------------------------------"

    $allAdrDeployments = Get-CMAutoDeploymentRuleDeployment -Id $adr.AutoDeploymentId

    foreach($adrDeployment in $allAdrDeployments)
    {
        Write-Output "     Collection: $($adrDeployment.CollectionName)"

        $deployomentInfo = [xml]$adrDeployment.DeploymentTemplate

        $supressRestarts = New-Object System.Collections.ArrayList($null)

        if($deployomentInfo.DeploymentCreationActionXML.SuppressServers -eq 'Checked')
        {
            $supressRestarts += "Servers"   
        }
        if($deployomentInfo.DeploymentCreationActionXML.SuppressWorkstations -eq 'Checked')
        {
            $supressRestarts += "Workstations"   
        }
        if($supressRestarts.Count -eq 0)
        {
            $supressRestarts += "None"
        }

        $ignoreMaintenaceWindows = New-Object System.Collections.ArrayList($null)

        if($deployomentInfo.DeploymentCreationActionXML.AllowRestart -eq $true)
        {
            $ignoreMaintenaceWindows += "Restarts"   
        }
        if($deployomentInfo.DeploymentCreationActionXML.AllowInstallOutSW -eq $true)
        {
            $ignoreMaintenaceWindows += "Update Installation"   
        }
        if($ignoreMaintenaceWindows.Count -eq 0)
        {
            $ignoreMaintenaceWindows += "None"
        }

        Write-Output "      - UTC: $($deployomentInfo.DeploymentCreationActionXML.Utc)"
        Write-Output "      - Available Offset: $($deployomentInfo.DeploymentCreationActionXML.AvailableDeltaDuration) $($deployomentInfo.DeploymentCreationActionXML.AvailableDeltaDurationUnits)"
        Write-Output "      - Deadline Offset: $($deployomentInfo.DeploymentCreationActionXML.Duration) $($deployomentInfo.DeploymentCreationActionXML.DurationUnits)"
        Write-Output "      - Suppress restarts: $($supressRestarts -join ',')"
        Write-Output "      - Deployment Enabled: $($deployomentInfo.DeploymentCreationActionXML.EnableDeployment)"
        Write-Output "      - WoL Enabled: $($deployomentInfo.DeploymentCreationActionXML.EnableWakeOnLan)"
        Write-Output "      - User Notification: $($deployomentInfo.DeploymentCreationActionXML.UserNotificationOption)"
        Write-Output "      - Ignore Maintenace Windows for: $($ignoreMaintenaceWindows -join ',')"
        Write-Output "      - Force rescan afer restart: $($deployomentInfo.DeploymentCreationActionXML.RequirePostRebootFullScan -eq 'Checked')"
        Write-Output ""

    }
    
}

Pop-Location
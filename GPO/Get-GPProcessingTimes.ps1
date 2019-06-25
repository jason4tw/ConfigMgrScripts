[CmdletBinding()] 
Param
(
    [ValidateSet('User', 'Computer', 'Both')]
    [string]$PolicyType = 'Both',
    [ValidateSet('Yes', 'No')]
    [string]$IncludeCSEInfo = 'Yes'
)

Function Get-CSEEvents
{
    [CmdletBinding()] 
    Param
    (
        [parameter(ParameterSetName="user")]
        [switch]$UserGPTime,
        [parameter(ParameterSetName="computer")]
        [switch]$ComputerGPTime,
        [switch]$IncludePolicyCSEInfo
    )

    $result = New-Object -TypeName psobject
    $result | Add-member -MemberType NoteProperty -Name SystemName -Value $env:COMPUTERNAME
 
    if($UserGPTime -eq $true)
    {
        $startEventID = 4005
        $endEventID = 8005

        $result | Add-Member -MemberType NoteProperty -Name "Type" -Value "User"
    }
    else
    {
        $startEventID = 4004
        $endEventID = 8004

        $result | Add-Member -MemberType NoteProperty -Name "Type" -Value "Computer"
    }

    $eventFilter = "*[System[EventID='$startEventID']]"
    $startEvents = Get-WinEvent -ProviderName Microsoft-Windows-GroupPolicy -FilterXPath $eventFilter -ErrorAction SilentlyContinue
    
    if($startEvents)
    {
        $activityId = $startEvents[0].ActivityId.GUID
        $principal = $startEvents[0].Properties[1].Value
        $isBackground = $startEvents[0].Properties[4].Value
        $isAsync = $startEvents[0].Properties[5].Value

        $result | Add-Member -MemberType NoteProperty -Name "Principal" -Value $principal
        $result | Add-Member -MemberType NoteProperty -Name "Background" -Value $isBackground
        $result | Add-Member -MemberType NoteProperty -Name "Asynch" -Value $isAsync

        $eventFilter = "*[System[EventID='$endEventID' and Correlation[@ActivityID='{$ActivityID}']]]"
        $endEvent = Get-WinEvent -ProviderName Microsoft-Windows-GroupPolicy -FilterXPath $eventFilter -ErrorAction SilentlyContinue

        if($endEvent)
        {
            $elapsedTime = $endEvent.Properties[0].Value
            $endTime = $endEvent.TimeCreated.ToUniversalTime().ToString("u")

            $result | Add-Member -MemberType NoteProperty -Name "ElapsedTime" -Value $elapsedTime
            $result | Add-Member -MemberType NoteProperty -Name "EndTime" -Value $endTime
        }

        if($IncludePolicyCSEInfo)
        {
            $eventFilter = "*[System[(EventID='4016' or EventID='5016' or EventID='6016' or EventID='7016') and Correlation[@ActivityID='{$ActivityID}']]]"
            $cseEvents = Get-WinEvent -ProviderName Microsoft-Windows-GroupPolicy -FilterXPath $eventFilter -ErrorAction SilentlyContinue

            if ($cseEvents)
            {
                $allCseResults = @()

                $cseStartEvents = $cseEvents | Where-Object { $_.Id -eq '4016' }

                foreach ($event in $cseStartEvents)
                {
                    $cseName = $event.Properties[1] | Select-Object -ExpandProperty Value
                    $cseGPOs = ($event.Properties[5]  | Select-Object -ExpandProperty Value).TrimEnd("`n") -replace "`n",", "

                    $cseResult = New-Object -TypeName psobject
                    $cseResult | Add-Member -MemberType NoteProperty -Name CSEName -Value $cseName
                    $cseResult | Add-Member -MemberType NoteProperty -Name GPOs -Value $cseGPOs

                    $allCseResults += $cseResult
                }

                $cseEndEvents = $cseEvents | Where-Object { $_.Id -ne '4016' }

                foreach ($event in $cseEndEvents)
                {
                    $cseName = $event.Properties[2] | Select-Object -ExpandProperty Value
                    $cseTime = $event.Properties[0] | Select-Object -ExpandProperty Value
                    $cseError = $event.Properties[1] | Select-Object -ExpandProperty Value

                    $cseResult = $allCseResults | Where-Object { $_.CSEName -eq $cseName }
                    $cseResult | Add-Member -MemberType NoteProperty -Name ElapsedTime -Value $cseTime
                    $cseResult | Add-Member -MemberType NoteProperty -Name ErrorCode -Value $cseError
                }

                $result | Add-Member -MemberType NoteProperty -Name CSEResults -Value $allCseResults
            }
        }
    }

    $result
}

if ($PolicyType -eq 'Computer' -or $PolicyType -eq 'Both')
{
    Get-CSEEvents -ComputerGPTime -IncludePolicyCSEInfo:$($IncludeCSEInfo -eq 'Yes')
}

if ($PolicyType -eq 'User' -or $PolicyType -eq 'Both')
{
    Get-CSEEvents -UserGPTime -IncludePolicyCSEInfo:$($IncludeCSEInfo -eq 'Yes')
}

#$allGPResults = @()

#$allGPResults += Get-CSEEvents -ComputerGPTime -IncludePolicyCSEInfo
#$allGPResults += Get-CSEEvents -UserGPTime -IncludePolicyCSEInfo

#$gpResult = New-Object -TypeName psobject
#$gpResult | Add-member -MemberType NoteProperty -Name SystemName -Value $env:COMPUTERNAME
#$gpResult | Add-member -MemberType NoteProperty -Name GPOProcessing -Value $allGPResults

#$gpResult

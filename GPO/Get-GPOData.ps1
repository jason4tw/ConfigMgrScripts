[CmdletBinding()]
param
(
    [xml]$GpoReportXml
)


function ConvertFrom-CanonicalPath {

    [cmdletbinding()]

    param(

    [Parameter(Mandatory,ValueFromPipeline=$True,ValueFromPipelineByPropertyName=$True)] 

    [ValidateNotNullOrEmpty()] 

    [string]$CanonicalPath

    )

    process {

        $obj = $CanonicalPath.Replace(',','\,').Split('/')

        [string]$DN = "OU=" + $obj[$obj.count - 1]

        for ($i = $obj.count - 2;$i -ge 1;$i--){$DN += ",OU=" + $obj[$i]}

        $obj[0].split(".") | ForEach-Object { $DN += ",DC=" + $_}

        return $DN

    }

}

#[xml]$gpoReport = Get-GPOReport -all -ReportType XML
#$gpoReport.Save("c:\users\admsandysj\desktop\gpoReport.xml")
#[xml]$gpoR = Get-Content .\gpoReport.xml

#$gpoReport.GPOS.GPO | where {$_.User.Enabled -eq $false -and $_.Computer.Enabled -eq $false} | Select Name
#$gpoReport.GPOS.GPO | where {$_.User.extensiondata -eq $null -and $_.Computer.ExtensionData -eq $null} | Select Name

#$gpoReport.GPOS.GPO | Select-Object -Property @{Name= 'GpoName'; Expression = {$_.Name}}, @{Name = 'WmiFilterName'; Expression = { $_.FilterName}}, @{Name = 'LinkPath'; Expression = { $_.LinksTo.SOMPath}}

$userHash = @{}
$computerHash = @{}

$gpoTotalCount = $GpoReportXml.GPOS.GPO.count
$gpoCounter = 0

[System.Collections.ArrayList]$gpoList = @()

Write-Progress -Id 1 -Activity "Gathering GPO data" -Status "0 of $gpoTotalCount" -PercentComplete 0

foreach ($gpo in $GpoReportXml.GPOS.GPO)
{
    # $gpoOut = ($gpo | Select-Object -Property @{Name = 'GpoName'; Expression = {$_.Name}}, @{Name = 'WmiFilterName'; Expression = { $_.FilterName}}, @{Name = 'LinkPath'; Expression = { $_.LinksTo.SOMPath}})

    $gpoCounter++

    Write-Progress -Id 1 -Activity "Gathering GPO data" -CurrentOperation $gpo.Name -Status "$gpoCounter of $gpoTotalCount" -PercentComplete ($gpoCounter / $gpoTotalCount * 100)

    [System.Collections.ArrayList]$links = @()

    $linkTotalCount = $gpo.LinksTo.SOMPath.Count
    $linkCounter = 0
    $totalUserCount = 0
    $totalComputerCount = 0

    Write-Progress -Id 2 -ParentId 1 -Activity "Gathering OU membership" -Status "0 of $linkTotalCount" -PercentComplete 0

    foreach ($ou in $gpo.LinksTo.SOMPath)
    {
        $linkCounter++

        Write-Progress -Id 2 -ParentId 1 -Activity "Gathering OU membership" -CurrentOperation $ou -Status "$linkCounter of $linkTotalCount" -PercentComplete ($linkCounter / $linkTotalCount * 100)

        if($ou.IndexOf('/') -gt 0)
        {
            if($userHash.ContainsKey($ou))
            {
                $uCount = $userHash."$ou"
                #Write-Output "Hashed User Count for $ou : $uCount"
            }
            else
            {
                $uCount = (Get-ADUser -Filter * -SearchBase (ConvertFrom-CanonicalPath $ou)).Count
                $userHash.Add($ou, $uCount)
            }

            if($computerHash.ContainsKey($ou))
            {
                $cCount = $computerHash."$ou"
                #Write-Output "Hashed Computer Count for $ou : $cCount"
            }
            else
            {
                $cCount = (Get-ADComputer -Filter * -SearchBase (ConvertFrom-CanonicalPath $ou)).Count
                $computerHash.Add($ou, $cCount)
            }

            if($uCount)
            {
                $totalUserCount += $uCount
            }

            if($cCount)
            {
                $totalComputerCount += $cCount
            }
        }
        else
        {
            $uCount = "All"
            $cCount = "All"
        }

        $gpoLink = [PSCustomObject]@{
            OU = $ou
            UserCount = $uCount
            ComputerCount = $cCount
        }

        $links += $gpoLink

        #$gpoLink
    }

    Write-Progress -Id 2 -ParentId 1 -Activity "Gathering OU membership" -Completed

    $containsComputerSettings = $gpo.Computer.ExtensionData.Count -ne 0
    $containsUserSettings = $gpo.User.ExtensionData.Count -ne 0

    $gpoOut = [PSCustomObject]@{
        GpoName = $gpo.Name
        WmiFilterName = $gpo.FilterName
        ComputerSettingsEnabled = $gpo.Computer.Enabled
        ContainsComputerSettings = $containsComputerSettings
        ComputerSettings = ($gpo.Computer.ExtensionData | Sort-Object -Property Name | Select-Object -Expand Name) -join ','
        TotalComputerCount = $totalComputerCount
        UserSettingsEnabled = $gpo.User.Enabled
        ContainsUserSettings = $containsUserSettings
        UserSettings = ($gpo.User.ExtensionData | Sort-Object -Property Name | Select-Object -Expand Name) -join ','
        TotalUserCount = $totalUserCount
        TotalLinkCount = $linkTotalCount
        GpoLinks = $links
    }

    $gpoList += $gpoOut

}

Write-Progress -Id 1 -Activity "Gathering GPO data" -Completed

$gpoList | ConvertTo-Json -Depth 3
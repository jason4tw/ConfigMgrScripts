$allGpos = Get-GPO -All
#$allGpos = Get-GPO -Name 'CC-W10-Allow WinRM Client Basic Authentication'
#$allGpos = Get-GPO -Name 'CC-W10a-IE Configuration'
#$allGpos = Get-GPO -Name 'CC-W10-1709-Configuration'

$gpoTotalCount = $allGpos.Count
$gpoCounter = 0
$allGpoPerms = @()

Write-Progress -Id 1 -Activity "Gathering GPO data" -Status "0 of $gpoTotalCount" -PercentComplete 0

foreach($gpo in $allGpos)
{

    $gpoCounter++
    Write-Progress -Id 1 -Activity "Gathering GPO data" -CurrentOperation $gpo.DisplayName -Status "$gpoCounter of $gpoTotalCount" -PercentComplete ($gpoCounter / $gpoTotalCount * 100)

    $gpoPermissions = Get-GPPermissions $gpo.DisplayName -All | Where-Object { $_.Permission -eq 'GpoApply' }
    $trustees = @()
    $trusteeCounts = @{}

    $permTotalCount = $gpoPermissions.Count
    $permCounter = 0

    Write-Progress -Id 2 -ParentId 1 -Activity "Gathering GPO permissions" -Status "0 of $permTotalCount" -PercentComplete 0

    foreach($permission in $gpoPermissions)
    {
        $permCounter++
        Write-Progress -Id 2 -ParentId 1 -Activity "Gathering GPO permisisons" -CurrentOperation $permission.Trustee.Name -Status "$permCounter of $permTotalCount" -PercentComplete ($permCounter / $permTotalCount * 100)

        $trusteeType = [string]$permission.Trustee.SidType

        $trustees += New-Object -TypeName PSObject -Property @{Trustee=$permission.Trustee.Name; TrusteeType=$trusteeType}

        $count = 0

        if($trusteeCounts.ContainsKey($trusteeType))
        {
            $count = $trusteeCounts[$trusteeType]
        }

        $trusteeCounts[$trusteeType] = $count + 1

    }

    $gpoPerms = [PSCustomObject] @{
            GPO = $gpo.DisplayName
            TotalTrusteeCount = $gpoPermissions.Count
            TrusteeCounts = [PSCustomObject] $trusteeCounts
            Trustees = $trustees
        }

    $allGpoPerms += $gpoPerms
}

Write-Progress -Id 1 -Activity "Gathering GPO data" -Completed

$allGpoPerms
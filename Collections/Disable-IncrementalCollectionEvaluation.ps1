[CmdLetBinding()]
param
(
    [Parameter(Mandatory=$True,HelpMessage="Site Server Name")]
    [string]$SiteServer,
    [Parameter(Mandatory=$True,
        HelpMessage="Text file containing collection IDs")]
    [ValidateScript({
        if(-Not ($_ | Test-Path) ){
            throw "File or folder does not exist" 
        }
        if(-Not ($_ | Test-Path -PathType Leaf) ){
            throw "The CollectionList argument must be a file. Folder paths are not allowed."
        }
        return $true
    })]
    [System.IO.FileInfo]$CollectionList,
    [switch]$WhatIf
)
    
$providerLocation = @(Get-WMIObject -ComputerName $SiteServer -Namespace "root\sms" -Class "SMS_ProviderLocation")[0]

$allCollectionIDs = Get-Content -Path $CollectionList

$count = 0
$total = $allCollectionIds.Count

Write-Progress -Id 1 -Activity "Disabling Incrememntal Collection Evaluation" -PercentComplete 0 -Status "$count of $total"

foreach ($collectionID in $allCollectionIDs)
{
    $collection = Get-WMIObject -ComputerName $providerLocation.Machine -Namespace "root\sms\site_$($providerLocation.SiteCode)" -Class "SMS_Collection" -Filter "CollectionID = '$collectionID'"

    $percent = ((++$count / $total ) * 100 )
    
    Write-Progress -Id 1 -Activity "Disabling Incrememntal Collection Evaluation" -PercentComplete $percent -Status "$count of $total" -CurrentOperation $collection.Name

    if($collection.RefreshType -eq 6)
    {
        if($WhatIf -eq $false)
        {
            Write-Output "Updating '$($collection.Name) ($($collection.CollectionID))'..."
            $collection.RefreshType = 2
            $collection.Put() | Out-Null
         }
         else
         {
            Write-Output "WhatIf: Updating '$($collection.Name) ($($collection.CollectionID))'..."
         }
    }
    else
    {
        Write-Output "Incremental Evaluation was not enabled on '$($collection.Name) ($($collection.CollectionID))'..."
    }
}

Write-Progress -Id 1 -Completed -Activity "Disabling Incrememntal Collection Evaluation"


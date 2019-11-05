<#
	.SYNOPSIS
		Renames the collections within a specified folder using regular expression pattern matching.
	
	.DESCRIPTION
        Renames the collections within a specified folder using regular expression pattern matching.
	
	.PARAMETER Name
        A pattern specifying the name of the collections to include. Normal file system wildcards (* and ?) can be used.
	.PARAMETER Folder
        The folder to include collections from.
    .PARAMETER Search
        The Search string for the regular expression.
    .PARAMETER Replace
        The Replace string for the regular expression.
    .PARAMETER Recurse
        Include sub-folders of the specified folder.	
    .PARAMETER SiteServer
        The name of the primary site server to connect to. By default, the local host will be used.
    .EXAMPLE
        .\Rename-CMCollections.ps1
        ...

    .NOTES
        Version 1.1
        Jason Sandys

        Version History
        - 1.1 (23 October 2018): Added insert prefix option, fixed parameter sets, added progress bar
        - 1.0 (27 September 2018): Initial Version

        Limitations and Issues
       
#>

[CmdletBinding()]
param
(
    [Parameter(Mandatory=$False, HelpMessage = 'The site server to connect to.')]
    [string] $SiteServer = $env:computername,
    
    [Parameter(Mandatory=$False, HelpMessage = 'The collection(s) to rename (PowerShell wildcards are allowed)')]
    [alias('Name')]
    [alias('NameFilter')]
    [alias('Collection')]
    [string] $CollectionNameFilter = '*',
    
    [Parameter(Mandatory=$True, HelpMessage = 'The folder to rename collections in.')]
    [alias('Folder')]
    [string] $FolderName,
    [Parameter(Mandatory=$False, HelpMessage = 'Recurse sub-folders.')]
    [switch] $Recurse,
    
    [Parameter(Mandatory=$True, HelpMessage = 'The regular expression search string.', ParameterSetName='Replace')]
    [string] $Search,
    [Parameter(Mandatory=$True, HelpMessage = 'The regular expression replace string.', ParameterSetName='Replace')]
    [string] $Replace,

    [Parameter(Mandatory=$True, HelpMessage = 'The prefix to insert at the beginning of the folder name.', ParameterSetName='Prefix')]
    [string] $Prefix,

    [switch] $WhatIf
)

function Process-DeviceCollectionsInFolder
{
    param
    (
	    [Parameter(Mandatory=$True)]
        [string] $FolderNodeID,
        [Parameter(Mandatory=$True)]
        [string] $FolderNodeName,
        [switch] $InsertPrefix,
        [switch] $ReplaceString
    )

    $folderCollections = Get-WMIObject -Computer $providerSystem -Namespace $providerNamespace -Class SMS_ObjectContainerItem `
        -Filter "ContainerNodeID='$FolderNodeID' And ObjectType='5000'"

    if($WhatIf -eq $true)
    {
        Write-Host -ForegroundColor Blue -BackgroundColor White "`nWhatIf Mode - No changes will be made"  
    }
    
    Write-Host "`nProcessing collections in $FolderNodeName ..."

    $collectionCounter = 0
    $collectionCount = $folderCollections.Count
    Write-Progress -Id 1 -Activity "Processing collections in $FolderNodeName" -Status "0 of $collectionCount"

    foreach($collectionContainerItem in $folderCollections)
    {
        $collection = Get-CMDeviceCollection -CollectionId $collectionContainerItem.InstanceKey
        $collectioncounter++
 
        Write-Progress -Id 1 -Activity "Processing collections in $FolderNodeName" -Status "$collectionCounter of $($folderCollections.Count)" -CurrentOperation $collection.Name -PercentComplete (($collectionCounter / $collectionCount) * 100)

        if($collection.Name -like "$CollectionNameFilter")
        {
            if ($ReplaceString)
            {
                $newName = $collection.name -replace $Search, $Replace
            }
            elseif ($InsertPrefix) 
            {
                $newName = "$Prefix" + $collection.Name
            }
            else
            {
                $newName = ''
            }

            if($newName -ne '')
            {
                Write-Host "  > Renaming `"$($collection.Name)`" to `"$newName`""

                if($WhatIf -eq $false)
                {
                    Set-CMDeviceCollection -CollectionId $collection.CollectionId -NewName "$newName" 
                }
            }

        }

    }

    Write-Progress -Id 1 -Activity "Processing collections in $FolderNodeName" -Status "$collectionCounter of $($folderCollections.Count)" -Completed

    if($Recurse)
    {
        $allSubFolders = Get-WMIObject -Computer $providerSystem -Namespace $providerNamespace -Class SMS_ObjectContainerNode `
            -Filter "ParentContainerNodeID='$FolderNodeID' And ObjectType='5000'"

        foreach($subFolder in $allSubFolders)
        { 
            Process-DeviceCollectionsInFolder -FolderNodeID $subFolder.ContainerNodeID -FolderNodeName "$FolderNodeName\$($subFolder.Name)" -InsertPrefix:$InsertPrefix -ReplaceString:$ReplaceString
        }
    }   
}

Import-Module (Join-Path (Split-Path $env:SMS_ADMIN_UI_PATH -parent) ConfigurationManager.psd1)
$siteCode = Get-PSDrive -PSProvider CMSITE

$providerLocation = Get-WMIObject -Computer $SiteServer -Namespace root\sms -Class SMS_ProviderLocation
$siteCode = $providerLocation.SiteCode
$providerSystem = $providerLocation.Machine
$providerNamespace = "root\sms\site_$siteCode"

Push-Location $siteCode":"

$topFolder = Get-WMIObject -Compute $providerSystem -Namespace $providerNamespace -Class SMS_ObjectContainerNode -Filter "Name='$FolderName' And ObjectType='5000'"

Process-DeviceCollectionsInFolder -FolderNodeID $topFolder.ContainerNodeID -FolderNodeName $topFolder.Name -InsertPrefix:($PSCmdlet.ParameterSetName -eq 'Prefix') -ReplaceString:($PSCmdlet.ParameterSetName -eq 'Replace')

Pop-Location

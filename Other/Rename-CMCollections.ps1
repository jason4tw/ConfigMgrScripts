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
        Version 1.0
        Jason Sandys

        Version History
        - 1.0 (27 September 2018): Initial Version

        Limitations and Issues
       
#>

[CmdletBinding(DefaultParameterSetName = 'Standard')]
param
(
    [Parameter(Mandatory=$False, HelpMessage = 'The collection(s) to rename (PowerShell wildcards are allowed)', ParameterSetName='Standard')]
    [Parameter(ParameterSetName='FolderFilter')]
    [string]$Name = '*',
    [Parameter(Mandatory=$True, HelpMessage = 'The regular expression search string.', ParameterSetName='Standard')]
    [Parameter(ParameterSetName='FolderFilter')]
    [string]$Search,
    [Parameter(Mandatory=$True, HelpMessage = 'The regular expression replace string.', ParameterSetName='Standard')]
    [Parameter(ParameterSetName='FolderFilter')]
    [string]$Replace,
    [Parameter(Mandatory=$True, HelpMessage = 'The folder to rename collections in.', ParameterSetName='FolderFilter')]
    [string]$Folder,
    [Parameter(Mandatory=$False, HelpMessage = 'Recurse sub-folders.', ParameterSetName='FolderFilter')]
    [switch]$Recurse,
    [Parameter(Mandatory=$False, HelpMessage = 'The site server to connect to.', ParameterSetName='FolderFilter')]
    [string]$SiteServer = $env:computername,
    [switch]$WhatIf
)

function Process-DeviceCollectionsInFolder
{
    param
    (
	    [Parameter(Mandatory=$True)]
        [string]$FolderNodeID,
        [Parameter(Mandatory=$True)]
        [string]$FolderNodeName
    )

    $memberCollections = Get-WMIObject -Computer $providerSystem -Namespace $providerNamespace -Class SMS_ObjectContainerItem `
        -Filter "ContainerNodeID='$FolderNodeID' And ObjectType='5000'"

    foreach($collectionContainerItem in $memberCollections)
    {
        $collection = Get-CMDeviceCollection -CollectionId $collectionContainerItem.InstanceKey
 
        if($collection.Name -like "$Name")
        {
            $newName = $collection.name -replace $Search, $Replace

            Write-Host "Renaming `"$($collection.name)`" to `"$newName`""

            if($WhatIf -eq $false)
            {
                Set-CMDeviceCollection -CollectionId $collection.CollectionId -NewName "$newName"
            }

        }
    }

    if($Recurse)
    {
        $allSubFolders = Get-WMIObject -Computer $providerSystem -Namespace $providerNamespace -Class SMS_ObjectContainerNode `
            -Filter "ParentContainerNodeID='$FolderNodeID' And ObjectType='5000'"

        foreach($subFolder in $allSubFolders)
        {
            Process-DeviceCollectionsInFolder -FolderNodeID $subFolder.ContainerNodeID -FolderNodeName "$FolderNodeName\$($subFolder.Name)"
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

$topFolder = Get-WMIObject -Compute $providerSystem -Namespace $providerNamespace -Class SMS_ObjectContainerNode -Filter "Name='$Folder' And ObjectType='5000'"

Process-DeviceCollectionsInFolder -FolderNodeID $topFolder.ContainerNodeID -FolderNodeName $topFolder

Pop-Location
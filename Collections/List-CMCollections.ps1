param
(
	[Parameter(Mandatory=$false,ValueFromPipeline)]
    [string] $TopLevelFolderName = 'OSD'
)
    
function Get-CollectionsInFolder
{

	param
	(
		[Parameter(Mandatory=$true,ValueFromPipeline)]
        [string] $FolderID
	)

    Get-WMIObject -Namespace root\sms\site_300 -Class SMS_ObjectContainerItem -Filter "ObjectTypeName = 'SMS_Collection_Device' And ContainerNodeID = '$FolderID'"
}

function Get-SubFolders
{

	param
	(
		[Parameter(Mandatory=$true,ValueFromPipeline)]
        [string] $FolderID
	)

    Get-WMIObject -Namespace root\sms\site_300 -Class SMS_ObjectContainerNode -Filter "ObjectTypeName = 'SMS_Collection_Device' And ParentContainerNodeID = '$FolderID'"
}

function List-CollectionsInFolder
{

	param
	(
		[Parameter(Mandatory=$true,ValueFromPipeline)]
        [string] $FolderName,
        [string] $FolderPath = "",
        [switch] $Recurse = $true
	)

    $folder = Get-WMIObject -Namespace root\sms\site_300 -Class SMS_ObjectContainerNode -Filter "ObjectTypeName = 'SMS_Collection_Device' And Name ='$FolderName'"
    
    if(-not $folder)
    {
        return
    }
    
    Write-Output "$FolderPath$FolderName"
    
    $immediateCollections = Get-CollectionsinFolder -FolderID $folder.ContainerNodeId

    foreach ($collection in $immediateCollections)
    {
        $collectionName = (Get-CMDeviceCollection -Id $collection.InstanceKey).Name
        $collectionInfo = Get-CMCollectionSetting -CollectionId $collection.InstanceKey

        Write-Output "  - $collectionName"

        foreach ($variable in $collectionInfo.CollectionVariables)
        {
            Write-Output "    * $($variable.Name) = $($variable.value)"
        }
        
    }

    Write-Output ""

    if($Recurse)
    {
        $subFolders = $folder.ContainerNodeId | Get-SubFolders
        $subFolders | ForEach-Object { $_.Name | List-CollectionsInFolder -FolderPath "$FolderPath$FolderName\" }
    }
}

$TopLevelFolderName | List-CollectionsInFolder

<#
	.SYNOPSIS
		Creates a json data file with all specified collections and their members.
	
	.DESCRIPTION
        Creates a json data file that can be consumed by Build-CMDefaultCollections.ps1. All selected collections along with their
members are included in the data file. Members include query, direct, include, and exclude rules.
	
	.PARAMETER Output
        The name a of a file to place to json into. This file will be overwritten. If no file is specified, the json is output to the console.
	.PARAMETER Name
        A pattern specifying the name of the collections to include. Normal file system wildcards (* and ?) can be used.
	.PARAMETER Folder
        The folder to include collections from.
	.PARAMETER Recurse
        Include sub-folders of the specified folder.	
    .PARAMETER SiteServer
        The name of the primary site server to connect to. By default, the local host will be used.
    .PARAMETER LimitingCollection
        The name of the limiting collection.
    .EXAMPLE
        .\Build-CMDefaultConfig.ps1 -ConfigFile .\Build-CMDefaultConfig.json -Collections
        Creates collections defined in the Build-CMDefaultConfig.json configuration file.
    .EXAMPLE
        .\Build-CMDefaultConfig.ps1 -ConfigFile .\Build-CMDefaultConfig.json -Collections -UpdateMembership
        Creates collections defined in the Build-CMDefaultConfig.json configuration file if they don't already exist and
updates the membership of collections defined in the Build-CMDefaultConfig.json configuration file.

    .NOTES
        Version 1.0
        Jason Sandys

        Version History
        - 1.0 (26 September 2018): Initial Version

        Limitations and Issues
        - Does not create json hierarchy for folders.
        
#>

[CmdletBinding(DefaultParameterSetName = 'Standard')]
param
(
    [Parameter(Mandatory=$False, HelpMessage = 'The file to dump the output to', ParameterSetName='Standard')]
    [Parameter(ParameterSetName='FolderFilter')]
    [string]$Output,
    [Parameter(Mandatory=$False, HelpMessage = 'The collection(s) to dump.', ParameterSetName='Standard')]
    [Parameter(ParameterSetName='FolderFilter')]
    [string]$Name = '*',
	[Parameter(Mandatory=$True, HelpMessage = 'The folder to dump collection(s) from.', ParameterSetName='FolderFilter')]
    [string]$Folder,
    [Parameter(Mandatory=$False, HelpMessage = 'Recurse sub-folders when dumping collections', ParameterSetName='FolderFilter')]
    [switch]$Recurse,
    [Parameter(Mandatory=$False, HelpMessage = 'The site server to connect to.', ParameterSetName='FolderFilter')]
    [string]$SiteServer = $env:computername

)

function Get-CMFTWCollectionMembership
{
    param
    (
	    [Parameter(Mandatory=$True)]
        [Object]$Collection,
        [Parameter(Mandatory=$False)]
        [string]$CollectionPath
    )

   
    $collectionjson = @"
    {
        "name":"",
        "limitingCollection":"",
        "schedule":"",
        "incremental":"",
        "queryRules":[
        ],
        "includeRules":"",
        "excludeRules":"",
        "directRules":""
    }
    
"@


    $collectionObject = $collectionjson | ConvertFrom-Json

    $collectionObject.name = $Collection.Name
    $collectionObject.limitingCollection = $Collection.LimitToCollectionName
    $collectionObject.schedule = "weekly"
    $collectionObject.incremental = "no"

    $includeRules = New-Object System.Collections.ArrayList("")
    $excludeRules = New-Object System.Collections.ArrayList("")
    $directRules = New-Object System.Collections.ArrayList("")

    foreach($rule in $Collection.CollectionRules)
    {
        if($rule.SmsProviderObjectPath -eq 'SMS_CollectionRuleQuery')
        {
            $queryRule = [PSCustomObject]@{
                $($rule.RuleName) = $($rule.QueryExpression)
            }

            $collectionObject.queryRules += $queryRule
        }
        elseif($rule.SmsProviderObjectPath -eq 'SMS_CollectionRuleIncludeCollection')
        {
            $includeRules.Add($rule.RuleName) | Out-Null
            
            #Write-Host " + ($($rule.IncludeCollectionID)) $($rule.RuleName)"
        }
        elseif($rule.SmsProviderObjectPath -eq 'SMS_CollectionRuleExcludeCollection')
        {
            $excludeRules.Add($rule.RuleName) | Out-Null

            #Write-Host " x ($($rule.ExcludeCollectionID)) $($rule.RuleName)"
        }
        elseif($rule.SmsProviderObjectPath -eq 'SMS_CollectionRuleDirect')
        {
            $directRules.Add($rule.RuleName) | Out-Null
        }
    }

    $collectionObject.includeRules = $includeRules -join ","
    $collectionObject.excludeRules = $excludeRules -join ","
    $collectionObject.directRules = $directRules -join ","

    $collectionObject

}

function Process-DeviceCollectionsInFolder
{
    param
    (
	    [Parameter(Mandatory=$True)]
        [string]$FolderNodeID,
        [Parameter(Mandatory=$True)]
        [string]$FolderNodeName
    )

    $folderjson = @"
    {
        "name":"",
        "prefix":"",
        "collections":[
        ]
    }
"@

    $folderObject = $folderjson | ConvertFrom-Json

    $memberCollections = Get-WMIObject -Computer $providerSystem -Namespace $providerNamespace -Class SMS_ObjectContainerItem `
        -Filter "ContainerNodeID='$FolderNodeID' And ObjectType='5000'"

    $folderObject.name = $FolderNodeName

    foreach($coll in $memberCollections)
    {
        $collection = Get-CMDeviceCollection -Id $coll.InstanceKey

        if($collection.Name -like "$Name")
        {
            $collJson = Get-CMFTWCollectionMembership -Collection $collection -CollectionPath $FolderNodeName
            $folderObject.collections += $collJson
        }
    }

    $outputJson.defaultitems.devicecollectionfolders += $folderObject

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

    $json = @"
    {
        "defaultitems":{
            "variables":[
	            {
		            "organizationName":"ConfigMgrFTW"
                }
            ],
            "schedules":[
	            {
		            "name":"daily",
		            "recurcount":"1",
		            "recurinterval":"days",
		            "start":"06:00"
	            },
	            {
		            "name":"weekly",
		            "start":"07:00",
		            "dayofweek":"Sunday"
	            },
	            {
		            "name":"simpledaily",
		            "recurcount":"1",
		            "recurinterval":"days"
	            },
	            {
		            "name":"patchtuesday",
		            "dayofweek":"Tuesday",
		            "WeekOrder":"2",
		            "start":"20:00"
	            }
            ],
            "devicecollectionfolders":[

            ]
        }
    }
 
"@

Import-Module (Join-Path (Split-Path $env:SMS_ADMIN_UI_PATH –parent) ConfigurationManager.psd1)
$siteCode = Get-PSDrive -PSProvider CMSITE

Push-Location $siteCode":"

$outputJson = $json | ConvertFrom-Json

if($Folder -ne "")
{
    $providerLocation = Get-WMIObject -Computer $SiteServer -Namespace root\sms -Class SMS_ProviderLocation
    $siteCode = $providerLocation.SiteCode
    $providerSystem = $providerLocation.Machine
    $providerNamespace = "root\sms\site_$siteCode"

    $fldr = Get-WMIObject -Computer $providerSystem -Namespace $providerNamespace -Class SMS_ObjectContainerNode -Filter "Name='$Folder' And ObjectType='5000'"

    Process-DeviceCollectionsInFolder -FolderNodeID $fldr.ContainerNodeID -FolderNodeName $Folder

}
else
{
    $allCollections = (Get-CMDeviceCollection | Where-Object {$_.Name -like "$Name"})

    foreach ($collection in $allCollections)
    {
        $collJson = Get-CMFTWCollectionMembership -Collection $collection

        $outputJson.defaultitems.devicecollectionfolders[0].collections += $collJson
    }
}

if($Output -eq "")
{
    $outputJson | ConvertTo-Json -Depth 7
}
else
{
    $outputJson | ConvertTo-Json -Depth 7 | Out-File -FilePath FileSystem::$Output
}
    
Pop-Location
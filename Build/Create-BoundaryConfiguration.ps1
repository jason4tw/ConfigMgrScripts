<#
	.SYNOPSIS
        Creates boundaries and boundary groups in ConfigMgr from a csv file.
	
	.DESCRIPTION
		Creates boundaries and boundary groups in ConfigMgr from a csv file with the following format:
        Prefix,Location,Type,SubnetID,SubnetAddresses,SubnetMask. Boundaries and boundary groups
        with subnets not in the csv will be deleted.
        
        A log file with all activity is generated in the current folder.
	
	.PARAMETER DataFile
        The input csv data file to use. If not specified, .\data.csv is used.
        
    .PARAMETER Prefix
        The prefix to use for boundary and boundary group names.
        
    .PARAMETER Cleanup
        Clean up empty or non-referenced boundaries, boundary groups, and collections.
        
    .PARAMETER ForceCleanup
        Forces the cleanup of boundaries and boundary groups even if they
        contain or are contained in other boundaries or boundary groups.

    .PARAMETER Restore
        Restores collections and boundary groups to their data file defined state by
        removing all collection query rules, boundary group members, and site systems not defined in the data file.

    .PARAMETER Collections
        If specified, also creates corresponding location and type collections.

    .PARAMETER FolderName
        The Admin Console subfolder to place the created device collections in (under Device Collections). 
        The folder specified will be created if it does not exist.
        There is no default value and a folder name must be specified.
        
    .PARAMETER LimitingCollectionID
		The collection ID of the collection to limit the created collections to. SMS00001 (All Systems) is the default.

    .EXAMPLE
        .\Create-BoundaryConfiguration.ps1 -DataFile .\data2.csv
        Creates boundaries and boundary groups defined in the data2.csv data file.
		
	.NOTES
		Version 2.2
        Jason Sandys

        Version History
        - 2.2 (20 December 2019):
            - Added ability to add comments to the category based collection.
        - 2.1 (17 December 2019):
            - Updated restore functionality to list boundaries in boundary groups and query rules in collections that don't
              exist in data file.
            - Updated boundary group and collection creation to not delete and re-create rules in the data file
            - Added site system addition to boundary groups and restore.
        - 2.0 (15 December 2019): Major overhaul including:
            - Dynamic column consumption
            - COllection folder creation and cleanup
            - Collection creation and cleanup
		- 1.0 (12 December 2019): Initial Version

        Dependencies:
        - ConfigurationManager PowerShell module and Admin console loaded locally.
        - Indented.Net.IP PowerShell module installed.
        
        Limitations and Issues
		- None.
#>

[CmdletBinding()]
param
(
	[Parameter(ParameterSetName='No Collections', HelpMessage = 'The data file to use.')]
    [Parameter(ParameterSetName = "Create Collections")]
	[ValidateScript({ Test-Path -PathType Leaf -Path $_ })]
	[Alias('data')]
    [string] $DataFile = '.\data.csv',

    [Parameter(ParameterSetName='No Collections', HelpMessage = 'The prefix to use for boundary and boundary group names.')]
    [Parameter(ParameterSetName = "Create Collections")]
    [string] $Prefix = 'Auto: ',

    [Parameter(ParameterSetName='No Collections', HelpMessage = 'Clean up empty or non-referenced boundaries, boundary groups, and collections.')]
    [Parameter(ParameterSetName = "Create Collections")]
    [switch] $Cleanup,

	[Parameter(ParameterSetName='No Collections', HelpMessage = 'Forces the cleanup of boundaries and boundary groups even if they contain or are contained in other boundaries or boundary groups.')]
    [Parameter(ParameterSetName = "Create Collections")]
    [switch] $ForceCleanup,

    [Parameter(ParameterSetName='No Collections', HelpMessage = 'Restores collections and boundary groups to their data file defined state by
        removing all collection query rules and boundary group members not defined in the data file.')]
    [Parameter(ParameterSetName = "Create Collections")]
    [switch] $Restore,

    [Parameter(ParameterSetName='Create Collections', Mandatory=$true, HelpMessage = 'If specified, also creates collections.')]
    [switch] $Collections,
    
    [Parameter(ParameterSetName='Create Collections', Mandatory=$true, HelpMessage = 'The folder to place the created collections in. The folder specified will be created in it does not exist.')]
    [Alias('folder')]
    [string] $FolderName,

    [Parameter(ParameterSetName='Create Collections', HelpMessage = 'The collection ID of the collection to limit the created collections to.')]
    [string] $LimitingCollectionID = 'SMS00001',

    [switch] $WhatIf
)

function Read-SubnetInfo
{
    [CmdletBinding()]
    param (
        [Parameter(ValueFromPipeline=$true)]
        [object] $data,
        [hashtable] $Subnets,
        [hashtable] $SiteSystems,
        [hashtable] $Columns,
        [string] $KeyColumn,
        [hashtable] $CategoryComments
    )

    process
    {
        $overlap = $false
        $subnetID = ($data.SubnetID).Trim()
        $subnetMask = ($data.SubnetMask).Trim()
        $siteSystemsForItem = ($data.SiteSystems).Trim()

        $subnetInfo = Get-NetworkSummary -IPAddress $subnetID -SubnetMask $subnetMask

        foreach ($col in $additionalColumns)
        {
            $value = $data.$col

            if($columnFilters.Contains($col))
            {
                $value = Invoke-Command -NoNewScope -ScriptBlock ([Scriptblock]::Create($columnFilters.Item($col)))
            }
                
            $subnetInfo | Add-Member -MemberType NoteProperty -Name $col -Value $value.Trim()
        }

        foreach ($col in $Columns.Keys)
        {
            $value = $data.$col

            if($columnFilters.Contains($col))
            {
                $value = Invoke-Command -NoNewScope -ScriptBlock ([Scriptblock]::Create($columnFilters.Item($col)))
            }

            $subnetInfo | Add-Member -MemberType NoteProperty -Name $col -Value $value.Trim()
        }

        foreach($previousSubnet in $Subnets.Keys)
        {
            $previousSubnetInfo = $Subnets.Item($previousSubnet)

            if(($subnetInfo.NetworkDecimal -ge $previousSubnetInfo.NetworkDecimal -and $subnetInfo.NetworkDecimal -le $previousSubnetInfo.BroadcastDecimal) `
                 -or ($subnetInfo.BroadcastDecimal -ge $previousSubnetInfo.NetworkDecimal -and $subnetInfo.BroadcastDecimal -le $previousSubnetInfo.BroadcastDecimal) `
                 -or ($previousSubnetInfo.NetworkDecimal -ge $subnetInfo.NetworkDecimal -and $previousSubnetInfo.NetworkDecimal -le $subnetInfo.BroadcastDecimal) `
                 -or ($previousSubnetInfo.BroadcastDecimal -ge $subnetInfo.NetworkDecimal -and $previousSubnetInfo.BroadcastDecimal -le $subnetInfo.BroadcastDecimal))
            {
                Write-Warning " Overlapping subnets found: '$subnetID' at '$($subnetInfo.$KeyColumn))' overlaps with '$($previousSubnetInfo.NetworkAddress)' at '$($previousSubnetInfo.$KeyColumn)'."
                $overlap = $true
            }
        }

        if($Subnets.Contains($subnetInfo.CIDRNotation))
        {
            Write-Warning " Duplicate subnet found: '$subnetID' at '$($subnetInfo.$KeyColumn)'"
        }
        elseif($overlap -ne $true)
        {
            $Subnets.Add($subnetInfo.CIDRNotation, $subnetInfo)

            foreach($col in $Columns.Keys)
            {
                if(($Columns.$col).Contains($subnetInfo.$col))
                {
                    ($Columns.$col.($subnetInfo.$col)) += ",$($subnetInfo.CIDRNotation)"
                }
                else
                {
                    ($Columns.$col).Add($subnetInfo.$col, $subnetInfo.CIDRNotation)
                }

                if($commentsByCategory.Contains($col))
                {
                    if(($CategoryComments.$col).Contains($subnetInfo.$col))
                    {
                        ($CategoryComments.$col).($subnetInfo.$col) += "$commentSeperator$($subnetInfo.($commentsByCategory.$col))"
                    }
                    else
                    {
                        ($CategoryComments.$col).Add($subnetInfo.$col, $subnetInfo.($commentsByCategory.$col))        
                    }
                }

                if($siteSystemsForItem.Length -gt 0)
                {
                    if(($SiteSystems.$col).Contains($subnetInfo.$col))
                    {
                        ($SiteSystems.$col.($subnetInfo.$col)) += ",$siteSystemsForItem"
                    }
                    else
                    {
                        ($SiteSystems.$col).Add($subnetInfo.$col, $siteSystemsForItem)
                    }
                }
            }
        }
    }
}

function New-Boundary
{
[CmdletBinding()]
param (
    [Parameter(ValueFromPipeline=$true)]
    [object] $Subnet
)
    process
    {
        $boundaryName = $Prefix + $ExecutionContext.InvokeCommand.ExpandString($boundaryNameTemplate)

        $range = $Subnet.Value.HostRange -replace ' ', ''

        $boundary = Get-CMBoundary -BoundaryName $boundaryName

        if(-not($boundary))
        {
            Write-Host " + Creating boundary: $boundaryName ($($Subnet.Value.CIDRNotation) = $range)"
            
            try
            {
                $boundary = New-CMBoundary -Name $boundaryName -Type IPRange -Value $range -WhatIf:$WhatIf
            }
            catch
            {
                Write-Warning " Could not create boundary."
                return
            }
        }
        else
        {
            Write-Host " = Boundary already exists: $boundaryName"
        }

        $boundary
    }
}

function Invoke-BoundaryCheck
{
    [CmdletBinding()]
    param (
        [hashtable] $Subnets,
        [switch] $Clean,
        [switch] $ForceClean
    )  
    
    Get-CMBoundary | Where-Object { $_.DisplayName -like "$Prefix*" } `
    | ForEach-Object {
        $start,$end = $_.Value -split '-'
        $boundarySubnet = ConvertTo-Subnet -Start $start -End $end

        if(-not($Subnets.Contains("$($boundarySubnet.NetworkAddress)/$($boundarySubnet.MaskLength)")))
        {
            if(($Clean) -and `
                ($_.GroupCount -eq 0 -or $ForceClean))
            {
                if($_.GroupCount -eq 0)
                {
                    Write-Host " x Deleting $($_.DisplayName) as it does not exist in the data file and is not a member of any boundary groups"
                }
                else
                {
                    Write-Host " x Deleting $($_.DisplayName) as it does not exist in the data file and ForceCleanup was specified."
                }

                Remove-CMBoundary -Name $_.DisplayName -Force -WhatIf:$WhatIf
            }
            else
            {
                Write-Host " * $($_.DisplayName) does not exist in the data file and is a member of $($_.GroupCount) boundary groups"            
            }
        }    
    }
}

function New-BoundaryGroup
{
[CmdletBinding()]
param (
    [Parameter(ValueFromPipeline=$true)]
    [System.Collections.DictionaryEntry] $Item,
    [string] $BoundaryGroupCategory,
    [hashtable] $Boundaries,
    [hashtable] $SiteSystems
)

    process
    {
        $boundaryGroupName = $ExecutionContext.InvokeCommand.ExpandString($boundaryGroupNamePrefix) + $Item.Name

        $boundaryGroup = Get-CMBoundaryGroup -Name $boundaryGroupName
        
        if(-not($boundaryGroup))
        {
            Write-Host " + Creating boundary group: $boundaryGroupName"
            $boundaryGroup = New-CMBoundaryGroup -Name $boundaryGroupName -WhatIf:$WhatIf
        }
        else
        {
            Write-Host " = Boundary group already exists: $boundaryGroupName"
        }

        if($boundaryGroup)
        {
            $currentBoundaries = (Get-CMBoundary -BoundaryGroupName $boundaryGroupName).DisplayName
            $autoBoundaries = New-Object -TypeName "System.Collections.ArrayList"

            $categorySubnets = $Item.Value -split ',' 

            foreach ($subnet in $categorySubnets)
            {
                if($Boundaries.ContainsKey($subnet))
                {
                    if($currentBoundaries -notcontains ($Boundaries.Item($subnet)).DisplayName)
                    {
                        Write-Host "   + Adding boundary '$(($Boundaries.Item($subnet)).DisplayName)' to boundary group"
                        Add-CMBoundaryToGroup -BoundaryGroupId $boundaryGroup.GroupId -BoundaryId ($Boundaries.Item($subnet)).BoundaryID -WhatIf:$WhatIf
                    }
                    else 
                    {
                        Write-Host "   = Boundary '$(($Boundaries.Item($subnet)).DisplayName)' already exists in boundary group"
                    }

                    $autoBoundaries.Add(($Boundaries.Item($subnet)).DisplayName) | Out-Null
                }
                else
                {
                    Write-Warning "   ! Boundary not found for subnet '$subnet'"
                }
            }

            foreach($boundaryName in $currentBoundaries)
            {
                if($autoBoundaries -notcontains $boundaryName)
                {
                    if($Restore)
                    {
                        Write-Host "   x Removing boundary '$boundaryName' as it exists in boundary group but not the data file"
                        Remove-CMBoundaryFromGroup -BoundaryGroupId $boundaryGroup.GroupID -BoundaryName $boundaryName -Force
                    }
                    else
                    {
                        Write-Host "   ~ Boundary '$boundaryName' exists in boundary group but not the data file"
                    }
                }
            }

            $currentSiteSystems = ((Get-WmiObject -Namespace "root\sms\site_$siteCode" `
             -Class SMS_BoundaryGroupSiteSystems `
             -Filter  "GroupID='$($boundaryGroup.GroupID)'").ServerNALPath)
            
            if($currentSiteSystems)
            {
                $currentSiteSystems = $currentSiteSystems | ForEach-Object { ($_ -split '\\')[5] }
            } 

            $desiredSiteSystems = ($SiteSystems.($Item.Name) -split ',' | Sort-Object -Unique)

            if(($SiteSystems.($Item.Name)).Length -gt 0)
            {
                foreach($siteSystem in $desiredSiteSystem)
                {
                    if($currentSiteSystems -notcontains $siteSystem)
                    {
                        Write-Host "   + Adding Site System '$siteSystem' to the boundary group"
                        Set-CMBoundaryGroup -Id $boundaryGroup.GroupID -AddSiteSystemServerName $siteSystem

                    }
                    else
                    {
                        Write-Host "   = Site System '$siteSystem' already exists in boundary group"
                    }
                }
            }

            foreach($siteSystem in $currentSiteSystems)
            {
                if($desiredSiteSystems -notcontains $siteSystem)
                {
                    if($Restore)
                    {
                        Write-Host "   - Removing Site System '$siteSystem' exists in boundary group but not the data file"                        
                        Set-CMBoundaryGroup -Id $boundaryGroup.GroupID -RemoveSiteSystemServerName $siteSystem
                    }
                    else
                    {
                        Write-Host "   ~ Site System '$siteSystem' exists in boundary group but not the data file"                        
                    }
                }
            }

            $boundaryGroup
        }
    }
}

function Invoke-BoundaryGroupCheck
{
    [CmdletBinding()]
    param (
        [hashtable] $BoundaryGroups,
        [switch] $Clean,
        [switch] $ForceClean
    )
    Get-CMBoundaryGroup | Where-Object { $_.Name -like "$Prefix*" } `
     | ForEach-Object {
        $boundaryGroupName = $_.Name
        if(-not($BoundaryGroups.Contains($boundaryGroupName)))
        {
            $boundaryGroup = $_

            if($Clean)
            {
                Get-CMBoundary -BoundaryGroupId $boundaryGroup.GroupID | Where-Object {$_.DisplayName -like "$Prefix*"} `
                 | ForEach-Object { Remove-CMBoundaryFromGroup -BoundaryId $_.BoundaryID -BoundaryGroupId $boundaryGroup.GroupID -Force -WhatIf:$WhatIf }

                $boundaryGroup = Get-CMBoundaryGroup -Id $boundaryGroup.GroupID

                if($boundaryGroup.MemberCount -eq 0 -or $ForceCleanup)
                {
                    if($boundaryGroup.MemberCount -eq 0)
                    {
                        Write-Host " x Deleting $($_.Name) as it does not exist in the data file and contains no boundaries"
                    }
                    else
                    {
                        Write-Host " x Deleting $($_.Name) as it does not exist in the data file and ForceCleanup was specified"                    
                    }
                    Remove-CMBoundaryGroup -Name $_.Name -Force -WhatIf:$WhatIf
                }
                else
                {
                    Write-Host " * $($_.Name) does not exist in the data file but contains $($boundaryGroup.MemberCount) other boundaries"
                }
            }
            else
            {
                Write-Host " * $($_.Name) does not exist in the data file but contains $($boundaryGroup.MemberCount) boundaries"
            }
        }    
    }
}

function New-ConsoleFolder
{

	[CmdletBinding()]
	param
	(
        [Parameter(ValueFromPipeline=$true)]
        [string] $FolderName,
        [string] $ParentPath
	)

	$fullFolderPath = Join-Path -Path $ParentPath -ChildPath $FolderName

	if(-not (Test-Path -Path $fullFolderPath -PathType Container))
	{
		Write-Host " + Creating folder: '$FolderName' in '$ParentPath' ..."
	
        $folder = New-Item -Path $ParentPath -Name $FolderName -ItemType Directory -WhatIf:$WhatIf
        $folder = Get-Item -Path (Join-Path -Path $ParentPath -ChildPath $FolderName)
	}
	else 
	{
        Write-Host " = Folder already exists: '$fullFolderPath'"
        $folder = Get-Item -Path $fullFolderPath
    }
    
    $folder
}

function New-Collection
{
	[CmdletBinding()]
	param
	(
        [Parameter(ValueFromPipeline=$true)]
        [System.Collections.DictionaryEntry] $Item,
        [string] $ItemType,
        [string] $Comments
	)

    process
    {
        $collectionName = $ExecutionContext.InvokeCommand.ExpandString($collectionNamePrefix) + $Item.Name

        $collection = Get-CMCollection -Name $collectionName

        if(-not($collection))
        {
            Write-Host " + Creating new collection: $collectionName ..."
            if($WhatIf -ne $true)
            {
                $collection = New-CMDeviceCollection `
                    -Name $collectionName `
                    -LimitingCollectionId $LimitingCollectionID `
                    -RefreshType Periodic `
                    -RefreshSchedule $updateSchedule
            }
        }
        else
        {
            Write-Host " = Collection already exists: $collectionName ..."
        }

        if($Comments.Length -gt 0)
        {
            Write-Host "   & Settings comments to '$Comments' ..."
            Set-CMCollection -InputObject $collection -Comment $Comments
        }

        $collection
    }
}

function Move-Collection
{
	[CmdletBinding()]
	param
	(
        [Parameter(ValueFromPipeline=$true)]
        [object] $Collection,
        [object] $ConsoleFolder

	)

    process
    {
        $collectionQuery = "SELECT InstanceKey FROM SMS_ObjectContainerItem WHERE ObjectType=""5000"" And ContainerNodeId=""$($ConsoleFolder.ContainerNodeID)"" And InstanceKey=""$($Collection.CollectionID)"""

        $collectionInFolder = (@(Get-WMIObject -Namespace "root\sms\site_$siteCode" -Query $collectionQuery).Count -eq 1)
        
        $y,$z,$consoleFolderPath = $ConsoleFolder.PSPath -Split '\\'

        $consoleFolderPath = "${siteCode}:\$($consoleFolderPath -join '\')"

        if((Test-Path -Path $consoleFolderPath -PathType Container) -and (-not $collectionInFolder))
        {
            Write-Host "   > Moving '$($Collection.Name)' to '$consoleFolderPath'..."
            Move-CMObject -FolderPath $consoleFolderPath -InputObject $Collection -WhatIf:$WhatIf
        }

        $Collection
    }
}

function Add-BoundaryGroupQueryRuleToCollection
{
	[CmdletBinding()]
	param
	(
        [Parameter(ValueFromPipeline=$true)]
        [object] $Collection,
		[object] $BoundaryGroup
	)

    $currentRules = (Get-CMDeviceCollectionQueryMembershipRule -CollectionId $Collection.CollectionID).RuleName

    if($currentRules -notcontains $BoundaryGroup.Name)
    {
        Write-Host "   + Adding query rule for '$($BoundaryGroup.Name)'"
        #boundaryGroupID is referenced in $queryTemplate
        $boundaryGroupID = $BoundaryGroup.GroupID

        $queryRule = $ExecutionContext.InvokeCommand.ExpandString($queryTemplate)
        Add-CMDeviceCollectionQueryMembershipRule `
            -CollectionId $Collection.CollectionID `
            -RuleName $BoundaryGroup.Name `
            -QueryExpression $queryRule
    }
    else
    {
        Write-Host "   = Query rule for '$($BoundaryGroup.Name)' already exists" 
    }

    foreach ($rule in $currentRules)
    {
        if($rule -ne $BoundaryGroup.Name)
        {
            if($Restore)
            {
                Write-Host "   x Removing query rule named '$rule' as it exists but is not in the data file"             
                Remove-CMDeviceCollectionQueryMembershipRule `
                 -CollectionId $Collection.CollectionID `
                 -RuleName $rule `
                 -Force
            }
            else
            {
                Write-Host "   ~ Query rule named '$rule' exists but not in the data file"             
            }
        }
    }
}

function Invoke-CollectionCheck
{
	[CmdletBinding()]
	param
	(
		[object] $ConsoleFolder,
        [hashtable] $ValidBoundaryGroups,
        [string] $ItemType,
        [switch] $Clean
	)

	$collectionQuery = "SELECT InstanceKey FROM SMS_ObjectContainerItem WHERE ObjectType=""5000"" And ContainerNodeId=""$($ConsoleFolder.ContainerNodeID)"""

	$collectionsInFolder = Get-WMIObject -Namespace "root\sms\site_$siteCode" -Query $collectionQuery

	foreach($collectionItem in $collectionsInFolder)
	{
		$collection = Get-CMDeviceCollection -Id $collectionItem.InstanceKey
		if($collection)
		{
            $prefixToReplace = $ExecutionContext.InvokeCommand.ExpandString($collectionNamePrefix)
            $collectionName = $collection.Name -replace "$prefixToReplace", ""

			if(-not($ValidBoundaryGroups.Contains($collectionName)))
            {
                if($Clean)
                {
                    Write-Host " x Deleting collection: $($collection.Name) ..."
                    Remove-CMDeviceCollection -Id $collection.CollectionID -Force -WhatIf:$WhatIf
                }
                else
                {
                    Write-Host " * Collection $($collection.Name) does not exist in the data file"
                }
			}
		}
	}
}

function Invoke-FolderCheck
{
    [CmdletBinding()]
	param
	(
        [hashtable] $Folders,
        [object] $ConsoleFolder,
        [switch] $Clean,
        [switch] $ForceClean
    )

    $y,$z,$consoleFolderPath = $ConsoleFolder.PSPath -Split '\\'
    $consoleFolderPath = "${siteCode}:\$($consoleFolderPath -join '\')"

    foreach($childFolder in (Get-ChildItem -Path $consoleFolderPath))
    {
        if(-not($Folders.Contains($childFolder.Name)))
        {
            $childFolderPath = Join-Path -Path $consoleFolderPath -ChildPath $childFolder.Name
            $collectionQuery = "SELECT InstanceKey FROM SMS_ObjectContainerItem WHERE ObjectType=""5000"" And ContainerNodeId=""$($childFolder.ContainerNodeID)"""
            $collectionsInFolder = Get-WMIObject -Namespace "root\sms\site_$siteCode" -Query $collectionQuery
            $collectionsInFolderCount = (@($collectionsInFolder).Count)
            
            if($Clean)
            {
                #$ItemType is reference in $CollectionNamePrefix
                $ItemType = $childFolder.Name
                $collectionNamePrefixToFind = $ExecutionContext.InvokeCommand.ExpandString($collectionNamePrefix)
                
                foreach($collection in $collectionsInFolder)
                {
                    if((Get-CMDeviceCollection -CollectionId $collection.InstanceKey).Name -like "$collectionNamePrefixToFind*")
                    {
                        Remove-CMDeviceCollection -CollectionId $collection.InstanceKey -Force
                    }
                }

                $collectionsInFolder = Get-WMIObject -Namespace "root\sms\site_$siteCode" -Query $collectionQuery
                $collectionsInFolderCount = (@($collectionsInFolder).Count)

                if($collectionsInFolderCount -eq 0 -or $ForceClean)
                {
                    if($collectionsInFolderCount -eq 0)
                    {
                        Write-Host " x Deleting folder '$($childFolder.Name)' as it contained no other collections ..."
                    }
                    else
                    {
                        Write-Host " x Deleting folder '$($childFolder.Name)' and all collections in it as ForceCheck was specified ..."
                        $collectionsInFolder | ForEach-Object { Remove-CMDeviceCollection -CollectionId $_.InstanceKey -Force }
                    }

                    Remove-Item -Path $childFolderPath -Force -WhatIf:$WhatIf
                }
                else
                {
                    Write-Host " * Folder $($childFolder.Name) does not exist in the data file but contains $collectionsInFolderCount other collections"                    
                }
            }
            else
            {
                Write-Host " * Folder $($childFolder.Name) does not exist in the data file and contains $collectionsInFolderCount collections"
            }
        }
    }
}

if($WhatIf -eq $true)
{
	Write-Host -ForegroundColor DarkBlue -BackgroundColor Yellow "Operating in WhatIf mode, no changes will be made."
}

#Load the Indented.Net.IP module
if (-not (Import-Module -Name Indented.Net.IP -PassThru))
{
    exit 1
}

#Load Configuration Manager PowerShell Module
if(-not(Import-Module -Name ($Env:SMS_ADMIN_UI_PATH.Substring(0, $Env:SMS_ADMIN_UI_PATH.Length - 5) + '\ConfigurationManager.psd1') -PassThru))
{
    exit 1
}

#Get SiteCode
$siteCode = Get-PSDrive -PSProvider CMSITE

$subnets = @{}
$boundaries = @{}
$boundaryGroups = @{}

# Configuration **************************************************************************************
$boundaryNameTemplate = '$($Subnet.Value.Location) ($($Subnet.Value.SubnetAddresses))'
$boundaryGroupNamePrefix = '$Prefix$BoundaryGroupCategory '
$collectionNamePrefix = '${ItemType}: '
$keyCategory = 'Location'
$boundaryGroupCategoryNames = 'Location','Type'
$additionalColumns = 'SubnetAddresses'
$columnFilters = @{'Type' = '("$value" -split "/")[0]'}
$queryTemplate = 'select SMS_R_System.ResourceId, SMS_R_System.ResourceType, SMS_R_System.Name, SMS_R_System.SMSUniqueIdentifier, SMS_R_System.ResourceDomainORWorkgroup, SMS_R_System.Client, SMS_G_System_BOUNDARYGROUPCACHE.BoundaryGroupIDs from  SMS_R_System inner join SMS_G_System_BOUNDARYGROUPCACHE on SMS_G_System_BOUNDARYGROUPCACHE.ResourceID = SMS_R_System.ResourceId where SMS_G_System_BOUNDARYGROUPCACHE.BoundaryGroupIDs like "%$boundaryGroupID%"'
$commentsByCategory = @{'Location' = 'Type'}
$commentSeperator = ', '
# End Configuration **********************************************************************************

$boundaryGroupCategories = @{}
$siteSystems = @{}

foreach($category in $boundaryGroupCategoryNames)
{
    $boundaryGroupCategories.Add($category, @{})
    $siteSystems.Add($category, @{})
}

$collectionComments = @{}

foreach($category in $commentsByCategory.Keys)
{
    $collectionComments.Add($category, @{})
}

Write-Host "Loading subnet, location, and type information from data file ..."
$config = Import-Csv -Path $DataFile
$config |  Read-SubnetInfo -Subnets $subnets `
 -Columns $boundaryGroupCategories `
 -KeyColumn $keyCategory `
 -SiteSystems $siteSystems `
 -CategoryComments $collectionComments

Push-Location $siteCode":"

Write-Host ""
Write-Host "Processing boundaries from subnets ..."
$subnets.GetEnumerator() | ForEach-Object { $subnetItem = $_; New-Boundary -Subnet $subnetItem  } `
 | ForEach-Object { $boundaries.Add($subnetItem.Value.CIDRNotation, $_) }

foreach($category in $boundaryGroupCategories.Keys)
{
    Write-Host ""
    Write-Host "Processing boundary groups based on $category ..."
    ($boundaryGroupCategories.$category).GetEnumerator() `
     | New-BoundaryGroup -BoundaryGroupCategory $category -Boundaries $boundaries -SiteSystems $siteSystems.$category `
     | ForEach-Object { $boundaryGroups.Add($_.Name, $_) }
}

Write-Host ""
Write-Host "Checking for stale boundary groups ..."
Invoke-BoundaryGroupCheck -BoundaryGroups $boundaryGroups -Clean:$Cleanup -ForceClean:$ForceCleanup

Write-Host ""
Write-Host "Checking for stale boundaries ..."
Invoke-BoundaryCheck -Subnets $subnets -Clean:$Cleanup -ForceClean:$ForceCleanup

if($Collections)
{
    $categoryCollectionFolders = @{}

    Write-Host ""
    Write-Host "Processing console folders for device collections ..."

    $updateSchedule = New-CMSchedule -Start (Get-Date -Hour 12 -Minute 0 -Second 0) -RecurCount 1 -RecurInterval Days
    $rootDeviceCollectionPath = "${siteCode}:\DeviceCollection"

    $targetFolder = New-ConsoleFolder -ParentPath $rootDeviceCollectionPath -FolderName $FolderName
    $targetFolderPath = Join-Path -Path $rootDeviceCollectionPath -ChildPath $FolderName
    
    if($targetFolder)
    {
        $boundaryGroupCategories.Keys `
         | ForEach-Object { $category = $_; New-ConsoleFolder -ParentPath $targetFolderPath -FolderName $category } `
         | ForEach-Object { $categoryCollectionFolders.Add($category, $_) }

        foreach($boundaryGroupCategory in $boundaryGroupCategories.Keys)
        {
            Write-Host ""
            Write-Host "Processing collections based on $boundaryGroupCategory ..."

            ($boundaryGroupCategories.$boundaryGroupCategory).GetEnumerator() `
             | ForEach-Object { $boundaryGroupName = "$Prefix$boundaryGroupCategory $($_.Name)"; `
             New-Collection -Item $_ -ItemType $boundaryGroupCategory -Comments $collectionComments.$boundaryGroupCategory.($_.Name) `
             | Move-Collection -ConsoleFolder $categoryCollectionFolders.$boundaryGroupCategory `
             | Add-BoundaryGroupQueryRuleToCollection -BoundaryGroup $boundaryGroups.Item($boundaryGroupName) }
        }
    }

    Write-Host ""
    Write-Host "Checking for stale collections ..."
    foreach($category in $boundaryGroupCategories.Keys)
    {
        Invoke-CollectionCheck -ConsoleFolder $categoryCollectionFolders.$category -ItemType $category -ValidBoundaryGroups $boundaryGroupCategories.$category -Clean:$Cleanup
    }

    Write-Host ""
    Write-Host "Checking for stale collection folders ..."
    Invoke-FolderCheck -ConsoleFolder $targetFolder -Folders $categoryCollectionFolders -Clean:$Cleanup -ForceClean:$ForceCleanup
}

Pop-Location
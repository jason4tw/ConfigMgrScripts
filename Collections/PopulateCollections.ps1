<#
	.DESCRIPTION
        Populates a set of collections with direct membership rules from a set of text files.
        
        A JSON-based configuration file maps the text files to the collections whose membership
        each defines.

        All existing direct membership rules are removed from the collection first.
        Use the -DoNotClean option to prevent this.
	
	.PARAMETER ConfigFile
        The input json configuration file to use. If not specified, .\config.json is used.

    .PARAMETER CollectionSet
        Specifies which set of collections defined in the JSON configuration file to modify.
        If not specified, all collection sets are updated.

    .PARAMETER CollectionName
        Specifies a single collection defined in the JSON configuration file to modify.
        If not specified, all collections in the specified collections sets are updated.

    .PARAMETER DoNotClean
        If specified, the script does not remove existing direct membership rules first.

    .EXAMPLE
        .\PopulateCollections.ps1
        Updates the collection membership of the all collections specified in the configuration file.

    .EXAMPLE
        .\PopulateCollections.ps1 -CollectionSet Updates
        Updates the collection membership of the Updates collections specified in the configuration file.

    .EXAMPLE
        .\PopulateCollections.ps1 -ConfigFile .\orgconfig.json -CollectionSet Updates
        Updates the collection membership of the Updates collections specified in the 'orgconfig.json' configuration file.

    .EXAMPLE
        .\PopulateCollections.ps1 -CollectionSet Software -DoNotClean
        Updates the collection membership of the Software collections without removing existing direct membership rules.

		
	.NOTES
		Version 1.00
        Jason Sandys
        Version History
        - 1.00 (24 February 2020): Initial Version

#>

[CmdletBinding()]
param
(
    [Parameter(HelpMessage = 'The configuration file to use.')]
    [ValidateScript( { Test-Path -PathType Leaf -Path $_ })]
    [Alias('configuration')]
    [Alias('config')]
    [string] $ConfigFile = '.\config.json',
    [string] $CollectionSet,
    [string] $CollectionName,
    [switch] $DoNotClean
)
if ($PSCommandPath -eq $null) { function GetPSCommandPath() { return $MyInvocation.ScriptName; } $PSCommandPath = GetPSCommandPath; }

#Load Configuration Manager PowerShell Module
if (-not(Import-Module -Name ($Env:SMS_ADMIN_UI_PATH.Substring(0, $Env:SMS_ADMIN_UI_PATH.Length - 5) + '\ConfigurationManager.psd1') -PassThru))
{
    Write-Error "Could not load Configuration Manager module."
    exit 1
}

Write-Host "Loaded Configuration Manager Module"
Write-Host ""

#Get SiteCode
$siteCode = Get-PSDrive -PSProvider CMSITE

$config = Get-Content -Path $ConfigFile -Raw | ConvertFrom-Json

if($CollectionSet)
{
    $collectionSetToProcess = $config.CollectionSets | Where-Object -FilterScript { $_.name -eq $CollectionSet }
}
else
{
    $collectionSetToProcess = $config.CollectionSets
}

if ($CollectionName)
{
    $collectionsToProcess = $collectionSetToProcess.collections | Where-Object -FilterScript { $_.name -eq $CollectionName }
}
else
{
    $collectionsToProcess = $collectionSetToProcess.collections   
}

$collectionTotalCount = @($collectionsToProcess).count
$collectionCount = 0

if ($collectionTotalCount -gt 0)
{
    Push-Location -Path "${siteCode}:"

    Write-Progress -Id 1 -Activity "Processing Collections" -PercentComplete 0

    foreach ($collection in $collectionsToProcess)
    {
        $coll = Get-CMDeviceCollection -Name $collection.name
        $collectionCount++
        $percentComplete = ($collectionCount / $collectionTotalCount) * 100

        if($coll)
        {
            Write-Progress -Id 1 -Activity "Processing Collections" -PercentComplete $percentComplete `
                -Status "$($collection.name) [$collectionCount of $collectionTotalCount]"

            $collectionFilepath = 'FileSystem::' + (Join-Path -Path (Split-Path -Path $PSCommandPath -Parent) -ChildPath $collection.filename)

            Write-Host $collectionFilepath

            if (Test-Path -Path $collectionFilepath -PathType Leaf)
            {
                if(-not $DoNotClean)
                {
                    Write-Progress -Id 1 -Activity "Processing Collections" -PercentComplete $percentComplete `
                        -Status "$($collection.name) [$collectionCount of $collectionTotalCount]" `
                        -CurrentOperation "Removing existing direct membership rules"

                    Get-CMDeviceCollectionDirectMembershipRule -InputObject $coll | `
                        ForEach-Object -Process { Remove-CMCollectionDirectMembershipRule -InputObject $coll -ResourceId $_.ResourceID -Force }
                }

                Write-Progress -Id 1 -Activity "Processing Collections" -PercentComplete $percentComplete `
                    -Status "$($collection.name) [$collectionCount of $collectionTotalCount]"

                $members = Get-Content -Path $collectionFilepath

                Write-Progress -Id 2 -ParentId 1 -Activity "Adding direct membership rules to $($collection.name)" -PercentComplete 0

                $deviceTotalCount = @($members).Count
                $deviceCount = 0

                $members | ForEach-Object -Process { 
                    $deviceCount++

                    if($_ -ne '' -and $_ -ne $null)
                    {
                            $device = Get-CMDevice -Name $_

                        if ($device)
                        {
                            Write-Progress -Id 2 -ParentId 1 -Activity "Adding direct membership rules to $($collection.name)" `
                                -PercentComplete (($deviceCount / $deviceTotalCount) * 100) `
                                -Status "$_ [$deviceCount of $deviceTotalCount]"

                            Add-CMDeviceCollectionDirectMembershipRule -InputObject $coll -Resource $device
                        }
                        else
                        {
                            Write-Warning "Not found. Resource: '$_'  Text file: '$($collection.filename)'  Collection: '$($collection.name)'."                                         
                        }
                    }
                }

                Write-Progress -Id 2 -ParentId 1 -Activity "Adding direct membership rules to $($collection.name)" -PercentComplete 0 -Complete
            }
            else
            {
                Write-Warning "Not found. Text file: '$($collection.filename)'  Collection: '$($collection.name)'."                 
            }
        }
        else
        {
            Write-Warning "Not found. Collection: '$($collection.name)'."    
        }
    }

    Write-Progress -Id 1 -Activity "Processing Collections" -PercentComplete 100 -Completed

    Pop-Location
}

<#
	.SYNOPSIS
		Creates standard objects and configurations in ConfigMgr from a supplied json configuration file.
	
	.DESCRIPTION
        Creates Collection Folders, Collections, Maintenance Windows, Client Settings, Update Packages, and 
        Automatic Deployment rules in ConfigMgr using a json file where these are all defined.
	
	.PARAMETER ConfigFile
        The input json configuration file to use. If not specified, .\config.json is used.
        
	.PARAMETER Collections
        Create or update collections defined in the configuration file.

    .PARAMETER MaintenanceWindows
        Create or update maintenance windows defined on collections in the configuration file.

	.PARAMETER ClientSettings
        Create Client Settings Packages defined in the configuration file.

	.PARAMETER AutomaticDeploymentRules
        Create Automatic Deployment Rules defined in the configuration file.
        
	.PARAMETER UpdatePackages
        Create Update (aka Deployment) Packages defined in the configuration file.

    .PARAMETER MaintenaceWindows
        Create Maintenance Windows defined in the configuration file.
        
	.PARAMETER Recreate
		Delete existing objects with the same names if they exist before trying to create those specified in the
        configuration file. If not specified and an object already exists with the same name, it will be skipped.
        
	.PARAMETER TypeFilter
        Specifies a schedule and template type to filter on in the configuration file.

    .EXAMPLE
        .\Build-CMDefaultConfiguration.ps1 -ConfigFile .\Build-CMDefaultConfiguration.json -Collections
        Creates collections defined in the json configuration file.

    .EXAMPLE
        .\Build-CMDefaultConfiguration.ps1 -ConfigFile .\config.json -Collections -UpdateMembership
        Creates collections defined in the json configuration file if they don't already exist and updates 
        the membership of collections defined in the json configuration file.
		
	.NOTES
		Version 2.0
        Jason Sandys

		Version History
        - 2.0 (12 January 2020): Heavily modified from verison 1.
            - Better template implmentation and json representation
            - Added maintenance windows
            - Added ability to create folder hierarchy
            - Added type filters for schedules and templates

        Limitations and Issues
        - None
#>

[CmdletBinding()]
param
(
	[Parameter(HelpMessage = 'The configuration file to use.')]
	[ValidateScript({ Test-Path -PathType Leaf -Path $_ })]
	[Alias('configuration')]
	[Alias('config')]
	[string] $ConfigFile = '.\config.json',
    [Alias('Type')]
    [string] $TypeFilter,
    [switch] $Collections,
    [Alias('MWs')]
    [switch] $MaintenanceWindows,
    [switch] $ClientSettings,
    [Alias('ADRs')]
	[switch] $AutomaticDeploymentRules,
	[switch] $Recreate,
	[switch] $UpdatePackages,
    [switch] $WhatIf

)

function Invoke-ProcessParameters
{
    param
    (
        [Parameter(Mandatory=$true,ValueFromPipeline)]
        [PSCustomObject] $Item,
        [string] $Filter,
        [switch] $ExcludeNameParam
    )

    if($null -eq $Item._type -or (($Item._type -split ',') -contains  $Filter))
    {
        $parameters = @{}

        $Item | Get-Member -Type NoteProperty | ForEach-Object {
            $paramName = $_.Name
            if($Item.$paramName -is [boolean] -or $Item.$paramName -is [array])
            {
                $paramValue = $Item.$paramName
            }
            else
            {
                $paramValue = $ExecutionContext.InvokeCommand.ExpandString($Item.$paramName).Trim()
            }				

            if((-not ($paramValue -is [string]) -and $null -ne $paramValue) -or `
                ($paramValue -is [string] -and -not ([string]::IsNullOrEmpty($paramValue))))
            {
                switch -wildcard ($paramName)
                {
                    'schedule'
                    {
                        $parameters.Add($paramName, $Schedules.Get_Item("$paramValue"))
                        break
                    }
                    '_*'
                    {
                        break
                    }
                    'true'
                    {
                        $parameters.Add($paramName, $true)
                        break
                    }
                    'false'
                    {
                        $parameters.Add($paramName, $false)
                        break
                    }
                    'name'
                    {
                        if(-not $ExcludeNameParam)
                        {
                            $parameters.Add($paramName, $paramValue)
                        }
                        break
                    }
                    default
                    {
                        $parameters.Add($paramName, $paramValue)
                    }
                }
            }
        }

        $parameters
    }
    else
    {
        $null    
    }
}

function Add-TemplateParameters
{
	param
	(
		[Parameter(Mandatory=$true)]
        [PSCustomObject] $Item,
        [string] $TemplateProperty = "_template",
        [PSCustomObject] $Templates,
        [string] $Filter,
		[array] $ExcludeParams
    )
    
    if($Item.$TemplateProperty -and $Templates)
    {
        $templateReferences = $Templates | `
            Where-Object { $_.name -eq $Item.$TemplateProperty -and ($null -eq $_._type -or ($_._type -split ',') -contains $Filter) }# | `
            #Select-Object -First 1

        foreach($template in $templateReferences)
        {
            $template.settings | Get-Member -Type NoteProperty | ForEach-Object {
                if($null -eq $Item.($_.Name))
                {
                    $Item | Add-Member -MemberType NoteProperty -Name $_.Name -Value $template.settings.($_.Name)
                }
            }
        }
    }
}

function New-CMFTWSchedule
{
	param
	(
		[Parameter(Mandatory=$true,ValueFromPipeline)]
		[PSCustomObject] $ScheduleInfo,
		[string] $Filter
	)

    begin
    {
        $schedulehash = @{}
        Write-Host "Creating schedules ..."
    }

    process
    {
        $parameters = Invoke-ProcessParameters -Item $ScheduleInfo -Filter $Filter -ExcludeNameParam

        if($parameters.Keys.Count -gt 0)
        {
            Write-Host "  + $($ScheduleInfo.name)"

            $schedulehash.Add($ScheduleInfo.name, (New-CMSchedule @parameters))
        }
    }

    end
    {
        Write-Host ""
        $schedulehash
    }
}

function New-CMFTWDeviceCollectionFolder
{
param
(
    [Parameter(Mandatory=$true)]
    [string] $ParentPath,
    [Parameter(Mandatory=$true)]
    [Hashtable] $Schedules,
    [Parameter(Mandatory=$true,ValueFromPipeline)]
    [PSCustomObject] $FolderInfo,
    [string] $Filter
)
    process
    {
        if($FolderInfo.name)
        {
			$newFolderPath = Join-Path -Path $ParentPath -ChildPath $FolderInfo.name

			if(-not (Test-Path -Path $newFolderPath -PathType Container))
			{
				Write-Host "  + Folder '$($FolderInfo.name)' at '$ParentPath'."

				New-Item -Path $ParentPath -Name $FolderInfo.name -ItemType Directory -WhatIf:$WhatIf
			}
			else
			{
				Write-Host "  = Folder '$($FolderInfo.name)' already exists at '$ParentPath'."			
			}
            
			if($FolderInfo.collections)
			{
                $FolderInfo.collections | `
                    New-CMFTWDeviceCollection -CollectionPrefix $FolderInfo.prefix `
                        -Schedules $Schedules -ParentPath $newFolderPath -Filter $Filter
			}

			if($FolderInfo.devicecollectionfolders)
			{
                $FolderInfo.devicecollectionfolders | `
                    New-CMFTWDeviceCollectionFolder -ParentPath $newFolderPath -Schedules $Schedules -Filter $Filter
			}
        }
    }
}

function New-CMFTWDeviceCollection
{
param
(
    [Parameter(Mandatory=$false)]
    [string] $ParentPath,
    [Parameter(Mandatory=$false)]
    [string] $CollectionPrefix,
    [Parameter(Mandatory=$true)]
    [Hashtable] $Schedules,
    [Parameter(Mandatory=$true,ValueFromPipeline)]
    [PSCustomObject] $CollectionInfo,
    [string] $Filter
)
    process
    {
        if($CollectionInfo.name)
        {
            $collection = $null
            $collectionName = $ExecutionContext.InvokeCommand.ExpandString("$CollectionPrefix$($CollectionInfo.name)")
            $collectionComment = $ExecutionContext.InvokeCommand.ExpandString("$($CollectionInfo.comment)")

            $collection = (Get-CMDeviceCollection -Name $collectionName)

            $collectionAlreadyExists = ($null -ne $collection)

            if(-not($collection))
			{
                Add-TemplateParameters -Item $CollectionInfo -Templates $_build.templates.collection -Filter $TypeFilter

				$limitingCollection = $ExecutionContext.InvokeCommand.ExpandString($CollectionInfo.limitingCollection)
				$limitingCollectionID = (Get-CMDeviceCollection -Name $limitingCollection).CollectionID

				if(-not ($limitingCollectionID))
				{
					$limitingCollectionID = (Get-CMDeviceCollection -Name "$CollectionPrefix$limitingCollection").CollectionID

					if(-not ($limitingCollectionID))
					{
						$limitingCollectionID = 'SMS00001'
					}
				}
			
                Write-Output "    + Collection '$collectionName' (limited to '$limitingCollection')."

                if($Schedules.ContainsKey($CollectionInfo.schedule))
				{
					if($CollectionInfo.incremental -eq 'yes')
					{
						$refreshType = 'Both'   
					}
					else 
					{
						$refreshType = 'Periodic'    
					}

					$collection = New-CMDeviceCollection -Name $collectionName `
                        -LimitingCollectionId $limitingCollectionID `
                        -RefreshType $refreshType `
                        -RefreshSchedule $Schedules.Get_Item($CollectionInfo.schedule) `
                        -Comment $collectionComment `
                        -WhatIf:$WhatIf
				}
				else
				{
					if($CollectionInfo.incremental -eq 'yes')
					{
						$refreshType = 'Continuous'    
					}
					else
					{
						$refreshType = 'None'
					}

					$collection = New-CMDeviceCollection -Name $collectionName `
                        -LimitingCollectionId $limitingCollectionID `
                        -RefreshType $refreshType `
                        -WhatIf:$WhatIf
				}

				if((Test-Path -Path $ParentPath -PathType Container))
				{
					Write-Output "      > Moving to $ParentPath."
					Move-CMObject -FolderPath $ParentPath -InputObject $collection -WhatIf:$WhatIf
				}

            }
            else
            {
        	    Write-Output "    = Collection '$collectionName' already exists."
            }

			if($UpdateMembership -or -not($collectionAlreadyExists))
			{
                $CollectionInfo.includeRules | `
                    Add-CMFTWDeviceCollectionReferenceRule -Collection $collection -RuleType Include

                $CollectionInfo.excludeRules | `
                    Add-CMFTWDeviceCollectionReferenceRule -Collection $collection -RuleType Exclude

                $CollectionInfo.directRules | `
                    Add-CMFTWDeviceCollectionDirectRule -Collection $collection

                $CollectionInfo.queryRules | `
                    Add-CMFTWDeviceCollectionQueryRule -Collection $collection
			}
        }
    }
}

function Add-CMFTWDeviceCollectionQueryRule
{
    param
    (
        [Parameter(Mandatory=$true)]
        [Object] $Collection,
        [Parameter(ValueFromPipeline)]
        [PSCustomObject] $RuleInfo  
    )

    process
    {
        if(-not($RuleInfo))
        {
            return
        } 

        $ruleName = $ExecutionContext.InvokeCommand.ExpandString($RuleInfo.name)
        if($ruleName -eq '.')
        {
            $ruleName = $Collection.Name
        }
        $ruleQuery = $ExecutionContext.InvokeCommand.ExpandString($RuleInfo.query)

        if(Get-CMDeviceCollectionQueryMembershipRule -Collection $Collection -RuleName $ruleName)
        {
            Write-Output "      = Query rule for $ruleName already exists."
        }
        else
        {
            Write-Output "      + Query rule: '$ruleName'"
            Add-CMDeviceCollectionQueryMembershipRule -Collection $Collection -RuleName $ruleName -QueryExpression $ruleQuery -WhatIf:$WhatIf
        }
    }
}

function Add-CMFTWDeviceCollectionReferenceRule
{
    param
    (
        [Parameter(Mandatory=$true)]
        [Object] $Collection,
        [Parameter(ValueFromPipeline)]
        [string] $RuleCollectionName,
        [ValidateSet('Include','Exclude')]
        [string] $RuleType
    )

    process
    {
        if(-not($RuleCollectionName))
        {
            return
        }
        
        $referenceCollectionName = $ExecutionContext.InvokeCommand.ExpandString($RuleCollectionName)

        if(-not(Get-CMDeviceCollection -Name $referenceCollectionName))
        {
            Write-Output "      ! The collecton '$referenceCollectionName' does not exist and cannot be referenced."
        }
        else
        {
            Write-Output "      + Creating new $RuleType rule for '$referenceCollectionName'"
        
            if($RuleType -eq 'Include')
            {
                Add-CMDeviceCollectionIncludeMembershipRule -Collection $collection -IncludeCollectionName $referenceCollectionName -WhatIf:$WhatIf
            }
            elseif($RuleType -eq 'Exclude')
            {
                Add-CMDeviceCollectionExcludeMembershipRule -Collection $collection -ExcludeCollectionName $referenceCollectionName -WhatIf:$WhatIf
            }
        }
    }
}

function Add-CMFTWDeviceCollectionDirectRule
{
    param
    (
        [Parameter(Mandatory=$true)]
        [Object] $Collection,
        [Parameter(ValueFromPipeline)]
        [string] $RuleResourceName
    )

    process
    {
        if(-not($RuleResourceName))
        {
            return
        }
        
        $referenceResourceName = $ExecutionContext.InvokeCommand.ExpandString($RuleResourceName)

        if(-not(Get-CMDevice -Name $referenceResourceName))
        {
            Write-Output "      ! The resource '$referenceResourceName' does not exist and cannot be referenced."
        }
        else
        {
            Write-Output "      + Creating new Direct rule for '$referenceResourceName'"
        
            Add-CMDeviceCollectionDirectMembershipRule -Collection $collection -Resource $referenceResourceName -WhatIf:$WhatIf
        }
    }
}

function New-CMFTWUpdatePackage
{
	param
	(
		[Parameter(Mandatory=$true,ValueFromPipeline)]
		[PSCustomObject] $PackageInfo
	)
	
	begin
	{
        Write-Host "Creating update packages ..."
	}
	
	process
	{
        if(Get-CMSoftwareUpdateDeploymentPackage -Name $PackageInfo.Name)
        {
            Write-Output "  = Update package '$($PackageInfo.Name)' already exists"
        }
        else
        {
            if($PackageInfo.Path.LastIndexOf('\') -eq ($PackageInfo.Path.length - 1))
            {
                $folderName = $PackageInfo.Name -replace '\s', ''
                $fullSourcePath = Join-Path -Path $PackageInfo.Path -ChildPath $folderName
            }
            else
            {
                $folderName = [System.IO.Path]::GetFileName($PackageInfo.Path)
                $fullSourcePath = $PackageInfo.Path
            }

            if(-not (Test-Path -Path "FileSystem::$fullSourcePath" -PathType Container))
            {
                Write-Output "  + Source folder '$fullSourcePath'"
                New-Item -Path "FileSystem::$fullSourcePath" -Name $folderName -ItemType Directory -WhatIf:$WhatIf | Out-Null
            }
            else
            {
                Write-Output "  = Source folder '$fullSourcePath' already exists"
            }

            Write-Output "  + Update package: '$($PackageInfo.Name)'"
            New-CMSoftwareUpdateDeploymentPackage -Name $PackageInfo.Name -Path $fullSourcePath -WhatIf:$WhatIf | Out-Null
        }
	}
	
	end
	{
        Write-Host ""
	}
}

function New-CMFTWADRDeployment
{
	param
	(
		[Parameter(Mandatory=$true,ValueFromPipeline)]
		[PSCustomObject] $ADRDeploymentInfo,
		[Parameter(Mandatory=$true)]
		[string] $ADRName,
		[string] $Filter
	)
	
	begin
	{
	}
	
	process
	{
		$adrDeploymentCollection = $ExecutionContext.InvokeCommand.ExpandString($ADRDeploymentInfo.CollectionName)

        Add-TemplateParameters -Item $ADRDeploymentInfo -Templates $_build.templates.updatedeployments -Filter $TypeFilter

        $parameters = Invoke-ProcessParameters -Item $ADRDeploymentInfo -Filter $Filter

        if($parameters.Keys.Count -gt 0)
        {
            if (Get-CMDeviceCollection -Name $adrDeploymentCollection)
            {
                Write-Host "    + Automatic Deployment Rule Deployment for '$adrDeploymentCollection'"
                New-CMAutoDeploymentRuleDeployment @parameters | out-null
            }
            else
            {
                Write-Warning "Could not find target collection '$($parameters.CollectionName)' trying to create Automatic Deployment Rule Deployment."
            }
		}
	}
	
	end
	{
	}
}

function New-CMFTWADR
{ 
	param
	(
		[Parameter(Mandatory=$true)]
		[Hashtable] $Schedules,
		[Parameter(Mandatory=$true,ValueFromPipeline)]
		[PSCustomObject] $ADRInfo,
		[string] $Filter
	)
	
	begin
	{
        Write-Host "Creating automatic deployment rules ..."
	}
	
	process
	{
		$adrName = $ExecutionContext.InvokeCommand.ExpandString($ADRInfo.Name)
        $currentADR = Get-CMSoftwareUpdateAutoDeploymentRule -Name $adrName -Fast
        $updatePackageName = $ExecutionContext.InvokeCommand.ExpandString($ADRInfo.DeploymentPackageName)
        $updatePackage = Get-CMSoftwareUpdateDeploymentPackage -Name $updatePackageName

        if($currentADR -and -not($Recreate))
        {
            Write-Output "  = Skipping Existing Automatic Deployment Rule: $adrName"
            return
        }
        elseif($currentADR -and $Recreate -and $updatePackage)
        {
            Write-Output "  - Removing Existing Automatic Deployment Rule: $adrName"
            Remove-CMSoftwareUpdateAutoDeploymentRule -Name $adrName -Force
        }
        
        if(-not($updatePackage))
        {
			Write-Output "  - Skipping '$adrName' because the '$updatePackageName' update package does not exist."	
        }
        else
        {				
            Add-TemplateParameters -Item $ADRInfo -Templates $_build.templates.automaticdeploymentrules -Filter $TypeFilter
            Add-TemplateParameters -Item $ADRInfo -TemplateProperty '_deploymentTemplate' -Templates $_build.templates.updatedeployments -Filter $TypeFilter

            $parameters = Invoke-ProcessParameters -Item $ADRInfo -Filter $Filter

            if($parameters.Keys.Count -gt 0)
            {
                if(Get-CMDeviceCollection -Name $parameters.CollectionName)
                {
                    Write-Host "  + Automatic Deployment Rule: $adrName"
                    New-CMSoftwareUpdateAutoDeploymentRule @parameters | out-null
                }
                else
                {
                    Write-Warning "Could not find target collection '$($parameters.CollectionName)' trying to create Automatic Deployment Rule: $adrName."
                }
            }
        }
			
        if(Get-CMSoftwareUpdateAutoDeploymentRule -Name $adrName -Fast)
        {
            if($ADRInfo._additionaldeployments)
            {
                $additionalDeployments = $ADRInfo._additionaldeployments | `
                    Where-Object { $null -eq $_._type -or (($_._type -split ',') -contains $Filter) }
                $additionaldeployments | New-CMFTWADRDeployment -ADRName $adrName -Filter $TypeFilter
            }
        }
	}
	
	end
	{
        Write-Host ""
	}
}

function New-CMFTWMaintenanceWindow
{
	param
	(
		[Parameter(Mandatory=$true)]
		[Hashtable] $Schedules,
		[Parameter(Mandatory=$true,ValueFromPipeline)]
		[PSCustomObject] $MWInfo
	)
	
	begin
	{
        Write-Host "Creating maintenance windows ..."
    }

    process
    {
        $mwDeploymentCollection = $ExecutionContext.InvokeCommand.ExpandString($MWInfo.CollectionName)

        Add-TemplateParameters -Item $MWInfo -Templates $_build.templates.maintenancewindows -Filter $TypeFilter

        $parameters = Invoke-ProcessParameters -Item $MWInfo -Filter $Filter

        if ($parameters.Keys.Count -gt 0)
        {
            if (Get-CMDeviceCollection -Name $mwDeploymentCollection)
            {
                Write-Host "    + Maintenance Window for '$mwDeploymentCollection'"
                New-CMMaintenanceWindow @parameters | out-null
            }
            else
            {
                Write-Warning "Could not find target collection '$($parameters.CollectionName)' trying to create Maintenance Window."
            }
        }

    }
    
    end
	{
        Write-Host ""
	}   
}

#Load Configuration Manager PowerShell Module
if(-not(Import-Module -Name ($Env:SMS_ADMIN_UI_PATH.Substring(0, $Env:SMS_ADMIN_UI_PATH.Length - 5) + '\ConfigurationManager.psd1') -PassThru))
{
    Write-Error "Could not load Configuration Manager module."
    exit 1
}

Write-Host "Loaded Configuration Manager Module"
Write-Host ""

#Get SiteCode
$siteCode = Get-PSDrive -PSProvider CMSITE

# Load configuration file and convert into PowerShell objects from json
$_build = (((Get-Content -Path $ConfigFile) -Join "`n") | ConvertFrom-Json).configuration

Push-Location "${siteCode}:"

# Create variables for each variable defined in the configuraiton file
$_build.variables | `
    Get-Member -MemberType NoteProperty | `
    ForEach-Object { Set-Variable -Name $_.Name -Value $_build.variables.($_.Name) }

if($_build.schedules)
{
	$schedules = $_build.schedules | New-CMFTWSchedule -Filter $TypeFilter
}

if($Collections -eq $true -and $_build.devicecollectionfolders)
{
    Write-Host "Creating device collection folders and collections ..."

    $_build.devicecollectionfolders | `
        New-CMFTWDeviceCollectionFolder -ParentPath "${siteCode}:\DeviceCollection" -Schedules $schedules -Filter $TypeFilter
    
    Write-Host ""

}

if($UpdatePackages -eq $true -and $_build.updatepackages)
{
	$_build.updatepackages | New-CMFTWUpdatePackage
}

if($AutomaticDeploymentRules -eq $true -and $_build.automaticdeploymentrules)
{	
    $_build.automaticdeploymentrules | `
        New-CMFTWADR -Schedules $schedules -Filter $TypeFilter
}

if($MaintenanceWindows -eq $true -and $_build.maintenancewindows)
{
    $_build.maintenancewindows | `
        New-CMFTWMaintenanceWindow -Schedules $schedules    
}

Pop-Location
<#
	.SYNOPSIS
		Creates standard objects and configurations in ConfigMgr from a supplied json configuration file.
	
	.DESCRIPTION
		Creates Collection Folders, Collections, Client Settings, Update Packages, and Automatic Deployment rules in ConfigMgr using
a json file where these are all defined.

Version 1.5
26 September 2018
	
	.PARAMETER ConfigFile
		The input json configuration file to use. If not specified, .\Build-CMDefaultConfig.json is used.
	.PARAMETER Collections
        Process and create or update collections defined in the configuration file.
	.PARAMETER ClientSettings
        Create Client Settings Packages defined in the configuration file.
	.PARAMETER ADRs
        Create Automatic Deployment Rules defined in the configuration file.	
	.PARAMETER UpdateMembership
        Update the membership of existing collections specified in the configuration file. If this is not specified and the collection already exists,
the membership of collection will not be updated.

    .EXAMPLE
        .\Build-CMDefaultConfig.ps1 -ConfigFile .\Build-CMDefaultConfig.json -Collections
        Creates collections defined in the Build-CMDefaultConfig.json configuration file.
    .EXAMPLE
        .\Build-CMDefaultConfig.ps1 -ConfigFile .\Build-CMDefaultConfig.json -Collections -UpdateMembership
        Creates collections defined in the Build-CMDefaultConfig.json configuration file if they don't already exist and
updates the membership of collections defined in the Build-CMDefaultConfig.json configuration file.
        
#>

[CmdletBinding()]
param
(
	[Parameter(HelpMessage = 'The configuration file to use.')]
	[ValidateScript({ Test-Path -PathType Leaf -Path $_ })]
	[Alias('config')]
	[string]$ConfigFile = '.\Build-CMDefaultConfig.json',
	[switch]$Collections,
	[switch]$ClientSettings,
	[switch]$ADRs,
	[switch]$UpdateMembership,
    [switch]$WhatIf

)

function Process-Parameters
{

	param
	(
		[Parameter(Mandatory=$true)]
        [PSCustomObject] $Object,
		[switch] $ExcludeNameParam,
		[array] $ExcludeParams
	)

	$command = ""
	
	$Object | Get-Member -Type NoteProperty | % {
				
		$paramName = $_.Name				
		$paramValue = $ExecutionContext.InvokeCommand.ExpandString($Object.$($_.Name))
		
		$paramValue = $paramValue.Trim()
		
		if($paramValue -and $paramValue -ne "" -and -not ($ExcludeParams -contains $paramName))
		{
			if ($paramValue -ieq 'true' -or $paramValue -ieq 'false')
			{
				$command += " -$($paramName) `$$paramValue"
			}
			
			elseif ($paramName -ieq 'schedule')
			{
				$command += " -$($paramName) `$Schedules.Get_Item(`"$paramValue`")"
			}
			
			elseif ($paramName -ieq 'type')
			{
				$command += " -$($paramValue)"
			}
			
			elseif ($paramName -ieq 'start')
			{
				$command += " -$($paramName) '$(Get-Date -Date((Get-Date).ToString('yyyy-MM-dd' + 'T' + $paramValue)))'"
			}

			elseif ($paramName -ine 'name' -or -not $ExcludeNameParam)
			{
				if($paramValue -match '\s' -and $paramValue[0] -ne "'" -and $paramValue[0] -ne '"')
				{
					$command += " -$($paramName) '$paramValue'"
				}
				else
				{
					$command += " -$($paramName) $paramValue"
				}
			}
		}
	}
	
	$command
}

function Create-DeviceCollectionFolder
{
param
(
    [Parameter(Mandatory=$true)]
    [string] $Path,
    [Parameter(Mandatory=$true)]
    [Hashtable] $Schedules,
    [Parameter(Mandatory=$true,ValueFromPipeline)]
    [PSCustomObject] $FolderInfo,
    [int] $TotalFolderCount 
)

    begin
    {
        $fldrCount = 1
    }

    process
    {
        if($TotalFolderCount)
        {
            Write-Progress -Activity "Creating Device Collection Folders" -Status "$fldrCount of $TotalFolderCount" -CurrentOperation $FolderInfo.name `
                -PercentComplete ($fldrCount++ / $TotalFolderCount * 100) -Id 1
        }
    
        if($FolderInfo.name)
        {

			if(-not (Test-Path -Path "$($Path)\$($FolderInfo.name)" -PathType Container))
			{
				Write-Verbose "+ Creating folder named $($_.name) at $Path."

				if($WhatIf -eq $false)
				{
					New-Item -Path $Path -Name $_.name -ItemType Directory
				}
			}
			else
			{
				Write-Verbose "= Folder named $($_.name) at $Path already exists."			
			}
            
			if($FolderInfo.collections)
			{
				$FolderInfo.collections | Create-DeviceCollection -Prefix $FolderInfo.prefix -Schedules $Schedules -Path "$($Path)\$($FolderInfo.name)" -TotalCollectionCount ($FolderInfo.collections | Measure-Object).Count
			}

			#if($FolderInfo.devicecollectionfolders)
			#{
			#    $FolderInfo.devicecollectionfolders | Create-DeviceCollectionFolder -Path "$($Path)\$($FolderInfo.name)" -Schedules $Schedules
			#}

        }
    }

    end
    {
        if($TotalFolderCount)
        {
            Write-Progress -Activity "Creating Device Collection Folders" -Completed -Id 1
        }
    }
}

function Create-DeviceCollection
{
param
(
    [Parameter(Mandatory=$false)]
    [string] $Path,
    [Parameter(Mandatory=$false)]
    [string] $Prefix,
    [Parameter(Mandatory=$true)]
    [Hashtable] $Schedules,
    [Parameter(Mandatory=$true,ValueFromPipeline)]
    [PSCustomObject] $CollectionInfo,
    [int] $TotalCollectionCount     
)
    begin
    {
        $collCount = 1
    }

    process
    {
        if($TotalCollectionCount)
        {
            Write-Progress -Activity "Creating or Updating Collections" -Status "$collCount of $TotalCollectionCount" -CurrentOperation $CollectionInfo.name `
                -PercentComplete ($collCount++ / $TotalCollectionCount * 100) -Id 2 -ParentId 1
        }
        
        if($CollectionInfo.name)
        {
            $refreshType = 'None'
            $collection = $null
            $theCollectionName = $ExecutionContext.InvokeCommand.ExpandString("$Prefix$($CollectionInfo.name)")

            $collection = (Get-CMDeviceCollection -Name $theCollectionName)


            $collectionAlreadyExists = ($collection -ne $null)

            if(-not($collection))
			{
			
				$limitingCollectionID = (Get-CMDeviceCollection -Name $CollectionInfo.limitingCollection).CollectionID

				if(-not ($limitingCollectionID))
				{
					$limitingCollectionID = (Get-CMDeviceCollection -Name "$Prefix$($CollectionInfo.limitingCollection)").CollectionID

					if(-not ($limitingCollectionID))
					{
						$limitingCollectionID = 'SMS00001'
					}
				}
			
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

					Write-Output " + Creating collection named '$theCollectionName' limited to $limitingCollectionID."

					if($WhatIf -eq $false)
					{
						$collection = New-CMDeviceCollection -Name $theCollectionName `
							-LimitingCollectionId $limitingCollectionID `
							-RefreshType $refreshType `
							-RefreshSchedule $Schedules.Get_Item($CollectionInfo.schedule)
					}

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

					Write-Output " + Creating collection named '$theCollectionName' limited to $limitingCollectionID."

					if($WhatIf -eq $false)
					{
						$collection = New-CMDeviceCollection -Name $theCollectionName `
							-LimitingCollectionId $limitingCollectionID `
							-RefreshType $refreshType
					}

				}

            }
            else
            {
        	    Write-Output " = Collection named '$theCollectionName' already exists."
            }

			if($UpdateMembership -or -not($collectionAlreadyExists))
			{
				if((Test-Path -Path $Path -PathType Container) -and ($WhatIf -eq $false))
				{
					Move-CMObject -FolderPath $Path -InputObject $collection
				}
				else
				{
					Write-Output "  -> Moving '$theCollectionName' to $Path."
				}

				
				if($CollectionInfo.queryRules)
				{
					$CollectionInfo.queryRules | Get-Member -Type NoteProperty | % {
						
						$rule = $ExecutionContext.InvokeCommand.ExpandString($CollectionInfo.queryRules.$($_.Name))
						
						if(Get-CMDeviceCollectionQueryMembershipRule -Collection $collection -RuleName $_.Name)
                        {
                            Write-Output " = Query rule for '$theCollectionName': $($_.Name) already exists."
                        }
                        else
                        {
                            Write-Output "  + Creating new query rule for '$theCollectionName': $rule"

						    if($WhatIf -eq $false)
						    {
							    Add-CMDeviceCollectionQueryMembershipRule -Collection $collection -RuleName $_.Name -QueryExpression $rule
						    }
                        }
					}
				}
					
				if($CollectionInfo.includeRules)
				{
					$CollectionInfo.includeRules -split "," | % {
						
						#$includeCollectionName = ($CollectionInfo.includeRules.$($_.Name))

                        if(-not(Get-CMDeviceCollection -Name $_))
                        {
                            Write-Output " x The collecton $_ does not exist  and cannot be included in $theCollectionName."
                        }
                        elseif(Get-CMDeviceCollectionIncludeMembershipRule -Collection $collection -IncludeCollectionName $_)
                        {
                            Write-Output " = Include rule for '$theCollectionName': $_ already exists."
                        }
                        else
                        {
						    Write-Output " + Creating new include rule for '$theCollectionName': $_"
						
						    if($WhatIf -eq $false)
						    {
							    Add-CMDeviceCollectionIncludeMembershipRule -Collection $collection -IncludeCollectionName $_
						    }
                        }
					}
				}
					
				if($CollectionInfo.excludeRules)
				{
					$CollectionInfo.excludeRules -split "," | % {
						
						#$excludeCollectionName = ($CollectionInfo.excludeRules.$($_.Name))
							
						if(-not(Get-CMDeviceCollection -Name $_))
                        {
                            Write-Output " x The collecton $_ does not exist and cannot be excluded from $theCollectionName."
                        }
                        elseif(Get-CMDeviceCollectionExcludeMembershipRule -Collection $collection -ExcludeCollectionName $_)
                        {
                            Write-Output " = Exclude rule for '$theCollectionName': $_ already exists."
                        }
                        else
                        {
                            Write-Output " + Creating new exclude rule for '$theCollectionName': $_"
						
						    if($WhatIf -eq $false)
						    {
							    Add-CMDeviceCollectionExcludeMembershipRule -Collection $collection -ExcludeCollectionName $_
						    }
                        }
					}
				}

				if($CollectionInfo.directRules)
				{
					$CollectionInfo.directRules -split "," | % {
						
						if(Get-CMDeviceCollectionDirectMembershipRule -Collection $collection -ResourceName $_)
                        {
                            Write-Output " = Direct rule for '$theCollectionName': $_ already exists."
                        }
                        else
                        {
                            $res = (Get-CMDevice -Name $_)
                            
                            if(-not($res))
                            {
                                Write-Output " x The resource '$_' does not exist"
                            }
                            else
                            {
                                Write-Output " + Creating new direct rule for '$theCollectionName': $_"
						
						        if($WhatIf -eq $false)
						        {
							        Add-CMDeviceCollectionDirectMembershipRule -Collection $collection -Resource $res
						        }
                            }
                        }
					}
				}
			}
        }

    }
    end
    {
        if($TotalCollectionCount)
        {
            Write-Progress -Activity "Creating or Updating Collections" -Completed -Id 2 -ParentId 1
        }
    }
}

function Set-ClientSettings
{
	param
	(
		[Parameter(Mandatory=$true)]
        [string] $Name,
		[Parameter(Mandatory=$true)]
		[Hashtable] $Schedules,
		[Parameter(Mandatory=$true,ValueFromPipeline)]
		[PSCustomObject] $ClientSettingsItemInfo,
		[int] $TotalClientSettingsItemCount 
	)

	begin
	{
		$csiCount = 0
	}
	
	process
	{
		if($TotalClientSettingsItemCount)
        {
            Write-Progress -Activity "Setting Client Settings" -Status "$csCount of $TotalClientSettingsItemCount" -CurrentOperation $ClientSettingsItemInfo.type `
                -PercentComplete ($csiCount++ / $TotalClientSettingsItemCount * 100) -Id 2 -ParentId 1
        }
		
		if($ClientSettingsItemInfo.type)
		{
		
			$commandline = Process-Parameters -Object $ClientSettingsItemInfo
						
			Write-Output "  + Updating client settings '$($ClientSettingsItemInfo.type)': $commandline"

			if($WhatIf -eq $false)
			{
                $commandline = "Set-CMClientSetting -Name `"$Name`" $commandline"
				Invoke-Expression -Command $commandline
			}
		}
	}

	end
	{
	    if($TotalClientSettingsCount)
        {
            Write-Progress -Activity "Setting Client Settings" -Completed -Id 2 -ParentId 1
        }
	}
}

function Create-ClientSettings
{
	param
	(
		[Parameter(Mandatory=$true)]
		[Hashtable] $Schedules,
		[Parameter(Mandatory=$true,ValueFromPipeline)]
		[PSCustomObject] $ClientSettingsInfo,
		[int] $TotalClientSettingsCount 
	)

	begin
	{
		$csCount = 0
	}
	
	process
	{
		if($TotalClientSettingsCount)
        {
            Write-Progress -Activity "Creating, Configuring, and Deploying Client Settings" -Status "$csCount of $TotalClientSettingsCount" -CurrentOperation $ClientSettingsInfo.name `
                -PercentComplete ($csCount++ / $TotalClientSettingsCount * 100) -Id 1
        }
		
		if($ClientSettingsInfo.name)
		{
	
			if(-not (Get-CMClientSetting -name $ClientSettingsInfo.name))
			{	
				Write-Output "+ Creating client settings named '$($ClientSettingsInfo.name)'."

				if($WhatIf -eq $false)
				{
					New-CMClientSetting -Name $ClientSettingsInfo.name -Type $ClientSettingsInfo.type
				}
			}
			else
			{
				Write-Output "= Client settings named '$($ClientSettingsInfo.name)' already exists."			
			}
			
			$_.settings | Set-ClientSettings -Name $ClientSettingsInfo.name -Schedules $Schedules -TotalClientSettingsItemCount ($ClientSettingsInfo.settings | Measure-Object).count
			
			if($ClientSettingsInfo.name -ne 'Default Client Settings' -and $ClientSettingsInfo.collection)
			{
				Write-Output ">> Deploying client settings named '$($ClientSettingsInfo.name)' to the collection named '$($ClientSettingsInfo.collection)."

				if($WhatIf -eq $false)
				{
					Start-CMClientSettingDeployment -ClientSettingName $ClientSettingsInfo.name -CollectionName $ClientSettingsInfo.collection
				}
			}
		}
	}
	
	end
	{
	    if($TotalClientSettingsCount)
        {
            Write-Progress -Activity "Creating, Configuring, and Deploying Client Settings" -Completed -Id 1
        }
	}
}

function Create-Schedule
{
	param
	(
		[Parameter(Mandatory=$true,ValueFromPipeline)]
		[PSCustomObject] $ScheduleInfo,
		[int] $TotalScheduleCount
	)

    begin
    {
        $schedule = $null
        $schedulehash = @{}
		$scheduleCount = 0
    }

    process
    {
		if($TotalScheduleCount)
        {
            Write-Progress -Activity "Creating Schedules" -Status "$scheduleCount of $TotalScheduleCount" -CurrentOperation $ScheduleInfo.name `
                -PercentComplete ($scheduleCount++ / $TotalScheduleCount * 100) -Id 1
        }
		
		$commandline = Process-Parameters -ExcludeNameParam -Object $ScheduleInfo 
		
		$commandline = "New-CMSchedule $commandline"
		$schedule = Invoke-Expression -Command $commandline

        $schedulehash.Add($ScheduleInfo.name, $schedule)
    }
    end
    {
		if($TotalScheduleCount)
        {
            Write-Progress -Activity "Creating Schedules" -Completed -Id 1
        }
		
        $schedulehash
    }
}

function Create-UpdatePackage
{
	param
	(
		[Parameter(Mandatory=$true,ValueFromPipeline)]
		[PSCustomObject] $PackageInfo,
		[int] $TotalPackageCount 
	)
	
	begin
	{
		$packageCount = 0
	}
	
	process
	{
		if($TotalPackageCount)
        {
            Write-Progress -Activity "Creating Update Packages" -Status "$packageCount of $TotalPackageCount" -CurrentOperation $ADRInfo.type `
                -PercentComplete ($packageCount++ / $TotalPackageCount * 100) -Id 1
        }
		
		$folderName = $PackageInfo.Name -replace '\s', ''
		$sourceFolderPath = $PackageInfo.Path
        $fullSourcePath = $sourceFolderPath + $folderName
		
		if($PackageInfo.Path.LastIndexOf('\') -ne ($PackageInfo.Path.length - 1))
		{
			$fullSourcePath = $PackageInfo.Path
            $lastSlash = $PackageInfo.Path.LastIndexOf('\') + 1
			$folderName = $PackageInfo.Path.Substring($lastSlash)
			$sourceFolderPath = $PackageInfo.Path.Substring(0, $lastSlash)
		}       
		
		if(-not (Test-Path $fullSourcePath))
		{
			Write-Output "  + Creating source folder named $folderName at $sourceFolderPath"
			
			if($WhatIf -eq $false)
			{
				New-Item -Path $sourceFolderPath -Name $folderName -ItemType Directory
			}	
		}
		
		Write-Output "  + Creating update package: $($PackageInfo.Name)"

		if($WhatIf -eq $false)
		{
			. ".\New-CMDeploymentPackage.ps1" -SiteServer $([System.Net.Dns]::GetHostByName((hostname)).HostName) -Name $PackageInfo.Name -SourcePath $fullSourcePath
		}
	}
	
	end
	{
		if($TotalPackageCount)
        {
            Write-Progress -Activity "Creating Update Packages" -Completed -Id 1
        }
	}
}

function Create-ADRDeployment
{
	param
	(
		[Parameter(Mandatory=$true,ValueFromPipeline)]
		[PSCustomObject] $ADRDeploymentInfo,
		[int] $TotalADRDeploymentCount,
		[Parameter(Mandatory=$true)]
		[string] $ADRName
	)
	
	begin
	{
		$adrDeploymentCount = 0
	}
	
	process
	{
		if($TotalADRDeploymentCount)
        {
            Write-Progress -Activity "Creating Automatic Deployment Rule Deployments" -Status "$adrDeploymentCount of $TotalADRDeploymentCount" -CurrentOperation $ADRDeploymentInfo.CollectionName `
                -PercentComplete ($adrDeploymentCount++ / $TotalADRDeploymentCount * 100) -Id 2 -ParentId 1
        }
		
		$commandline = Process-Parameters -Object $ADRDeploymentInfo               
		Write-Output "  + Creating Automatic Deployment Rule Deployment: $commandline"

		if($WhatIf -eq $false)
		{
			$commandline = "New-CMAutoDeploymentRuleDeployment -name '$ADRName' $commandline"
			Invoke-Expression -Command $commandline
		}
	}
	
	end
	{
		if($TotalADRDeploymentCount)
        {
            Write-Progress -Activity "Creating Automatic Deployment Rule Deployments" -Completed -Id 2
        }
	}
}

function Create-ADR
{
	param
	(
		[Parameter(Mandatory=$true)]
		[Hashtable] $Schedules,
		[Parameter(Mandatory=$true,ValueFromPipeline)]
		[PSCustomObject] $ADRInfo,
		[int] $TotalADRCount 
	)
	
	begin
	{
		$adrCount = 0
	}
	
	process
	{
		if($TotalADRCount)
        {
            Write-Progress -Activity "Creating Automatic Deployment Rules" -Status "$adrCount of $TotalADRCount" -CurrentOperation $ADRInfo.type `
                -PercentComplete ($adrCount++ / $TotalADRCount * 100) -Id 1
        }
		
		$commandline = Process-Parameters -Object $ADRInfo -ExcludeParams "additionaldeployments"              
		Write-Output "  + Creating Automatic Deployment Rule: $commandline"

		if($WhatIf -eq $false)
		{
			$commandline = "New-CMSoftwareUpdateAutoDeploymentRule $commandline"
			Invoke-Expression -Command $commandline
		}
		
		if($ADRInfo.additionaldeployments)
		{
			$ADRInfo.additionaldeployments | Create-ADRDeployment -ADRName $ADRInfo.Name -TotalADRDeploymentCount ($ADRInfo.additionaldeployments | Measure-Object).Count
		}
	}
	
	end
	{
		if($TotalADRCount)
        {
            Write-Progress -Activity "Creating Automatic Deployment Rules" -Completed -Id 1
        }
	}
}

if($WhatIf -eq $true)
{
	$VerbosePreference = "Continue"
}

#Load Configuration Manager PowerShell Module
Import-module ($Env:SMS_ADMIN_UI_PATH.Substring(0, $Env:SMS_ADMIN_UI_PATH.Length - 5) + '\ConfigurationManager.psd1')

#Get SiteCode
$siteCode = Get-PSDrive -PSProvider CMSITE

$buildConfig = ((Get-Content -Path $ConfigFile) -Join "`n")
	
$buildObjects = ($buildConfig | ConvertFrom-Json)

Push-Location $siteCode":"

$buildObjects.defaultitems.variables | Get-Member -MemberType NoteProperty | `
    % { Set-Variable -Name $_.Name -Value $buildObjects.defaultitems.variables.($_.Name) }

if($buildObjects.defaultitems.schedules)
{
	$schedules = ($buildObjects.defaultitems.schedules | Create-Schedule)
}

if($Collections -eq $true -and $buildObjects.defaultitems.devicecollectionfolders)
{
	$buildObjects.defaultitems.devicecollectionfolders | Create-DeviceCollectionFolder -Path "$($siteCode):\DeviceCollection" -Schedules $schedules -TotalFolderCount ($buildObjects.defaultitems.devicecollectionfolders | Measure-Object).Count
}

if($ClientSettings -eq $true -and $buildObjects.defaultitems.clientsettings)
{
	$buildObjects.defaultitems.clientsettings | Create-ClientSettings -Schedules $schedules -TotalClientSettingsCount ($buildObjects.defaultitems.clientsettings | Measure-Object).Count
}

if($ADRs -eq $true)
{	
    Pop-Location

	if($buildObjects.defaultitems.updatepackages)
	{
		$buildObjects.defaultitems.updatepackages | Create-UpdatePackage -TotalPackageCount ($buildObjects.defaultitems.updatepackages | Measure-Object).Count
	}

	Push-Location $siteCode":"

    if($buildObjects.defaultitems.automaticdeploymentrules)
	{
		$buildObjects.defaultitems.automaticdeploymentrules | Create-ADR -Schedules $schedules -TotalADRCount ($buildObjects.defaultitems.automaticdeploymentrules | Measure-Object).Count
	}
}

Pop-Location
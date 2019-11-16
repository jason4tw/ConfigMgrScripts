<#
	.SYNOPSIS
		Creates standard objects and configurations in ConfigMgr from a supplied json configuration file.

	.DESCRIPTION
		Creates Collection Folders, Collections, Client Settings, Update Packages, and Automatic Deployment rules in ConfigMgr using
a json file where these are all defined.

	.PARAMETER ConfigFile
		The input json configuration file to use. If not specified, .\CMDefaultConfig.json is used.
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

	.NOTES
		Version 1.5
        Jason Sandys

		Version History
		- 1.51 (27 September 2018): Added subfolder creation capabilities
        - 1.5 (26 September 2018): Initial Version

        Limitations and Issues
		- Does not create json hierarchy for folders.
#>

[CmdletBinding()]
param
(
	[Parameter(HelpMessage = 'The configuration file to use.')]
	[ValidateScript({ Test-Path -PathType Leaf -Path $_ })]
	[Alias('config')]
	[string]$ConfigFile = '.\CMDefaultConfig.json',
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

	$Object | Get-Member -Type NoteProperty | ForEach-Object {

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
                if($paramValue -match "^\d\d:\d\d$")
                {
                    #Only specified hour and minutes for schedule, using today's date.
                    [string]$startTime = (Get-Date -Date((Get-Date).ToString('yyyy-MM-dd' + 'T' + $paramValue))).ToString()
                }else{
                    [string]$startTime = ([datetime]$paramValue).ToString()
                }
				$command += " -$($paramName) '$startTime'"
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
		}else
        {
            if($paramName -eq "NonRecurring"){
                $command += " -$($paramName)"
            }
        }
	}

	$command
}

function New-CMFTWDeviceCollectionFolder
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
			$fullFlderPath = "$($Path)\$($FolderInfo.name)"
			$folderPath = Split-Path -Path $fullFlderPath
			$folderName = Split-Path -Leaf -Path $fullFlderPath

			if(-not (Test-Path -Path $fullFlderPath -PathType Container))
			{
				Write-Output "+ Creating folder named $folderName at $folderPath."

				if($WhatIf -eq $false)
				{
					New-Item -Path $folderPath -Name $folderName -ItemType Directory
				}
			}
			else
			{
				Write-Output "= Folder named $folderName at $folderPath already exists."
			}

			if($FolderInfo.collections)
			{
				$FolderInfo.collections | New-CMFTWDeviceCollection -Prefix $FolderInfo.prefix -Schedules $Schedules -Path $fullFlderPath -TotalCollectionCount ($FolderInfo.collections | Measure-Object).Count
			}
			#Uncommenting the next 4 lines allowed me to nest device collection folders just fine, not sure why they were commented in the first place.
			if($FolderInfo.devicecollectionfolders)
			{
			    $FolderInfo.devicecollectionfolders | New-CMFTWDeviceCollectionFolder -Path "$($Path)\$($FolderInfo.name)" -Schedules $Schedules
			}

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

#Allows the creation of maintenance windows for specific device collections with the JSON file
function New-CMFTWMaintenanceWindow
{
param
(
    [Parameter(Mandatory=$true)]
    [Hashtable] $Schedules,
    [Parameter(Mandatory=$true,ValueFromPipeline)]
    [PSCustomObject] $MWInfo,
    [int] $TotalMWCount
)
    begin
    {
        $MWCount = 1
    }

    process
    {
        if($TotalMWCount)
        {
            Write-Progress -Activity "Creating or Updating Collections" -Status "$MWCount of $TotalMWCount" -CurrentOperation $MWInfo.name `
                -PercentComplete ($MWCount++ / $TotalMWCount * 100) -Id 2 -ParentId 1
        }

        if($MWInfo.Name)
        {
            if(($null -ne $MWInfo.schedule) -and ($Schedules.ContainsKey($MWInfo.schedule)))
            {
                $theSchedule = $Schedules.Get_Item($MWInfo.schedule)
                if(($null -ne $MWInfo.collectionname) -and ($null -ne $theSchedule))
                {
                    $currentMaintenanceWindows = Get-CMMaintenanceWindow -CollectionName $($MWInfo.CollectionName) | Select-Object -ExpandProperty Name

                    if($currentMaintenanceWindows -notcontains $($MWInfo.Name))
                    {


                        $commandline = Process-Parameters -Object $MWInfo
                        Write-Output " + Creating maintenance window named '$($MWInfo.Name)' for collection '$($MWInfo.CollectionName)'"
                        $commandline = "New-CMMaintenanceWindow $commandline"
			            Invoke-Expression -Command $commandline

                    }else
                    {
                        Write-Output " = A maintenance window named '$($MWInfo.Name)' already exists for collection '$($MWInfo.CollectionName)'."
                    }
                }
            }else
            {
                Write-Output " ! Could not find a schedule named '$($MWInfo.schedule)' in the hash table, skipping maintenance windows creation."
            }
        }
    }
}

function New-CMFTWDeviceCollection
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
            #$refreshType = 'None'
            $collection = $null
            $theCollectionName = $ExecutionContext.InvokeCommand.ExpandString("$Prefix$($CollectionInfo.name)")

            $collection = (Get-CMDeviceCollection -Name $theCollectionName)


            $collectionAlreadyExists = ($null -ne $collection)

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

				if(($null -ne $CollectionInfo.schedule) -and ($Schedules.ContainsKey($CollectionInfo.schedule)))
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

				if((Test-Path -Path $Path -PathType Container))
				{
					Write-Output "  -> Moving '$theCollectionName' to $Path."

					if ($WhatIf -eq $false)
					{
						Move-CMObject -FolderPath $Path -InputObject $collection
					}
				}

            }
            else
            {
        	    Write-Output " = Collection named '$theCollectionName' already exists."
            }

			if($UpdateMembership -or -not($collectionAlreadyExists))
			{

				if($CollectionInfo.queryRules)
				{
					$CollectionInfo.queryRules | Get-Member -Type NoteProperty | ForEach-Object {

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
                    $CollectionInfo.includeRules | Get-Member -Type NoteProperty | ForEach-Object {

						$includeCollectionName = $ExecutionContext.InvokeCommand.ExpandString($CollectionInfo.includeRules.$($_.Name))

                        if(-not(Get-CMDeviceCollection -Name $includeCollectionName))
                        {
                            Write-Output " x The collecton $includeCollectionName does not exist  and cannot be included in $theCollectionName."
                        }
                        elseif(Get-CMDeviceCollectionIncludeMembershipRule -Collection $collection -IncludeCollectionName $includeCollectionName)
                        {
                            Write-Output " = Include rule for '$theCollectionName': $includeCollectionName already exists."
                        }
                        else
                        {
						    Write-Output " + Creating new include rule for '$theCollectionName': $includeCollectionName"

						    if($WhatIf -eq $false)
						    {
							    Add-CMDeviceCollectionIncludeMembershipRule -Collection $collection -IncludeCollectionName $includeCollectionName
						    }
                        }
					}
				}

				if($CollectionInfo.excludeRules)
				{
					#$CollectionInfo.excludeRules -split "," | ForEach-Object {
					$CollectionInfo.excludeRules | Get-Member -Type NoteProperty | ForEach-Object {

						#$excludeCollectionName = ($CollectionInfo.excludeRules.$($_.Name))
                        $excludeCollectionName = $ExecutionContext.InvokeCommand.ExpandString($CollectionInfo.excludeRules.$($_.Name))

						if(-not(Get-CMDeviceCollection -Name $excludeCollectionName))
                        {
                            Write-Output " x The collecton $excludeCollectionName does not exist and cannot be excluded from $theCollectionName."
                        }
                        elseif(Get-CMDeviceCollectionExcludeMembershipRule -Collection $collection -ExcludeCollectionName $excludeCollectionName)
                        {
                            Write-Output " = Exclude rule for '$theCollectionName': $excludeCollectionName already exists."
                        }
                        else
                        {
                            Write-Output " + Creating new exclude rule for '$theCollectionName': $excludeCollectionName"

						    if($WhatIf -eq $false)
						    {
							    Add-CMDeviceCollectionExcludeMembershipRule -Collection $collection -ExcludeCollectionName $excludeCollectionName
						    }
                        }
					}
				}

				if($CollectionInfo.directRules)
				{
					$CollectionInfo.directRules -split "," | ForEach-Object {

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

                if($CollectionInfo.comment)
                {
                    Write-Output " + Setting collection comment to `"$($CollectionInfo.comment)`""
                    Set-CMDeviceCollection -InputObject $collection -Comment "$($CollectionInfo.comment)"
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

function New-CMFTWClientSettings
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

function New-CMFTWSchedule
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

function New-CMFTWUpdatePackage
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

		#If there is a variable in the name of the update package, expand it
        $PackageInfo.Name = $ExecutionContext.InvokeCommand.ExpandString($PackageInfo.Name)

		$folderName = $PackageInfo.Name -replace '\s', ''
        $split = $PackageInfo.Path -split "\\"
        $sourceFolderPath = ""
        foreach($match in $split){
            if($match[0] -eq "$"){
                #Expanding variable
                $matchValue = "$(Invoke-Expression -Command $match)"
            }else{
                $matchValue = $match
            }
            if($sourceFolderPath -eq ""){
                $sourceFolderPath = $matchValue
            }else{
                $sourceFolderPath = Join-Path $sourceFolderPath $matchValue
            }
        }
        $fullSourcePath = Join-Path $sourceFolderPath $folderName

		if(-not (Test-Path $fullSourcePath))
		{
			Write-Output "  + Creating source folder named $folderName at $sourceFolderPath"

			if($WhatIf -eq $false)
			{
				New-Item -Path $sourceFolderPath -Name $folderName -ItemType Directory
			}
		}else{
            Write-Output " = Source folder `"$fullSourcePath`" already exists."
        }

		Push-Location $siteCode":"
        if(-not (Get-CMSoftwareUpdateDeploymentPackage -Name $($PackageInfo.Name))){

            Write-Output " + Creating update package: $($PackageInfo.Name)"

		    if($WhatIf -eq $false)
		    {
                New-CMSoftwareUpdateDeploymentPackage -Name $($PackageInfo.Name) -Path $fullSourcePath | Out-Null
		    }
        }else{
            Write-Output " = Update package '$($PackageInfo.Name)' already exists."
        }
        #Enable Binary Differential Replication (BDR) on the package
        $SiteServer = Get-CMSite | Select-Object -ExpandProperty ServerName
        $updatePackage = Get-CMSoftwareUpdateDeploymentPackage -Name $($PackageInfo.Name)

        Pop-Location

        $wmiPackage = Get-WmiObject -Namespace "root\SMS\site_AA1" -ComputerName $SiteServer -Query "SELECT * FROM SMS_SoftwareUpdatesPackage WHERE PackageID = '$($updatePackage.PackageID)'"
        $wmiPackage = [wmi]$wmiPackage.__PATH
        if($wmiPackage.PkgFlags -ne 83886080){
            Write-Output " + Enabling Binary Differential Replication (BDR) for the update package."
            $wmiPackage.PkgFlags = 83886080
            $wmiPackage.Put() | Out-Null
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

function New-CMFTWADRDeployment
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

		if($WhatIf -eq $false)
		{

            $RuleID = Get-CMAutoDeploymentRule -Name $ADRName -Fast | Select-Object -ExpandProperty AutoDeploymentID
            $currentDeployments = Get-CMAutoDeploymentRuleDeployment | Where-Object {$_.RuleID -eq $RuleID}
            if($currentDeployments.Collectionname -contains $ADRDeploymentInfo.CollectionName){
                Write-Output " ! This ADR already has a deployment for collection '$($ADRDeploymentInfo.CollectionName)', skipping deployment."
            }else{
                $commandline = Process-Parameters -Object $ADRDeploymentInfo
		        Write-Output "  + Creating Automatic Deployment Rule Deployment: $commandline"

                $commandline = "New-CMAutoDeploymentRuleDeployment -name '$ADRName' $commandline"
			    $ADRDeployment = Invoke-Expression -Command $commandline

                #Source: https://smsagent.blog/2019/01/31/creating-adr-deployments-in-sccm-with-powershell/
                # Update the deployment with some additional params not available in the cmdlet
                [xml]$DT = $ADRDeployment.DeploymentTemplate

                # If any update in this deployment requires a system restart, run updates deployment evaluation cycle after restart
                If ($null -eq $DT.DeploymentCreationActionXML.RequirePostRebootFullScan)
                {
                    $NewChild = $DT.CreateElement("RequirePostRebootFullScan")
                    [void]$DT.SelectSingleNode("DeploymentCreationActionXML").AppendChild($NewChild)
                }
                $DT.DeploymentCreationActionXML.RequirePostRebootFullScan = "Checked"
                $ADRDeployment.DeploymentTemplate = $DT.OuterXml
                $ADRDeployment.Put()
            }
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

function New-CMFTWADR
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
		if(-not (Get-CMSoftwareUpdateAutoDeploymentRule -Name $($ADRInfo.Name) -Fast)){
            $commandline = Process-Parameters -Object $ADRInfo -ExcludeParams "additionaldeployments"
		    Write-Output "  + Creating Automatic Deployment Rule: $commandline"

		    if($WhatIf -eq $false)
		    {
			    $commandline = "New-CMSoftwareUpdateAutoDeploymentRule $commandline"
			    $ADRDeployment = Invoke-Expression -Command $commandline

                #Source: https://smsagent.blog/2019/01/31/creating-adr-deployments-in-sccm-with-powershell/
                # Update the deployment with some additional params not available in the cmdlet
                [xml]$DT = $ADRDeployment.DeploymentTemplate

                # If any update in this deployment requires a system restart, run updates deployment evaluation cycle after restart
                If ($null -eq $DT.DeploymentCreationActionXML.RequirePostRebootFullScan)
                {
                    $NewChild = $DT.CreateElement("RequirePostRebootFullScan")
                    [void]$DT.SelectSingleNode("DeploymentCreationActionXML").AppendChild($NewChild)
                }
                $DT.DeploymentCreationActionXML.RequirePostRebootFullScan = "Checked"
                $ADRDeployment.DeploymentTemplate = $DT.OuterXml
                $ADRDeployment.Put()
		    }
        }else{
            Write-Output " = Automatic Deployment Rule '$($ADRInfo.Name)' already exists."
        }

		if($ADRInfo.additionaldeployments)
		{
			$ADRInfo.additionaldeployments | New-CMFTWADRDeployment -ADRName $ADRInfo.Name -TotalADRDeploymentCount ($ADRInfo.additionaldeployments | Measure-Object).Count
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
    ForEach-Object { Set-Variable -Name $_.Name -Value $buildObjects.defaultitems.variables.($_.Name) }

if($buildObjects.defaultitems.schedules)
{
	$schedules = ($buildObjects.defaultitems.schedules | New-CMFTWSchedule)
}

if($Collections -eq $true -and $buildObjects.defaultitems.devicecollectionfolders)
{
	$buildObjects.defaultitems.devicecollectionfolders | New-CMFTWDeviceCollectionFolder -Path "$($siteCode):\DeviceCollection" -Schedules $schedules -TotalFolderCount ($buildObjects.defaultitems.devicecollectionfolders | Measure-Object).Count
}

if($Collections -eq $true -and $buildObjects.defaultitems.maintenancewindows)
{
    $buildObjects.defaultitems.maintenancewindows | New-CMFTWMaintenanceWindow -Schedules $schedules -TotalMWCount ($buildObjects.defaultitems.maintenancewindows | Measure-Object).Count
}

if($ClientSettings -eq $true -and $buildObjects.defaultitems.clientsettings)
{
	$buildObjects.defaultitems.clientsettings | New-CMFTWClientSettings -Schedules $schedules -TotalClientSettingsCount ($buildObjects.defaultitems.clientsettings | Measure-Object).Count
}

if($ADRs -eq $true)
{
    Pop-Location

	if($buildObjects.defaultitems.updatepackages)
	{
		$buildObjects.defaultitems.updatepackages | New-CMFTWUpdatePackage -TotalPackageCount ($buildObjects.defaultitems.updatepackages | Measure-Object).Count
	}

	Push-Location $siteCode":"

    if($buildObjects.defaultitems.automaticdeploymentrules)
	{
		$buildObjects.defaultitems.automaticdeploymentrules | New-CMFTWADR -Schedules $schedules -TotalADRCount ($buildObjects.defaultitems.automaticdeploymentrules | Measure-Object).Count
	}
}

Pop-Location
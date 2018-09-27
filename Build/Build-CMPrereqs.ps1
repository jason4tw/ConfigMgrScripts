[CmdletBinding()]
param
(
	[Parameter(Position = 0, HelpMessage = 'Make the site system a DP')]
	[switch]$DP,
	[Parameter(Position = 1, HelpMessage = 'PXE enable the DP')]
	[switch]$PXE = $true,
	[Parameter(Position = 2, HelpMessage = 'Make the site system an MP')]
	[switch]$MP,
	[Parameter(Position = 3, HelpMessage = 'Make the site system a SUP')]
	[switch]$SUP,
	[Parameter(Position = 4, HelpMessage = 'Location to store WSUS Updates (will be created if does not exist).')]
	[string]$WSUSDir,
	[Parameter(Position = 5, HelpMessage = 'The SQL Server instance to use.', Mandatory = $False)]
	[string]$SQLInstance = [System.Net.Dns]::GetHostByName((hostname)).HostName,
	[Parameter(Position = 6, HelpMessage = 'Make the site system an Endpoint Protection Point')]
	[switch]$EPPP,
	[Parameter(Position = 7, HelpMessage = 'Make the site system an AI Sync Point')]
	[switch]$AISP,
	[Parameter(Position = 8, HelpMessage = 'Make the site system an Application Catalog Web site and webservice Point')]
	[switch]$AppCat
)

Function Start-Countdown 
{   
    Param
    (
        [Int32]$Seconds = 10,
        [string]$Message = "Pausing ..."
    )
    ForEach ($elapsed in (1..$Seconds))
    {   
        Write-Progress -Id 1 -Activity $Message -Status "Pausing for $Seconds seconds" -CurrentOperation "$($Seconds - $elapsed) left" -PercentComplete (($elapsed / $Seconds) * 100)
        Start-Sleep -Seconds 1
    }
    
    Write-Progress -Id 1 -Activity $Message -Status "Completed" -PercentComplete 100 -Completed
}

Function Decline-Updates
{
    Param
    (
        [string]$Text,
		[switch]$Superseded,
		[Microsoft.UpdateServices.Administration.IUpdateServer]$WSUSserver
    )
	
	$i = 0
	
	if($Superseded)
	{
		Write-Output " + Declining Superseded updates."	
		
		$allUpdates = $WSUSserver.GetUpdates()
		$countSupersededAll = 0
		
		foreach ($update in $allUpdates)
		{
			if(!$update.IsDeclined -and $update.IsSuperseded)
			{
				$countSupersededAll++
			}
		}

		foreach ($update in $allUpdates) 
		{
            
			if (!$update.IsDeclined -and $update.IsSuperseded)
			{
				$i++
				$percentComplete = "{0:N2}" -f (($i / $countSupersededAll) * 100)
				
				Write-Progress -Activity "Declining Superseded Updates" -CurrentOperation "$($update.Title)" -Status "($i of $countSupersededAll)" -PercentComplete $percentComplete
				
				try 
				{
					$update.Decline()
				}
				catch [System.Exception]
				{
					Write-Host "Failed to decline update $($update.Id.UpdateId.Guid). Error:" $_.Exception.Message
				}
			}
        } 

		Write-Progress -Activity "Declining Superseded Updates" -Completed

	}
	else
	{
		Write-Output " + Declining $Text updates."
			
		$approveState = 'Microsoft.UpdateServices.Administration.ApprovedStates' -as [type]
		
		$updateScope = New-Object Microsoft.UpdateServices.Administration.UpdateScope -Property @{
			TextIncludes = $Text
			ApprovedStates = $approveState::Any
		}
		
		$allFoundUpdates = $WSUSserver.GetUpdates($updateScope)
		$foundUpdateCount = $allFoundUpdates.count
		
		ForEach ($update in $allFoundUpdates)
		{
			$i++
			
			Write-Progress -Activity "Declining $text updates" -CurrentOperation "$($update.Title)" -Status "($i of $foundUpdateCount)" -PercentComplete (($i / $foundUpdateCount) * 100)
			
			$update.Decline()
		}
		
		Write-Progress -Activity "Declining $text Updates" -Completed
	}
}

$windowsFeatures = @()

if($MP -eq $true)
{
	$windowsFeatures += @("NET-Framework-45-Features","NET-Framework-45-Core","RSAT","RSAT-Role-Tools","Web-Server","Web-WebServer","Web-Common-Http","Web-Static-Content","Web-Default-Doc","Web-App-Dev","Web-ISAPI-Ext","Web-ISAPI-Filter","Web-Security","Web-Windows-Auth","Web-Mgmt-Tools","Web-Mgmt-Console","Web-Mgmt-Compat","Web-Metabase","Web-WMI","BITS","BITS-IIS-Ext")
}

if($DP -eq $true)
{
	$windowsFeatures += @("RSAT","RSAT-Role-Tools","Web-Server","Web-WebServer","Web-Common-Http","Web-Static-Content","Web-Default-Doc","Web-App-Dev","Web-ISAPI-Ext","Web-ISAPI-Filter","Web-Security","Web-Windows-Auth","Web-Mgmt-Tools","Web-Mgmt-Console","Web-Mgmt-Compat","Web-Metabase","Web-WMI","RDC")
	
	if($PXE -eq $true)
	{
		$windowsFeatures += @("WDS","WDS-Deployment")
	}
}

if($SUP -eq $true)
{
	$windowsFeatures = @("Web-Server","Web-WebServer","Web-Common-Http","Web-Static-Content","Web-Default-Doc","UpdateServices-Services","UpdateServices-DB")
}

if($EPPP -eq $true)
{
	$windowsFeatures += @("NET-Framework-45-Features","NET-Framework-45-Core")
}

if($AISP -eq $true)
{
	$windowsFeatures += @("NET-Framework-45-Features","NET-Framework-45-Core")
}

if($AppCat -eq $true)
{
	$windowsFeatures += @("NET-Framework-45-Features","NET-Framework-45-Core","Web-Server","Web-WebServer","Web-Common-Http","Web-Static-Content","Web-Default-Doc","Web-App-Dev","NET-HTTP-Activation","NET-Non-HTTP-Activ","ASP.NET 4.6","Web-Metabase")
}


$windowsFeatureList = ($windowsFeatures | Sort-Object | Get-Unique)

$windowsFeatureCount = $windowsFeatureList.Count
$i = 0

$windowsFeatureList | ForEach-Object {
	$i++
	Write-Progress -id 1 -Activity "Windows Feature Installation" -Status "$($_), $($i) of $($windowsFeatureCount)" -PercentComplete (($i / $windowsFeatureCount)*100)
	Write-Output "   + Installing: $($_)"
	
	Add-WindowsFeature $_  -IncludeManagementTools | Out-Null
}

Write-Progress -id 1 -Activity "Windows Feature Installation" -Completed

if($SUP -eq $true)
{
	# Create WSUS Updates folder if doesn't exist
	if (-not (Test-Path $WSUSDir))
	{
		Write-Output ' + Creating WSUS content folder.'
		New-Item $WSUSDir -type directory | Out-null
	}
	
	Start-Countdown -Seconds 120

	# Run WSUS Post-Configuration
	Write-Output ' + Running WSUS Post-install.'

	Push-Location "C:\Program Files\Update Services\Tools"
			
	.\wsusutil.exe postinstall SQL_INSTANCE_NAME=$SQLInstance CONTENT_DIR=$WSUSDir
	
	Start-Countdown -Seconds 60
	
	Write-Output ' + Configuring WSUS.'

	# Get WSUS Server Object
	$wsus = Get-WSUSServer
	 
	# Connect to WSUS server configuration
	$wsusConfig = $wsus.GetConfiguration()
	 
	# Set to download updates from Microsoft Updates
	Set-WsusServerSynchronization -SyncFromMU
	 
	# Set Update Languages to English and save configuration settings
	$wsusConfig.AllUpdateLanguagesEnabled = $false
	$wsusConfig.SetEnabledUpdateLanguages("en")
	$wsusConfig.Save()
	
	Write-Output ' + Synchronizing categories.'
	
	# Get WSUS Subscription and perform initial synchronization to get latest categories
	$subscription = $wsus.GetSubscription()
	$subscription.StartSynchronizationForCategoryOnly()
	
	While ($subscription.GetSynchronizationStatus() -ne 'NotProcessing') {
		Write-Host "." -NoNewline
		Start-Sleep -Seconds 30
	}

	Write-Host ""
	
	Write-Output ' + Configuring Products and Classifications.'

	Get-WsusProduct | Set-WsusProduct -Disable
	
	Get-WsusProduct | Where-Object {
		$_.Product.Title -in ('Forefront Endpoint protection 2010','Silverlight','Windows Defender','Windows 7','Windows 10','Windows Server 2012 R2','Windows Server 2016','Office 365 Client')
	} | Set-WsusProduct
	
	Get-WsusClassification | Set-WsusClassification -Disable
	
	Get-WsusClassification | Where-Object {
		$_.Classification.Title -in ('Critical Updates','Definition Updates', 'Feature Packs', 'Security Updates', 'Service Packs', 'Update Rollups', 'Updates', 'Upgrades')
	} | Set-WsusClassification
	
	Write-Output ' + Synchronizing updates.'
	
	$subscription.StartSynchronization()
	
	#while($subscription.GetSynchronizationProgress().ProcessedItems -eq 0)
	#{
	#	Start-Sleep -Seconds 5
	#}
	
	do
	{
		Start-Sleep -Seconds 5
		
		$processedItems = $subscription.GetSynchronizationProgress().ProcessedItems
		$totalItems = $subscription.GetSynchronizationProgress().TotalItems
		
		if($totalItems > 0)
		{
			Write-Progress -Activity "Synchronizing WSUS" -Status "($processedItems of $totalItems)" -PercentComplete (($processedItems / $totalItems) * 100)
		}
		
	} while($processedItems -lt $totalItems -or $subscription.GetSynchronizationStatus() -eq 'Running')
	
	#while ($subscription.GetSynchronizationProgress().ProcessedItems -ne $subscription.GetSynchronizationProgress().TotalItems)
	#{
	#	Write-Progress -Activity "Synchronizing WSUS" -Status "$($subscription.GetSynchronizationProgress().ProcessedItems) of $($subscription.GetSynchronizationProgress().TotalItems)" -PercentComplete #($subscription.GetSynchronizationProgress().ProcessedItems * 100/ ($subscription.GetSynchronizationProgress().TotalItems))
	#	Start-Sleep -Seconds 5
	#}
	
	Write-Progress -Completed -Activity "Synchronizing WSUS"
	
	Decline-Updates -Text "Itanium" -WSUSserver $wsus
	Decline-Updates -Text "ia64" -WSUSserver $wsus
	Decline-Updates -Text "Technical Preview" -WSUSserver $wsus
	Decline-Updates -Text "Beta" -WSUSserver $wsus
	Decline-Updates -Superseded -WSUSserver $wsus
	
	Pop-Location
}

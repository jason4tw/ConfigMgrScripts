[CmdletBinding()]
param
(
	[Parameter(HelpMessage = 'The server to install the role on.',Mandatory = $False)]
	[string]$ServerName = [System.Net.Dns]::GetHostByName((hostname)).HostName,	
	[Parameter(ParameterSetName = "DP", HelpMessage = 'Make the site system a DP')]
	[switch]$DP,
	[Parameter(ParameterSetName = "DP", HelpMessage = 'PXE enable the DP')]
	[switch]$PXE = $true,
	[Parameter(ParameterSetName = "MP", HelpMessage = 'Make the site system an MP')]
	[switch]$MP,
	[Parameter(ParameterSetName = "SUP", HelpMessage = 'Make the site system a SUP')]
	[switch]$SUP,
	[Parameter(ParameterSetName = "SUP", HelpMessage = 'Location to store WSUS Updates (will be created if does not exist).', Mandatory = $True)]
	[string]$WSUSDir,
	[Parameter(ParameterSetName = "SUP", HelpMessage = 'The SQL Server instance to use.', Mandatory = $False)]
	[string]$SQLInstance = [System.Net.Dns]::GetHostByName((hostname)).HostName,
	[Parameter(HelpMessage = 'Make the site system an Endpoint Protection Point')]
	[switch]$EPPP,
	[Parameter(HelpMessage = 'Make the site system an AI Sync Point')]
	[switch]$AISP,
	[Parameter(HelpMessage = 'Make the site system an Application Catalog Web site and webservice Point')]
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

Function SUP-Sync
{
	
	Sync-CMSoftwareUpdate -FullSync $True

	Start-Countdown -Seconds 30

	Write-Output " ... Waiting for synchronization ..."
	
	$syncSuccess = $null
	
	do
	{
		Start-Sleep 30
		$syncSuccess = Get-Content -Path "$($env:SMS_LOG_PATH)\wsyncmgr.log" -tail 10 | Where-Object { $_ -like 'Sync succeeded*' }
	} while (-not $syncSuccess)
}

if (-not(Get-Module ConfigurationManager))
{
    import-module ($Env:SMS_ADMIN_UI_PATH.Substring(0, $Env:SMS_ADMIN_UI_PATH.Length - 5) + '\ConfigurationManager.psd1')

    if (-not(Get-Module ConfigurationManager))
    {
        Write-Error "Could not load Configuration Manager module."
        Exit 1
    }
}

$PSD = Get-PSDrive -PSProvider CMSite
$siteCode = $PSD.Name
$siteServer = $PSD.Root

Write-Output "Installing prerequsities"

#Invoke-Command -ComputerName $ServerName -FilePath .\CMBuild-Prereqs.ps1 -ArgumentList $DP, $PXE, $MP, $SUP, $WSUSDir, $SQLInstance

Push-Location "$($PSD):"

if (-not (Get-CMSiteSystemServer -SiteSystemServerName $ServerName -SiteCode $siteCode))
{

    Write-Output "Installing $ServerName as a site system."

    New-CMSiteSystemServer -ServerName $ServerName -SiteCode $siteCode

    if (-not (Get-CMSiteSystemServer -SiteSystemServerName $ServerName -SiteCode $siteCode))
    {
        Write-Error "The Site System $($ServerName) has not been created. Please check the logs for further information"
		
		Pop-Location

        exit 1
    }
}
else
{
    Write-Output "$ServerName is already a site system."
}

if($DP -eq $true)
{
	if (-not (Get-CMDistributionPoint -SiteSystemServerName $ServerName -SiteCode $siteCode))
	{
		Write-Output "Installing $ServerName as a Distribution Point."

        $certExpire = [DateTime]::Now.AddYears(5)

		Add-CMDistributionPoint -SiteSystemServerName $ServerName -SiteCode $siteCode -EnableContentValidation:$true -EnablePXE:$($PXE) -AllowPxeResponse:$($PXE) -CertificateExpirationTimeUtc $certExpire -EnableUnknownComputerSupport:$true

		if (-not (Get-CMDistributionPoint -SiteSystemServerName $ServerName -SiteCode $siteCode))
		{
			Write-Error "The Distribution Point role has not been added to $($ServerName). Please check the logs for further information"

		}
	}
}

if($MP -eq $true)
{
	if (-not (Get-CMManagementPoint -SiteSystemServerName $ServerName -SiteCode $siteCode))
	{
		Write-Output "Installing $ServerName as a Management Point."

		Add-CMManagementPoint -SiteSystemServerName $ServerName -SiteCode $siteCode

		if (-not (Get-CMManagementPoint -SiteSystemServerName $ServerName -SiteCode $siteCode))
		{
			Write-Error "The Management Point role has not been added to $($ServerName). Please check the logs for further information"

		}
	}
}

if($EPPP -eq $true)
{
	if (-not (Get-CMEndpointProtectionPoint -SiteSystemServerName $ServerName -SiteCode $siteCode))
	{
		Write-Output "Installing $ServerName as a Management Point."

		Add-CMEndpointProtectionPoint -SiteSystemServerName $ServerName -SiteCode $siteCode -LicenseAgreed $true -ProtectionService AdvancedMembership

		if (-not (Get-CMEndpointProtectionPoint -SiteSystemServerName $ServerName -SiteCode $siteCode))
		{
			Write-Error "The Management Point role has not been added to $($ServerName). Please check the logs for further information"

		}
	}
}

if($AISP -eq $true)
{
	if (-not (Get-CMAssetIntelligenceSynchronizationPoint -SiteSystemServerName $ServerName -SiteCode $siteCode))
	{
		Write-Output "Installing $ServerName as a Management Point."

		Add-CMAssetIntelligenceSynchronizationPoint -SiteSystemServerName $ServerName -SiteCode $siteCode

		if (-not (Get-CMAssetIntelligenceSynchronizationPoint -SiteSystemServerName $ServerName -SiteCode $siteCode))
		{
			Write-Error "The Management Point role has not been added to $($ServerName). Please check the logs for further information"

		}
	}
}

if($AppCat -eq $true)
{
	if (-not (Get-CMApplicationCatalogWebServicePoint -SiteSystemServerName $ServerName -SiteCode $siteCode))
	{
		Write-Output "Installing $ServerName as a Management Point."

		Add-CMApplicationCatalogWebServicePoint -SiteSystemServerName $ServerName -SiteCode $siteCode

		if (-not (Get-CMApplicationCatalogWebServicePoint -SiteSystemServerName $ServerName -SiteCode $siteCode))
		{
			Write-Error "The Management Point role has not been added to $($ServerName). Please check the logs for further information"

		}
	}
	
	if (Get-CMApplicationCatalogWebServicePoint -SiteSystemServerName $ServerName -SiteCode $siteCode -and -not (Get-CMApplicationCatalogWebsitePoint -SiteSystemServerName $ServerName -SiteCode $siteCode))
	{
		Write-Output "Installing $ServerName as a Management Point."

		Add-CMApplicationCatalogWebsitePoint -SiteSystemServerName $ServerName -SiteCode $siteCode -ApplicationWebServicePointServerName $ServerName

		if (-not (Get-CMApplicationCatalogWebsitePoint -SiteSystemServerName $ServerName -SiteCode $siteCode))
		{
			Write-Error "The Management Point role has not been added to $($ServerName). Please check the logs for further information"

		}
	}
}

if($SUP -eq $true)
{
	if (-not (Get-CMSoftwareUpdatePoint -SiteSystemServerName $ServerName -SiteCode $siteCode))
	{
		Write-Output "Installing $ServerName as a Software Update Point."

		Add-CMSoftwareUpdatePoint -SiteSystemServerName $ServerName -SiteCode $siteCode -ClientConnectionType Intranet -WsusIisPort 8530 -WsusIisSslPort 8531

		if (-not (Get-CMSoftwareUpdatePoint -SiteSystemServerName $ServerName -SiteCode $siteCode))
		{
			Write-Error "The Software Update Point role has not been added to $($ServerName). Please check the logs for further information"

		}
		else
		{
			Write-Output " + Removing products, classifications, and languages."

			$SUPComp = Get-CMSoftwareUpdatePointComponent
			Set-CMSoftwareUpdatePointComponent -InputObject $SUPComp -RemoveUpdateClassification 'Critical Updates','Definition Updates', 'Feature Packs', 'Security Updates', 'Service Packs', 'Tools', 'Update Rollups', 'Updates', 'Upgrades'
			Set-CMSoftwareUpdatePointComponent -InputObject $SUPComp -RemoveCompany Microsoft
			Set-CMSoftwareUpdatePointComponent -InputObject $SUPComp -RemoveLanguageUpdateFile French, German, Japanese, Russian, 'Chinese (Simplified, China)'
			Set-CMSoftwareUpdatePointComponent -InputObject $SUPComp -RemoveLanguageSummaryDetail French, German, Japanese, Russian, 'Chinese (Simplified, China)'

			Write-Output " + Configuring synchronization."

			$syncSched = New-CMSchedule -DayOfWeek Tuesday -Start "$(Get-Date -Format d) 6:00 PM"
			Set-CMSoftwareUpdatePointComponent -InputObject $SUPComp -Schedule $syncSched
			Set-CMSoftwareUpdatePointComponent -InputObject $SUPComp -EnableCallWsusCleanupWizard $true
		
			Start-Countdown -Seconds 300

			Write-Output " + Initiating synchronization."

			SUP-Sync

			Write-Output " + Adding products and classifications."

			Set-CMSoftwareUpdatePointComponent -InputObject $SUPComp -AddUpdateClassification 'Critical Updates','Definition Updates', 'Feature Packs', 'Security Updates', 'Service Packs', 'Update Rollups', 'Updates', 'Upgrades'
			Set-CMSoftwareUpdatePointComponent -InputObject $SUPComp -AddProduct 'Forefront Endpoint protection 2010','Silverlight','Windows Defender','Windows 7','Windows 10','Windows Server 2012 R2','Windows Server 2016','Office 365 Client'

			Write-Output " + Initiating synchronization."

			SUP-Sync
		}
	}
}

Pop-Location
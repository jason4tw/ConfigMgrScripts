<#
	.SYNOPSIS
		Imports drivers for a specific location and creates a package for them.
	
	.DESCRIPTION
        Imports drivers for a specific location and creates a package for them.
	
	.PARAMETER Model
        The model name for the drivers being imported.
	.PARAMETER OS
        The OS for the drivers being imported.
    .PARAMETER Vendor
        The Vendor for the drivers being imported.
    .PARAMETER Architecture
        The Architecture for the drivers being imported.
    .PARAMETER ImportSource
        The source location to import the drivers from. This should *not* be the same as the package source location. 
This location and the dirvers in it should be considered imutable once the drivers are imported into ConfigMgr.
    .PARAMETER PackageSourceRoot
        The location of the package source folder. Driver files will be copied to this location.
    .EXAMPLE
        .\Import-CMDrivers.ps1 -Model 8020 -Vendor Dell -Architecure x64 -ImportSource \\cm301\ConfigMgr\Import\Drivers\Dell\8020\x64
    .EXAMPLE
        .\Import-CMDrivers.ps1 -Model 8020 -Vendor Dell -Architecure x64 -ImportSource \\cm301\ConfigMgr\Import\Drivers\Dell\8020\x64 -PackageSourceRoot \\cm301\ConfigMgr\Content\Drivers

    .NOTES
        Version 1.0
        Jason Sandys

        Version History
        - 1.0 (28 September 2018): Initial Version

        Limitations and Issues
       
#>

[CmdletBinding()]
Param
(
   [Parameter(Mandatory=$true)]
        [string]$Model,
   [Parameter(Mandatory=$true)]
        [ValidateSet("Win7","Win8,Win10")] 
        [string]$OS,   
   [Parameter(Mandatory=$true)]
        [ValidateSet("HP","Lenovo","Dell","Panasonic")] 
        [string]$Vendor,   
   [Parameter(Mandatory=$true)]
        [ValidateSet("x86","x64")] 
        [string]$Architecture,
   [Parameter(Mandatory=$true)]
        [string]$ImportSource,
   [Parameter(Mandatory=$false)]
        [string]$PackageSourceRoot = "\\cm301\ConfigMgr\Content\Drivers"

)  

$packageName = "$Vendor $Model - $OS $Architecture"
$packageSourceLocation = "$PackageSourceRoot\$Vendor\$Model\$OS-$Architecture"

# Verify Driver Source exists.
Write-Host "Checking for " $ImportSource

If (Get-Item $ImportSource -ErrorAction SilentlyContinue)
{
    $presentLocation = (Get-Location)

    # Get driver files
    $infFiles = Get-ChildItem -Path $ImportSource -Recurse -Filter "*.inf"

    # Import ConfigMgr module
    Import-Module ($Env:SMS_ADMIN_UI_PATH.Substring(0,$Env:SMS_ADMIN_UI_PATH.Length-5) + '\ConfigurationManager.psd1')

    $PSD = Get-PSDrive -PSProvider CMSite

    Push-Location "$($PSD):"

    $driverPackage = Get-CMDriverPackage -Name $packageName

    If ($driverPackage)
    {
        Write-Host "$packageName Already exists."
    }
    else
    {
        Write-Host "Creating new Driver Package: " $packageName
        
        If (Get-Item FileSystem::$packageSourceLocation -ErrorAction SilentlyContinue)
        {
            Write-Warning "$packageSourceLocation already exists… "
            Set-Location $presentLocation
            Exit 1
        }
        else
        {
            Write-Host "Creating Driver package source directory $packageSourceLocation"
            New-Item -ItemType Directory FileSystem::$packageSourceLocation | Out-Null
        }

        $driverPackage = New-CMDriverPackage -Name $packageName -Path $packageSourceLocation
        #Get-CMDriverPackage -Name $packageName
    }

    $modelCategory = Get-CMCategory -Name $Model

    if(-not $modelCategory)
    {
        $modelCategory = New-CMCategory -CategoryType DriverCategories -Name $Model
    }

    $architectureCategory = Get-CMCategory -Name $Architecture

    if(-not $architectureCategory)
    {
        $architectureCategory = New-CMCategory -CategoryType DriverCategories -Name $Architecture
    }

    $osCategory = Get-CMCategory -Name $OS
        
    if(-not $osCategory)
    {
        $osCategory = New-CMCategory -CategoryType DriverCategories -Name $OS 
    }
        
    $vendorCategory = Get-CMCategory -Name $Vendor

    if(-not $vendorCategory)
    {
        $vendorCategory = New-CMCategory -CategoryType DriverCategories -Name $Vendor 
    }
        
    $categories = @((Get-CMCategory -Name $Model), (Get-CMCategory -Name $Architecture), (Get-CMCategory -Name $OS), (Get-CMCategory -Name $Vendor))

    $totalInfCount = $infFiles.count
    $driverCounter = 0

    $driversIds = @()
    $driverSourcePaths = @()
        
    foreach($driverFile in $infFiles)
    {
        Write-Progress -Id 1 -Activity "Importing Drivers" -CurrentOperation "Importing: $($driverFile.Name)" -Status "($($driverCounter + 1) of $totalInfCount)" -PercentComplete ($driverCounter++ / $totalInfCount * 100)
        Write-Output $driverFile.FullName
            
        try
        {
            $importedDriver = (Import-CMDriver -UncFileLocation $driverFile.FullName -ImportDuplicateDriverOption AppendCategory -EnableAndAllowInstall $True -AdministrativeCategory $categories -UpdateDistributionPointsforDriverPackage $False | Select *)
            #$importedDriver = (Import-CMDriver -UncFileLocation $driverFile.FullName -ImportDuplicateDriverOption AppendCategory -EnableAndAllowInstall $True -AdministrativeCategory $categories -DriverPackage $driverPackage -UpdateDistributionPointsforDriverPackage $False | Select *)
            
            if($importedDriver)
            {
               #Write-Progress -Id 1 -Activity "Importing Drivers" -CurrentOperation "Adding to `"$packageName`" driver package: $($driverFile.Name):" -Status "($driverCounter of $totalInfCount)" -PercentComplete ($driverCounter / $totalInfCount * 100)
               # Add-CMDriverToDriverPackage -DriverId $importedDriverID -DriverPackageName $packageName

               $driverContent = Get-WmiObject -Namespace "root\sms\site_$($PSD)" -class SMS_CIToContent -Filter "CI_ID='$($importedDriver.CI_ID)'"

               $driversIds += $driverContent.ContentID
               $driverSourcePaths += $importedDriver.ContentSourcePath

            }
        }
        catch
        {
        }
    }

    Write-Progress -Id 1 -Activity "Importing Drivers" -Completed

    #$rawDriverPackage = Get-WmiObject -Namespace "root\sms\site_$($PSD)" -ClassName SMS_DriverPackage -Filter "PackageID='$($driverPackage.PackageID)'"

    #$result = $rawDriverPackage.AddDriverContent($driversIds, $driverSourcePaths, $false)

    #if($result.ReturnValue -ne 0)
    #{
    #     Write-Error "Could not import: $($result.ReturnValue)."
    #}

    Set-Location $presentLocation
}
else
{
    Write-Warning "Driver Source not found. Cannot continue"
}

Pop-Location

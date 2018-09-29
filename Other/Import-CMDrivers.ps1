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
    .PARAMETER CreateConsoleFolders
        This switch instructs the script to create folders in the console for the imported drivers.

    .EXAMPLE
        .\Import-CMDrivers.ps1 -Model 8020 -Vendor Dell -Architecure x64 -ImportSource \\cm301\ConfigMgr\Import\Drivers\Dell\8020\x64
    .EXAMPLE
        .\Import-CMDrivers.ps1 -Model 8020 -Vendor Dell -Architecure x64 -ImportSource \\cm301\ConfigMgr\Import\Drivers\Dell\8020\x64 -PackageSourceRoot \\cm301\ConfigMgr\Content\Drivers
    .EXAMPLE
        .\Import-CMDrivers.ps1 -Model 8020 -Vendor Dell -Architecure x64 -CreateConsoleFolders -ImportSource \\cm301\ConfigMgr\Import\Drivers\Dell\8020\x64

    .NOTES
        Version 1.2
        Jason Sandys

        Version History
        - 1.0 (28 September 2018): Initial Version
        - 1.1 (28 Septmeber 2018): Corrected OS validation syntax error.
                                    Fixed script to preroply add drivers to package.
                                    Removed old, stale code.
        - 1.1.1 (28 September 2018): Fixed order of file enumeration and changing to the CM provider drive
        - 1.2 (29 September 2018): Added console folder creation

        Limitations and Issues

       
#>

[CmdletBinding()]
Param
(
   [Parameter(Mandatory=$true)]
        [string]$Model,
   [Parameter(Mandatory=$true)]
        [ValidateSet("Win7","Win8","Win10")] 
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
        [string]$PackageSourceRoot = "\\cm301\ConfigMgr\Content\Drivers",
    [Parameter(Mandatory=$false)]
        [switch]$CreateConsoleFolders

)

$packageName = "$Vendor $Model - $OS $Architecture"
$packageSourceLocation = "$PackageSourceRoot\$Vendor\$Model\$OS-$Architecture"

# Verify Driver Source exists.
Write-Host "Checking for import source location: $ImportSource"

If (-not(Get-Item $ImportSource -ErrorAction SilentlyContinue))
{
    Write-Warning "Driver Source not found. Cannot continue."
}
else
{
    # Import ConfigMgr module
    Import-Module ($Env:SMS_ADMIN_UI_PATH.Substring(0,$Env:SMS_ADMIN_UI_PATH.Length-5) + '\ConfigurationManager.psd1')

    $PSD = Get-PSDrive -PSProvider CMSite

    # Get driver files
    $infFiles = Get-ChildItem -Path FileSystem::$ImportSource -Recurse -Filter "*.inf"
    Write-Host "Found $($infFiles.count) .inf files in import source location."
    
    Push-Location "$($PSD):"

    $driverPackage = Get-CMDriverPackage -Name $packageName

    If ($driverPackage)
    {
        Write-Host "$packageName already exists."
    }
    else
    {
        Write-Host "Creating new Driver Package: $packageName."
        
        If (Get-Item FileSystem::$packageSourceLocation -ErrorAction SilentlyContinue)
        {
            Write-Warning "    $packageSourceLocation already exists."
            Pop-Location
            Exit 1
        }
        else
        {
            Write-Host "Creating Driver package source directory $packageSourceLocation."
            New-Item -ItemType Directory FileSystem::$packageSourceLocation | Out-Null
        }

        $driverPackage = New-CMDriverPackage -Name $packageName -Path $packageSourceLocation
    }

    if ($CreateConsoleFolders)
    {
        $folderPath = "$($PSD.Name):\Driver\$Vendor\$Model\$OS-$Architecture"
        Write-Host "Creating console folder at $folderPath."

        if(-not(Test-Path -Path "$($PSD.Name):\Driver\$Vendor\$Model" -PathType Container))
        {
            New-Item -Path "$($PSD.Name):\Driver" -Name $Vendor -ItemType Directory
        }
        
        if(-not(Test-Path -Path "$($PSD.Name):\Driver\$Vendor\$Model" -PathType Container))        
        {
            New-Item -Path "$($PSD.Name):\Driver\$Vendor\" -Name $Model -ItemType Directory
        }
        
        if(-not(Test-Path -Path $folderPath -PathType Container))
        {
            $newFolder = New-Item -Path "$($PSD.Name):\Driver\$Vendor\$Model" -Name $OS-$Architecture -ItemType Directory
        }

        $newFolder = Get-Item $folderPath     
    }
 
    $modelCategory = Get-CMCategory -Name $Model

    if(-not $modelCategory)
    {
        Write-Host "Creating category: $Model."
        $modelCategory = New-CMCategory -CategoryType DriverCategories -Name $Model
    }

    $architectureCategory = Get-CMCategory -Name $Architecture

    if(-not $architectureCategory)
    {
        Write-Host "Creating category: $Architecture."
        $architectureCategory = New-CMCategory -CategoryType DriverCategories -Name $Architecture
    }

    $osCategory = Get-CMCategory -Name $OS
        
    if(-not $osCategory)
    {
        Write-Host "Creating category: $osCategory."
        $osCategory = New-CMCategory -CategoryType DriverCategories -Name $OS 
    }
        
    $vendorCategory = Get-CMCategory -Name $Vendor

    if(-not $vendorCategory)
    {
        Write-Host "Creating category: $Vendor."
        $vendorCategory = New-CMCategory -CategoryType DriverCategories -Name $Vendor 
    }
        
    $categories = @((Get-CMCategory -Name $Model), (Get-CMCategory -Name $Architecture), (Get-CMCategory -Name $OS), (Get-CMCategory -Name $Vendor))

    $totalInfCount = $infFiles.count
    $driverCounter = 0

    foreach($driverFile in $infFiles)
    {
        Write-Progress -Id 1 -Activity "Importing Drivers" -CurrentOperation "Importing: `"$($driverFile.Name)`"" -Status "($($driverCounter + 1) of $totalInfCount)" -PercentComplete (++$driverCounter / $totalInfCount * 100)
        Write-Host "Importing driver from $($driverFile.FullName)"
            
        try
        {
            $importedDriver = Import-CMDriver -UncFileLocation $driverFile.FullName -ImportDuplicateDriverOption AppendCategory -EnableAndAllowInstall $True -AdministrativeCategory $categories -UpdateDistributionPointsforDriverPackage $False
            
            if($importedDriver -and $driverPackage)
            {
               Write-Progress -Id 1 -Activity "Importing Drivers" -CurrentOperation "Adding `"$($driverFile.Name)`" to package: `"$packageName`"" -Status "($driverCounter of $totalInfCount)" -PercentComplete ($driverCounter / $totalInfCount * 100)
               Write-Host "Adding `"$($driverFile.Name)`" to package: `"$packageName`"."
               Add-CMDriverToDriverPackage -Driver $importedDriver -DriverPackage $driverPackage
            }

            if($importedDriver -and $newFolder)
            {
                Write-Progress -Id 1 -Activity "Importing Drivers" -CurrentOperation "Moving: `"$($driverFile.Name)`" to subfolder named `"$Vendor\$Model\$OS-$Architecture`"" -Status "($($driverCounter + 1) of $totalInfCount)" -PercentComplete ($driverCounter / $totalInfCount * 100)
                Write-Host "Moving: `"$($driverFile.Name)`" to subfolder named `"$Vendor\$Model\$OS-$Architecture`"."
                Move-CMObject -FolderPath $folderPath -InputObject $importedDriver
            }
        }
        catch
        {
        }
    }

    Write-Progress -Id 1 -Activity "Importing Drivers" -Completed

    Pop-Location

}



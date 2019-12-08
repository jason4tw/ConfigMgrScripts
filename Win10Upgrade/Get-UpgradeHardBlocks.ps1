<#
	.SYNOPSIS
		Gets hard blocks from Windows upgrade compatibility check XML file.
	
	.DESCRIPTION
		Gets hard blocks from Windows upgrade compatibility check XML file.
	
	.PARAMETER LocationToSearch
        The location to search for the upgrade compatibility files. 
        By default, C:\$WINDOWS.~BT\Sources\Panther\ is searched.

    .PARAMETER Programs
        Include program hard blocks. 

    .PARAMETER Devices
        Include device hard blocks.

    .EXAMPLE
        .\Get-UpgradeHardBlocks.ps1 -Programs -Devices
        Gets program and device hard blocks from the Windows upgrade compatibility check XML file in the default location.
		
	.NOTES
		Version 1.0
        Jason Sandys

        - 1.0 (3 December 2019): Initial Version

        Limitations and Issues
		- None.
#>

[CmdletBinding()]
param
(
	[Parameter(HelpMessage = 'The configuration file to use.')]
	[Alias('config')]
	[string] $LocationToSearch = $env:SystemDrive + '\$WINDOWS.~BT\Sources\Panther\',
	[switch] $Programs,
	[switch] $Devices

)

if(Test-Path -PathType Leaf -Path $LocationToSearch)
{
    $compat = get-childitem "$LocationToSearch\*compat*.xml" | Select-Object -Last 1
    
    if($compat)
    {
        $compatxml = (Get-Content -Path $compat.FullName) -as [xml] 
        $deviceHardBlocks = $compatxml.CompatReport.Devices.Device | Where-Object { $_.CompatibilityInfo.BlockingType -eq 'Hard' }

        $deviceHardBlocks | ForEach-Object {
            $result = New-Object -TypeName psobject
            $result | Add-member -MemberType NoteProperty -Name SystemName -Value $env:COMPUTERNAME
            $result | Add-member -MemberType NoteProperty -Name Class -Value $_.Class
            $result | Add-member -MemberType NoteProperty -Name ClassGuid -Value $_.ClassGuid
            $result | Add-member -MemberType NoteProperty -Name DeviceInstanceId -Value $_.DeviceInstanceId
            $result | Add-member -MemberType NoteProperty -Name Manufacturer -Value $_.Manufacturer
            $result | Add-member -MemberType NoteProperty -Name Model -Value $_.Model

            $result
        }
    }
}


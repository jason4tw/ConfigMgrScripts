<#
	.SYNOPSIS
        Checks DNS for the system names in a data file.
	
	.DESCRIPTION
		Checks DNS for the system names in a data file including forward and reverse lookups.
        
        A log file with all activity is generated in the current folder.
	
	.PARAMETER DataFile
        The input data file to use.

    .EXAMPLE
        .\IPCheck.ps1 -DataFile .\sys.txt
        Performs forward and reverse name lookups on the systems listed in sys.txt.
		
	.NOTES
		Version 1.1
        Jason Sandys

        Version History
        - 1.1 (15 December 2019): Added progress tracking and pipeline input
		- 1.0 (Unknown): Initial Version

        Dependencies:
        - None.
        
        Limitations and Issues
		- None.
#>

[CmdletBinding()]
param
(
	[Parameter(HelpMessage = 'The data file to use.')]
	[ValidateScript({ Test-Path -PathType Leaf -Path $_ })]
	[Alias('data')]
    [string] $DataFile
)

function Invoke-DNSLookup
{
    [CmdletBinding()]
    param
    (
        [Parameter(ValueFromPipeline, HelpMessage = 'The name of the system to lookup in DNS.')]
        [string] $Name
    )

    process
    {
        $systemName = $Name.ToLower().Trim()
    
        $systemResult = New-Object PSObject
        $systemResult | Add-Member -MemberType NoteProperty -Name Name -Value $systemName
    
        $systemResult | Add-Member -MemberType NoteProperty -Name IP -Value "-"
        $systemResult.IP = ([System.Net.DNS]::GetHostEntry($systemName).AddressList | Select-Object -First 1).IPAddressToString
    
        $systemResult | Add-Member -MemberType NoteProperty -Name Reverse -Value "-"
        $systemResult.Reverse = ([System.Net.DNS]::GetHostEntry($systemResult.IP) | Select-Object -First 1).HostName
        $systemResult.Reverse = ($systemResult.Reverse -split "[.]")[0].ToLower().Trim()
    
        $firstOctet = ($systemResult.IP -split "[.]")[0].Trim()
        if ($systemResult.Reverse -eq $firstOctet)
        {
            $systemResult.Reverse = "-"
        }
            
        $systemResult | Add-Member -MemberType NoteProperty -Name Status -Value "-"
        
        if ($systemResult.IP -eq "-")
        {
            $sys.Status = "Could not Resolve IP"
        }
        elseif ($systemResult.Reverse -eq "-")
        {
            $systemResult.Status = "IP Address not found in reverse zone"
        }
        elseif ($systemResult.Name -ne $systemResult.Reverse)
        {
            $systemResult.Status = "IP registered to another system"
        }
        else
        {
            $systemResult.Status = "OK"
        }
        
        $systemResult
    }
}

$ErrorActionPreference = "silentlycontinue"

if($DataFile)
{
    $output = New-Object -TypeName "System.Collections.ArrayList" 

    $systemNameList = Get-Content -Path $DataFile

    $total = ($systemNameList | Measure-Object).Count
    $count = 0

    Write-Progress -Id 1 -Activity "Performing DNS Lookups" -PercentComplete 0

    foreach ($systemName in $systemNameList)
    {
        $count++

        Write-Progress -Id 1 -Activity "Performing DNS Lookups" `
         -CurrentOperation $systemName `
         -Status "($count / $total)" `
         -PercentComplete (($count / $total) * 100)

        $result = Invoke-DNSLookup -Name $systemName

        $output.Add($result) | Out-Null
    }

    Write-Progress -Id 1 -Activity "Performing DNS Lookups" `
        -CurrentOperation $systemName `
        -Completed

    $output
}
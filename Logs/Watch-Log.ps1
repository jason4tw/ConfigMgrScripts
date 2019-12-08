[CmdletBinding()]
Param(
	[Parameter(Mandatory=$True)]
    [string[]] $Locations,
	
    [Parameter(Mandatory=$False)]
    [alias('Log')]
    [string] $LogFilename = "ccm.log",
	
	[Parameter(Mandatory=$True)]
    [alias('Text')]
    [string] $SearchText,
    
    [Parameter(Mandatory=$False)]
    [alias('Capture')]
    [string] $CapturePath = (Convert-Path .)
)
function Get-HostName
{
    param ([string] $FilePath)
    
    $FilePath | select-string -pattern "(?<=\\\\)(.*?)(?=\\)" `
    | Select-Object -ExpandProperty Matches | Select-Object -ExpandProperty Value
}

foreach ($location in $Locations)
{
    $path = Join-Path -Path $location -ChildPath $LogFilename
    $hostname = Get-HostName -FilePath $path

    if(Test-Path -Path $path -PathType Leaf)
    {
        $outfilename = Join-Path -Path $CapturePath -ChildPath "$hostname-$LogFilename"
        
        Write-Host "Begin monitoring '$path', output captured to '$outfilename' ..."

        Start-Job -Name "$hostname-$LogFilename" -ArgumentList $path,$SearchText,$outfilename -ScriptBlock {
            param($path, $searchText, $outfilename)

            "Begin Watching $path for '$searchText'" | Out-File -FilePath $outfilename -Force
            Get-Content -Path $path -Wait | Where-Object {$_ -match $searchText} | Out-File -FilePath $outfilename -Append
        }
    }
    else
    {
        Write-Error "Could not open $LogFilename at $location."
    }
}
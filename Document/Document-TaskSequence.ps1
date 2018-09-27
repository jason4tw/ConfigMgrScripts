[CmdletBinding()]
param
(
    [Parameter(Mandatory = $true)]
    [ValidateScript({(Test-Path $_ -PathType 'Container') -and ($_.Substring(0,2) -eq '\\')})]
    [string] $Location,
    [Parameter(Mandatory = $true)]
    [alias("TSID")]
    [string] $TaskSequenceID
)

function Export-TaskSequence 
{
    param 
    (
        [Parameter(Mandatory = $true)]
        [ValidateScript({(Test-Path $_ -PathType 'Container') -and ($_.Substring(0,2) -eq '\\')})]
        [string] $Path,
        [Parameter(Mandatory = $true)]
        [alias("TSID")]
        [string] $TaskSequenceID
    )            
 
    $currentLocation = Get-Location
    Set-Location "$((Get-PSDrive -PSProvider CMSite).Name):"
 
    $taskSequences = Get-CMTaskSequence -TaskSequencePackageId $TaskSequenceID
       
    foreach ($ts in $taskSequences) 
    {
        write-host "`nTask Sequence: $($ts.Name)"
        $tsXMLPath = ($Path + "$($ts.Name).xml") -replace ':', ''
                
        write-host "File Path: $tsXMLPath"
        Set-Location $currentLocation
                
        Write-Output '<?xml-stylesheet type="text/xsl" href="tsDocumentorv2.xsl"?>' | Out-File $tsXMLPath
        $ts.Sequence | Out-File $tsXMLPath -Append
    }

    Set-Location $currentLocation
}
$xslPath = Join-Path -Path $(Get-Location) -ChildPath 'tsDocumentorv2.xsl'

if(-not(Test-Path $xslPath -PathType 'Leaf'))
{
    Write-Error "The tsDocumentorv2.xsl does not exist"
    exit 2
}

$date = Get-Date -UFormat %m%d%Y_%H%M%S
$backupLocation = "$Location\$date\"

if (-not(Get-Module ConfigurationManager))
{ 
    Import-Module ($env:SMS_ADMIN_UI_PATH.Substring(0, $env:SMS_ADMIN_UI_PATH.Length - 5) + '\ConfigurationManager.psd1')
}

if (-not(Get-Module ConfigurationManager))
{
    Write-Error "Could not load the Configuration Manager module."
}
else
{
    New-Item $backupLocation -ItemType directory | out-null
 
    $xsl = (Get-ChildItem $psscriptroot -File tsDocumentorv2.xsl).FullName
    Copy-Item $xsl -Destination $backupLocation -Verbose
 
    Export-TaskSequence -Path $backupLocation -TaskSequenceID $TaskSequenceID       
}
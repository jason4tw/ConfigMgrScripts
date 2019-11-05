param
(
	[Parameter(Mandatory = $false)]
	[string]$Path,
	[Parameter(Mandatory = $false)]
	[ValidatePattern('[a-zA-Z0-9.-s]')]
	[alias("OS")]
	[string]$OperatingSystem = "Windows",
	[switch]$UseExisting
)

# Set the OSDCaptureDestination task sequence variable 
$TSEnv = New-Object -ComObject Microsoft.SMS.TSEnvironment;

if ($UseExisting -eq $false -and $Path -ne $null)
{
	# Generate the dynamic file path and log it 
	$filePath = '{0}\{1}-{2}.wim' -f $Path, $OperatingSystem, (Get-Date -Format 'yyyy-MM-dd hh.mm.ss');

}
else
{
	$currentPath = $TSEnv.Value('OSDCaptureDestination');
	$currentPath = "\\cm100.lab100.configmgrftw.com\ConfigMgr\ImageCapture\Win101511Ent(x64).wim"
	
	$filePath = '{0}\{1}-{2}{3}' -f [System.IO.Path]::GetDirectoryName($currentPath),
		[System.IO.Path]::GetFileNameWithoutExtension($currentPath),
		(Get-Date -Format 'yyyy-MM-dd hh.mm.ss'),
		[System.IO.Path]::GetExtension($currentPath)
}

$TSEnv.Value('OSDCaptureDestination') = $filePath;

$logText = 'Capture file path is: {0}' -f $filePath;
Set-Content -Path $env:WinDir\Temp\OSDCapture.log -Value $logText;
Write-Host -Object $logText;
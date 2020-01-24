<#
	.SYNOPSIS
		Parses Windows 10 upgrade compat xml files to find hard-blockers.
	
	.DESCRIPTION
        Parses Windows 10 upgrade compat xml files to find hard-blockers.
	
    .PARAMETER Folder
        The location where the compat xml files are located.

    .EXAMPLE
        .\Parse-Win10CompatXML.ps1 
		
	.NOTES
		Version 1.22
        Jason Sandys

        Version History
        - 1.22 (23 January 2020): Modified PSCommandPath shim to use $MyInvocation.ScriptName
        - 1.21 (23 January 2020): Modifed to use older WIM cmdlets for PowerShell 2.0 compatibility
        - 1.2 (21 January 2020): 
            - Updated to store hard blocks in WMI instead of just returning the information to stdout
            - Added logging
        - 1.1 (14 January 2020): Updated to create custom objects with compat info that can be converted to JSON
        - 1.0 (12 January 2020): Initial version

        Limitations and Issues
        - None
#>

[CmdletBinding()]
param
(
	[Parameter(Mandatory=$false,HelpMessage='The location where the compat xml files are located.')]
	[ValidateScript({ Test-Path -PathType Container -Path $_ })]
    [string] $Folder = 'C:\$WINDOWS.~BT\Sources\Panther\',
    [Parameter(Mandatory=$false)]
        [string] $Namespace = 'ITLocal',
    [Parameter(Mandatory=$false)]
        [string] $ClassPrefix = 'Win10_'
)

function Convert-HashToString
{
    param
    (
        [Parameter(Mandatory = $true)]
        [System.Collections.Hashtable]
        $Hash
    )
    $hashstr = "@{"
    $keys = $Hash.keys
    foreach ($key in $keys)
    {
        $v = $Hash[$key]
        if ($key -match "\s")
        {
            $hashstr += "`"$key`"" + "=" + "`"$v`"" + ";"
        }
        else
        {
            $hashstr += $key + "=" + "`"$v`"" + ";"
        }
    }
    $hashstr += "}"
    return $hashstr
}
function Add-LogMsg
{
    Param
    (
        [Parameter(Mandatory=$false)]
            [string] $LogPath,
        [Parameter(Mandatory=$false)]
            [string] $Message,
        [Parameter(Mandatory=$false)]
            [ValidateSet('Info','Warning','Error')]
            [string] $MessageType = 'Info',
        [switch] $AddHeader
    )

    if(-not($script:logfilename))
    {
        $script:componentName = [System.IO.Path]::GetFileNameWithoutExtension($PSCommandPath)
        $script:componentFilename = [System.IO.Path]::GetFileName($PSCommandPath)
        $script:logFilename = $script:componentName + '.log'
    }

    if($LogPath)
    {
        if(Test-Path -Path $LogPath -PathType Folder)
        {
            $log = Join-Path -Path $LogPath -ChildPath $script:logFilename
        }
        else
        {
            $log = Join-Path -Path $env:TEMP -ChildPath $script:logFilename
        }
    }
    elseif($script:logfile)
    {
        $log = $script:logfile
    }
    else
    {
        $logFolder = Get-NativeRegStringValue -Hive 'HKLM' -Key 'SOFTWARE\Microsoft\SMS\Client\Configuration\Client Properties' -ValueName 'Local SMS Path'

        if($logFolder)
        {
            $logFolder = Join-Path -Path $logFolder -ChildPath 'Logs'     
        }
        else
        {
            $logFolder = $env:TEMP
        }
        
        $script:logfile = Join-Path -Path $logFolder -ChildPath $script:logFilename
        $log = $script:logfile
    }

    if($AddHeader)
    {
        Add-LogItem -LogPath $log -Message '--------------------------------------------------------------------------------------------------------------' -MessageType 'Info' -Component $script:componentName -Filename $script:componentFilename        
        Add-LogItem -LogPath $log -Message "Running $PSCommandPath" -MessageType 'Info' -Component $script:componentName -Filename $script:componentFilename
    }

    if($Message)
    {
        Add-LogItem -LogPath $log -Message $Message -MessageType $MessageType -Component $script:componentName -Filename $script:componentFilename
    }
}
function Add-LogItem
{
    Param
    (
        [Parameter(Mandatory=$true)]
            [string] $LogPathandFilename,
        [Parameter(Mandatory=$true)]
            [string] $Message,
        [Parameter(Mandatory=$false)]
        [ValidateSet('Info','Warning','Error')]
            [string] $MessageType = 'Info',
        [Parameter(Mandatory=$false)]
            [string] $Component,
        [Parameter(Mandatory=$false)]
            [string] $Filename
    )

    switch($MessageType)
    {
        'Info' { $msgType = 1 }
        'Warning' { $msgType = 2 }
        'Error' { $msgType = 3 }
    }

    $msg = "<![LOG[$Message]LOG]!><time=`"$(Get-Date -Format 'HH:mm:ss').000+0`" date=`"$(Get-Date -Format 'MM-dd-yyyy')`" "
    $msg += "component=`"$Component`" context=`"`" type=`"$msgType`" thread=`"$pid`" file=`"$Filename`">"

    Add-Content -Path $LogPathandFilename -Value $msg -ErrorAction SilentlyContinue
}
function Get-NativeRegStringValue
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory=$false)]
            [string] $Hive = 'HKLM',
        [Parameter(Mandatory=$false)]
            [string] $Key = 'SOFTWARE\MicrosoftWindows\CurrentVersion',
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
            [string] $ValueName
    )

    if([System.Environment]::Is64BitOperatingSystem -and -not([System.Environment]::Is64BitProcess))
    {
        $wbemPropertySet = New-Object -ComObject 'WbemScripting.SWbemNamedValueSet'
        $wbemPropertySet.Add('__ProviderArchitecture', 64) | Out-null
        $wbemPropertySet.Add('__RequiredArchitecture', $true) | Out-null

        $wbemLocator = New-Object -ComObject 'Wbemscripting.SWbemLocator'
        $wbemServices = $wbemLocator.ConnectServer('.', 'root\Default', $null, $null, $null, $null, $null, $wbemPropertySet)
        $stdRegProv = $wbemServices.Get('stdRegProv')

        $params = $stdRegProv.Methods_['GetStringValue'].Inparameters

        switch ($Hive)
        {            
            'HKCR' { $hiveVal = "&h80000000" } #HKEY_CLASSES_ROOT             
            'HKCU' { $hiveVal = "&h80000001" } #HKEY_CURRENT_USER            
            'HKLM' { $hiveVal = "&h80000002" } #HKEY_LOCAL_MACHINE             
            'HKU'  { $hiveVal = "&h80000003" } #HKEY_USERS                             
            'HKCC' { $hiveVal = "&h80000005" } #HKEY_CURRENT_CONFIG            
            'HKDD' { $hiveVal = "&h80000006" } #HKEY_DYN_DATA                            
        }
        
        $params.Properties_['Hdefkey'].Value = $hiveVal
        $params.Properties_['sSubKeyName'].Value = $Key
        $params.Properties_['sValueName'].Value = $ValueName

        $result = $stdRegProv.ExecMethod_.Invoke("GetStringValue", $params, $null, $wbemPropertySet)
        $value = $result.Properties_['sValue'].Value
    }
    else
    {
        $value = (Get-ItemProperty -Path "${Hive}:$Key").$ValueName
    }

    $value
}
function Get-WMINamespace
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
            [string] $NamespaceName,
        [switch] $Create
    )

    $parentNamespace = Split-Path -Path $NamespaceName -Parent
    $leafNamespace = Split-Path -Path $NamespaceName -Leaf

    #$namespace = Get-CimInstance -Namespace $parentNamespace -ClassName __namespace | Where-Object -FilterScript { $_.Name -eq $leafNamespace }
    $namespace = Get-WmiObject -Namespace $parentNamespace -Class __namespace | Where-Object -FilterScript { $_.Name -eq $leafNamespace }

    if(-not($namespace))
    {
        Add-LogMsg -Message "Namespace '$NamespaceName' does not exist ..."

        if($Create)
        {
            try
            {
                #$namespace = New-CimInstance -Namespace $parentNamespace -ClassName __namespace -Property @{ Name = $leafNamespace }
                $rootNamespace = [wmiclass]"${parentNamespace}:__namespace"
                $namespace = $rootNamespace.CreateInstance()
                $namespace.Name = $leafNamespace
                [void] $namespace.Put()
                Add-LogMsg -Message "Successfully created the '$NamespaceName' namespace..."
            }
            catch
            {
                $namespace = $null
                Add-LogMsg -Message "Failed to create the '$NamespaceName' namespace: $_" -MessageType 'Error'
            }
        }
    }
    else
    {
        Add-LogMsg -Message "Namespace '$NamespaceName' exists."
    }

    $namespace
}
function Get-WMIClass
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
            [string] $NamespaceName,
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
            [string] $ClassName,
        [Parameter(Mandatory=$true)]
            [hashtable] $ClassProperties,
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
            [string[]] $KeyProperties,
        [switch] $Create,
        [switch] $RemoveExistingInstances
    )

    $class = $null

    if(Get-WMINamespace -NamespaceName $NamespaceName -Create)
    {
        #$class = Get-CimClass -Namespace $NamespaceName | Where-Object -FilterScript { $_.CimClassName -eq $ClassName }
        $class = Get-WMIObject -Namespace $NamespaceName -list | Where-Object -FilterScript { $_.Name -eq $ClassName }

        if(-not($class))
        {
            Add-LogMsg -Message "Class '$ClassName' does not exist in the '$NamespaceName' namespace ..."

            if($Create)
            {
                try
                {
                    $class = New-Object System.Management.ManagementClass("$NamespaceName", [string]::Empty, $null)
                    $class["__CLASS"] = $ClassName

                    foreach ($property in $ClassProperties.Keys)
                    {
                        $class.Properties.Add($property, $ClassProperties.$property, $false)
                        
                        if($KeyProperties -contains $property)
                        {
                            $class.Properties[$property].Qualifiers.Add("Key", $true)
                        }
                    }

                    [void] $class.Put()

                    Add-LogMsg -Message "Successfully created the '$ClassName' class in the $NamespaceName namespace."                   
                }
                catch
                {
                    $class = $null
                    Add-LogMsg -Message "Failed to create the '$ClassName' class in the $NamespaceName namespace: $_" -MessageType 'Error'
                }
            }

        }
        else
        {
            Add-LogMsg -Message "Class '$ClassName' exists in the '$NamespaceName' namespace."

            if($RemoveExistingInstances)
            {
                try
                {
                    #Get-CimInstance -Namespace $NamespaceName -ClassName $ClassName | Remove-CimInstance
                    Get-WMIObject -Namespace $NamespaceName -Class $ClassName | Remove-WMIObject
                    Add-LogMsg -Message "Removed all existing instances in the '$ClassName' class."
                }
                catch
                {
                    Add-LogMsg -Message "Failed to remove existing instances in the '$ClassName' class: $_"
                }
            }
        }
    }

    $class
}
function Add-WMIObject
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory=$true,ValueFromPipeline=$true)]
        [ValidateNotNullOrEmpty()]
            [System.Xml.XmlElement] $CompatItem,
        [Parameter(Mandatory=$true)]
            [string] $FullNamespaceName,
        [Parameter(Mandatory=$true)]
            [hashtable] $PropertyMap,
        [Parameter(Mandatory=$true)]
            [string] $ClassName
    )

    process
    {
        # $instanceData = @{}

        # foreach($property in $PropertyMap.Keys)
        # {
        #     $instanceData.Add($property, $_.($PropertyMap.$property))
        # }

        # try
        # {
            #New-CimInstance -Namespace $FullNamespaceName -ClassName $ClassName -Property $instanceData | Out-Null
            $class = [wmiclass]"${FullNamespaceName}:$ClassName"
            $object = $class.CreateInstance()
            foreach($property in $PropertyMap.Keys)
            {
                $object.$property = $_.($PropertyMap.$property)
            }
            [void] $object.Put()
            #Add-LogMsg -Message "Created new '$ClassName' object for $(Convert-HashToString -Hash $instanceData)"
            Add-LogMsg -Message "Created new '$ClassName' object for $(Convert-HashToString -Hash $PropertyMap)"
        # }
        # catch
        # {
        #     Add-LogMsg -Message "Failed to create new '$ClassName' object: $_" -MessageType 'Error'
        # }
    }
}

if ($PSCommandPath -eq $null) { function GetPSCommandPath() { return $MyInvocation.ScriptName; } $PSCommandPath = GetPSCommandPath; }

$deviceCompatProperties = @{ 'DeviceClass' = [System.Management.CimType]::String
                             'DeviceInstanceId' = [System.Management.CimType]::String
                             'Manufacturer' = [System.Management.CimType]::String
                             'Model' = [System.Management.CimType]::String }
$deviceCompatKeyProperties = 'DeviceInstanceId'
$deviceCompatPropertyMap = @{ 'DeviceClass' = 'Class'
                              'DeviceInstanceId' = 'DeviceInstanceId'
                              'Manufacturer' = 'Manufacturer'
                              'Model' = 'Model' }
$deviceCompatClassName = ($ClassPrefix + 'DeviceCompatBlocker')

$programCompatProperties = @{ 'ProgramName' = [System.Management.CimType]::String }
$programCompatPropertyMap = @{ 'ProgramName' = 'Name' }
$programCompatKeyProperties = 'ProgramName'
$programCompatClassName = ($ClassPrefix + 'ProgramCompatBlocker')

Add-LogMsg -AddHeader

$compatXMLFiles = Get-ChildItem -Path (Join-Path -Path $Folder -ChildPath 'compat*.xml')

if($Namespace.StartsWith('root'))
{
    $fullNamespaceName = $Namespace
}
else
{
    $fullNamespaceName = Join-Path -Path 'root' -ChildPath $Namespace
}

if((Get-WMIClass -NamespaceName $fullNamespaceName -ClassName $deviceCompatClassName `
    -ClassProperties $deviceCompatProperties -KeyProperties $deviceCompatKeyProperties `
    -Create -RemoveExistingInstances) -and `
   (Get-WMIClass -NamespaceName $fullNamespaceName -ClassName $programCompatClassName `
    -ClassProperties $programCompatProperties -KeyProperties $programCompatKeyProperties `
    -Create -RemoveExistingInstances))
{
    $compatXMLFiles | ForEach-Object {
        [xml]$compatXML = Get-Content -Path $_

        $compatXML.CompatReport.Devices.Device | `
         Where-Object -FilterScript { $_.CompatibilityInfo.BlockingType -eq 'Hard'} | `
         Add-WMIObject -FullNamespaceName $fullNamespaceName -ClassName $deviceCompatClassName `
          -PropertyMap $deviceCompatPropertyMap

        $compatxml.CompatReport.Programs.Program | `
         Where-object -filterscript { $_.Compatibilityinfo.BlockingType -eq 'hard'} | `
         Add-WMIObject -FullNamespaceName $fullNamespaceName -ClassName $programCompatClassName `
          -PropertyMap $programCompatPropertyMap
    }
}

Add-LogMsg -Message 'Finished'

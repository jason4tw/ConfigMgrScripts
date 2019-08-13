<#
.SYNOPSIS
    Enables Windows Defender and other advanced security features on Windows 10 during OS Deployment with ConfigMgr.

.EXAMPLE
    .\Enable-WinDefAdvancedSecurity.ps1

.NOTES
    Version history:
    1.0.0 - (2016-06-08) Script created.
    1.0.1 - (2016-08-10) Script updated to support Windows 10 version 1607 that no longer required the Isolated User Mode feature, since it's embedded in the hypervisor.
    1.1.0 - (2018-04-03) Script renamed updated to include features beyond just Credential Guard.

    Based on the Enable-CredentialGuard.ps1 script by Nickolaj Andersen.

    FileName:    Enable-WinDefAdvancedSecurity.ps1
    Author:      Jason Sandys
    Created:     2016-06-08
    Updated:     2018-04-03
    Version:     1.1.0
#>

[CmdletBinding()]

param(
	[parameter(HelpMessage="Enable Device Guard Virtualization Based Security (VBS)")]
	[ValidateNotNullOrEmpty()]
    [alias("VBS")]
	[switch]$DeviceGuardVBS,

	[parameter(HelpMessage="Enable Credential Guard")]
	[ValidateNotNullOrEmpty()]
    [alias("CredGuard")]
	[switch]$CredentialGuard,

	[parameter(HelpMessage="Enable Hypervisor Code Integrity (HVCI)")]
	[ValidateNotNullOrEmpty()]
    [alias("HVCI")]
	[switch]$HypervisorCodeIntegrity,

	[parameter(HelpMessage="Enable Application Guard")]
	[ValidateNotNullOrEmpty()]
    [alias("AppGuard")]
	[switch]$ApplicationGuard,

	[parameter(HelpMessage="Enable Smart Screen")]
	[ValidateNotNullOrEmpty()]
	[switch]$SmartScreen
)

Begin {
    # Construct TSEnvironment object
    try {
        $TSEnvironment = New-Object -ComObject Microsoft.SMS.TSEnvironment -ErrorAction Stop
        $LogFolder = $TSEnvironment.Value("_SMSTSLogPath")
    }
    catch [System.Exception] {
        Write-Host -Message "Unable to construct Microsoft.SMS.TSEnvironment object"
        $LogFolder = $env:TEMP

    }
}
Process {
    # Functions
    function Write-CMLogEntry {
	    param(
		    [parameter(Mandatory=$true, HelpMessage="Value added to the smsts.log file.")]
		    [ValidateNotNullOrEmpty()]
		    [string]$Value,

		    [parameter(Mandatory=$true, HelpMessage="Severity for the log entry. 1 for Informational, 2 for Warning and 3 for Error.")]
		    [ValidateNotNullOrEmpty()]
            [ValidateSet("1", "2", "3")]
		    [string]$Severity,

		    [parameter(Mandatory=$false, HelpMessage="Name of the log file that the entry will written to.")]
		    [ValidateNotNullOrEmpty()]
		    [string]$FileName = "Enable-WinDefAdvancedSecurity.log"
	    )
	    # Determine log file location
        $LogFilePath = Join-Path -Path $Script:LogFolder -ChildPath $FileName

        # Construct time stamp for log entry
        $Time = -join @((Get-Date -Format "HH:mm:ss.fff"), "+", (Get-WmiObject -Class Win32_TimeZone | Select-Object -ExpandProperty Bias))

        # Construct date for log entry
        $Date = (Get-Date -Format "MM-dd-yyyy")

        # Construct context for log entry
        $Context = $([System.Security.Principal.WindowsIdentity]::GetCurrent().Name)

        # Construct final log entry
        $LogText = "<![LOG[$($Value)]LOG]!><time=""$($Time)"" date=""$($Date)"" component=""WindowsDefenderAdvancedSecurity"" context=""$($Context)"" type=""$($Severity)"" thread=""$($PID)"" file="""">"
	
	    # Add value to log file
        try {
	        Add-Content -Value $LogText -LiteralPath $LogFilePath -ErrorAction Stop
        }
        catch [System.Exception] {
            Write-Warning -Message "Unable to append log entry to EnableCredentialGuard.log file"
        }
    }

    function Set-RegistryValue {
        param(
    		    [parameter(Mandatory=$true, HelpMessage="Path to the registry key where the value will be created or updated.")]
		        [ValidateNotNullOrEmpty()]
		        [string]$KeyPath,
    		    [parameter(Mandatory=$true, HelpMessage="The name of the registry value to create or update.")]
		        [ValidateNotNullOrEmpty()]
		        [string]$ValueName,
    		    [parameter(Mandatory=$true, HelpMessage="The type of the registry value to create.")]
		        [ValidateSet("DWord","String","Binary")] 
		        [string]$ValueType,
    		    [parameter(Mandatory=$true, HelpMessage="The actual value to set.")]
		        [ValidateNotNullOrEmpty()]
		        [string]$Value
        )

            Write-CMLogEntry -Value "Adding the value $ValueName to $KeyPath as $ValueType with data $Value" -Severity 1
            New-ItemProperty -Path $KeyPath -Name $ValueName -PropertyType $ValueType -Value $Value -Force

    }

    function Create-RegistryKey {
        param(
    		    [parameter(Mandatory=$true, HelpMessage="Path to the registry key that will be created if necessary.")]
		        [ValidateNotNullOrEmpty()]
		        [string]$KeyPath
        )

        if (-not(Test-Path -Path $KeyPath)) {
            Write-CMLogEntry -Value "Creating $KeyPath registry key" -Severity 1
            New-Item -Path $KeyPath -ItemType Directory -Force
        }

    }

    if($DeviceGuardVBS -eq $true -or $CredentialGuard -eq $true -or $HypervisorCodeIntegrity -eq $true) {
        Write-CMLogEntry -Value "Starting configuration of Device Guard Virtualizatoin Based Security (VBS)" -Severity 1

        # Enable required Windows Features
        try {
            Enable-WindowsOptionalFeature -FeatureName Microsoft-Hyper-V-HyperVisor -Online -All -LimitAccess -NoRestart -ErrorAction Stop
            Write-CMLogEntry -Value "Successfully enabled Microsoft-Hyper-V-HyperVisor feature" -Severity 1
        }
        catch [System.Exception] {
            Write-CMLogEntry -Value "An error occured when enabling Microsoft-Hyper-V-HyperVisor feature, see DISM log for more information" -Severity 3

            exit 1
        }
    
        # For version older than Windows 10 version 1607 (build 14939), add the IsolatedUserMode feature as well
        if ([int](Get-WmiObject -Class Win32_OperatingSystem).BuildNumber -lt 14393) {
            try {
                Enable-WindowsOptionalFeature -FeatureName IsolatedUserMode -Online -All -LimitAccess -NoRestart -ErrorAction Stop
                Write-CMLogEntry -Value "Successfully enabled IsolatedUserMode feature" -Severity 1
            }
            catch [System.Exception] {
                Write-CMLogEntry -Value "An error occured when enabling IsolatedUserMode feature, see DISM log for more information" -Severity 3

                exit 1
            }
        }

        # Add required registry key for Device Guard VBS
        $DeviceGuardRegistryKeyPath = "HKLM:\SYSTEM\CurrentControlSet\Control\DeviceGuard"
        Create-RegistryKey -KeyPath $DeviceGuardRegistryKeyPath

        # Add registry value RequirePlatformSecurityFeatures - 1 for Secure Boot only, 3 for Secure Boot and DMA Protection
        Set-RegistryValue -KeyPath $DeviceGuardRegistryKeyPath -ValueName RequirePlatformSecurityFeatures -ValueType DWord -Value 1

        # Add registry value EnableVirtualizationBasedSecurity - 1 for Enabled, 0 for Disabled
        Set-RegistryValue -KeyPath $DeviceGuardRegistryKeyPath -ValueName EnableVirtualizationBasedSecurity -ValueType DWord -Value 1

        # Add registry value Locked - 1 for Enabled, 0 for Disabled
        Set-RegistryValue -KeyPath $DeviceGuardRegistryKeyPath -ValueName Locked -ValueType DWord -Value 1
    
        Write-CMLogEntry -Value "Successfully enabled Device Guard VBS" -Severity 1
    }

    if($CredentialGuard -eq $true) {
        Write-CMLogEntry -Value "Starting configuration for Credential Guard" -Severity 1

        # Add registry value LsaCfgFlags - 1 enables Credential Guard with UEFI lock, 2 enables Credential Guard without lock, 0 for Disabled
        Set-RegistryValue -KeyPath "HKLM:\SYSTEM\CurrentControlSet\Control\Lsa" -ValueName LsaCfgFlags -ValueType DWord -Value 1

        Write-CMLogEntry -Value "Successfully enabled Credential Guard" -Severity 1
    }

    if($HypervisorCodeIntegrity -eq $true) {
        Write-CMLogEntry -Value "Starting configuration for Hypervisor Code Integrity (HVCI)" -Severity 1

        # Add required registry key for HVCI
        $HVCIRegistryKeyPath = "HKLM:\SYSTEM\CurrentControlSet\Control\DeviceGuard\Scenarios\HypervisorEnforcedCodeIntegrity"
        Create-RegistryKey -KeyPath $HVCIRegistryKeyPath

        Set-RegistryValue -KeyPath $HVCIRegistryKeyPath -ValueName Enabled -ValueType DWord -Value 1
        Set-RegistryValue -KeyPath $HVCIRegistryKeyPath -ValueName Locked -ValueType DWord -Value 1

        Write-CMLogEntry -Value "Successfully enabled Hypervisor Code Integrity (HVCI)" -Severity 1
    }

    if($ApplicationGuard -eq $true) {
        Write-CMLogEntry -Value "Starting configuration for Application Guard" -Severity 1
    
        # Enable required Windows Features
        try {
            Enable-WindowsOptionalFeature -FeatureName Windows-Defender-ApplicationGuard -Online -All -LimitAccess -NoRestart -ErrorAction Stop
            Write-CMLogEntry -Value "Successfully enabled Windows-Defender-ApplicationGuard feature" -Severity 1
        }
        catch [System.Exception] {
            Write-CMLogEntry -Value "An error occured when enabling Windows-Defender-ApplicationGuard feature, see DISM log for more information" -Severity 3

            exit 1
        }

        Write-CMLogEntry -Value "Successfully enabled Application Guard" -Severity 1
    }

    if($SmartScreen -eq $true) {
        Write-CMLogEntry -Value "Starting configuration for Smart Screen" -Severity 1

        Create-RegistryKey -KeyPath "HKLM:\SOFTWARE\Policies\Microsoft\MicrosoftEdge\PhishingFilter"
        Create-RegistryKey -KeyPath "HKLM:\SOFTWARE\Policies\Microsoft\Internet Explorer\PhishingFilter"
        Create-RegistryKey -KeyPath "HKLM:\SOFTWARE\Policies\Microsoft\Windows\System"
    
        Set-RegistryValue -KeyPath "HKLM:\SOFTWARE\Policies\Microsoft\MicrosoftEdge\PhishingFilter" -ValueName EnabledV9 -ValueType DWord -Value 1
        Set-RegistryValue -KeyPath "HKLM:\SOFTWARE\Policies\Microsoft\Internet Explorer\PhishingFilter" -ValueName EnabledV9 -ValueType DWord -Value 1
        Set-RegistryValue -KeyPath "HKLM:\SOFTWARE\Policies\Microsoft\Windows\System" -ValueName EnableSmartScreen -ValueType DWord -Value 1
        Set-RegistryValue -KeyPath "HKLM:\SOFTWARE\Policies\Microsoft\Windows\System" -ValueName ShellSmartScreenLevel -ValueType String -Value Warn

        Write-CMLogEntry -Value "Successfully enabled Smart Screen" -Severity 1
    }
}

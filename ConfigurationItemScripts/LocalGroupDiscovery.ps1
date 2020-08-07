$logfile = "C:\Windows\Temp\LocalGroupDiscovery.log"

Add-Content -path $logfile -value "`n`r******** Start: [$([DateTime]::Now)]. ********"

if ((Test-Connection -ComputerName "cm1.lab1.configmgrftw.com" -Quiet) -eq $true)
{

    $isCompliant = $true

    $localSystemName = $env:computername
    $localAdminGroupName = "Administrators"

    $validLocalAdminName = "vanpersie"
    $validLocalCSTName = "fifa"
    $validDomainAdmins = "Domain Admins"
    $validAdminSuffix = "_local"

    $validLocalAdminGroupLookup = @{
                                "OU=Computers,OU=Region1,DC=lab1,DC=configmgrftw,DC=com" = "rg1_local_admin";
                                "OU=Computers,OU=Region2,DC=lab1,DC=configmgrftw,DC=com" = "rg2_local_admin";
                                "OU=Workstations,OU=Region3,DC=lab1,DC=configmgrftw,DC=com" = "rg3_local_admins";
                                "OU=Workstations,OU=Region4,DC=lab1,DC=configmgrftw,DC=com" = "rg4_local_admins";
								"OU=Workstations,OU=LAB1,DC=lab1,DC=configmgrftw,DC=com" = "lab1_local_admins";
                                }

    $validLocalServiceGroupLookup = @{
                                "OU=Computers,OU=Region1,DC=lab1,DC=configmgrftw,DC=com" = "rg1_Local_Admin_Service_Accounts";
                                "OU=Computers,OU=Region2,DC=lab1,DC=configmgrftw,DC=com" = "rg2_Local_Admin_Service_Accounts";
                                "OU=Workstations,OU=Region3,DC=lab1,DC=configmgrftw,DC=com" = "rg3_Local_Admin_Service_Accounts";
                                "OU=Workstations,OU=Region4,DC=lab1,DC=configmgrftw,DC=com" = "rg4_Local_Admin_Service_Accounts";
								"OU=Workstations,OU=LAB1,DC=lab1,DC=configmgrftw,DC=com" = "lab1_Local_Admin_Service_Accounts";
                                }

    $systemDistinguishedName = (Get-ItemProperty -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Group Policy\State\Machine" -Name Distinguished-Name).'Distinguished-Name'

    $systemOU = $systemDistinguishedName.Substring($systemDistinguishedName.IndexOf(",") + 1)

    $expectedStandardMembers = 5

    $validLocalAdminGroupLookup.keys | ForEach-Object {
        if($systemOU.contains($_))
        {
            $validLocalAdminGroup = $validLocalAdminGroupLookup[$_]
        }
    }

    $validLocalServiceGroupLookup.keys | ForEach-Object {
        if($systemOU.contains($_))
        {
            $validLocalServiceGroup = $validLocalServiceGroupLookup[$_]
        }
    }

    Add-Content -path $logfile -value "System DN: $systemDistinguishedName"
    Add-Content -path $logfile -value "Expected local admin name: $validLocalAdminName"
    Add-Content -path $logfile -value "Expected local cs group name: $validLocalCSTName"
    Add-Content -path $logfile -value "Expected local admin group name: $validLocalAdminGroup"
    Add-Content -path $logfile -value "Expected local service group name: $validLocalServiceGroup"
    Add-Content -path $logfile -value "Examining local admin group `"$localAdminGroupName`" ..." 

    $validMemberCount = 0

    $group= [ADSI]"WinNT://$localSystemName/$localAdminGroupName,group"

    $members = @($group.psbase.Invoke("Members"))

    $memberNames = @()

    $members | ForEach-Object {
    
        $mbrName = $_.GetType().InvokeMember("Name", 'GetProperty', $null, $_, $null)
        $mbrPath = $_.GetType().InvokeMember("AdsPath", 'GetProperty', $null, $_, $null)

        $memberNames = $memberNames + $mbrName

        if($mbrName -eq $validLocalAdminName -or 
            $mbrName -eq $validDomainAdmins -or 
            $mbrName -eq $validLocalCSTName -or 
            $mbrName -eq $validLocalAdminGroup -or 
            $mbrName -eq $validLocalServiceGroup)

        {
            $isCompliant = $isCompliant -and $true
            $validMemberCount += 1

            Add-Content -path $logfile -value "`t-- Found valid core member ($validMemberCount): $mbrName"
        }
        elseif($mbrName.EndsWith($validAdminSuffix))
        {
            $isCompliant = $isCompliant -and $true
            Add-Content -path $logfile -value "`t-- Found valid member: $mbrName"
        }
        else
        {
            $isCompliant = $isCompliant -and $false
            Add-Content -path $logfile -value "`txx Found invalid member: $mbrName"
        }
    }

    if($isCompliant -eq $true -and $validMemberCount -eq $expectedStandardMembers)
    {
        Add-Content -path $logfile -value "+ System is Compliant"
        Write-Host "Compliant"
    }
    else
    {
        Add-Content -path $logfile -value "X System is Non-Compliant"
        Write-Host "Non-Compliant"
    }
}
else
{
    Add-Content -path $logfile -value "Not connected to internal network." 
}
Add-Content -path $logfile -value "******** Finish: [$([DateTime]::Now)]. ********`n`r`n`r" 

param (

 [parameter(Mandatory=$true)]
 [string]$Path,

[parameter(Mandatory=$true)]
 [string]$Value
)

try {

    if(Get-ItemProperty -Path $Path | Select-Object -ExpandProperty $Value -ErrorAction Stop | Out-Null)
    {
        return $true
    }
    else
    {
        return $false
    }

 }

catch {

return $false

}

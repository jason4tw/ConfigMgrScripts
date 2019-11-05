$clientSettings = Get-CMClientSetting

foreach($clientSettingPkg in $clientSettings)
{
    Write-Output $clientSettingPkg.Name
    
    foreach($agent in $clientSettingPkg.AgentConfigurations)
    {
        Write-Output " - $($agent.SmsProviderObjectPath)"

        foreach($setting in $agent.PropertyList.Keys)
        {
            Write-Output ("   - $setting : " + $agent.PropertyList["$setting"])
        }

        Write-Output ""
    }

    Write-Output ""
}
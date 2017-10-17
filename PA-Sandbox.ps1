# A group of functions for testing things against Palo Alto Panorama

Function Enable-TLS1.2()
{
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
}

Function Set-APIKey
{
    $script:apiKey = "<YOUR_PALO_API_KEY_HERE>"
}

Function Set-Domain
{
    $script:domain = "my.palo-instance.here"
}

Function Get-AllPANOSAdmins()
{
    $response = Invoke-RestMethod "https://$script:domain/api/?type=op&cmd=<show><admins><all></all></admins></show>&key=$script:apiKey" -Method GET
    return $response.response.result.admins.entry
}

Function Get-AllPANOSDevices()
{
    $response = Invoke-RestMethod "https://$script:domain/api/?type=op&cmd=<show><devices><all></all></devices></show>&key=$apiKey" -Method GET
    #return $response.response.result.devices.entry.name
    return $response
}

Function Get-DeviceNetMask
{
    $response = Invoke-RestMethod "https://$script:domain/api/?type=config&action=get&xpath=/config/shared/address/entry[@name='SomeSharedObjectNameHere']/ip-netmask&key=$apiKey" -Method GET
    return $response.response.InnerText
}

Function Get-AllDeadSharedDevices()
{
    $response = Invoke-RestMethod "https://$script:domain/api/?type=config&action=get&xpath=/config/shared/address&key=$apiKey"
    return $response
}

Function Create-TestSharedDevice()
{
    $response = Invoke-RestMethod "https://$script:domain/api/?key=$apiKey&type=config&action=set&xpath=/config/shared/address/entry[@name='API - Test Thingy 123']&element=<ip-netmask>0.0.0.0/23</ip-netmask>" -Method Post
    return $response
}

Enable-TLS1.2
Set-APIKey
Set-Domain

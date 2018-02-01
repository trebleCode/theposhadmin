Function Enable-TLS1.2()
{
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
}

Function Trust-AllCertsPolicy()
{
add-type @"
    using System.Net;
    using System.Security.Cryptography.X509Certificates;
    public class TrustAllCertsPolicy : ICertificatePolicy {
        public bool CheckValidationResult(
            ServicePoint srvPoint, X509Certificate certificate,
            WebRequest request, int certificateProblem) {
            return true;
        }
    }
"@
[System.Net.ServicePointManager]::CertificatePolicy = New-Object TrustAllCertsPolicy

} 

Function Set-APIKey{

    $script:apiKey = "MYAPIKEY"

}

Function Get-AllPANOSAdmins()
{
    $response = Invoke-RestMethod "https://panorama.bam.bamroot.net/api/?type=op&cmd=<show><admins><all></all></admins></show>&key=$script:apiKey" -Method GET
    return $response.response.result.admins.entry
}

Function Get-AllPANOSDevices()
{
    $response = Invoke-RestMethod "https://panorama.bam.bamroot.net/api/?type=op&cmd=<show><devices><all></all></devices></show>&key=$apiKey" -Method GET
    #return $response.response.result.devices.entry.name
    return $response
}

Function Get-DeviceNetMask($IP,$ObjectName)
{
    $response = Invoke-RestMethod "https://$IP/api/?type=config&action=get&xpath=/config/shared/address/entry[@name='$ObjectName']/ip-netmask&key=$apiKey" -Method GET
    return $response.response.InnerText
}

Function Get-AllSharedObjects($IP)
{
    $response = Invoke-RestMethod "https://$IP/api/?type=config&action=get&xpath=/config/shared/address&key=$apiKey"
    return $response
}

Function Create-TestSharedDevice($IP,$NewObjectName,$Netmask)
{
    $response = Invoke-RestMethod "https://$IP/api/?key=$apiKey&type=config&action=set&xpath=/config/shared/address/entry[@name='$NewObjectName']&element=<ip-netmask>$NetMask</ip-netmask>" -Method Post
    return $response
}

Function Get-DevDesc($IP,$ObjectName)
{
    $response = Invoke-RestMethod "https://$IP/api/?key=$apiKey&type=config&action=get&xpath=/config/shared/address/entry[@name='$ObjectName']/description"
    return $response
}

Function Get-LocalDeviceSecurityPolicies($IP)
{
    $response = Invoke-RestMethod "https://$IP/api/?key=$apiKey&type=config&action=get&xpath=/config/devices/entry[@name='localhost.localdomain']/vsys/entry[@name='vsys1']/rulebase/security"
    return $response.response.result.security.rules.entry
}

Function Generate-GlobalProtectCertificate($IP, $Username, $ExpiryDays, $Digest, $RSABits, $SignedBy)
{
     $certRequest = Invoke-RestMethod "https://$IP/api/?key=$apiKey&type=op&cmd=<request><certificate><generate><name>$CertificateName<signed-by>$SignedBy<algorithm><certificate-name>$Username<days-till-expiry>$ExpiryDays<digest>$Digest<RSA><rsa-nbits>$RSABits</rsa-nbits></RSA></digest><days-till-expiry><certificate-name></algorithm></name></signed-by></generate></certificate></request>" -Method Post
     return $certRequest
}

Function Get-AllAddressGroups($IP,$DeviceGroupName)
{
    $deviceGroups = @()

    foreach($group in $deviceGroups)
    {
        $memberGroups = Invoke-RestMethod "https://$IP//api/?type=config&action=get&key=$apiKey&xpath=/config/devices/entry[@name='localhost.localdomain']/device-group/entry[@name='$group']/address-group" -Method Get

        if($memberGroups.response.result -ne $NULL)
        {    
            foreach($member in $memberGroups)
            {
                Write-Host "here"
            }
        }
        else
        {
            Write-Output "Device group $group was empty"
        }
    }


    $addressGroups = Invoke-RestMethod "https://$IP/api/?type=config&action=get&key=$apiKey&xpath=/config/devices/entry[@name='localhost.localdomain']/device-group/entry[@name='$DeviceGroupName]/address-group" -Method Get
    return $addressGroups
}


Function Get-AllLocalAddresses($IP,$DeviceGroup)
{
    $deviceGroups = @()
    $localAddresses = Invoke-RestMethod "https://$IP/api/?type=config&action=get&key=$apiKey&xpath=/config/devices/entry[@name='localhost.localdomain']/vsys/entry[@name='vsys1']/address" -Method Get
    return $localAddresses.response.result.address.entry
}


Enable-TLS1.2
Trust-AllCertsPolicy
Set-APIKey
$data = Get-AllLocalAddresses -IP "1.2.3.4"

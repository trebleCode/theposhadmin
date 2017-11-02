$pmpAPIKey = "Your-PasswordManagerPro-API-Key-Here"
$pmpServer = "https://<PMP SERVER ADDRESS>:<PMP SERVER PORT>"

Function Get-OwnedResources()
{
    # Get the details of resources owned by the API account making the query

    $request =  Invoke-RestMethod "$pmpServer/restapi/json/v1/resources?AUTHTOKEN=$pmpAPIKey" -Method Get

    return $request.operation.details
}

Function Get-ResourceAccountList($ResourceID)
{
    # Get a list of accounts under a specified resource

    $request = Invoke-RestMethod "$pmpServer/restapi/json/v1/resources/$ResourceID/accounts?AUTHTOKEN=$pmpAPIKey" -Method Get
    return $request.operation.details.'ACCOUNT LIST'
}

Function Get-ResourceAccountDetails($ResourceID, $AccountID)
{
    # Get detailed information on a specific account within a resource

    $request = Invoke-RestMethod "$pmpServer/restapi/json/v1/resources/$ResourceID/accounts/$($AccountID)?AUTHTOKEN=$pmpAPIKey" -Method Get
    return $request.operation.details
}

Function Get-ResourceAccountID($ResourceName, $AccountName)
{
    # Retrieve a resource's account's ID

    $request = Invoke-RestMethod "$pmpServer/restapi/json/v1/resources/resourcename/$ResourceName/accounts/$($AccountName)?AUTHTOKEN=$pmpAPIKey" -Method Get
    return $request.operation.details.password
}


Function Get-ResourceAccountPassword($ResourceID, $AccountID)
{
    # Retrieve an account password

    $request = Invoke-RestMethod "$pmpServer/restapi/json/v1/resources/$ResourceID/accounts/$($AccountID)/password?AUTHTOKEN=$pmpAPIKey" -Method Get
    return $request.operation.details.password
}

Function Add-NewCertificateCheckType()
{
    try
    {
        $checkTypeExists = [ServerCertificateValidationCallback]
        
        # if type exists, continue, else load into session

        if($checkTypeExists -ne $NULL)
        {
            Write-Host "`nCustom type already loaded, continuing processing..." -ForegroundColor Green
            Out-Null
        }
    }
    catch [Exception]
    {
        if($Error[0].TargetObject.FullName -eq "ServerCertificateValidationCallback")
        {
            # Load new type into session
Add-Type @"
    using System;
    using System.Net;
    using System.Net.Security;
    using System.Security.Cryptography.X509Certificates;
    public class ServerCertificateValidationCallback
    {
        public static void Ignore()
        {
            ServicePointManager.ServerCertificateValidationCallback += 
                delegate
                (
                    Object obj, 
                    X509Certificate certificate, 
                    X509Chain chain, 
                    SslPolicyErrors errors
                )
                {
                    return true;
                };
        }
    }
"@
    Write-Host "`nNew custom type [ServerCertificateValidationCallback] has been loaded..." -ForegroundColor Green
        }
    }
}

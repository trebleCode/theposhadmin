Function ConnectTo-RemoteDomain
{
<#
    .SYNOPSIS
    Connects to a remote domain controller and establishes a new PS drive

    .EXAMPLE
    ConnectTo-RemoteDomain -Domain
#>
[cmdletBinding()]
param(
    [ValidateNotNullOrEmpty()]
    [String]$Domain
)

    BEGIN{}

    PROCESS
    {
        $userCred = $ENV:USERNAME

        do
        {
            $targetdomaindc = Get-ADDomainController -DomainName $Domain -Discover
            $targetdcname = $($targetdomaindc.hostname)
            Write-Host "`nEnter $Domain credentials to get available Domain Controllers" -ForegroundColor Yellow
            $DCs = Get-ADDomainController -Filter * -Server $targetdcname -Credential (Get-Credential $Domain\$userCred)
        }
        until ($DCs -ne $NULL)

        $i = 0

        do
        {
            $testConnection = Test-Connection -ComputerName "$($DCs.Name[$i] + "." + $Domain)" -Count 2
            $currentDC = $DCs.Name[$i]
            $i++
        }
        until($testConnection -ne $NULL)

        $checkDrives = Get-PSDrive
        if ($checkDrives.Name -notcontains $Domain)
        {
            Write-Host "Enter $Domain credentials to connect to an available Domain Controller" -ForegroundColor Yellow
            try
            {
                New-PSDrive -Name $Domain.split(".")[0] -PSProvider ActiveDirectory -Server $($currentDC + "." + $Domain) -Credential (Get-Credential $Domain\$userCred) -Root "//RootDSE/" -Scope Global
                $newDrive = $Domain.split(".")[0] + ":\"    
            }
            catch [Exception]
            {
                Write-Output "`nDrive with name $($Domain.split(".")[0]) already exists"
            }
            
        }
        
        Write-Output "`nChanging location to PS drive $($Domain.Split(".")[0]):\"
        cd $newDrive
        pwd
    }

    END{}
}

ConnectTo-RemoteDomain -Domain "my.domain.com"

<#
.SYNOPSIS
   Resets specified user password with an optional Unlock parameter
.DESCRIPTION
   Resets domain user password using a secure string and allows the
   executor the option to unlock the specified account after the password
   reset attempt is made. The Unlock value is True by default
.EXAMPLE
   .\Reset-UserPassword_v1_0.ps1 -Identity Calvin -NewPassword 4w4$0m3
.EXAMPLE
   .\Reset-UserPassword_v1_0.ps1 -Identity Hobbes -NewPassword 133tz0r -Unlock $True
.EXAMPLE
    .\Reset-UserPassword_v1_0.ps1 -Identity Rosalyn -NewPassword Pr3ttySw33t -Unlock $False
#>


[CmdletBinding()]
Param
(
    [Parameter(Mandatory=$True,
                ValueFromPipelineByPropertyName=$True,
                Position=0)]
    [string]$Identity,

    [Parameter(Mandatory=$True)]
    [string]$NewPassword,

    [Parameter(Mandatory=$False)]
    [bool]$Unlock = $True
)

Begin
{
    Import-Module ActiveDirectory

    Function Show-Banner()
    {
        Write-Host =
@'
###############################
#                             #
#    Password Reset Script    #
#                             #
###############################

'@
    }

    Function Reset-Password()
    {
        $SecPaswd = ConvertTo-SecureString �String $NewPassword �AsPlainText �Force
        try
        {
    }

    Function Unlock-User()
    {
        try
        {
            $Error[0].Exception
        }  
    }
}

Process
{
    Show-Banner
    Reset-Password
    if($Unlock -eq $true)
    {
        Unlock-User
    }
}
End
{
    Write-Host "End of script."
}
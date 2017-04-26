<#
.SYNOPSIS
    Takes string as an input and outputs
    as keystrokes
.EXAMPLE
    .\AutoKeyboard -AppName "InternetExplorer" -Keys @("one set of keys","12345","TAB","ENTER")
#>


[cmdletBinding()]
Param
(
    [Parameter(Mandatory=$true)]
    [ValidateNotNullOrEmpty()]
    [string]$AppName,

    [Parameter(Mandatory=$true)]
    [ValidateNotNullOrEmpty()]
    [Array]$Keys
)
BEGIN
{
    Function Start-WShell
    {
        $script:Shell = New-Object -COM “WScript.Shell”
    }

    Function Send-Keystrokes()
    {
        foreach($key in $Keys)
        {
            $Script:Shell.SendKeys($Key)
        }
    }

    Function Set-Focus($Name)
    {
        $script:Shell.AppActivate("$Name")
    }
}

PROCESS
{
    
    Start-WShell
    Set-Focus -Name "Internet Explorer"
    Send-Keystrokes
}

END{}

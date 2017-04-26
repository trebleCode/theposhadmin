Function Change-ToScriptDir()
{
    try
    {
        Set-Location $ENV:USERPROFILE\Desktop\PS
    }
    catch [exception]
    {
        Write-Host "There was an error: "
        Write-Host $Error[0].Exception
    }
}

Function Customize-Host
{
    if($Host.Name -eq "Windows PowerShell ISE Host")
    {
        $Host.UI.RawUI.WindowTitle = "PowerScott ISE"
    }
    else
    {
        $Host.UI.RawUI.WindowTitle = "PowerScott"
    }   
}

Function Close-AllIE()
{
    Get-Process iexplore | Foreach-Object { if($_.CloseMainWindow()){ $_.CloseMainWindow()} else {$_.CloseMainWindow()} } 
}

Function Start-IE()
{
    $IE = New-Object -ComObject InternetExplorer.Application
    $IE.Visible = $True
}


Function New-SWRandomPassword() 
{
    <#
    .Synopsis
       Generates one or more complex passwords designed to fulfill the requirements for Active Directory
    .DESCRIPTION
       Generates one or more complex passwords designed to fulfill the requirements for Active Directory
    .EXAMPLE
       New-SWRandomPassword
       C&3SX6Kn

       Will generate one password with a length between 8  and 12 chars.
    .EXAMPLE
       New-SWRandomPassword -MinPasswordLength 8 -MaxPasswordLength 12 -Count 4
       7d&5cnaB
       !Bh776T"Fw
       9"C"RxKcY
       %mtM7#9LQ9h

       Will generate four passwords, each with a length of between 8 and 12 chars.
    .EXAMPLE
       New-SWRandomPassword -InputStrings abc, ABC, 123 -PasswordLength 4
       3ABa

       Generates a password with a length of 4 containing atleast one char from each InputString
    .EXAMPLE
       New-SWRandomPassword -InputStrings abc, ABC, 123 -PasswordLength 4 -FirstChar abcdefghijkmnpqrstuvwxyzABCEFGHJKLMNPQRSTUVWXYZ
       3ABa

       Generates a password with a length of 4 containing atleast one char from each InputString that will start with a letter from 
       the string specified with the parameter FirstChar
    .OUTPUTS
       [String]
    .NOTES
       Written by Simon Wåhlin, blog.simonw.se
       I take no responsibility for any issues caused by this script.
    .FUNCTIONALITY
       Generates random passwords
    .LINK
       http://blog.simonw.se/powershell-generating-random-password-for-active-directory/
   
    #>
    [CmdletBinding(DefaultParameterSetName='FixedLength',ConfirmImpact='None')]
    [OutputType([String])]
    Param
    (
        # Specifies minimum password length
        [Parameter(Mandatory=$false,
                   ParameterSetName='RandomLength')]
        [ValidateScript({$_ -gt 0})]
        [Alias('Min')] 
        [int]$MinPasswordLength = 8,
        
        # Specifies maximum password length
        [Parameter(Mandatory=$false,
                   ParameterSetName='RandomLength')]
        [ValidateScript({
                if($_ -ge $MinPasswordLength){$true}
                else{Throw 'Max value cannot be lesser than min value.'}})]
        [Alias('Max')]
        [int]$MaxPasswordLength = 12,

        # Specifies a fixed password length
        [Parameter(Mandatory=$false,
                   ParameterSetName='FixedLength')]
        [ValidateRange(1,2147483647)]
        [int]$PasswordLength = 8,
        
        # Specifies an array of strings containing character groups from which the password will be generated.
        # At least one char from each group (string) will be used.
        [String[]]$InputStrings = @('abcdefghijkmnpqrstuvwxyz', 'ABCEFGHJKLMNPQRSTUVWXYZ', '23456789', '!"#%&'),

        # Specifies a string containing a character group from which the first character in the password will be generated.
        # Useful for systems which requires first char in password to be alphabetic.
        [String] $FirstChar,
        
        # Specifies number of passwords to generate.
        [ValidateRange(1,2147483647)]
        [int]$Count = 1
    )
    Begin {
        Function Get-Seed{
            # Generate a seed for randomization
            $RandomBytes = New-Object -TypeName 'System.Byte[]' 4
            $Random = New-Object -TypeName 'System.Security.Cryptography.RNGCryptoServiceProvider'
            $Random.GetBytes($RandomBytes)
            [BitConverter]::ToUInt32($RandomBytes, 0)
        }
    }
    Process {
        For($iteration = 1;$iteration -le $Count; $iteration++){
            $Password = @{}
            # Create char arrays containing groups of possible chars
            [char[][]]$CharGroups = $InputStrings

            # Create char array containing all chars
            $AllChars = $CharGroups | ForEach-Object {[Char[]]$_}

            # Set password length
            if($PSCmdlet.ParameterSetName -eq 'RandomLength')
            {
                if($MinPasswordLength -eq $MaxPasswordLength) {
                    # If password length is set, use set length
                    $PasswordLength = $MinPasswordLength
                }
                else {
                    # Otherwise randomize password length
                    $PasswordLength = ((Get-Seed) % ($MaxPasswordLength + 1 - $MinPasswordLength)) + $MinPasswordLength
                }
            }

            # If FirstChar is defined, randomize first char in password from that string.
            if($PSBoundParameters.ContainsKey('FirstChar')){
                $Password.Add(0,$FirstChar[((Get-Seed) % $FirstChar.Length)])
            }
            # Randomize one char from each group
            Foreach($Group in $CharGroups) {
                if($Password.Count -lt $PasswordLength) {
                    $Index = Get-Seed
                    While ($Password.ContainsKey($Index)){
                        $Index = Get-Seed                        
                    }
                    $Password.Add($Index,$Group[((Get-Seed) % $Group.Count)])
                }
            }

            # Fill out with chars from $AllChars
            for($i=$Password.Count;$i -lt $PasswordLength;$i++) {
                $Index = Get-Seed
                While ($Password.ContainsKey($Index)){
                    $Index = Get-Seed                        
                }
                $Password.Add($Index,$AllChars[((Get-Seed) % $AllChars.Count)])
            }
            Write-Output -InputObject $(-join ($Password.GetEnumerator() | Sort-Object -Property Name | Select-Object -ExpandProperty Value))
        }
    }
}


Function Select-Domain([string]$Name)
{

    $userCred = $ENV:USERNAME

    do
    {
        $targetdomaindc = Get-ADDomainController -DomainName $Name -Discover
        $targetdcname = $($targetdomaindc.hostname)
        Write-Host "`nEnter $Name credentials to get available Domain Controllers" -ForegroundColor Yellow
        $DCs = Get-ADDomainController -Filter * -Server $targetdcname -Credential (Get-Credential $Name\$userCred)
    }
    until ($DCs -ne $NULL)

    $i = 0

    do
    {
        $testConnection = Test-Connection -ComputerName $DCs.Name[$i] -Count 2
        $currentDC = $DCs.Name[$i]
        $i++
    }
    until($testConnection -ne $NULL)

    $checkDrives = Get-PSDrive
    if ($checkDrives.Name -notcontains $Name)
    {
        Write-Host "Enter $Name credentials to connect to an available Domain Controller" -ForegroundColor Yellow
        New-PSDrive -Name $Name -PSProvider ActiveDirectory -Server $currentDC -Credential (Get-Credential $Name\$userCred) -Root "//RootDSE/" -Scope Global
    }

    $newDrive = $Name + ":\"

    cd $newDrive
}

Function Start-NotepadPlusPlus()
{
    try
    {
        Start-Process notepad++.exe
    }
    catch [exception]
    {
        Write-Host "There was an error: " -NoNewline
        Write-Host $Error[0].Exception
    }
}

Function Empty-RecycleBin()
{

    Write-Host "Beginning cleanup routine..."

    $Shell = New-Object -ComObject Shell.Application
    $Recycler = $Shell.NameSpace(0xa)

    $totalItems = @($Recycler.Items()).Count
    $successCount = 0

    foreach($item in $Recycler.Items())
    {
        try
        {
            Remove-Item -Path $item.Path -Confirm:$false -Force -Recurse
            $successCount ++
        }
        catch [Exception]
        {
            Write-Host "Error attempting to remove " -NoNewline
            Write-Host $item -ForegroundColor Red
        }
    }

    if($successCount -eq $totalItems)
    {
        Write-Host "`nAll items successfully removed from Recycle Bin" -ForegroundColor Green
    }
}

Function Show-Banner
{
    
    $versionBanner = 
@'
        ____  _____      ______
       / __ \/ ___/_   _/ ____/
      / /_/ /\__ \ | / /___ \  
     / ____/___/ / |/ /___/ /  
    /_/    /____/|___/_____/
   
'@

    Write-Host $versionBanner -ForegroundColor Yellow
    Write-Host "         " $Host.Version
    Write-Host " "
}

Function Control-Apps([switch]$Start,[switch]$Stop)
{
    $apps = @("communicator.exe","outlook.exe","C:\Program Files (x86)\Microsoft\Remote Desktop Connection Manager\RDCMan.exe","powershell_ISE.exe", `
                "iexplore.exe","notepad++.exe","EXCEL.exe","explorer.exe")
    if($Start)
    {
        foreach($app in $apps)
        {
            if($app -eq "iexplore.exe")
            {
                $IE = New-Object -ComObject InternetExplorer.Application
                $IE.Visible = $True
            }
            else
            {
                try
                {
                    Start-Process $app
                }
                catch [Exception] 
                {
                    Write-Host "Error starting application: " -NoNewline
                    Write-Host $app -ForegroundColor Yellow
                }
            }
        }
    }
    elseif($Stop)
    {
        foreach($app in $apps)
        {
            Write-Output "Closing $app"
            if ($app -notmatch "RDCMan.exe")
            {
                taskkill /IM $app /F 
            }

            else
            {
                $app = $app.ToString().Replace("C:\Program Files (x86)\Microsoft\Remote Desktop Connection Manager\","")
                taskkill /IM $app /F 
            }

            if($app -eq "explorer.exe")
            {
                Write-Host "Restarting Windows Explorer"
                Start-Process $app
            }
        }
    }
}

Function Get-ADGroupDescription($Identity,[bool]$CopyToClipboard=$true)
{
    try
    {
        $description = $(Get-ADGroup -Identity $Identity -Properties Description | Select Description).Description
        
        if($CopyToClipboard -eq $true)
        {
            $description | clip
        }
        else
        {
            Write-Host "`nGroup: $Identity`nDescription: $description"
        }
    }
    catch [Exception]
    {
        "Could not find group '$Identity'"
    }
}

Function prompt {"PS "+[DateTime]::Now.ToString("hh:mm:ss")+">"}

Function RunAs-ElevatedLocalAdmin($Domain)
{
    Start-Process powershell.exe -Credential $(Get-Credential -UserName $Domain\$ENV:USERNAME) -NoNewWindow -ArgumentList “Start-Process powershell.exe -Verb runAs”
}

Customize-Host
Show-Banner

Set-Alias -Name cdps -Value Change-ToScriptDir
Set-Alias -Name xie -Value Close-AllIE
Set-Alias -Name newie -Value Start-IE
Set-Alias -Name genp -Value New-SWRandomPassword
Set-Alias -Name npp -Value Start-NotepadPlusPlus
Set-Alias -Name empty -Value Empty-RecycleBin
Set-Alias -Name apps -Value Control-Apps
Set-Alias -Name gdesc -Value Get-ADGroupDescription

cdps
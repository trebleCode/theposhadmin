## AMCP Tools

Function Add-WindowsForms()
{
    try
    {
        Add-Type -AssemblyName System.Windows.Forms
    }
    catch [Exception]
    {
        Write-Output "Failed to load Windows Forms assembly"
    }
}

Function Get-NowTime()
{
    $now = [DateTime]::Now.ToString("MM-dd-yyyy_hh-mm-ss")
    return $now
}

Function Get-CSVData()
{     
    Add-WindowsForms

    do {
        Write-Host "`nPlease select the CSV file to import in the dialog"
        Start-Sleep -Seconds 1
        $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{InitialDirectory = [Environment]::GetFolderPath('Desktop')}
    
        [void]$FileBrowser.ShowDialog()
     
        Write-Host "`nFile selected: " -NoNewline  
        Write-Host $FileBrowser.FileNames -ForegroundColor Yellow 
    
        $CSVFileName = $FileBrowser.FileName
        if ($CSVFileName.EndsWith(".csv"))
        {
            $choice = Read-Host "Are you sure this is the correct file? (y/n)"
            $choice = $choice.ToLower()
        }
        elseif($script:CSVFileName -eq "")
        {
            $quitChoice = Read-Host "A CSV file was not selected. Quit (y/n)?"
            if($quitChoice.ToLower() -eq "y")
            {
                Write-Host "Exiting function"
                Exit
            }
        }
        else
        {
            Write-Host "The file selected is not a CSV file."
            Write-Host "Restarting file selection loop."
        }
    }
    Until ($choice -eq "y")

    $CSV = Import-CSV $CSVFileName
    Write-Host "Returning object [0]File & [1]Path"
    return [PSCustomObject]@{File=$CSV;Path="$CSVFileName"}
}

Function Test-CSVHeaders([String]$CSVPath,[Array]$Headers)
{
    $pathTest = Test-Path $CSVPath

    if($pathTest -eq $True)
    {
        $matchedHeaderCount = 0
        $totalHeaderCount = $Headers.Count
        $CSV = Import-CSV -Path $CSVPath
        $CSVMembers = $CSV | Get-Member

        foreach($header in $Headers)
        {
            if($CSVMembers.Name -match $header)
            {
                $matchedHeaderCount++
            }
            else
            {
                Write-Output "No header $header found in file"
            }
        }
        if($matchedHeaderCount -eq $totalHeaderCount)
        {
            Write-Output "CSV header validation successful"
            return $True
        }
    }
    else
    {
        Write-Output "CSV header validation unsuccessful"
        return $False
    }
}

Function Test-UsersFromCSV($Path, $Header)
{
    $pathTest = Test-Path -Path $Path
    if($pathTest -eq $True)
    {
        $counter = 0
        [Array]$usersFound = $null
        [Array]$usersNotFound = $null
        $csv = Import-CSV -Path $Path    

        if($(Test-CSVHeaders -CSVPath $Path -Headers $Header) -eq $True)
        {
            foreach($user in $csv.$Header)
            {
                $test = Test-User -Identity $user
                if($test -ne $null)
                {
                    $usersFound += $test.$Header
                    $counter++
                }
                else
                {
                    $usersNotFound += $user
                }
            }
        }
        Write-Host "Returning object [0]users found & [1]users not found"
        return @($usersFound,$usersNotFound)
    }
    else
    {
        Write-Output "Path test to $Path failed"
    }
}

Function Test-User($Identity)
{
    try
    {
        $user = Get-ADUser -Identity $Identity
        if($user -ne $Null -and $user -ne "")
        {
            return $user
        }
        else
        {
            return $null
        }
    }
    catch [Exception]
    {
        return $null
    }
}

Function Compare-UserCounts($ReferenceCount, $DifferenceCount)
{
    if($ReferenceCount -eq $DifferenceCount)
    {
        Write-Output "`nAll users accounted for!"
    }
    else
    {
        Write-Output "`nCounts not equal.`nReferenceCount: $ReferenceCount, DifferenceCount: $DifferenceCount"
    }
}

Function Show-UsersNotFound($Set)
{
    Write-Output "`nThe following users were not found in AD"
    $Set | Format-List
}

Function Get-GroupMemberEmails($Identity)
{
    $results = $null
    $group = @(Get-ADGroup -Identity $Identity -Properties Members)

    if($group.Members.Count -ne 0)
    {
        foreach($member in $Group.Members)
        {
            if($(Get-ADObject $member).ObjectClass-eq "user")
            {
                $userInfo = Get-ADUser $member -Properties Name,sAMAccountName,EmailAddress
                [Array]$results += $userInfo
            }
        }
        $results | Format-Table -AutoSize
    }
    else
    {
        Write-Output "`nGroup $Identity has no members"
    }
}

Function Get-BulkGroupOwners($InputObject)
{
    $accountsNotFound = @()
    $accountsWithNoOwner = @()
    $accountsWithOwners = @()

    foreach($identity in $InputObject)
    {
        try
        {
            $personalPager = (Get-ADUser $identity -properties personalpager).personalpager
            
            Write-Host "Found account: " -NoNewline
            
            Write-Host $identity -ForegroundColor Green 
            try
            {
                $accountOwner = Get-ADUser -Filter {EmployeeID -eq $personalPager -and employeeType -eq "E"} -properties sAMAccountName,employeeID,employeetype,displayname,extensionattribute15,Enabled,Manager | select sAMAccountName,employeeID,employeetype,displayname,extensionattribute15,Enabled,Manager
                
                Write-Host "Found account owner: " -NoNewline
                Write-Host $accountOwner.sAMAccountName -ForegroundColor Green 
                
                [Array]$accountsWithOwner = [Array]$accountsWithOwner + [PSCustomObject]@{
                                                                    serviceAccount = $identity;
                                                                    PersonalPager = $personalPager;
                                                                    Owner = $accountOwner.sAMAccountName;
                                                                    displayname = $accountOwner.displayname;
                                                                    employeeID = $accountOwner.employeeID;
                                                                    employeetype = $accountOwner.employeeType;
                                                                    Code = "FoundWithOwner";
                                                                    Enabled = $accountOwner.Enabled;
                                                                    Manager = $accountOwner.Manager;
                                                                    }

                

            }
            catch [Exception]
            {
                Write-Host "Account owner not found for: " -NoNewline
                Write-Host $identity -ForegroundColor Red

                [Array]$accountsWithNoOwner =  [Array]$accountsWithNoOwner + [PSCustomObject]@{
                                                                                                serviceAccount = $identity;
                                                                                                PersonalPager = $personalPager;
                                                                                                Owner = $accountOwner.sAMAccountName;
                                                                                                displayname = $accountOwner.displayname;
                                                                                                employeeID = $accountOwner.employeeID;
                                                                                                employeetype = $accountOwner.employeeType;
                                                                                                Code = "FoundWithoutOwner";
                                                                                                Enabled = $accountOwner.Enabled;
                                                                                                Manager = $accountOwner.Manager;
                                                                                                }
            }
        }
        catch [Exception]
        {
            Write-Host "Account not found " -NoNewline
            Write-Host $identity -ForegroundColor Red

            [Array]$accountsNotFound = [Array]$accountsNotFound + [PSCustomObject]@{
                                                                                    serviceAccount = $identity;
                                                                                    PersonalPager = $personalPager;
                                                                                    Owner = $accountOwner.sAMAccountName;
                                                                                    displayname = $accountOwner.displayname;
                                                                                    employeeID = $accountOwner.employeeID;
                                                                                    employeetype = $accountOwner.employeeType;
                                                                                    Code = "NotFound";
                                                                                    Enabled = $accountOwner.Enabled;
                                                                                    Manager = $accountOwner.Manager;
                                                                                    }
        }
    }

    Write-Host "`nAdding results..."
    $combinedResults = $accountsWithOwner + $accountsWithNoOwner + $accountsNotFound 
    $combinedResults | Select * | Sort Code | Format-Table -AutoSize -Wrap

    do
    {
        try
        {
            $exportChoice = Read-Host "Export to CSV? (y/n)"
            if($exportChoice.ToLower() -ne "y" -and $exportChoice.ToLower() -ne "n")
            {
                Write-Output "Selection invalid"
                $exportChoice = $null
            }
        }
        catch [Exception]
        {
            $exportChoice = $null
        }
        
    }
    until($exportChoice.ToLower() -eq "y" -or $exportChoice.ToLower() -eq "n")

    if($exportChoice.ToLower() -eq "y")
    {
        $rightNow = Get-NowTime
        $combinedResults | Export-CSV -Path $ENV:USERPROFILE\Desktop\BulkAccountOwners_$rightNow.csv -NoTypeInformation
        Write-Output "`nResults exported to $ENV:USERPROFILE\Desktop\BulkAccountOwners_$rightNow.csv"
    }
    else
    {
        Out-Null
    }
}

Function Set-ServiceAccountPassword($Identity, $NewPassword, [bool]$PasswordNeverExpires, [bool]$UserMustChangeAtLogon)
{
    $results = @()
    
    try
    {
        Set-ADAccountPassword -Identity $Identity -Reset -NewPassword (ConvertTo-SecureString -String $NewPassword -AsPlainText -Force)
        Write-Host "Successfully changed password for acccount: " -NoNewline
        Write-Host $Identity -ForegroundColor Green
        $results += [PSCustomObject]@{SamAccountName = $Identity;Status = "SUCCESS"}

        if($PasswordNeverExpires -eq $True -and $UserMustChangeAtLogon -eq $True)
        {
            try
            {
                Set-ADUser -Identity $Identity -PasswordNeverExpires:$True -ChangePasswordAtLogon:$True
            }
            catch [Exception]
            {

            }
        }
        elseif($PasswordNeverExpires -eq $False -and $UserMustChangeAtLogon -eq $True)
        {
            try
            {
                Set-ADUser -Identity $Identity -PasswordNeverExpires:$False -ChangePasswordAtLogon:$True
            }
            catch [Exception]
            {

            }
        }
        elseif($PasswordNeverExpires -eq $True -and $UserMustChangeAtLogon -eq $False)
        {
            try
            {
                Set-ADUser -Identity $Identity -PasswordNeverExpires:$True -ChangePasswordAtLogon:$False
            }
            catch [Exception]
            {

            }
        }
        elseif($PasswordNeverExpires -eq $False -and $UserMustChangeAtLogon -eq $False)
        {
            Out-Null
        }
    }
    catch [Exception]
    {
        Write-Host "Failed to change password for acccount: " -NoNewline
        Write-Host $Identity -ForegroundColor Red
        $results += [PSCustomObject]@{SamAccountName = $Identity;Status = "FAIL"}
    }






    Write-Host "`nRESULTS"
    $results | Sort | Format-Table -AutoSize
}

Function New-TestAccount($Identity, $Description, $Password, [string]$Owner, [switch]$Core, [switch]$Reserved)
{
    # Check the account does not already exist in AD
    try
    {
        $userTest = Get-ADUser $Identity
        if($userTest -ne $NULL)
        {
            Write-Host "Account $Identity already exists!"
            Write-Host "Exiting"
            Continue
        }
    }
    catch [Exception]
    {
        # Create the new AD account if not found by above test
        try
        {
            New-ADUser $Identity -Path "OU=Secondary Accounts,OU=TIAA-CREF Users,DC=ad,DC=tiaa-cref,DC=org" -description $Description -Givenname $Identity -Surname $Identity -UserPrincipalName "$Identity@ad.tiaa-cref.org" -scriptPath "login.cmd"
            Write-Host "Successfully created account: $Identity"

            # Set the password for the new account
            try
            {
                Set-ADAccountPassword $Identity -reset -NewPassword (ConvertTo-SecureString -AsPlainText $Password -Force)
                Write-Host "Successfully set password for: $Identity" -ForegroundColor Green
            }
            catch [Exception]
            {
                Write-Host "Unable to set password for: $Identity" -ForegroundColor Red
            }

            # Set the employeeType attribute
            try
            {
                if($Core)
                {
                    Set-ADUser $Identity -replace @{employeetype="TC"}
                    Write-Host "Successfully set employeeType for: $Identity" -ForegroundColor Green
                }
                elseif($Reserved)
                {
                    Set-ADUser $Identity -replace @{employeetype="TR"}
                    Write-Host "Successfully set employeeType for: $Identity" -ForegroundColor Green
                }
            }
            catch [Exception]
            {
                Write-Host "Unable to set employeeType for: $Identity" -ForegroundColor Red
            }

            # Set personalPager attribute
            try
            {
                Set-ADUser $Identity -replace @{personalpager=$Owner}
                Write-Host "Successfully set personalPager for: $Identity" -ForegroundColor Green
            }
            catch [Exception]
            {
                Write-Host "Unable to set personalPager for: $Identity" -ForegroundColor Red
            }

            # Set ChangePasswordAtLogon
            try
            {
                Set-ADUser $Identity -ChangePasswordAtLogon $True
                Write-Host "Successfully set ChangePasswordAtLogon for: $Identity" -ForegroundColor Green
            }
            catch [Exception]
            {
                Write-Host "Unable to set ChangePasswordAtLogon for: $Identity" -ForegroundColor Red
            }

            # Enable the account
            try
            {
                Enable-ADAccount -Identity $Identity
                Write-Host "Successfully enabledUser: $Identity" -ForegroundColor Green
            }
            catch [Exception]
            {
                Write-Host "Unable to enable user: $Identity" -ForegroundColor Red
            }
        }
        catch [Exception]
        {
            Write-Host "Unable to create new AD user: $Identity" -ForegroundColor Red
            Write-Host "Exiting"
            Exit
        }
    }
}

Function Confirm-CSVExport($InputObject,$Filename)
{
    do
    {
        try
        {
            $exportChoice = Read-Host "Export to CSV? (y/n)"
            if($exportChoice.ToLower() -ne "y" -and $exportChoice.ToLower() -ne "n")
            {
                Write-Output "Selection invalid"
                $exportChoice = $null
            }
        }
        catch [Exception]
        {
            $exportChoice = $null
        }
        
    }
    until($exportChoice.ToLower() -eq "y" -or $exportChoice.ToLower() -eq "n")

    if($exportChoice.ToLower() -eq "y")
    {
        $rightNow = Get-NowTime
        $InputObject | Export-CSV -Path $ENV:USERPROFILE\Desktop\$FileName-$rightNow.csv -NoTypeInformation
        Write-Output "`nResults exported to $ENV:USERPROFILE\Desktop\$FileName-$rightNow.csv"
    }
    else
    {
        Out-Null
    }
}

Function Get-CommonManagerOfGroupMembers([string]$Group)
{
    $managers = @()
    $results = @()

    try
    {
        $groupmembers = @(Get-ADGroupMember -Identity $Group)
        $groupFound = $True
    }
    catch [Exception]
    {
        Write-Host "Group " -NoNewline
        Write-Host $Group -NoNewline -ForegroundColor Yellow
        Write-Host "Not found"
        $groupFound = $False
    }
    
    if($groupmembers.Count -gt 0)
    {
        foreach($member in $groupmembers)
        {
            $manager = (Get-ADUser $member -Properties Manager).Manager

            if($manager -ne $Null)
            {
                $managers += $manager
            }
        }
    }
    else
    {
        Write-Host "Group " -NoNewline
        Write-Host $Group -NoNewline -ForegroundColor Yellow
        Write-Host " has no members"
    }
    
    $groupedManagers = $managers | Group-Object | Sort Count -Descending

    # return the top 3 most common managers

    if($groupedManagers.Count -gt 3 -or $groupedManagers.Count -eq 3)
    {
        foreach($item in $groupedManagers[0..2])
        {
            $managerData = Get-ADUser $item.Name -Properties Name,SamAccountName,EmailAddress,GivenName,Surname
            $result = [PSCustomObject]@{Name = $($managerData.GivenName + " " + $managerData.Surname);
                                        sAMAccountName = $managerData.sAMAccountName;
                                        EmailAddress = $managerData.EmailAddress;
                                        Count = $item.Count
                                        }
            $results += $result
        }
        
    }
    elseif($groupedManagers.Count -eq 2)
    {
        foreach($item in $groupedManagers[0..1])
        {
            $managerData = Get-ADUser $item.Name -Properties Name,SamAccountName,EmailAddress,GivenName,Surname
            $result = [PSCustomObject]@{Name = $($managerData.GivenName + " " + $managerData.Surname);
                                        sAMAccountName = $managerData.sAMAccountName;
                                        EmailAddress = $managerData.EmailAddress;
                                        Count = $item.Count
                                        }
            $results += $result
        }
    }
    elseif($groupedManagers.Count -eq 1)
    {
        $managerData = Get-ADUser $groupedManagers.Name -Properties Name,SamAccountName,EmailAddress,GivenName,Surname
        $result = [PSCustomObject]@{Name = $($managerData.GivenName + " " + $managerData.Surname);
                                    sAMAccoutName = $managerData.sAMAccountName;
                                    EmailAddress = $managerData.EmailAddress;
                                    Count = $item.Count
                                    }
    }

    $results | Format-Table -AutoSize -Wrap

    return $results
}

<#
.SYNOPSIS
Analyzes a set of home directories
and returns the most common one in use
if available

.DESCRIPTION
    - User enters the username of an AD account.
    - Script looks up the user's manager and identifies
      all direct reports of that manager.
    - Home directories of all direct reports are compared
      against the initially specified user
    - Results are printed to the console.

.EXAMPLE
.\Analyze-HomeDrives.ps1

Enter a username to search: a123456

Searching for user a123456

Getting manager's direct reports

Partial (25%) user match
Most Common Home Directory: \\SomeServer\SomeShare
#>

[cmdletBinding()]
Param()

BEGIN
{
    Function Set-Dependencies()
    {
        Import-Module ActiveDirectory
    }

    Function Show-Banner()
    {
@'
----------------------------------
  Home Directory Analyzer Script
----------------------------------
'@
    }

    Function Get-InitialUser()
    {
        $userInput = Read-Host "`nEnter a username to search"

        try
        {
            Write-Output "`nSearching for user $userInput"
            $userData = Get-ADUser "$userInput" -Properties Manager,Name,SamAccountName,City,State,StreetAddress,HomeDirectory
            $userManagerPath = $userData.Manager
            $userManagerName = $userData.Manager.Replace("CN=","").Split(",")[0]

            if($userData.HomeDirectory -eq $NULL)
            {
                Write-Output "User $userInput has no HomeDirectory"
                $noHomeDirectory = $True
            }
            elseif($userData.HomeDirectory -ne $NULL)
            {
                $noHomeDirectory = $false
            }
        }
        catch [Exception]
        {
            Write-Output "User $userInput not found"
            break
        }

        if($noHomeDirectory -ne $True)
        {
            try
            {
                Write-Output "`nGetting manager's direct reports"
                $usersToCompare = Get-ManagerDirectReports -Name $userManagerPath
            }
            catch [Exception]
            {
                $Error[0].Exception.Message
            }
            try
            {
                Compare-Users -ReferenceUser $userData -DifferenceUsers $usersToCompare
            }
            catch [Exception]
            {
                $Error[0].Exception.Message
            }
        }

        Run-Again
    }

    Function Get-ManagerDirectReports($Name)
    {
        $directReports = Get-ADUser -Filter {Manager -eq $Name} -Properties Name,SamAccountName,City,State,StreetAddress,HomeDirectory | Select Name,SamAccountName,City,State,StreetAddress,HomeDirectory
        return $directReports
    }

    Function Compare-Users($ReferenceUser, [Array]$DifferenceUsers)
    {
        $perfectMatches = 0

        $totalCityStateMatch = 0
        $totalStateMatch = 0

        $totalMatches = 0
        $totalDoesNotMatch = 0

        Write-Output "`nSearching direct report information"

        foreach($user in $DifferenceUsers)
        {
            # Skip the username entered by the user
            if($user.SamAccountName -eq $ReferenceUser.SamAccountName)
            {
                Out-Null
            }
            else
            {
                # City, State, and StreetAddress match
                if(($user.City -eq $ReferenceUser.City) -and 
                    ($user.State -eq $ReferenceUser.State) -and
                    ($user.StreetAddress -eq $ReferenceUser.StreetAddress))
                {
                    $perfectMatches++
                    $totalMatches++
                }
                # City and State match
                if(($user.City -eq $ReferenceUser.City) -and 
                    ($user.State -eq $ReferenceUser.State) -and
                    ($user.StreetAddress -ne $ReferenceUser.StreetAddress))
                {
                    $totalCityStateMatch++
                    $totalMatches++
                }
                # State matches
                if(($user.City -ne $ReferenceUser.City) -and 
                    ($user.State -eq $ReferenceUser.State))
                {
                    $totalStateMatch++
                    $totalMatches++
                }
                # Not matched
                if($user.State -ne $ReferenceUser.State)
                {
                    $totalDoesNotMatch++
                }
            }
        }

        # 1:1 Perfect Match
        if($perfectMatches -eq $DifferenceUsers.Count)
        {
            Write-Host "`n100% match" -ForegroundColor Green
            Write-Host "Home Directory: $($ReferenceUser.HomeDirectory.Trim($ReferenceUser.SamAccountName))"
        }

        #  More matches than unmatched
        if(($totalMatches -lt $DifferenceUsers.Count) -and ($totalMatches -gt $totalDoesNotMatch))
        {
            Write-Host "`nPartial ($((($totalMatches/$DifferenceUsers.Count)*100).ToString(00.00))%) user match" -ForegroundColor Yellow
            Write-Host "Most Common Home Directory: $($ReferenceUser.HomeDirectory.Trim($ReferenceUser.SamAccountName))"
        }

        # More unmatched than matches
        if(($totalMatches -lt $DifferenceUsers.Count) -and ($totalMatches -lt $totalDoesNotMatch))
        {
            # If (City+State > State only) AND (City+State > Total unmatched)

            if(($totalCityStateMatch -gt $totalStateMatch) -and $totalCityStateMatch -gt $totalDoesNotMatch)
            {
                $mostMatchedHD = $DifferenceUsers | ? {($_.City -eq $ReferenceUser.City -and $_.State -eq $ReferenceUser.State) -and $_.HomeDirectory -ne $NULL} | Select HomeDirectory,SamAccountName -First 1
                Write-Host "`nPartial ($((($totalMatches/$DifferenceUsers.Count)*100).ToString(00.00))%) user match" -ForegroundColor Yellow
                Write-Host "Most Common Home Directory: $($mostMatchedHD.HomeDirectory.Trim($mostMatchedHD.SamAccountName))"
            }
            # If (City+State > State only) AND (City+State < Total unmatched)

            if(($totalCityStateMatch -gt $totalStateMatch) -and $totalCityStateMatch -lt $totalDoesNotMatch)
            {
                # Get the most common state and group together
                $mostCommonState = $($DifferenceUsers.State | Group | Sort Count -Descending | Select Name -First 1).Name
                $mostMatchedHD = $DifferenceUsers | ? {($_.State -eq $mostCommonState) -and $_.HomeDirectory -ne $NULL} | Select HomeDirectory,SamAccountName -First 1
                
                Write-Host "`nPartial ($((($totalMatches/$DifferenceUsers.Count)*100).ToString(00.00))%) user match" -ForegroundColor Yellow
                Write-Host "Most Common Home Directory: $($mostMatchedHD.HomeDirectory.Trim($mostMatchedHD.SamAccountName))"
            }

            # If (State only > City+State) AND (State only > Total unmatched)

          if(($totalCityStateMatch -lt $totalStateMatch) -and $totalStateMatch -gt $totalDoesNotMatch)
            {
                $mostMatchedHD = $DifferenceUsers | ? {($_.City -ne $ReferenceUser.City -and $_.State -eq $ReferenceUser.State) -and $_.HomeDirectory -ne $NULL} | Select HomeDirectory,SamAccountName -First 1
                Write-Host "`nPartial ($((($totalMatches/$DifferenceUsers.Count)*100).ToString(00.00))%) user match" -ForegroundColor Yellow
                Write-Host "Most Common Home Directory: $($mostMatchedHD.HomeDirectory.Trim($mostMatchedHD.SamAccountName))"
            }
            
            # If (State only > City+State) AND (State only < Total unmatched)

            if(($totalCityStateMatch -lt $totalStateMatch) -and $totalStateMatch -lt $totalDoesNotMatch)
            {
                # Get the most common state and group together
                $mostCommonState = $($DifferenceUsers.State | Group | Sort Count -Descending | Select Name -First 1).Name
                $mostMatchedHD = $DifferenceUsers | ? {($_.State -eq $mostCommonState) -and $_.HomeDirectory -ne $NULL} | Select HomeDirectory,SamAccountName -First 1
                
                Write-Host "`nPartial ($((($totalMatches/$DifferenceUsers.Count)*100).ToString(00.00))%) user match" -ForegroundColor Yellow
                Write-Host "Most Common Home Directory: $($mostMatchedHD.HomeDirectory.Trim($mostMatchedHD.SamAccountName))"
            }

            # If perfect matches are the only match but less than total unmatched

            if($perfectMatches -gt 0 -and $totalCityStateMatch -eq 0 -and $totalStateMatch -eq 0 -and $perfectMatches -lt $totalDoesNotMatch)
            {
                # Get the most common state and group together
                $mostCommonState = $($DifferenceUsers.State | Group | Sort Count -Descending | Select Name -First 1).Name
                $mostMatchedHD = $DifferenceUsers | ? {($_.State -eq $mostCommonState) -and $_.HomeDirectory -ne $NULL} | Select HomeDirectory,SamAccountName -First 1
                
                Write-Host "`nPartial ($((($totalMatches/$DifferenceUsers.Count)*100).ToString(00.00))%) user match" -ForegroundColor Yellow
                Write-Host "Most Common Home Directory: $($mostMatchedHD.HomeDirectory.Trim($mostMatchedHD.SamAccountName))"
            }
        }

        # No matches

        if($totalMatches -eq 0)
        {
            Write-Host "`nNo matches found" -ForegroundColor Red
        }
    }

    Function Run-Again()
    {
        do
        {
            $choice = Read-Host "`nRun again? (y/n)"

            if($choice.ToLower() -ne "y" -and $choice.ToLower() -ne "n")
            {
                Write-Output "Invalid choice"
            }
        } until ($choice.ToLower() -eq "y" -or $choice.ToLower() -eq "n")

        if($choice.ToLower() -eq "y")
        {
            Get-InitialUser
        }
        else
        {
            Write-Output "End of script"
        }
    }
}

PROCESS
{
    Set-Dependencies
    Show-Banner
    Get-InitialUser
}

END{}
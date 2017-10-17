BEGIN
{
    Function Set-Dependencies()
    {
        try
        {
            Import-Module ActiveDirectory
        }
        catch [Exception]
        {
            Write-Warning "AD Module for Powershell not installed.`nEnding script."
            Exit
        }
    }

    Function Set-ScriptArrays()
    {
        $script:secGroups = New-Object System.Collections.ArrayList
        $script:distGroups = New-Object System.Collections.ArrayList
        $script:combinedGroups = New-Object System.Collections.ArrayList
    }

    Function Get-Groups($Type)
    {
       $groupQuery = Get-ADGroup -Filter {groupCategory -eq $Type} -Properties Name,groupCategory,DistinguishedName,extensionAttribute9,extensionAttribute10 | Select Name,groupCategory,DistinguishedName,extensionAttribute9,extensionAttribute10
       if($groupQuery -ne $NULL)
       {
           
           return $groupQuery
       }
       else
       {
           "`nQuery for $Type was NULL"
       }
    }

    Function Add-ToResults($InputObject)
    {
        foreach($group in $InputObject)
        {
            if($group.groupCategory -eq "Security")
            {
                $script:secGroups = $script:secGroups + [PSCustomObject]@{Name=$group.Name;
                                                                            Att9=$group.extensionAttribute9;
                                                                            Att10=$group.extensionAttribute10;
                                                                            DN=$group.DistinguishedName;
                                                                            Category=$group.groupCategory
                                                                         }
            }
            if($group.groupCategory -eq "Distribution")
            {
                $script:distGroups = $script:secGroups + [PSCustomObject]@{Name=$group.Name;
                                                                            Att9=$group.extensionAttribute9;
                                                                            Att10=$group.extensionAttribute10;
                                                                            DN=$group.DistinguishedName;
                                                                            Category=$group.groupCategory
                                                                           }
            }
        }
    }

    Function Export-Results($InputObject)
    {
        try
        {
            $InputObject | Export-CSV -Path $ENV:USERPROFILE\Desktop\$([DateTime]::Now.ToString("hh-mm-ss_MM-dd-yyyy_") +"GroupInfo.csv") -NoTypeInformation
            Write-Host "`n[SUCCESS] File successfuly exported to $ENV:USERPROFILE\Desktop"
        }
        catch [Exception]
        {
            Write-Host "`nERROR: Export failed" -ForegroundColor Red
            Write-Host "ErrorMessage: $($Error[0].Exception)"
            Write-Host "InnerException: $($Error[0].Exception.InnerException)"
        }
    }
}

PROCESS
{
    Set-Dependencies
    Set-ScriptArrays

    Write-Host "`nGetting security groups..."
    $secQuery = Get-Groups -Type "Security"

    Write-Host "Getting distribution groups..."
    $distQuery = Get-Groups -Type "Distribution"

    Write-Host "`nAdding security groups to results..."
    Add-ToResults -InputObject $secQuery

    Write-Host "Adding distribution groups to results..."
    Add-ToResults -InputObject $distQuery

    Write-Host "`nCombining results..."
    $script:combinedGroups = $script:secGroups + $script:distGroups

    Export-Results -InputObject $script:combinedGroups
}

END
{

}

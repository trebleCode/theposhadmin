BEGIN

{
    Function Set-ScriptVars
    {
        $script:managerArray = @()
        $script:LevelCount = 1
    }

    function Get-TopLevelReports([string]$Identity, $Level)
    {
        $script:Level = $Level

        TRY
	    {
            $directReports = Get-Aduser -identity $Identity -Properties directreports,Name
            Write-Host "Getting direct reports for:" $directReports.Name
            foreach($directReport in $directReports.DirectReports)
            {
                $userData = Get-ADUser -Identity $directReport -Properties Name, mail, manager, telephoneNumber, samAccountName, DirectReports | Select-Object -Property Name, telephoneNumber, Mail, SamAccountName, Manager, DirectReports
                $userData2 = [PSCustomObject] @{Level=$Script:LevelCount;Name=$userData.Name;Mail=$userData.Mail;Phone=$userData.telephoneNumber;Username=$userData.samaccountname;City=$userData.City;State=$userData.State;Manager=$userData.Manager.TrimStart("^CN=").Split(",")[0]} 
                $script:managerArray += $userData2
            }
            $script:LevelCount++
        }
        CATCH
	    {
		    Write-Verbose -Message "Error encountered"
		    Write-Verbose -Message $Error[0].Exception.Message
        }
    }

    Function Get-DirectReports($Identity)
    {
        $subdirectReports = Get-Aduser -identity $Identity -Properties directreports | select directreports
        if($subdirectReports.directReports.Count -ne 0)
        {
    
            $totalDR = $subdirectReports.DirectReports.Count
            $j = 1
            $percentComplete = (($j / $totalDR) * 100).ToString("00.00")
            foreach($subdirectReport in $subdirectReports.DirectReports)
            {
                
                Write-Progress -Activity "Gathering data" -Status "$j of $totalDR Complete" -PercentComplete $percentComplete
                $subuserData = Get-ADUser -Identity $subdirectReport -Properties Name, mail, manager, telephoneNumber, samAccountName, City, State, DirectReports | Select-Object -Property Name, telephoneNumber, Mail, SamAccountName, City, State, Manager, DirectReports
                $subuserData2 = [PSCustomObject] @{Level=$Script:LevelCount;Name=$subuserData.Name;Mail=$subuserData.Mail;Phone=$subuserData.telephoneNumber;UserName=$subuserData.samaccountname;City=$subuserData.City;State=$subuserData.State;Manager=$subuserData.Manager.TrimStart("^CN=").Split(",")[0]} 
                $script:managerArray += $subuserData2
                $j++
                $percentComplete = (($j / $totalDR) * 100).ToString("00.00")
            }
        }
    }    

    Function Export-Data
    {
        $fileNamePrefix = "DirectReportData-"
        $dateTime = Get-Date -Format _MM-dd-yyyy_hh-mm-ss
        $script:managerArray | Export-CSV -Path $env:USERPROFILE\Desktop\$fileNamePrefix$script:TopLevelManager$dateTime.csv -NoTypeInformation
    }
}

PROCESS
{
    Set-ScriptVars
    $script:TopLevelManager = Read-Host "Enter top level manager samaccountname"
    Get-TopLevelReports -Identity $script:TopLevelManager -Level 3
    $i = 1
    
    do
    {
        Write-Host "Tier Count: $i"
        $currentSet = $script:managerArray | ? {$_.Level -eq $i}

        foreach($item in $currentSet)
        {
            Write-Host "`tGetting Direct Reports for:" $item.Name
            Get-DirectReports -Identity $item.username
        }
        $i++
        $Script:LevelCount++
    }
    until ($script:LevelCount -eq $Script:Level + 1)

    Export-Data
}

END
{
}
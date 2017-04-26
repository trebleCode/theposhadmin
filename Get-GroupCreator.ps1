BEGIN
{
    Function Import-Dependencies
    {
        Write-Output "`nAttempting dependency imports"

        try
        {
            Import-Module ActiveDirectory
        }
        catch [exception] 
        {
            $Error[0]
            Write-Warning "ActiveDirectory Module for Powershell not installed."
            Write-Output "Exiting script."
            Exit
        }

        try
        {
            Add-Type -AssemblyName System.Windows.Forms
        }
        catch [exception]
        {
            $Error[0]
            Write-Warning "Windows Forms assembly could not be loaded."
            Write-Output "Exiting script."
            Exit
        }

        Write-Output "Imports successful"
    }
    Function Show-Banner
    {
        Write-Host "#=========================================#" -ForegroundColor Yellow
        Write-Host "#    Group Owner Identification Script    #" -ForegroundColor Yellow
        Write-Host "#          " -ForegroundColor Yellow -NoNewline
        Write-Host "    Version 1.0" -ForegroundColor White -NoNewline
        Write-Host "                #" -ForegroundColor Yellow
        Write-Host "#=========================================#" -ForegroundColor Yellow
        Write-Host " "

        Write-Host "NOTE: " -ForegroundColor Red -NoNewline
        Write-Host "This script requires a CSV file with"
        Write-Host "      a header entitled " -NoNewline 
        Write-Host "samaccountname" -ForegroundColor Yellow
        Write-Host " "
    }

    Function Clear-ScriptVars
    {
        $script:allResults = $NULL
        $script:CSVFileName
        $script:CSV = $NULL
    }

    Function Get-CSVData()
    {
        
        do {
            Write-Host "`nPlease select the CSV file to import in the dialog"
            Start-Sleep -Seconds 1
            $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{InitialDirectory = [Environment]::GetFolderPath('Desktop')}
    
            [void]$FileBrowser.ShowDialog()
     
            Write-Host "`nFile selected: " -NoNewline  
            Write-Host $FileBrowser.FileNames -ForegroundColor Yellow 
    
            $script:CSVFileName = $FileBrowser.FileName
            if ($script:CSVFileName.EndsWith(".csv"))
            {
                $choice = Read-Host "Are you sure this is the correct file? (y/n)"
                $choice = $choice.ToLower()
            }
            elseif($script:CSVFileName -eq "")
            {
                $quitChoice = Read-Host "A CSV file was not selected. Quit (y/n)?"
                if($quitChoice.ToLower() -eq "y")
                {
                    Write-Host "Exiting script"
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

        $script:CSV = Import-CSV $script:CSVFileName
    }

    Function Validate-CSVData()
    {
        Write-Host "`nValidating CSV data..."

        # Validate that the CSV data exists.

        if ($Script:csv -eq $NULL) 
        {
            Write-Output "`nCSV file either has no samaccountname header or contains no data. Exiting script"
            Exit   
        }

        # Validate that the proper columns are not empty and that the header names are correct.

        if ($script:CSV.samaccountname -ne $NULL)
        {
            Write-Output "`nCSV has been validated with samaccountname field."
        }
    }

    Function Get-GroupOwners
    {
        $foundGroups = @()
        $notFoundGroups = @()
        $foundGroupswithNoOwner = @()

        $i = 0
        $totalTargets = @($script:CSV.samaccountname).Count

        $groupNames = $script:CSV.samaccountname

        foreach ($group in $groupNames)
        {
            $i++
            $percentComplete = (($i/$totalTargets) * 100)
            $percentCompleteRound = (($i/$totalTargets) * 100).ToString("0.00")

            Write-Progress -Activity "Searching for group $group" -Status "$i of $totalTargets ($percentCompleteRound complete)" -PercentComplete $percentComplete

            $searchTarget =  [ADSI](([ADSISearcher]"(name=$group)").FindOne().Path)
            $searchTargetOwner = $searchTarget.PSBase.ObjectSecurity.Owner

            if ($searchTarget -ne $NULL -and $searchTargetOwner -ne $NULL)
            {
                # get the AD user properties to grab the displayname

                try
                {
                    $targetDisplayName = Get-ADUser $searchTargetOwner.ToString().Trim("AD-ENT\") -Properties * | Select DisplayName
                    $targetDisplayName = $targetDisplayName.DisplayName

                    Write-Output "`tSuccessfully found information for $group. Adding to results"

                    $foundResult = [PSCustomObject]@{Group=$group;Owner=$searchTargetOwner;DisplayName=$targetDisplayName}
                    [Array]$foundGroups = $foundGroups + $foundResult
                }
                catch [exception]
                {
                    if ($targetDisplayName -eq $NULL)
                    {
                        Write-Output "`tError attempting to find displayname for for owner of: $group"

                        # add to results even if displayname returns null

                        $foundResult = [PSCustomObject] @{Group=$group;Owner=$searchTargetOwner;DisplayName=$NULL}
                        [Array]$foundGroups = $foundGroups + $foundResult
                    }
                    
                    $Error[0] 
                }
            }

            elseif ($searchTarget -ne $NULL -and $searchTargetOwner -eq $NULL)
            {
                Write-Output "`tGroup target found, but owner is NULL"
                $foundNoOwnerResult = [PSCustomObject] @{Group=$group;Owner=$NULL;DisplayName=$NULL}
                [Array]$foundGroupsWithNoOwner = $foundGroupsWithNoOwner + $foundNoOwnerResult
            }

            elseif ($searchTarget -eq $NULL -and $searchTargetOwner -eq $NULL)
            {
                Write-Output "`tGroup target not found"
                $notFoundResult = [PSCustomObject] @{Group=$group;Owner=$NULL;DisplayName=$NULL}
                [Array]$notFoundGroups = $notFoundGroups + $notFoundResult
            }
        }

        $script:allResults = @()
        $script:allResults += $foundGroups
        $script:allResults += $foundGroupswithNoOwner
        $script:allResults += $notFoundGroups
    }

    Function Export-Results
    {
        # Create unique file name

        $exportPath = "$ENV:USERPROFILE\Desktop"
        $namePrefix = "GroupOwnerSearch_"
        $dateTime = Get-Date -Format MM-dd-yyyy_hh-mm-ss
        $fileName = $namePrefix + $dateTime + ".csv"
        $fullFileName = "$ENV:USERPROFILE\Desktop\$fileName"


        try
        {
            $script:allResults | Export-CSV -Path $fullFileName -NoTypeInformation
            Write-Output "`nResults successfully exported to $fullFileName"
        }
        catch [exception]
        {
            Write-Output "`nThere was an error attempting to output the results"
            $Error[0]
        }
    }

    Function Prompt-ReRun
    {
        do
        {
            $openChoice = Read-Host "`nRun again? (y/n)"
            $openChoice = $openChoice.ToLower()
        } until($openChoice -eq "y" -or $openChoice -eq "n")
        
        if($openChoice -ne "y" -and $openChoice -ne "n")
        {
            Write-Host "Invalid entry"
        }
        elseif($openChoice -eq "y")
        {
            Run-Stack
        }
        else
        {
            Out-Null
        }
    }

    Function Run-Stack
    {
        Clear-ScriptVars
        Get-CSVData
        Validate-CSVData
        Get-GroupOwners
        Export-Results
        Prompt-ReRun
    }
}

PROCESS
{
    Show-Banner
    Import-Dependencies
    Run-Stack
}

END
{
    Write-Host "`nEnd of script!" -ForegroundColor Yellow
}
BEGIN
{
    $ErrorActionPreference = "Stop"
    Add-Type -AssemblyName System.Windows.Forms
    
    Function Get-CSVData()
    {
        Write-Host "Initializing script-level variables..."
        
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
                $quitChoice = Read-Host "No file was selected. Quit (y/n)?"
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

        Write-Host "Creating log file..."

        [string]$logFileBaseName = "./GroupMembershipFinderLog"
        $logDate = Get-Date -Format MM-dd-yyyy_hh-mm-ss
        [string]$script:logFileName = $logFileBaseName + "_$logDate" + ".txt"
        
        try {
            $script:logFile = New-Item -Path $ENV:USERPROFILE\Desktop\$script:logFileName -ItemType File
        }
        catch [Exception] {
            Append-ToLogFile -Message $_.Exception
            Write-Host "`nUnable to create log file at this location. Please check the directory permissions. Exiting script." -ForegroundColor Yellow
            Start-Sleep -Seconds 2
            Exit
        }
    }

    Function Log-Error($Message)
    {
        "`n$Message" >> $script:logFile
    }

    Function Get-UserMemberships
    {
        $allUserData = @()
        $now = Get-Date -Format MM-dd-yyyy_hh-mm-ss
        $dataFile = New-Item -ItemType File -Path $ENV:USERPROFILE\Desktop\GroupMembershipFinderData_$now.txt

        foreach($user in $script:CSV)
        {
            try
            {
                $user = $user.member_id

                Write-Host "`nSearching for user " -NoNewline
                Write-Host "$user" -ForegroundColor Yello

                $userData = Get-ADPrincipalGroupMembership $user | select name | Sort-Object name -Descending 
                [PSCustomObject]$customResult = @{User=$user;MemberOf=$userData.Name}
                [array]$allUserData = [array]$allUserData + $customResult

                Write-Host "`tSuccessfully gathered information for " -ForegroundColor Green -NoNewline
                Write-Host "$user" -ForegroundColor Yellow
            }
            catch [exception]
            {
                Write-Host "There was an error searching for $user"
                Write-Host "Error has been logged."
                Log-Error -Message $Error.Exception[0]
            }
        }

        
        foreach($record in $allUserData)
        {
            $record.User >> $dataFile
            foreach($group in $record.MemberOf)
            {
                "`t"+$group >> $dataFile
            }
        }

        Write-Host "`nUser data successfully exported to " -NoNewline
        Write-Host "$ENV:USERPROFILE\Desktop\GroupMembershipInfo_$now.csv" -ForegroundColor Yellow
    }

}

PROCESS
{
    Get-CSVData
    Get-UserMemberships
}
END
{
    Write-Host "End of script"
}
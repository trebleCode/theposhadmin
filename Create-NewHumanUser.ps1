[cmdletBinding()]
Param(
    [Parameter(Mandatory=$false)]
    [string]$Group,

    [Parameter(Mandatory=$true)]
    [ValidateNotNullOrEmpty]
    [String]$userOUPath,

    [Parameter(Mandatory=$true)]
    [ValidateNotNullOrEmpty]
    [String]$userUPN,

    [Parameter(Mandatory=$true)]
    [ValidateNotNullOrEmpty]
    [SecureString]$Password
)

BEGIN
{
    Function Set-Modules()
    {
        Import-Module ActiveDirectory
        Add-Type -AssemblyName System.Windows.Forms
    }

    Function Set-Vars()
    {
        Write-Host "Initializing script-level variables..."
        $script:usersNotCreated = @()
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

    Function Validate-CSV()
    {
        # check both required headers

        if($script:CSV.displayname -eq $NULL -and $script:CSV.samaccountname -eq $NULL)
        {
            Write-Output "CSV does not contain 'displayname' and 'samaccountname' headers. Please add these and rerun the script."
            Exit
        }

        #check displayname header

        elseif($script:CSV.displayname -eq $NULL -and $script:CSV.samaccountname -ne $NULL)
        {
            Write-Output "CSV does not contain a 'displayname' header. Please add this and rerun the script."
            Exit
        }

        #check samaccountname header

        elseif($script:CSV.samaccountname -eq $NULL -and $script:CSV.displayname -ne $NULL)
        {
            Write-Output "CSV does not contain a 'samaccountname' header. Please add this and rerun the script."
            Exit
        }

        else
        {
            Write-Output "CSV file has been successfully validated."
        }
    }
}

PROCESS
{
    Set-Modules
    Set-Vars
    Get-CSVData
    Validate-CSV

    foreach($row in $script:CSV)
    {
        $name = $row.DisplayName
        $trimmedName = $row.DisplayName.TrimStart('r')
        $givenName = $trimmedName.Split(" ")[0]
        $surname = $trimmedName.Split(" ")[1]
        $upn = $row.SamAccountName + $script:userUPN
        
        Write-Host "Creating user " -NoNewLine
        Write-Host "$name" -ForegroundColor Yellow

        try
        {
            New-ADUser -Name $name -DisplayName $row.DisplayName -GivenName $givenName -Surname $surname -SamAccountName $row.SamAccountName -UserPrincipalName $upn -Path $userOUPath -Enabled $True -AccountPassword (ConvertTo-SecureString -String $Password -AsPlainText -Force)
            if($Group -ne $NULL -and $Group -ne "")
            {
                Add-ADGroupMember -Identity $Group -Members $row.SamAccountName
                Write-Host "Successfully added " -NoNewline -ForegroundColor Green
                Write-Host $row.SamAccountName -ForegroundColor Yellow -NoNewline
                Write-Host " to group " -NoNewline -ForegroundColor Green
                Write-Host $Group -ForegroundColor Yellow
            }
        }
        catch [exception]
        {
            $Error[0]
            $script:usersNotCreated += $name
        }
    }
}

END
{
    if($script:usersNotCreated.Count -gt 0)
    {
        Write-Output "`nUnable to create new accounts for the following users"
        Write-Output $script:usersNotCreated
    }
    Write-Output "Script complete."
}
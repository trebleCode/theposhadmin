BEGIN
{
    Function Set-ScriptVars
    {
        $script:results = New-Object System.Collections.ArrayList
    }

    Function Get-CSV
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
            else
            {
                Write-Host "The file selected is not a CSV file."
                Write-Host "Restarting file selection loop."
            }
        }
        Until ($choice -eq "y")

        $script:CSV = Import-CSV $script:CSVFileName
    }

    Function Get-ParentGroup($Name)
    {
        $parentGroup = $Name

        $nestedGroups = @(Get-ADGroupMember -Identity $Name -Recursive | ? { $_.objectClass -eq "group" })

        if($nestedGroups.Count -eq 0)
        {
            $result = [PSCustomObject] @{ParentGroup=$Name;NestedGroups = $NULL}
            $script:results += $result
        }
        elseif($nestedGroups.Count -gt 0)
        {
            foreach($group in $nestedGroups)
            {
                Write-Host "`tSearching nested group " -NoNewline
                Write-Host "$group" -ForegroundColor Yellow
                $result = [PSCustomObject] @{ParentGroup=$Name;NestedGroups = $group}
                $script:results += $result
                Get-ParentGroup $group
            }
        }
    }
}

PROCESS
{
    Set-ScriptVars
    Get-CSV
    foreach($group in $script:CSV.group)
    {
        Write-Host "Searching group " -NoNewline
        Write-Host "$group" -ForegroundColor Yellow
        Get-ParentGroup $group
    }
}

END
{
    $script:results
}
﻿BEGIN
{
    Function Set-ScriptVars()
    {
        Add-Type -AssemblyName System.Windows.Forms
    }

    Function Show-Banner
    {
@'

  _____                   _____        _        
 / ____|                 |  __ \      | |       
| (___   ___ __ _ _ __   | |  | | __ _| |_ __ _ 
 \___ \ / __/ _` | '_ \  | |  | |/ _` | __/ _` |
 ____) | (_| (_| | | | | | |__| | (_| | || (_| |
|_____/ \___\__,_|_| |_| |_____/ \__,_|\__\__,_|
                                                
                                                
 _____                            _            
|_   _|                          | |           
  | |  _ __ ___  _ __   ___  _ __| |_ ___ _ __ 
  | | | '_ ` _ \| '_ \ / _ \| '__| __/ _ \ '__|
 _| |_| | | | | | |_) | (_) | |  | ||  __/ |   
|_____|_| |_| |_| .__/ \___/|_|   \__\___|_|   
                | |                            
                |_|                            


'@
    }

    Function Select-File($FileType)
    {
        ## Select file via selection dialog

        do {
            if($FileType -eq "xlsx")
            {
                Write-Host "`nPlease select the Excel file to import in the dialog"
            }
            elseif($FileType -eq "txt")
            {
                Write-Host "`nPlease select the Prescan or Postscan text file to import in the dialog"
            }
            
            Start-Sleep -Seconds 1
            $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{InitialDirectory = [Environment]::GetFolderPath('Desktop')}
    
            [void]$FileBrowser.ShowDialog()
     
            Write-Host "`nFile selected: " -NoNewline  
            Write-Host $FileBrowser.FileNames -ForegroundColor Yellow 
    
            $FileName = $FileBrowser.FileName
            if ($FileName.EndsWith(".$FileType"))
            {
                $selectionValid = $True
            }
            elseif($FileName -eq "")
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
                Write-Host "The file selected is not a .$FileType file."
                Write-Host "Restarting file selection loop."
                $selectionValid = $False
            }
        } until ($selectionValid -eq $True)

        if($FileType -eq "txt")
        {
            $Script:TextFile = $FileName
            $Script:TextParentPath = (Get-Item $FileName).Directory.FullName
        }
        elseif($FileType -eq "xlsx")
        {
            $Script:ExcelFile = $FileName
            $Script:ExcelParentPath = (Get-Item $FileName).Directory.FullName
        }
    }

    Function Open-Excel($Sheet)
    {
        $ExcelPath = $Script:ExcelFile
        $Script:Excel = New-Object -ComObject Excel.Application
        $Script:Excel.Visible = $False
        $Script:Excel.UserControl = $False
        $Script:Excel.Interactive = $True
        $Script:Excel.DisplayAlerts = $False

        $Script:ExcelWorkBook = $Script:Excel.Workbooks.Open($ExcelPath)
        $Script:ExcelWorkSheet = $Script:Excel.WorkSheets.item($Sheet)
        $Script:ExcelWorkSheet.Activate()
    }

    Function Get-TextContent()
    {
        $Script:TextContent = Get-Content $Script:TextFile 
    }

    Function Release-Ref ($ref) 
    { 
        ([System.Runtime.InteropServices.Marshal]::ReleaseComObject( 
        [System.__ComObject]$ref) -gt 0) | Out-Null
        [System.GC]::Collect() 
        [System.GC]::WaitForPendingFinalizers() 
    } 
    Function Copy-TextData()
    {       
        # create a CSV from the scan data

        $Script:TextContent = Get-Content $Script:TextFile

        foreach($line in $Script:TextContent)
        {
            if($line -eq "CSV was validated without errors." -or $line -eq "")
            {
                Out-Null
            }
            else
            {
                $i = 0
                $values = $line -split ","
                $result = [PSCustomObject]@{Server=$values[0];`
                                            Role=$values[1];`
                                            Object=$values[2];`
                                            Type=$values[3];`
                                            Path=$values[4]
                                           }

                [Array]$results = $results + $result
            }
        }
        $csvName = "scanData_" + "$(@(000..999) | Get-Random)"
        $results | Export-CSV -Path "$ENV:USERPROFILE\Desktop\$csvName.csv" -NoTypeInformation
        $csvPath = $(Get-Item $ENV:USERPROFILE\Desktop\$csvName.csv).VersionInfo.FileName

        # Remove header generated by hashtable
        # and skip the next two lines

        $tempContent = Get-Content $csvPath
        $replacementContent = $tempContent | Select -Skip 3
        Set-Content $csvPath -Value $replacementContent

        # create temporary workbook and save as xlsx
        $tempXL = New-Object -ComObject Excel.Application
        $tempXL.Visible = $False
        $tempXL.UserControl = $False
        $tempXL.Interactive = $True
        $tempXL.DisplayAlerts = $False

        $tempWB = $tempXL.WorkBooks.Open("$csvPath")
        $tempWS = $tempWB.WorkSheets
        
        $convertedName = $csvPath.Replace(".csv",".xlsx")
        $tempWB.SaveAs($convertedName,1)
        $tempWB.Saved = $True

        $tempRange = $tempWB.Worksheets.Item(1).UsedRange
        $tempRange.Copy()

        if($Script:logSelection -eq "Prescan")
        {
            $permRange = $Script:ExcelWorkBook.Worksheets.Item(2)
        }
        else
        {
            $permRange = $Script:ExcelWorkBook.Worksheets.Item(3)
        }

        $subRange = $permRange.Range("A2","E2")
        $permRange.Paste($subRange)
        $permRange.Columns.AutoFit()

        $Script:ExcelWorkBook.Save()
        $Script:ExcelWorkBook.Saved = $True
        $Script:Excel.Quit()

        $tempWB.Save()
        $tempWB.Saved = $True
        $tempXL.Quit()

        Release-Ref($Script:ExcelWorkSheet)
        Release-Ref($tempWS)

        Release-Ref($Script:ExcelWorkBook)
        Release-Ref($tempWB)

        Release-Ref($Script:Excel)
        Release-Ref($tempXL)

        Remove-Item $csvPath -Force
        Get-Item $convertedName | Remove-Item -Force
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
            $Script:TextFile = $NULL
            $Script:logSelection = $NULL
            Run-Selection
        }
        else
        {
            Out-Null
        }
    }

    Function Run-Selection
    {
        Select-File -FileType "xlsx"
        Select-File -FileType "txt"
        if($Script:TextFile -match "Prescan")
        {
            Open-Excel -Sheet "Prescan"
            $Script:logSelection = "Prescan"
        }
        elseif($Script:TextFile -match "Postscan")
        {
            Open-Excel -Sheet "Postscan"
            $Script:logSelection = "Postscan"
        }
    
        Get-TextContent
        Copy-TextData
        Prompt-ReRun
    }
}

PROCESS
{
    Set-ScriptVars
    Show-Banner
    Run-Selection
}

END
{

}
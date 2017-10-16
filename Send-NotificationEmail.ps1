<#
    .SYNOPSIS
        Imports an excel file and emails a manager listing the direct reports

    .DESCRIPTION
        - Imports an excel file through a file selection dialog
            Column 1 - Manager Name
            Column 2 - Manager's email
            Column 3 - Direct report name
            Column 4 - Direct report's email

        - Iterates through each manager and associates direct reports
          to build an array of Manager -> Direct Reports

        - Outputs a GridView object so user can visually confirm appropriate data

        - Asks for confirmation before sending

        - Emails manager the HTML email listing the manager's direct reports

        - All successful and failed sends are logged to $ENV:USERPROFILE\Desktop

#>

Function Set-Dependencies()
{
    [Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
    $script:reportArray = $NULL
}

Function Setup-Log()
{
    $now = [DateTime]::Now.ToString("MM-dd-yyyy_hh-mm-ss")
    $prefixName = "EmailSendLog_"
    $extension = ".txt"
    $logName = $prefixName + $now + $extension
    $script:log = New-Item -ItemType File -Path $ENV:USERPROFILE\Desktop -Name $logName
}

Function Append-ToLog([String]$Message)
{
    "`n$Message"  >> $script:log
}

Function Kill-ExistingExcel
{
    Get-Process | ? {$_.ProcessName -eq "EXCEL"} | Stop-Process
}
    
Function Select-File()
{
    ## Select file via selection dialog

    do {

        Write-Host "`nPlease select the Excel file to import in the dialog"
            
        Start-Sleep -Seconds 1
        $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{InitialDirectory = [Environment]::GetFolderPath('Desktop')}
    
        [void]$FileBrowser.ShowDialog()
     
        Write-Host "`nFile selected: " -NoNewline  
        Write-Host $FileBrowser.FileNames -ForegroundColor Yellow 
    
        $FileName = $FileBrowser.FileName
        if ($FileName.EndsWith(".xlsx"))
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
            Write-Host "The file selected is not a .XLSX file."
            Write-Host "Restarting file selection loop."
            $selectionValid = $False
        }
    } until ($selectionValid -eq $True)

    $Script:ExcelFile = $FileName
    $Script:ExcelParentPath = (Get-Item $FileName).Directory.FullName

}

Function Open-Excel($Sheet)
{
    $ExcelPath = $Script:ExcelFile
    $Script:Excel = New-Object -ComObject Excel.Application
    $Script:Excel.Visible = $True
    $Script:Excel.UserControl = $False
    $Script:Excel.Interactive = $True
    $Script:Excel.DisplayAlerts = $False

    $Script:ExcelWorkBook = $Script:Excel.Workbooks.Open($ExcelPath)
    $Script:ExcelWorkSheet = $Script:Excel.WorkSheets.item("Sheet1")
    $Script:ExcelWorkSheet.Activate()
}

Function Create-DirectReportArray()
{        
    $rows = $script:ExcelWorksheet.Rows
    $usedRows = $script:ExcelWorksheet.UsedRange.Rows 
    $usedRowCount = $script:ExcelWorksheet.UsedRange.Rows.Count
    $rowCount = 1

    foreach($row in $usedRows)
    {
        Write-Host "Processing row: $($row.Row) | Row Count: $($rowCount)"

        # Skip first row
        if($row.Row -eq 1)
        {
            Out-Null
        }
        else
        {
            $values = $row.Value2

            if($values[1,1] -ne $NULL)
            {
                $rowData = [PSCustomObject]@{
                                                MGR=$values[1,1];
                                                MGREM=$values[1,2];
                                                UN=$values[1,3];
                                                UNEM=$values[1,4]
                                            }
                $previousManagerName = $values[1,1]
                $previousManagerEmail = $values[1,2]

                $script:reportArray = [Array]$script:reportArray + $rowData
            }
            elseif($values[1,1] -eq $NULL)
            {
                $rowData = [PSCustomObject]@{
                                                MGR=$previousManagerName;
                                                MGREM=$previousManagerEmail;
                                                UN=$values[1,3];
                                                UNEM=$values[1,4]
                                            }
                $script:reportArray = [Array]$script:reportArray + $rowData
                        
            }
        }

        $rowCount++
    }
}

Function Create-Table()
{
    $script:tabName = "Managers & Direct Reports"

    # Create Table object
    $script:table = New-Object System.Data.DataTable “$script:tabName”

    # Create first column
    $script:col1 = New-Object System.Data.DataColumn Manager,([string])
    $script:col2 = New-Object System.Data.DataColumn DirectReports,([string])
    $script:table.columns.add($script:col1)
    $script:table.columns.add($script:col2)
}

Function Add-RowToTable($Value1,$Value2)
{
    # Create new row
    $newRow = $script:table.NewRow()

    # Add value to row
    $newRow.Manager = $Value1
    $newRow.DirectReports = $Value2

    # Add row to table
    $script:table.Rows.Add($newRow)
}


Function Get-UniqueManagerEmails
{

    $i = 0

    # get the set of unique manager emails
    $uniqueMgrEmails = $script:reportArray | Select MGREM -Unique
    $uniqueMgrCount = $uniqueMgrEmails.Count

    # build table data to display to user

    foreach($manager in $uniqueMgrEmails)
    {
        $users = $script:reportArray | ? {$_.MGREM -eq $uniqueMgrEmails[$i].MGREM} | Select UN,UNEM,MGREM

        if($users.Count -gt 1)
        {
            $multipleUsers = [pscustomobject]@{manager=$manager.MGREM;drs=$users.UN -join ','}
            Add-RowToTable -Value1 $multipleUsers.manager -Value2 $multipleUsers.drs
        }
        else
        {
            Add-RowToTable -Value1 $manager.MGREM -Value2 $users.UN
        }
   
        $i++
    }

    # show summary for visual confirmation

    Create-SummaryCheck -Input $script:table

    # confirm with user that data is correct

    Confirm-Send

    # build a list of users per manager

    $i = 0

    foreach($manager in $uniqueMgrEmails)
    {
        $users = $script:reportArray | ? {$_.MGREM -eq $uniqueMgrEmails[$i].MGREM} | Select UN,UNEM,MGREM

        $body = 
@"
<html>
<body>
Hi, 
<br>
<br>
YOUR MESSAGE GOES HERE
<br>
<br> $(foreach($user in $users){"`n<b>$($user.UN)</b><br>"})
<br>
<br>Thanks!
<br>

<br>YOUR NAME GOES HERE
</body>


</html>
"@
        try
        {
            Send-Email -To $manager.MGREM -From "YOUREMAIL@YOURDOMAIN.com" -Body $body
            Append-ToLog -Message "[SUCCESS] Successfully sent message to $($manager.MGREM)"
        }
        catch [Exception]
        {
            Append-ToLog -Message "[FAILURE] Failed to send message to $($manager.MGREM)"
        }

        $i++
    }
}

Function Send-Email($To, $From, $Body)
{
    Send-MailMessage -To $To -From $From -BodyAsHTML $Body -SmtpServer "<YOUR SMTP SERVER>" -Port 25 -Subject "<YOUR SUBJECT>"
}

Function Create-SummaryCheck()
{
    $script:table | Out-GridView
}

Function Confirm-Send()
{
    $choice = Read-Host "Are you sure you want to send emails to the listed recipients? (y/n)"

    if($choice.ToLower() -eq "y")
    {
        Write-Host "Preparing to send emails" -ForegroundColor Green
        Append-ToLog -Message "[SELECTION] User chose to proceed with emailing"
    }
    elseif($choice.ToLower() -eq $null)
    {
        Write-Host "Emails will not be sent. Ending script." -ForegroundColor Red
        Append-ToLog -Message "[SELECTION] User chose not to proceed with emailing. Ending script."
        Exit
    }
    else
    {
        Write-Host "Invalid input" -ForegroundColor Yellow
        Confirm-Send
    }

}

Set-Dependencies
Setup-Log
Kill-ExistingExcel
Select-File
Open-Excel
Create-DirectReportArray
Create-Table
Get-UniqueManagerEmails

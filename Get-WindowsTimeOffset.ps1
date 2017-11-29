<#
  .SUMMARY
  Scans a log file at C:\Windows\Temp\w32time.log
  Parses the localClockOffset results and sends
  email based alerts with HTML tables (depending
  on results).
#>
[cmdletbinding()]  
    Param(  
        [Parameter(Mandatory = $True,  Position = 0)]
        [ValidateNotNullOrEmpty()]  
        [string]$Computer = $(hostname),

        [Parameter( Mandatory = $True, Position = 1)]
        [ValidateNotNullOrEmpty()]
        [decimal]$LowThreshold = 0.5,

        [Parameter(Mandatory = $True, Position = 2)]
        [ValidateNotNullOrEmpty()]
        [decimal]$HighThreshold = 1.0
    )
BEGIN
{
    Function Check-Thresholds()
    {
        if($LowThreshold -gt $HighThreshold)
        {
            Write-Output "Low threshold cannot exceed high threshold`nRe-run script with appropriate parameter values."
            Exit
        }
    }

    Function Get-CurrentTime()
    {
        try
        {
            $now = [DateTime]::Now.ToString("MM-dd-yyyy_hh-mm-ss")
            Log-Action -Message "Getting current time...`n$now"
            return $now
        }
        catch [Exception]
        {
            Log-Error
        }
    }

    Function Set-ActionLogFile()
    {
        try
        {
            $script:actionLogFile = New-Item -Path C:\temp\ActionLog-TimeOffsetScript_$currentTime.log
        }
        catch [Exception]
        {
            Log-Error
        }
    }


    Function Get-TimeLog()
    {
        try
        {
            $logContent = Get-Content C:\Windows\temp\w32time.log -Encoding Unicode
            Log-Action -Message "Getting content of C:\Windows\temp\w32time.log"
            return $logContent
        }
        catch [Exception]
        {
            Log-Error
        }
    }

    Function Parse-TimeLogEvents($Content)
    {
        try
        {
            # Get the two most recent LocalClockOffset entries in the log
            $localClockData = $Content |  ? {$_ -match "localclockoffset"} | Select -Last 2
            Log-Action -Message "Getting content of localClockOffset entries"

            # parse as strings and convert to decimal for later

            $latestEntry = $localClockData[1].ToString().Split(" ")[-1].Split(":")[-1]
            $latestEntry = $latestEntry.Substring(0,$latestEntry.LastIndexOf("s"))
            [decimal]$latestEntry = $latestEntry.Substring(0,5)
            Log-Action -Message "Parsed localClockOffsetData entry 1 of 2"


            $previousEntry = $localClockData[0].ToString().Split(" ")[-1].Split(":")[-1]
            $previousEntry = $previousEntry.Substring(0,$previousEntry.LastIndexOf("s"))
            [decimal]$previousEntry = $previousEntry.Substring(0,5)
            Log-Action -Message "Parsed localClockOffsetData entry 2 of 2"

            $newestEntries = [PSCustomObject]@{LatestEntry=$latestEntry;PreviousEntry=$previousEntry}
            return $newestEntries
        }
        catch [Exception]
        {
            Log-Error
        }
    }

    Function Set-ErrorLogFile()
    {
        try
        {
            $script:errorLogFile = New-Item -Path C:\temp\ErrorLog-TimeOffsetScript_$currentTime.log
        }
        catch [Exception]
        {
            Log-Error
        }
    }

    Function Log-Error()
    {
        "ERROR | $($Error[0].Exception)" >> $script:errorLogFile
    }

    Function Log-Action($Message)
    {
        try
        {
            "ACTION | $Message" >> $script:actionLogFile
        }
        catch [Exception]
        {
            Log-Error
        }
    }

    Function Determine-CurrentEntryState($Data)
    {
        Log-Action -Message "Checking threshold entry values"
        # both entries emergency level high
        if($Data.LatestEntry -gt $HighThreshold -and $Data.PreviousEntry -gt $HighThreshold)
        {
            $status = "emergency_alert_latest_high_previous_high"
        }
        # previous entry above low threshold and under high,
        # latest entry above high
        elseif($Data.LatestEntry -gt $HighThreshold -and ($Data.PreviousEntry -lt $HighThreshold -and $Data.PreviousEntry -gt $LowThreshold))
        {
            $status = "priority_alert_latest_high_previous_medium"
        }
        # previous alert under both
        # latest entry above high
        elseif($Data.PreviousEntry -lt $LowThreshold -and $Data.LatestEntry -gt $HighThreshold)
        {
            $status = "priority_alert_latest_high_previous_none"
        }

        # previous entry above high threshold
        # latest entry under high, above low
        elseif(($Data.LatestEntry -lt $HighThreshold -and $Data.LatestEntry -gt $LowThreshold) -and $Data.PreviousEntry -gt $HighThreshold)
        {
            $status = "priority_alert_latest_medium_previous_high"
        }

        # previous entry under both
        # latest entry above low, under high
        elseif(($Data.LatestEntry -lt $HighThreshold -and $Data.LatestEntry -gt $LowThreshold) -and ($Data.PreviousEntry -lt $LowThreshold))
        {
            $status = "minor_alert_latest_medium_previous_none"
        }
        # Latest none, previous high
        elseif($Data.LatestEntry -lt $LowThreshold -and $Data.LatestEntry -gt $HighThreshold)
        {
            $status = "reset_alert_latest_none_previous_high"
        }

        # Latest none, previous medium
        elseif($Data.LatestEntry -lt $LowThreshold -and ($Data.PreviousEntry -lt $HighThreshold -and $Data.PreviousEntry -gt $LowThreshold))
        {
            $status = "reset_alert_latest_none_previous_medium"
        }
        # All good
        elseif($Data.LatestEntry -lt $LowThreshold -and $Data.PreviousEntry -lt $LowThreshold)
        {
            $status = "no_alert_latest_none_previous_none"
        }
        Log-Action -Message "Status determined...$status"
        return $status
    }

    Function Select-Alert($Type,$Data)
    {
        try
        {

        # begin building table

            $emailBody =
@"
<h1><strong>Current Ethereal Offset Thresholds</strong></h1>
<table style="height: 138px; width: 546px;">
<tbody>
<tr style="height: 16px;">
<td style="width: 318px; height: 16px;"><strong>Current Drift From Datum:</strong></td>
"@
    if($Data.LatestEntry -gt $LowThreshold -and $Data.LatestEntry -gt $HighThreshold)
    {
        # red
        $emailBody +=
@"
<td style="width: 267px; height: 16px; text-align: left;padding: 0px 10px" bgcolor="#ff0000">$($Data.LatestEntry)s</span></td>
</tr>

"@
    }
    elseif($Data.LatestEntry -gt $LowThreshold -and $Data.LatestEntry -lt $HighThreshold)
    {
        # yellow
        $emailBody +=
@"
<td style="width: 267px; height: 16px; text-align: left;padding: 0px 10px" bgcolor="#ffff00">$($Data.LatestEntry)s</span></td>
</tr>
"@
    }
    elseif($Data.LatestEntry -lt $LowThreshold -and $Data.LatestEntry -lt $HighThreshold)
    {
        # green
        $emailBody +=
@"
<td style="width: 267px; height: 16px; text-align: left;padding: 0px 10px" bgcolor="#00ff00">$($Data.LatestEntry)s</span></td>
</tr>
"@
    } 

    # Previous Drift from Datum row
    $emailBody +=
@"
<tr style="height: 18px;">
<td style="width: 318px; height: 18px;"><strong>Previous Drift From Datum:</strong></td>
"@

    if($Data.PreviousEntry -gt $LowThreshold -and $Data.PreviousEntry -gt $HighThreshold)
    {
        # red
        $emailBody +=
@"
<td style="width: 267px; height: 16px; text-align: left;padding: 0px 10px" bgcolor="#ff0000">$($Data.PreviousEntry)s</span></td>
</tr>

"@
    }
    elseif($Data.PreviousEntry -gt $LowThreshold -and $Data.PreviousEntry -lt $HighThreshold)
    {
        # yellow
        $emailBody +=
@"
<td style="width: 267px; height: 16px; text-align: left;padding: 0px 10px" bgcolor="#ffff00">$($Data.PreviousEntry)s</span></td>
</tr>
"@
    }
    elseif($Data.PreviousEntry -lt $LowThreshold -and $Data.PreviousEntry -lt $HighThreshold)
    {
        # green
        $emailBody +=
@"
<td style="width: 267px; height: 16px; text-align: left;padding: 0px 10px" bgcolor="#00ff00">$($Data.PreviousEntry)s</span></td>
</tr>
"@
    }

    # Status Code and Machine Name rows
    $emailBody +=
@"
<tr style="height: 18px;">
<td style="width: 318px; height: 18px;"><strong>Status Code:</strong></td>
<td style="width: 267px; height: 18px;">$Type</td>
</tr>
<tr style="height: 18px;">
<td style="width: 318px; height: 18px;"><strong>Machine Name:</strong></td>
<td style="width: 267px; height: 18px;">$Computer</td>
</tr>
<tr style="height: 18px;">
<td style="width: 318px; height: 18px;"><strong>Time Checked:</strong></td>
<td style="width: 267px; height: 18px;">$currentTime</td>
</tr>
</tbody>
</table>
"@
            switch($Type)
            {
                "emergency_alert_latest_high_previous_high"
                {
                    $emailTitle = "WARNING: Ethereal Offsets Thresholds High"
                    $Priority = "High"
                    break
                }

                "priority_alert_latest_high_previous_medium"
                {
                    $emailTitle = "WARNING: Ethereal Offsets Thresholds High"
                    $Priority = "High"
                    break
                }

                "priority_alert_latest_high_previous_none"
                {
                    $emailTitle = "WARNING: Ethereal Offsets Thresholds High"
                    $Priority = "High"
                    break
                }

                "priority_alert_latest_medium_previous_high"
                {
                    $emailTitle = "ALERT: Ethereal Offsets Thresholds Increasing"
                    $Priority = "Normal"
                    break
                }

                "minor_alert_latest_medium_previous_none"
                {
                    $emailTitle = "ALERT: Ethereal Offsets Thresholds Increasing"
                    $Priority = "Normal"
                    break
                }

                "reset_alert_latest_none_previous_high"
                {
                    $emailTitle = "ALERT: Ethereal Offsets Thresholds Reset"
                    $Priority = "Low"
                    break
                }

                "reset_alert_latest_none_previous_medium"
                {
                    $emailTitle = "ALERT: Ethereal Offsets Thresholds Reset"
                    $Priority = "Low"
                    break
                }

                "no_alert_latest_none_previous_none"
                {
                    # don't need to send an alert
                    Log-Action "Status non-alertive, not sending email"
                    break
                }
            }

            if($Type -ne "no_alert_latest_none_previous_none")
            {
                Log-Action "Attempting to send alert message..."
                Send-AlertMessage -Title $emailTitle -Body $emailBody -Priority $Priority
            }
        }
        catch [Exception]
        {
            Log-Error
        }
    }

    Function Send-AlertMessage($Title,$Body,$Priority)
    {
       try
       {
            Send-MailMessage -SmtpServer my.smtp.nocredentialneeded.server -BodyAsHtml $Body -Subject $Title -Port 25 -Priority $Priority -To "scott.sweeney@barings.com" -From "scott.sweeney@barings.com"
        }
        catch [Exception]
        {
            Log-Error
        }
    }
}

PROCESS
{
    Check-Thresholds
    $currentTime = Get-CurrentTime
    Set-ErrorLogFile
    Set-ActionLogFile
    $timeData = Get-TimeLog
    $script:parsedData = Parse-TimeLogEvents -Content $timeData
    $currentState = Determine-CurrentEntryState -Data $script:parsedData
    Select-Alert -Type $currentState -Data $script:parsedData
    
}

END
{
    Log-Action "Script completed. Any available errors will be logged to $script:errorLogFile"
}

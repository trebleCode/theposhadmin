<#
  .SYNOPSIS
  Builds off of Get-NtpTime script from Chris J. Warwick
  on TechNet: https://gallery.technet.microsoft.com/scriptcenter/Get-Network-NTP-Time-with-07b216ca
  and includes configurable default thresholds and HTML-based email alerts

#>

[cmdletbinding()]  
    Param(  
        [Parameter( Mandatory = $false, Position = 0)]
        [ValidateNotNullOrEmpty()]
        [decimal]$LowThreshold = 0.5,

        [Parameter(Mandatory = $false, Position = 1)]
        [ValidateNotNullOrEmpty()]
        [decimal]$HighThreshold = 1.0
)
BEGIN
{
    Function Set-RunstackCounter
    {
        $script:runstackCounter = 0
    }
    Function Get-NtpTime {
    [CmdletBinding()]
    [OutputType()]
    Param (
        [String]$Server = 'pool.ntp.org',
        [Int]$MaxOffset = 10000,     # (Milliseconds) Throw exception if network time offset is larger
        [Switch]$NoDns               # Do not attempt to lookup V3 secondary-server referenceIdentifier    
)


    # NTP Times are all UTC and are relative to midnight on 1/1/1900
    $StartOfEpoch=New-Object DateTime(1900,1,1,0,0,0,[DateTimeKind]::Utc)   


    Function OffsetToLocal($Offset) {
    # Convert milliseconds since midnight on 1/1/1900 to local time
        $StartOfEpoch.AddMilliseconds($Offset).ToLocalTime()
    }


    # Construct a 48-byte client NTP time packet to send to the specified server
    # (Request Header: [00=No Leap Warning; 011=Version 3; 011=Client Mode]; 00011011 = 0x1B)

    [Byte[]]$NtpData = ,0 * 48
    $NtpData[0] = 0x1B    # NTP Request header in first byte


    $Socket = New-Object Net.Sockets.Socket([Net.Sockets.AddressFamily]::InterNetwork,
                                            [Net.Sockets.SocketType]::Dgram,
                                            [Net.Sockets.ProtocolType]::Udp)
    $Socket.SendTimeOut = 2000  # ms
    $Socket.ReceiveTimeOut = 2000   # ms

    Try {
        $Socket.Connect($Server,123)
    }
    Catch {
        Write-Error "Failed to connect to server $Server"
        Throw 
    }


# NTP Transaction -------------------------------------------------------

        $t1 = Get-Date    # t1, Start time of transaction... 
    
        Try {
            [Void]$Socket.Send($NtpData)
            [Void]$Socket.Receive($NtpData)  
        }
        Catch {
            Write-Error "Failed to communicate with server $Server"
            Throw
        }

        $t4 = Get-Date    # End of NTP transaction time

# End of NTP Transaction ------------------------------------------------

    $Socket.Shutdown("Both") 
    $Socket.Close()

# We now have an NTP response packet in $NtpData to decode.  Start with the LI flag
# as this is used to indicate errors as well as leap-second information

    # Check the Leap Indicator (LI) flag for an alarm condition - extract the flag
    # from the first byte in the packet by masking and shifting 

    $LI = ($NtpData[0] -band 0xC0) -shr 6    # Leap Second indicator
    If ($LI -eq 3) {
        Throw 'Alarm condition from server (clock not synchronized)'
    } 

    # Decode the 64-bit NTP times

    # The NTP time is the number of seconds since 1/1/1900 and is split into an 
    # integer part (top 32 bits) and a fractional part, multipled by 2^32, in the 
    # bottom 32 bits.

    # Convert Integer and Fractional parts of the (64-bit) t3 NTP time from the byte array
    $IntPart = [BitConverter]::ToUInt32($NtpData[43..40],0)
    $FracPart = [BitConverter]::ToUInt32($NtpData[47..44],0)

    # Convert to Millseconds (convert fractional part by dividing value by 2^32)
    $t3ms = $IntPart * 1000 + ($FracPart * 1000 / 0x100000000)

    # Perform the same calculations for t2 (in bytes [32..39]) 
    $IntPart = [BitConverter]::ToUInt32($NtpData[35..32],0)
    $FracPart = [BitConverter]::ToUInt32($NtpData[39..36],0)
    $t2ms = $IntPart * 1000 + ($FracPart * 1000 / 0x100000000)

    # Calculate values for t1 and t4 as milliseconds since 1/1/1900 (NTP format)
    $t1ms = ([TimeZoneInfo]::ConvertTimeToUtc($t1) - $StartOfEpoch).TotalMilliseconds
    $t4ms = ([TimeZoneInfo]::ConvertTimeToUtc($t4) - $StartOfEpoch).TotalMilliseconds
 
    # Calculate the NTP Offset and Delay values
    $Offset = (($t2ms - $t1ms) + ($t3ms-$t4ms))/2
    $Delay = ($t4ms - $t1ms) - ($t3ms - $t2ms)

    # Make sure the result looks sane...
    If ([Math]::Abs($Offset) -gt $MaxOffset) {
        # Network server time is too different from local time
        Throw "Network time offset exceeds maximum ($($MaxOffset)ms)"
    }

    # Decode other useful parts of the received NTP time packet

    # We already have the Leap Indicator (LI) flag.  Now extract the remaining data
    # flags (NTP Version, Server Mode) from the first byte by masking and shifting (dividing)

    $LI_text = Switch ($LI) {
        0    {'no warning'}
        1    {'last minute has 61 seconds'}
        2    {'last minute has 59 seconds'}
        3    {'alarm condition (clock not synchronized)'}
    }

    $VN = ($NtpData[0] -band 0x38) -shr 3    # Server version number

    $Mode = ($NtpData[0] -band 0x07)     # Server mode (probably 'server')
    $Mode_text = Switch ($Mode) {
        0    {'reserved'}
        1    {'symmetric active'}
        2    {'symmetric passive'}
        3    {'client'}
        4    {'server'}
        5    {'broadcast'}
        6    {'reserved for NTP control message'}
        7    {'reserved for private use'}
    }

    # Other NTP information (Stratum, PollInterval, Precision)

    $Stratum = [UInt16]$NtpData[1]   # Actually [UInt8] but we don't have one of those...
    $Stratum_text = Switch ($Stratum) {
        0                            {'unspecified or unavailable'}
        1                            {'primary reference (e.g., radio clock)'}
        {$_ -ge 2 -and $_ -le 15}    {'secondary reference (via NTP or SNTP)'}
        {$_ -ge 16}                  {'reserved'}
    }

    $PollInterval = $NtpData[2]              # Poll interval - to neareast power of 2
    $PollIntervalSeconds = [Math]::Pow(2, $PollInterval)

    $PrecisionBits = $NtpData[3]      # Precision in seconds to nearest power of 2
    # ...this is a signed 8-bit int
    If ($PrecisionBits -band 0x80) {    # ? negative (top bit set)
        [Int]$Precision = $PrecisionBits -bor 0xFFFFFFE0    # Sign extend
    } else {
        # ..this is unlikely - indicates a precision of less than 1 second
        [Int]$Precision = $PrecisionBits   # top bit clear - just use positive value
    }
    $PrecisionSeconds = [Math]::Pow(2, $Precision)

    # Determine the format of the ReferenceIdentifier field and decode
    
    If ($Stratum -le 1) {
        # Response from Primary Server.  RefId is ASCII string describing source
        $ReferenceIdentifier = [String]([Char[]]$NtpData[12..15] -join '')
    }
    Else {

        # Response from Secondary Server; determine server version and decode

        Switch ($VN) {
            3       {
                        # Version 3 Secondary Server, RefId = IPv4 address of reference source
                        $ReferenceIdentifier = $NtpData[12..15] -join '.'

                        If (-Not $NoDns) {
                            If ($DnsLookup =  Resolve-DnsName $ReferenceIdentifier -QuickTimeout -ErrorAction SilentlyContinue) {
                                $ReferenceIdentifier = "$ReferenceIdentifier <$($DnsLookup.NameHost)>"
                            }
                        }
                        Break
                    }

            4       {
                        # Version 4 Secondary Server, RefId = low-order 32-bits of  
                        # latest transmit time of reference source
                        $ReferenceIdentifier = [BitConverter]::ToUInt32($NtpData[15..12],0) * 1000 / 0x100000000
                        Break
                    }

            Default {
                        # Unhandled NTP version...
                        $ReferenceIdentifier = $Null
                    }
        }
    }


    # Calculate Root Delay and Root Dispersion values
    
    $RootDelay = [BitConverter]::ToInt32($NtpData[7..4],0) / 0x10000
    $RootDispersion = [BitConverter]::ToUInt32($NtpData[11..8],0) / 0x10000


    # Finally, create output object and return

    $NtpTimeObj = [PSCustomObject]@{
        NtpServer = $Server
        NtpTime = OffsetToLocal($t4ms + $Offset)
        Offset = $Offset
        OffsetSeconds = [Math]::Round($Offset/1000, 3)
        Delay = $Delay
        t1ms = $t1ms
        t2ms = $t2ms
        t3ms = $t3ms
        t4ms = $t4ms
        t1 = OffsetToLocal($t1ms)
        t2 = OffsetToLocal($t2ms)
        t3 = OffsetToLocal($t3ms)
        t4 = OffsetToLocal($t4ms)
        LI = $LI
        LI_text = $LI_text
        NtpVersionNumber = $VN
        Mode = $Mode
        Mode_text = $Mode_text
        Stratum = $Stratum
        Stratum_text = $Stratum_text
        PollIntervalRaw = $PollInterval
        PollInterval = New-Object TimeSpan(0,0,$PollIntervalSeconds)
        Precision = $Precision
        PrecisionSeconds = $PrecisionSeconds
        ReferenceIdentifier = $ReferenceIdentifier
        RootDelay = $RootDelay
        RootDispersion = $RootDispersion
        Raw = $NtpData   # The undecoded bytes returned from the NTP server
    }

    # Set the default display properties for the returned object
    [String[]]$DefaultProperties =  'NtpServer', 'NtpTime', 'OffsetSeconds', 'NtpVersionNumber', 
                                    'Mode_text', 'Stratum', 'ReferenceIdentifier'

    # Create the PSStandardMembers.DefaultDisplayPropertySet member
    $ddps = New-Object Management.Automation.PSPropertySet('DefaultDisplayPropertySet', $DefaultProperties)

    # Attach default display property set and output object
    $PSStandardMembers = [Management.Automation.PSMemberInfo[]]$ddps 
    $NtpTimeObj | Add-Member -MemberType MemberSet -Name PSStandardMembers -Value $PSStandardMembers -PassThru
}

    Function Get-CurrentTime()
    {
        # used for setting log files with a run-time timestamp
        try
        {
            $now = [DateTime]::Now.ToString("MM-dd-yyyy_hh-mm-ss")
            return $now
        }
        catch [Exception]
        {
            Log-Event -Type "ERROR" -Message "$($Error[0].Exception)"
        }
    }

    Function Set-LogFile($Time)
    {
        $script:LogFile = New-Item -Path C:\temp\TimeOffsetScript_$Time.log
    }

    Function Log-Event($Type, $Message, $AdditionalMessage="")
    {

        "$Type | $Message`n$AdditionalMessage" >> $script:LogFile
    }

    Function Select-Alert($ReferenceServer, $Status)
    {
        try
        {
            Log-Event -Type "ACTION" -Message "Building email response HTML"

        # begin building table
        # Time server

        # Reference server name
            $emailBody =
@"
<table style="height: 120px; border: 1; border-style: solid; border-color: black"> 
<tbody>
<tr>
<td style="width: 172px;"><strong>NTP Server</strong></td>
<td style="width: 326px;; border-color: black">&nbsp;$($ReferenceServer.NtpServer)</td>
</tr>
<tr>
"@


        # Drift 
        # red
        if($Status -eq "HIGH")
        {
            $emailBody +=
@"
<td style="width: 172px;"><strong>Drift</strong></td>
<td style="width: 326px;" bgcolor="#ff0000">&nbsp;$($ReferenceServer.OffsetSeconds)s</td>
</tr>
<tr>
"@
        }

        # yellow
        elseif($Status -eq "MEDIUM")
        {
            $emailBody +=
@"
<td style="width: 172px;"><strong>Drift</strong></td>
<td style="width: 326px;" bgcolor="#fff000">&nbsp;$($ReferenceServer.OffsetSeconds)s</td>
</tr>
<tr>
"@
        }

        # green
        else
        {
            $emailBody +=
@"
<td style="width: 172px;"><strong>Drift</strong></td>
<td style="width: 326px;" bgcolor="#00ff00">&nbsp;$($ReferenceServer.OffsetSeconds)s</td>
</tr>
<tr>
"@
        }

        # ReferenceIdentifier

        $emailBody +=
@"
<td style="width: 172px;"><strong>Reference Identifier</strong></td>
<td style="width: 326px;">&nbsp;$($ReferenceServer.ReferenceIdentifier)</td>
</tr>
<tr>
"@

        
        # determine appropriate email priority and title

        if($Status -eq "HIGH")
        {
            $emailTitle = "NTP Offset Monitor - Status EMERGENCY"
            $emailPriority = "High"
            $resyncAttempt = Try-W32TMReSync

            $emailBody += 
@"
<td style="width: 172px;"><strong>Resync Attempt</strong></td>
<td style="width: 326px;">&nbsp;$($resyncAttempt)</td>
</tr>
<tr>
"@
        }
        elseif($Status -eq "MEDIUM")
        {
            $emailTitle = "NTP Offset Monitor - Status WARNING"
            $emailPriority = "Normal"
            $resyncAttempt = Try-W32TMReSync

            $emailBody += 
@"
<td style="width: 172px;"><strong>Resync Attempt</strong></td>
<td style="width: 326px;">&nbsp;$($resyncAttempt)</td>
</tr>
<tr>
"@
        }
        elseif($Status -eq "LOW")
        {
            $emailTitle = "NTP Offset Monitor - Status CLEAR"
            $emailPriority = "Low"
            $emailBody +=
@"
<td style="width: 172px;"><strong>Resync Attempt</strong></td>
<td style="width: 326px;">&nbsp;Not Required</td>
</tr>
<tr>
"@
        }

        Log-Event -Type "ACTION" -Message "Email HTML completed building - returning data to calling function"

        return ($emailTitle,$emailBody,$emailPriority)

        }
        catch [Exception]
        {
            Log-Event -Type "ERROR" -Message "$($Error[0].Exception)" -AdditionalMessage "Unable to construct HTML message"
        }
    }

    Function Try-W32TMReSync()
    {
        try
        {
            $tryResync = Invoke-Command -ComputerName localhost -ScriptBlock {w32tm /resync /computer:timeserver.bam.bamroot.net} -ErrorAction Stop
            Log-Event -Type "ACTION" -Message "Resync attempt successful"
            return "Success"
        }
        catch [exception]
        {
            Log-Event -Type "ERROR" -Message "Resync attempt failed"
            return "Failed"
        }
    }

    Function Send-AlertMessage($Title,$Body,$Priority)
    {
       try
       {
            $recipients = @("recipient1@example.com", "recipient2@example.com")
            Log-Event -Type "ACTION" -Message "Attempting to send email message"
            if($Title -match "CLEAR")
            {
                try
                {
                    Send-MailMessage -SmtpServer smtpmail.bam.bamroot.net -BodyAsHtml $Body -Subject $Title -Port 25 -Priority $Priority -To $recipients[0] -From "scott.sweeney@barings.com"
                    Log-Event -Type "ACTION" -Message "Message of NTP values sent successfully to $($recipients[0]) since it was a CLEAR status."
                }
                catch [Exception]
                {
                    Log-Event -Type "ERROR" -Message "$($Error[0].Exception)" -AdditionalMessage "Failed to send email. Target was only to $($recipients[0]) since it was a CLEAR status."
                }
            }
            else
            {
                try
                {
                    Send-MailMessage -SmtpServer smtpmail.bam.bamroot.net -BodyAsHtml $Body -Subject $Title -Port 25 -Priority $Priority -To $recipients -From "scott.sweeney@barings.com"
                    Log-Event -Type "ACTION" -Message "Message of NTP values sent successfully to $($recipients[0]) and $($recipients[1]) since it was a WARNING or EMERGENCY status." 
                }
                catch [Exception]
                {
                    Log-Event -Type "ERROR" -Message "$($Error[0].Exception)" -AdditionalMessage "Failed to send email. Target was to $($recipients[0]) and $($recipients[1]) since it was a WARNING or EMERGENCY status."
                }
            }    
        }
        catch [Exception]
        {
            Log-Event -Type "ERROR" -Message "$($Error[0].Exception)" -AdditionalMessage "Failed to send email"
        }
    }

    Function Check-ServerThresholds($ReferenceServer)
    {

        Log-Event -Type "ACTION" -Message "Attempt threshold check of $($ReferenceServer.NtpServer)"
        if($ReferenceServer.OffsetSeconds -match "-")
        {
            $ReferenceServer.OffsetSeconds = [Math]::Abs($ReferenceServer.OffsetSeconds)
        }

        if($ReferenceServer.OffsetSeconds -gt $LowThreshold -and $ReferenceServer.OffsetSeconds -gt $HighThreshold)
        {
            Log-Event -Type "ACTION" -Message "$($ReferenceServer.NtpServer) found with Offset: $($ReferenceServer.OffsetSeconds) | Offset: HIGH"
            $RefStatus = "HIGH"
        }
        elseif($ReferenceServer.OffsetSeconds -gt $LowThreshold -and $ReferenceServer.OffsetSeconds -lt $HighThreshold)
        {
            Log-Event -Type "ACTION" -Message "$($ReferenceServer.NtpServer) found with Offset: $($ReferenceServer.OffsetSeconds) | Offset: MEDIUM"
            $RefStatus = "MEDIUM"
        }
        elseif($ReferenceServer.OffsetSeconds -lt $LowThreshold -and $ReferenceServer.OffsetSeconds -lt $HighThreshold)
        {
            Log-Event -Type "ACTION" -Message "$($ReferenceServer.NtpServer) found with Offset: $($ReferenceServer.OffsetSeconds) | Offset: LOW"
            $RefStatus = "LOW"
        }

        return @($RefStatus)

    }

    Function Check-ScriptThresholds()
    {
        if($LowThreshold -gt $HighThreshold)
        {
            Log-Event -Type "ACTION" -Message "Low threshold cannot exceed high threshold`nScript will exit."
            Exit
        }
    }

    Function Run-Stack([switch]$Reloop)
    {
        Check-ScriptThresholds

        if($script:runstackCounter -lt 2)
        {
            if(!$Reloop)
            {
                $currentTime = Get-CurrentTime
                $actionLog = Set-LogFile -Time $currentTime

                Log-Event -Type "ACTION" -Message "Current time obtained"
                Log-Event -Type "ACTION" -Message "Log file set"
            }

            try
            {
                Log-Event -Type "ACTION" -Message "Attempting get of NTP data from timeserver.bam.bamroot.net"
                $timeServerData = Get-NtpTime -Server "timeserver.bam.bamroot.net" -MaxOffset 1000
            
            }
            catch [Exception]
            {
                Log-Event -Type "ERROR" -Message "$($Error[0].Exception)" -AdditionalMessage "Timeserver.bam.bamroot.net unavailable"
            }

            if($timeServerData -eq $NULL)
            {
                Log-Event -Type "ERROR" -Message "One or more queried timeserver were unreachable at runtime. Unable to process offset."
            
                try
                {
                    Send-AlertMessage -Title "NTP Offset Monitor - Status UNAVAILABLE" -Body "timeserver.bam.bamroot.net was unreachable at runtime. Script was unable to process offset times. A re-run will be attempted. Check the log for further details." -Priority "High"
                    Log-Event -Type "ACTION" -Message "Message of server query failure sent successfully"
                }
                catch [Exception]
                {
                    Log-Event -Type "ERROR" -Message "$($Error[0].Exception)" -AdditionalMessage "Failed to send notification email for script error"
                }

                try
                {
                    Log-Event -Type "ACTION" -Message "Attempting to re-run script after failure"
                    $script:runstackCounter++
                    Run-Stack -Reloop   
                }
                catch [Exception]
                {
                    Log-Event -Type "ERROR" -Message "$($Error[0].Exception)" -AdditionalMessage "Failed to re-run script stack. Exiting script"
                    Exit
                }
            }
            else
            {
                try
                {
                    $thresholdStatus = Check-ServerThresholds -ReferenceServer $timeServerData
                }
                catch [Exception]
                {
                    Log-Event -Type "ERROR" -Message "$($Error[0].Exception)"
                }

                try
                {
                    Log-Event -Type "ACTION" -Message "Determining appropriate threshold classification"
                    $alertType = Select-Alert -ReferenceServer $timeServerData -Status $thresholdStatus
                }
                catch [Exception]
                {
                    Log-Event -Type "ERROR" -Message "$($Error[0].Exception)"
                }

                try
                {
                    Send-AlertMessage -Title $alertType[0] -Body $alertType[1] -Priority $alertType[2]
                }
                catch [Exception]
                {
                    Log-Event -Type "ERROR" -Message "$($Error[0].Exception)"
                }
            }
        }
    }
}

PROCESS
{
    Set-RunstackCounter
    Run-Stack
}

END{}

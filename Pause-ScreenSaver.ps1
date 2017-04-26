BEGIN
{
    $ErrorActionPreference = "SilentlyContinue"

    Function Pause-ScreenSaver
    {
         Param(
         [Bool]$Continue = $True, #If This Is True, The Function Will Run Forever
         [Int]$DefaultSleepTime = 60 #Time To Wait Before Re-Looping
         )

        $Shell = New-Object -COM “WScript.Shell”
        $writeInstance = 0

        # While True, send key every 4.5 minutes

        While ($Continue -EQ $True)
        {
            $date = [DateTime]::Now
            Write-Host "`nCurrent Time: $date"

            $writeInstance++
            Write-Host "Current write: " -NoNewline
            Write-Host $writeInstance -ForegroundColor Yellow

            # Set focus to the Powershell window and send the key
            # Note that Set-ForegroundWindow is a PSCX cmdlet

            Write-Host "`nSending F15 key"
            Set-ForegroundWindow (Get-Process Powershell).MainWindowHandle
            $Shell.SendKeys(“F15”) | Out-Null
            Write-Host "Starting sleep interval"
            Start-Sleep -Seconds $DefaultSleepTime

        }
    }
 }
 PROCESS
{
    Pause-ScreenSaver -Continue $True -DefaultSleepTime 270
}

END
{

}
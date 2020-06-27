# Sleeps X seconds and displays a progress bar
function Start-SleepWithProgress {
    <#
.Synopsis
   # Sleeps X seconds and displays a progress bar
.EXAMPLE
   Start-SleepWithProgress -SleepTime 10 -Message "Closing Window"
.EXAMPLE
   Start-SleepWithProgress -SleepTime 20
#>
    [CmdletBinding()]
    [Alias()]
    Param
    (
        # Enter number of seconds to Sleep
        [Parameter(Mandatory = $true)]
        $SleepTime,

        # Enter message to display
        [String]
        $Message
    )

    if (!$Message) {$Message = "Taking a Break"}

    # Loop Number of seconds you want to sleep
    For ($i = 0; $i -le $SleepTime; $i++) {
        $Timeleft = ($SleepTime - $i);

        # Progress bar showing progress of the sleep
        Write-Progress -Activity $Message -CurrentOperation "$Timeleft More Seconds" -PercentComplete (($i / $SleepTime) * 100);

        # Sleep 1 second
        Start-Sleep 1
    }

    Write-Progress -Completed -Activity $Message
}

# Setup a new O365 Powershell Session
Function New-CleanO365Session {
    param(
        [Parameter(Mandatory = $true)]
        $Credfile
    )

    $CredFromFile = Import-Clixml $Credfile
    $CredFromFile.Password = $CredFromFile.Password | ConvertTo-SecureString
    $Credential = New-Object System.Management.Automation.PSCredential($CredFromFile.username, $CredFromFile.Password)
    #Note add this sentting to the session for the PSBOX
    # $proxysettings = New-PSSessionOption -ProxyAccessType IEConfig


    if ( Get-PSSession ) {
        Write-Log "Removing all PS Sessions"
        Get-PSSession | Remove-PSSession -Confirm:$false

        # Sleep 15s to allow the sessions to tear down fully
        Write-Log ("Taking a Break 15 seconds for Session Tear Down")
        Start-SleepWithProgress -SleepTime 15
    }

    # Force Garbage collection just to try and keep things more agressively cleaned up due to some issue with large memory footprints
    [System.GC]::Collect()

    # Clear out all errors
    $Error.Clear()

    # Create the session
    Write-Log "Creating new PS Session"

    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://ps.outlook.com/powershell/" -Credential $Credential -Authentication Basic -AllowRedirection -ErrorAction SilentlyContinue # -SessionOption $proxysettings 

    # Check for an error while creating the session
    if ($Error.Count -gt 0) {

        Write-Log "[ERROR] - Error while setting up session"
        Write-log $Error[0].Exception.Message

        # Increment our error count so we abort after so many attempts to set up the session
        $ErrorCount++

        # if we have failed to setup the session > 3 times then we need to abort because we are in a failure state
        if ($ErrorCount -gt 3) {

            Write-log "[ERROR] - Failed to setup session after multiple tries"
            Write-log "[ERROR] - Aborting Script"
            Stop-Transcript
            if ((Get-Content $TFile) -contains "No New Items" -and (((Get-ChildItem $TFile).Length | Measure-Object -Sum).Sum / 1kb) -lt "1.5") { Remove-Item $TFile }
            Exit

        }

        # If we are not aborting then sleep 60s in the hope that the issue is transient
        Write-Log "Taking a Break 60s so that issue can potentially be resolved"
        Start-SleepWithProgress -sleeptime 60

        # Attempt to set up the sesion again
        New-CleanO365Session -Credfile $CredFile
    } else {
        $ErrorCount = 0
    }

    $null = Import-PSSession $Session -AllowClobber -WarningAction SilentlyContinue
    $null = Connect-MsolService -Credential $Credential
    # Set the Start time for the current session
    Set-Variable -Scope script -Name SessionStartTime -Value (Get-Date)
}

# Verifies that the connection is healthy
# Goes ahead and resets it every $ResetSeconds number of seconds either way
Function Test-O365Session {
    Write-Log "Testing Session"
    # Get the time that we are working on this object to use later in testing
    $ObjectTime = Get-Date

    # Reset and regather our session information
    $SessionInfo = $null
    $SessionInfo = Get-PSSession | Where-Object {$_.State -eq 'Opened'}

    # Make sure we found a session
    if ($null -eq $SessionInfo) {
        Write-Log "[INFO] - No Session Found"
        Write-log "Recreating Session"
        New-CleanO365Session -Credfile $EOPCreds
    }
    # If we have looped thru objects for an amount of time gt our reset seconds then tear the session down and recreate it
    elseif (($ObjectTime - $SessionStartTime).totalseconds -gt $ResetSeconds) {
        Write-Log ("Session Has been active for greater than " + $ResetSeconds + " seconds" )
        Write-Log "Rebuilding Connection"

        # Estimate the throttle delay needed since the last session rebuild
        # Amount of time the session was allowed to run * our activethrottle value
        # Divide by 2 to account for network time, script delays, and a fudge factor
        # Subtract 15s from the results for the amount of time that we spend setting up the session anyway
        [int]$DelayinSeconds = ((($ResetSeconds * $ActiveThrottle) / 2) - 15)

        # If the delay is >15s then sleep that amount for throttle to recover
        if ($DelayinSeconds -gt 0) {

            Write-Log ("Taking a Break " + $DelayinSeconds + " addtional seconds to allow throttle recovery")
            Start-SleepWithProgress -SleepTime $DelayinSeconds
        }
        # If the delay is <15s then the sleep already built into New-CleanO365Session should take care of it
        else {
            Write-Log ("Active Delay calculated to be " + ($DelayinSeconds + 15) + " seconds no addtional delay needed")
        }

        # new O365 session and reset our object processed count
        New-CleanO365Session -Credfile $EOPCreds
    } else {
        # If session is active and it hasn't been open too long then do nothing and keep going
    }

    # If we have a manual throttle value then sleep for that many milliseconds
    if ($ManualThrottle -gt 0) {
        Write-log ("Taking a Break " + $ManualThrottle + " milliseconds")
        Start-Sleep -Milliseconds $ManualThrottle
    }

    # Force Garbage collection just to try and keep things more agressively cleaned up due to some issue with large memory footprints
    [System.GC]::Collect()
}

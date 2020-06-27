Function Write-Log {
    [CmdletBinding()]
    Param
    (
        # Enter Message to log
        [Parameter(Mandatory)]
        $String,

        # Log file to save Data
        [String]
        $LogFile,

        # No Verbose switch. - Do not print to console.
        [Switch]
        $NoOutput
    )

    [String]$Date = Get-Date -Format G

    <#Switch -Wildcard ($string) {

        "Someting*" { "Do Something" }

        } #>

    # Write everything to our log file
    if ($Global:LogFile) {
        ("[" + $Date + "] - " + $String) | Out-File -FilePath $Global:LogFile -Append
    }
    # If NoVerbose then supress cmd output
    if (!$Global:LogFile -or !$NoOutput) {
        ("[" + $Date + "] - " + $String) | Write-Output
    }
}

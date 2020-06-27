Function Find-Data($Data) {
    $Found = $false

    if ($Data -notmatch "\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}") {
        $FilterPolicies = Get-HostedContentFilterPolicy
        $i = 0
        While ($Found -eq $false -and $FilterPolicies.Count -gt $i) {
            Foreach ($FilterPolicy in $FilterPolicies) {
                if ($FilterPolicy.AllowedSenderDomains.Domain -contains $Data -or ($FilterPolicy.AllowedSenders.Sender | Select-Object -ExpandProperty Address) -contains $Data -or ($FilterPolicy.BlockedSenderDomains.Domain) -contains $Data -or ($FilterPolicy.BlockedSenders | Select-Object -ExpandProperty Sender | Select-Object -ExpandProperty Address) -contains $Data) {
                    $Global:Fltr = $FilterPolicy
                    $Found = $True
                    $i++
                    Break
                } else {
                    $Found = $false
                    $i++
                }
            }
        }
    } else {
        $ConnectionFilterPolicy = Get-HostedConnectionFilterPolicy
        if (!$ConnectionFilterPolicy.Count) {
            $ConnectionFilterPolicy | Add-Member "Count" -NotePropertyValue "1"
        }
        $i = 0
        While ($Found -eq $false -and $ConnectionFilterPolicy.Count -gt $i) {
            Foreach ($IP in $ConnectionFilterPolicy) {
                if ($IP.IPAllowList -contains $Data`
               -or ($IP.IPBlockList -contains $Data)) {
                    $Global:Fltr = $IP.Name
                    $Found = $True
                    $i++
                    Break
                } else {
                    $Found = $false
                    $i++
                }
            }
        }
    }
    Return $Found
}

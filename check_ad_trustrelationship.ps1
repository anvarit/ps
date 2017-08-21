$localCredential = Get-Credential

@(Get-AdComputer -Filter *).foreach({

    $output = @{ ComputerName = $_.Name }

    if (-not (Test-Connection -ComputerName $_.Name -Quiet -Count 1)) { $output.Status = 'Offline'
        } else {

        $trustStatus = Invoke-Command -ComputerName $_.Name -ScriptBlock { Test-ComputerSecureChannel } -Credential $localCredential
        $output.Status = $trustStatus
    }

    [pscustomobject]$output

})
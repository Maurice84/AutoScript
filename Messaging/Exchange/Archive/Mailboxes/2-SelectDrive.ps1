Function Messaging-Exchange-Archive-Mailboxes-2-SelectDrive {
    # ============
    # Declarations
    # ============
    $Task = "Select a drive"
    If ($global:ExchangeServer -ne $env:computername) {
        $Task += " (on $global:ExchangeServer)"
    }
    $Task += " for the mailbox export(s)"
    # =========
    # Execution
    # =========
    $global:StartDate = Get-Date 01-01-1900
    $global:EndDate = Get-Date
    Script-Index-Objects -CurrentTask $Task -FunctionName ($MyInvocation.MyCommand).Name -Server $global:ExchangeServer -Type "Drive" -Functie "Selecteren"
    If ($global:Drive) {
        $global:TitelDrive = $global:Drive
    }
    Else {
        $global:TitelDrive = "None (no archive export)"
    }
    # ==========
    # Finalizing
    # ==========
    Set-Variable $global:VarHeaderName -Value $global:TitelDrive
    Set-Variable $global:VarHeaderName -Value ("Write-Host -NoNewLine ('- ' + '$Task' + ': ') -ForegroundColor Gray; Write-Host '" + (Get-Variable -Name $global:VarHeaderName).Value + "' -ForegroundColor Yellow")
    Invoke-Expression -Command $FunctionTaskNames[[int]$FunctionTaskNames.IndexOf(($MyInvocation.MyCommand).Name) + 1]
}
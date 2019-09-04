Function Messaging-Exchange-Export-Mailboxes-3-SelectDrive {
    # ============
    # Declarations
    # ============
    $Task = "Select a drive"
    If ($global:ExchangeServer -ne $env:computername) {
        $Task += " (on $global:ExchangeServer) "
    }
    $Task += "as location for the export(s) of the mailbox(es)"
    # =========
    # Execution
    # =========
    Script-Module-SetHeaders -CurrentTask $Task -Name ($MyInvocation.MyCommand).Name
    Script-Index-Objects -CurrentTask $Task -FunctionName ($MyInvocation.MyCommand).Name -Server $global:ExchangeServer -Type "Drive" -Functie "Selecteren"
    # ==========
    # Finalizing
    # ==========
    Set-Variable $global:VarHeaderName -Value $global:Drive
    Set-Variable $global:VarHeaderName -Value ("Write-Host -NoNewLine ('- ' + '$Task' + ': ') -ForegroundColor Gray; Write-Host '" + (Get-Variable -Name $global:VarHeaderName).Value + "' -ForegroundColor Yellow")
    Invoke-Expression -Command $FunctionTaskNames[[int]$FunctionTaskNames.IndexOf(($MyInvocation.MyCommand).Name) + 1]
}
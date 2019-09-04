Function Identity-ActiveDirectory-Purge-Accounts-1-SelectAccount {
    # ============
    # Declarations
    # ============
    $Task = "Please select domain account(s) for DEFINITIVE purging"
    # =========
    # Execution
    # =========
    Script-Module-SetHeaders -CurrentTask $Task -Name ($MyInvocation.MyCommand).Name
    Script-Module-ImportMicrosoft -Name ($MyInvocation.MyCommand).Name -Module "ActiveDirectory"
    Script-Index-Objects -CurrentTask $Task -FunctionName ($MyInvocation.MyCommand).Name -Type "AccountRecycleBin" -Functie "Markeren"
    $global:ArrayAccounts = $global:Array
    # ==========
    # Finalizing
    # ==========
    Set-Variable $global:VarHeaderName -Value ([string]$global:Aantal + " (" + ($global:ArrayAccounts.Name -join ", ") + ")")
    Set-Variable $global:VarHeaderName -Value ("Write-Host -NoNewLine ('- ' + '$Task' + ': ') -ForegroundColor Gray; Write-Host '" + (Get-Variable -Name $global:VarHeaderName).Value + "' -ForegroundColor Yellow")
    Invoke-Expression -Command $FunctionTaskNames[[int]$FunctionTaskNames.IndexOf(($MyInvocation.MyCommand).Name) + 1]
}
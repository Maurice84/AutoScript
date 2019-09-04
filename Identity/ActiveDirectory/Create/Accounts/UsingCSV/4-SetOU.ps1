Function Identity-ActiveDirectory-Create-Accounts-UsingCSV-4-SetOU {
    # ============
    # Declarations
    # ============
    $Task = "Select an OU where the domain account(s) must be created"
    # =========
    # Execution
    # =========
    Script-Module-SetHeaders -CurrentTask $Task -Name ($MyInvocation.MyCommand).Name
    Script-Index-OU
    If ($global:OUPath -like "*Users*") {
        $global:Profiel = "Domein gebruiker"
    }
    If ($global:OUPath -like "*External*") {
        $global:Profiel = "Beheerder account"
    }
    If ($global:OUPath -like "*Mailbox*") {
        $global:Profiel = "Mailbox account"
    }
    If ($global:OUPath -like "*Service*") {
        $global:Profiel = "Service account"
    }
    If ($global:OUPath -like "*Test*") {
        $global:Profiel = "Test gebruiker"
    }
    # ==========
    # Finalizing
    # ==========
    Set-Variable $global:VarHeaderName -Value $global:Subtree
    Set-Variable $global:VarHeaderName -Value ("Write-Host -NoNewLine ('- ' + '$Task' + ': ') -ForegroundColor Gray; Write-Host '" + (Get-Variable -Name $global:VarHeaderName).Value + "' -ForegroundColor Yellow")
    Invoke-Expression -Command $FunctionTaskNames[[int]$FunctionTaskNames.IndexOf(($MyInvocation.MyCommand).Name) + 1]
}
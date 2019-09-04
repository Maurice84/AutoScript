Function Identity-ActiveDirectory-Create-Accounts-2-Duplicate {
    # ============
    # Declarations
    # ============
    $Task = "Select an account or user as a duplicate"
    # =========
    # Execution
    # =========
    Script-Module-SetHeaders -CurrentTask $Task -Name ($MyInvocation.MyCommand).Name
    If ($global:Profiel -eq "Profielkopie") {
        Script-Index-Objects -CurrentTask $Task -FunctionName ($MyInvocation.MyCommand).Name -Type "Account" -Functie "Selecteren"
        # ==========
        # Finalizing
        # ==========
        Set-Variable $global:VarHeaderName -Value ($global:DisplayName + " " + $global:Customer)
        Set-Variable $global:VarHeaderName -Value ("Write-Host -NoNewLine ('- ' + '$Task' + ': ') -ForegroundColor Gray; Write-Host '" + (Get-Variable -Name $global:VarHeaderName).Value + "' -ForegroundColor Yellow")
    }
    Invoke-Expression -Command $FunctionTaskNames[[int]$FunctionTaskNames.IndexOf(($MyInvocation.MyCommand).Name) + 1]
}
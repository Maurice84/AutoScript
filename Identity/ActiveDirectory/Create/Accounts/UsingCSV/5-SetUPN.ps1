Function Identity-ActiveDirectory-Create-Accounts-UsingCSV-5-SetUPN {
    # ============
    # Declarations
    # ============
    $Task = "Select an UPN suffix for the domain account(s)"
    # =========
    # Execution
    # =========
    Script-Module-SetHeaders -CurrentTask $Task -Name ($MyInvocation.MyCommand).Name
    Script-Index-Objects -CurrentTask $Task -FunctionName ($MyInvocation.MyCommand).Name -Type "UPNSuffix" -Functie "Selecteren"
    # ==========
    # Finalizing
    # ==========
    Set-Variable $global:VarHeaderName -Value $global:UPNSuffix
    Set-Variable $global:VarHeaderName -Value ("Write-Host -NoNewLine ('- ' + '$Task' + ': ') -ForegroundColor Gray; Write-Host '" + (Get-Variable -Name $global:VarHeaderName).Value + "' -ForegroundColor Yellow")
    Invoke-Expression -Command $FunctionTaskNames[[int]$FunctionTaskNames.IndexOf(($MyInvocation.MyCommand).Name) + 1]
}
Function Identity-ActiveDirectory-Create-Accounts-5-SetUPN {
    # ============
    # Declarations
    # ============
    $Task = "Select a UPN suffix for the new account"
    # =========
    # Execution
    # =========
    Script-Module-SetHeaders -CurrentTask $Task -Name ($MyInvocation.MyCommand).Name
    If ($global:Profiel -ne "Profielkopie") {
        Script-Index-Objects -CurrentTask $Task -FunctionName ($MyInvocation.MyCommand).Name -Type "UPNSuffix" -Functie "Selecteren"
    }
    Else {
        $global:UPNSuffix = $UserPrincipalName.Split('@')[1]
    }
    # ==========
    # Finalizing
    # ==========
    Set-Variable $global:VarHeaderName -Value $global:UPNSuffix
    Set-Variable $global:VarHeaderName -Value ("Write-Host -NoNewLine ('- ' + '$Task' + ': ') -ForegroundColor Gray; Write-Host '" + (Get-Variable -Name $global:VarHeaderName).Value + "' -ForegroundColor Yellow")
    Invoke-Expression -Command $FunctionTaskNames[[int]$FunctionTaskNames.IndexOf(($MyInvocation.MyCommand).Name) + 1]
}
Function Identity-ActiveDirectory-Modify-Accounts-1-SelectAccount {
    # ============
    # Declarations
    # ============
    $Task = "Select a domain account to modify"
    # =========
    # Execution
    # =========
    Script-Module-SetHeaders -CurrentTask $Task -Name ($MyInvocation.MyCommand).Name
    Script-Module-ImportMicrosoft -Name ($MyInvocation.MyCommand).Name -Module "ActiveDirectory"
    Script-Index-Objects -CurrentTask $Task -FunctionName ($MyInvocation.MyCommand).Name -Type "Account" -Functie "Selecteren"
    # ==========
    # Finalizing
    # ==========
    Invoke-Expression -Command $FunctionTaskNames[[int]$FunctionTaskNames.IndexOf(($MyInvocation.MyCommand).Name) + 1]
}
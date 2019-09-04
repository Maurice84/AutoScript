Function Identity-ActiveDirectory-Overview-Accounts-1-GetData {
    # =========
    # Execution
    # =========
    Script-Module-SetHeaders -Name ($MyInvocation.MyCommand).Name
    Script-Module-ImportMicrosoft -Name ($MyInvocation.MyCommand).Name -Module "ActiveDirectory"
    Script-Index-Objects -FunctionName ($MyInvocation.MyCommand).Name -Type "AccountBulk" -Functie "Overzicht"
}
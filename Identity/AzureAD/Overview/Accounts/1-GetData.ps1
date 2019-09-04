Function Identity-AzureAD-Overview-Accounts-1-GetData {
    # =========
    # Execution
    # =========
    Script-Module-SetHeaders -CurrentTask $Task -Name ($MyInvocation.MyCommand).Name
    Script-Module-ImportMicrosoft -Name ($MyInvocation.MyCommand).Name -Module "CredentialManager"
    Script-Module-ImportMicrosoft -Name ($MyInvocation.MyCommand).Name -Module "AzureAD"
    Script-Module-ImportMicrosoft -Name ($MyInvocation.MyCommand).Name -Module "MSOnline"
    Script-Connect-Office365 -Module "AzureAD" -Name ($MyInvocation.MyCommand).Name
    Script-Index-Objects -FunctionName ($MyInvocation.MyCommand).Name -Filter * -Type "Office365AccountBulk" -Functie "Overzicht"
    # ==========
    # Finalizing
    # ==========
    Get-Variable -Exclude $global:StartupVariables | Remove-Variable -ErrorAction SilentlyContinue; Script-Disconnect-Server; Invoke-Expression -Command $global:MenuNameCategory
}

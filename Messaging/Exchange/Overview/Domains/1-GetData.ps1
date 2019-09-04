Function Messaging-Exchange-Overview-Domains-1-GetData {
    # =========
    # Execution
    # =========
    Script-Module-SetHeaders -Name ($MyInvocation.MyCommand).Name
    Script-Module-ImportMicrosoft -Name ($MyInvocation.MyCommand).Name -Module "Exchange"
    Script-Index-Objects -FunctionName ($MyInvocation.MyCommand).Name -Type "Emaildomein" -Functie "Overzicht"
}
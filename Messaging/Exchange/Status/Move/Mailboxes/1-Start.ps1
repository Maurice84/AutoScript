Function Messaging-Exchange-Status-Move-Mailboxes-1-Start {
    # =========
    # Execution
    # =========
    Script-Module-SetHeaders -Name ($MyInvocation.MyCommand).Name
    Script-Module-ImportMicrosoft -Name ($MyInvocation.MyCommand).Name -Module "Exchange"
    Script-Module-SetHeaders -Name ($MyInvocation.MyCommand).Name
    Messaging-Exchange-Status-Results -Selectie "Move"
}
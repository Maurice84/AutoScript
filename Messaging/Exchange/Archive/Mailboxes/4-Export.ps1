Function Messaging-Exchange-Archive-Mailboxes-4-Export {
    # ============
    # Declarations
    # ============
    $Task = "Exporting selected mailbox(es)"
    # =========
    # Execution
    # =========
    Script-Module-SetHeaders -CurrentTask $Task -Name ($MyInvocation.MyCommand).Name
    Messaging-Exchange-Status-Request -Selectie "Export" -Type "Archive"
    Invoke-Expression -Command $FunctionTaskNames[[int]$FunctionTaskNames.IndexOf(($MyInvocation.MyCommand).Name) + 1]
}
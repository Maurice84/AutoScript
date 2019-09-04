Function Messaging-Exchange-Export-Mailboxes-5-Start {
    # =========
    # Execution
    # =========
    Script-Module-SetHeaders -Name ($MyInvocation.MyCommand).Name
    Write-Host "- Exporting selected mailbox(es)..." -ForegroundColor Gray
    Write-Host
    Messaging-Exchange-Status-Request -Selectie "Export"
}
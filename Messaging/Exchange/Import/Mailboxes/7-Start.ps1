Function Messaging-Exchange-Import-Mailboxes-7-Start {
    # =========
    # Execution
    # =========
    Script-Module-SetHeaders -CurrentTask $Task -Name ($MyInvocation.MyCommand).Name
    Write-Host "- Importing selected PST-file(s)..." -ForegroundColor Gray
    Write-Host
    Messaging-Exchange-Status-Request -Selectie "Import"
}
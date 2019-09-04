Function Messaging-Exchange-Restore-Mailboxes-5-Start {
    # =========
    # Execution
    # =========
    Script-Module-SetHeaders -CurrentTask $Task -Name ($MyInvocation.MyCommand).Name
    Write-Host -NoNewLine "- Restoring mail-items from $SourceName to $TargetName" -ForegroundColor Gray
    If ($TargetFolder -ne "n.a.") {
        Write-Host " with folder name: $TargetFolder..." -ForegroundColor Gray
    }
    Else {
        Write-Host "..." -ForegroundColor Gray
    }
    Write-Host
    Messaging-Exchange-Status-Request -Selectie "Restore"
}
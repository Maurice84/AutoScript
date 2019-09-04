Function Messaging-Exchange-Restore-Mailboxes-1-SelectSource {
    # ============
    # Declarations
    # ============
    $Task = "Select a mailbox to recover"
    # =========
    # Execution
    # =========
    Script-Module-SetHeaders -CurrentTask $Task -Name ($MyInvocation.MyCommand).Name
    Script-Module-ImportMicrosoft -Name ($MyInvocation.MyCommand).Name -Module "Exchange"
    Script-Module-SetHeaders -CurrentTask $Task -Name ($MyInvocation.MyCommand).Name
    Script-Index-Objects -CurrentTask $Task -FunctionName ($MyInvocation.MyCommand).Name -Type "MailboxHerstel" -Functie "Selecteren"
    $global:MailboxRestore = $global:Array
    $global:SourceName = $MailboxRestore.Name
    $global:SourceGUID = $MailboxRestore.Identity
    $global:SourceDB = $MailboxRestore.Database
    # ==========
    # Finalizing
    # ==========
    Set-Variable $global:VarHeaderName -Value $global:SourceName
    Set-Variable $global:VarHeaderName -Value ("Write-Host -NoNewLine ('- ' + '$Task' + ': ') -ForegroundColor Gray; Write-Host '" + (Get-Variable -Name $global:VarHeaderName).Value + "' -ForegroundColor Yellow")
    Invoke-Expression -Command $FunctionTaskNames[[int]$FunctionTaskNames.IndexOf(($MyInvocation.MyCommand).Name) + 1]
}
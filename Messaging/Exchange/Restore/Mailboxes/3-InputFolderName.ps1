Function Messaging-Exchange-Restore-Mailboxes-3-InputFolderName {
    # ============
    # Declarations
    # ============
    $Task = "Select a target folder for the restore"
    # =========
    # Execution
    # =========
    Script-Module-SetHeaders -CurrentTask $Task -Name ($MyInvocation.MyCommand).Name
    Write-Host
    Write-Host -NoNewLine "  Please enter a target folder name (for mailbox $TargetName) for the restore (i.e.: "; Write-Host -NoNewLine ("Restore " + $FileFormat) -ForegroundColor Yellow; Write-Host "),"
    Write-Host -NoNewLine "  or leave empty to import into the root of the mailbox: "
    $global:TargetFolder = Read-Host
    If (!$global:TargetFolder) {
        $global:TargetFolder = "n.a."
    }
    # ==========
    # Finalizing
    # ==========
    Set-Variable $global:VarHeaderName -Value $global:TargetFolder
    Set-Variable $global:VarHeaderName -Value ("Write-Host -NoNewLine ('- ' + '$Task' + ': ') -ForegroundColor Gray; Write-Host '" + (Get-Variable -Name $global:VarHeaderName).Value + "' -ForegroundColor Yellow")
    Invoke-Expression -Command $FunctionTaskNames[[int]$FunctionTaskNames.IndexOf(($MyInvocation.MyCommand).Name) + 1]
}
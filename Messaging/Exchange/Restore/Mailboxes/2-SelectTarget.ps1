Function Messaging-Exchange-Restore-Mailboxes-2-SelectTarget {
    # ============
    # Declarations
    # ============
    $Task = "Select a destination mailbox for the restore mail-items"
    # =========
    # Execution
    # =========
    Script-Module-SetHeaders -CurrentTask $Task -Name ($MyInvocation.MyCommand).Name
    Script-Index-Objects -CurrentTask $Task -FunctionName ($MyInvocation.MyCommand).Name -Type "Mailbox" -Functie "Selecteren"
    $TargetName = $global:Array.Name
    $TargetAlias = $global:Array.Alias
    # ==========
    # Finalizing
    # ==========
    Set-Variable $global:VarHeaderName -Value $global:TargetName
    Set-Variable $global:VarHeaderName -Value ("Write-Host -NoNewLine ('- ' + '$Task' + ': ') -ForegroundColor Gray; Write-Host '" + (Get-Variable -Name $global:VarHeaderName).Value + "' -ForegroundColor Yellow")
    Invoke-Expression -Command $FunctionTaskNames[[int]$FunctionTaskNames.IndexOf(($MyInvocation.MyCommand).Name) + 1]
}
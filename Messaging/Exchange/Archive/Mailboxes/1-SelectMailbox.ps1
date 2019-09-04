Function Messaging-Exchange-Archive-Mailboxes-1-SelectMailbox ([array]$Objecten) {
    # ============
    # Declarations
    # ============
    $Task = "Select mailbox(es) to archive/delete"
    # =========
    # Execution
    # =========
    Script-Module-SetHeaders -CurrentTask $Task -Name ($MyInvocation.MyCommand).Name
    Script-Module-ImportMicrosoft -Name ($MyInvocation.MyCommand).Name -Module "Exchange"
    Script-Module-SetHeaders -CurrentTask $Task -Name ($MyInvocation.MyCommand).Name
    If (!$Objecten) {
        Script-Index-Objects -CurrentTask $Task -FunctionName ($MyInvocation.MyCommand).Name -Type "MailboxBulk" -Functie "Markeren"
    }
    Else {
        Script-Index-Objects -CurrentTask $Task -FunctionName ($MyInvocation.MyCommand).Name -Objecten $Objecten -Type "MailboxBulk" -Functie "Markeren"
    }
    $global:Mailboxes = $global:Array
    If ($Mailboxes.Count -gt 1) {
        $global:AantalMailboxes = $global:Aantal
    }
    Else {
        $global:AantalMailboxes = $Mailboxes.Name
    }
    $global:TitelMailboxes = $AantalMailboxes
    if ($SelectieSizeFormat) {
        $global:TitelMailboxes += ": $SelectieSizeFormat"
    }
    # ==========
    # Finalizing
    # ==========
    Set-Variable $global:VarHeaderName -Value $global:TitelMailboxes
    Set-Variable $global:VarHeaderName -Value ("Write-Host -NoNewLine ('- ' + '$Task' + ': ') -ForegroundColor Gray; Write-Host '" + (Get-Variable -Name $global:VarHeaderName).Value + "' -ForegroundColor Yellow")
    Invoke-Expression -Command $FunctionTaskNames[[int]$FunctionTaskNames.IndexOf(($MyInvocation.MyCommand).Name) + 1]
}
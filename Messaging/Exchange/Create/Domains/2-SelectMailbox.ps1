Function Messaging-Exchange-Create-Domains-2-SelectMailbox {
    # ============
    # Declarations
    # ============
    $Task = "Select a mailbox to apply the mail domain(s)"
    # =========
    # Execution
    # =========
    Script-Module-SetHeaders -CurrentTask $Task -Name ($MyInvocation.MyCommand).Name
    Script-Index-Objects -CurrentTask $Task -FunctionName ($MyInvocation.MyCommand).Name -Type "Mailbox" -Functie "Selecteren"
    $global:NameEmaildomein = $Emaildomein
    $global:Email = $Email.Split('@')[0] + "@" + $Emaildomein
    # ==========
    # Finalizing
    # ==========
    Set-Variable $global:VarHeaderName -Value ($Name + $Customer)
    Set-Variable $global:VarHeaderName -Value ("Write-Host -NoNewLine ('- ' + '$Task' + ': ') -ForegroundColor Gray; Write-Host '" + (Get-Variable -Name $global:VarHeaderName).Value + "' -ForegroundColor Yellow")
    Invoke-Expression -Command $FunctionTaskNames[[int]$FunctionTaskNames.IndexOf(($MyInvocation.MyCommand).Name) + 1]
}
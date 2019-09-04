Function Messaging-Exchange-Move-Mailboxes-6-SelectMailbox {
    # ============
    # Declarations
    # ============
    $Task = "Select the source mailbox(es) which you want to move"
    # =========
    # Execution
    # =========
    Script-Module-SetHeaders -CurrentTask $Task -Name ($MyInvocation.MyCommand).Name
    If ($ServerSource -eq "Office 365 (Exchange Online)") {
        Script-Connect-Office365 -Module "ExchangeOnline" -Name ($MyInvocation.MyCommand).Name -Task $Task
    }
    Write-Host
    Script-Index-Objects -CurrentTask $Task -FunctionName ($MyInvocation.MyCommand).Name -Type "MailboxExportBulk" -Functie "Markeren"
    $global:Mailboxes = $global:Array
    # ==========
    # Finalizing
    # ==========
    Set-Variable $global:VarHeaderName -Value ($OUFormat + " " + [string]$global:Mailboxes.Count + " mailbox(es)" + $global:SelectieSizeFormat)
    Set-Variable $global:VarHeaderName -Value ("Write-Host -NoNewLine ('- ' + '$Task' + ': ') -ForegroundColor Gray; Write-Host '" + (Get-Variable -Name $global:VarHeaderName).Value + "' -ForegroundColor Yellow")
    Invoke-Expression -Command $FunctionTaskNames[[int]$FunctionTaskNames.IndexOf(($MyInvocation.MyCommand).Name) + 1]
}
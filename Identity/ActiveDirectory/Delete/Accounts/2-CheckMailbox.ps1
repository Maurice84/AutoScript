Function Identity-ActiveDirectory-Delete-Accounts-2-CheckMailbox {
    # ============
    # Declarations
    # ============
    $Task = "Select mailbox(es) to archive"
    # =========
    # Execution
    # =========
    Script-Module-SetHeaders -CurrentTask $Task -Name ($MyInvocation.MyCommand).Name
    $CheckMailbox = $global:ArrayAccounts | Where-Object { $_.msExchRecipientTypeDetails }
    If ($CheckMailbox) {
        Write-Host
        Do {
            Write-Host -NoNewLine "  Caution:" -ForegroundColor Red; Write-Host -NoNewLine " There are mailbox(es) detected which are not archived, would you like to archive these? (Y/N): "
            [string]$InputChoice = Read-Host
            $Inputkey = @("Y", "N") -contains $InputChoice
            If (!$Inputkey) {
                Write-Host "  Please use the letters above as input" -ForegroundColor Red
            }
        } Until ($Inputkey)
        Switch ($InputChoice) {
            "Y" {
                Messaging-Exchange-Archive-Mailboxes-1-SelectMailbox -Objecten $CheckMailbox
            }
        }
    }
    Invoke-Expression -Command $FunctionTaskNames[[int]$FunctionTaskNames.IndexOf(($MyInvocation.MyCommand).Name) + 1]
}
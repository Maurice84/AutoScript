Function Messaging-Exchange-Create-Mailboxes-5-SetPermissionsAccount {
    # ============
    # Declarations
    # ============
    $Task = "Select domain account(s) to grant Full Access permissions to the mailbox"
    # =========
    # Execution
    # =========
    Script-Module-SetHeaders -CurrentTask $Task -Name ($MyInvocation.MyCommand).Name
    Do {
        Write-Host -NoNewLine "Would you like to grant Full Access permissions to this mailbox? (Y/N): "
        $Choice = Read-Host
        $Input = @("Y", "N") -contains $Choice
        If (!$Input) {
            Write-Host "Please use the letters above as input" -ForegroundColor Red
        }
    } Until ($Input)
    Switch ($Choice) {
        "Y" {
            Script-Module-SetHeaders -CurrentTask $Task -Name ($MyInvocation.MyCommand).Name
            Script-Index-Objects -CurrentTask $Task -FunctionName ($MyInvocation.MyCommand).Name -Type "AccountsBulk" -Functie "Markeren"
            [array]$global:RechtenArray = $global:Array
            $global:TitelRechtenAccounts = [string]$RechtenArray.Length + " account(s)"
            $global:MailboxRechten = "Ja"
        }
        "N" {
            $global:TitelRechtenAccounts = "n.a."
            $global:MailboxRechten = "Nee"
        }
    }
    # ==========
    # Finalizing
    # ==========
    Set-Variable $global:VarHeaderName -Value $global:TitelRechtenAccounts
    Set-Variable $global:VarHeaderName -Value ("Write-Host -NoNewLine ('- ' + '$Task' + ': ') -ForegroundColor Gray; Write-Host '" + (Get-Variable -Name $global:VarHeaderName).Value + "' -ForegroundColor Yellow")
    if ($global:MailboxRechten -eq "Ja") {
        Invoke-Expression -Command $FunctionTaskNames[[int]$FunctionTaskNames.IndexOf(($MyInvocation.MyCommand).Name) + 1]
    }
    else {
        Invoke-Expression -Command $FunctionTaskNames[[int]$FunctionTaskNames.IndexOf(($MyInvocation.MyCommand).Name) + 2]
    }
}
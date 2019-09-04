Function Messaging-Exchange-Archive-Mailboxes-5-Start {
    # =========
    # Execution
    # =========
    Script-Module-SetHeaders -Name ($MyInvocation.MyCommand).Name
    If ($Drive) {
        Write-Host -NoNewLine "- Export status of selected mailbox(es): " -ForegroundColor Gray; Write-Host ($global:AantalMailboxes + ": succesvol geëxporteerd") -ForegroundColor Yellow
    }
    Do {
        Write-Host
        If ($Drive) {
            Write-Host "  Would you like to delete the mailbox(es) due to successful export(s)?"
        }
        Else {
            Write-Host -NoNewLine "  WARNING:" -ForegroundColor Red; Write-Host -NoNewLine " Are you sure you want to delete the selected mailbox(es) "; Write-Host -NoNewLine "WITHOUT EXPORT" -ForegroundColor Red; Write-Host "?"
        }
        Write-Host -NoNewLine "  Use "; Write-Host -NoNewLine "N" -ForegroundColor Yellow; Write-Host -NoNewLine " to return to the previous menu (Y/N): "
        [string]$Choice = Read-Host
        $Input = @("Y", "N") -contains $Choice
        If (!$Input) {
            Write-Host "Please use the letters above as input" -ForegroundColor Red
        }
    } Until ($Input)
    Switch ($Choice) {
        "Y" {
            Write-Host
            ForEach ($Mailbox in $Mailboxes) {
                $MailboxSamAccountName = $Mailbox.SamAccountName
                $MailboxAlias = $Mailbox.Alias
                $ExtensionAttribute1 = Script-Connect-Server -Module "ActiveDirectory" -Command ([scriptblock]::Create(`
                            "(Get-ADUser '$MailboxSamAccountName' -Properties * -Server '$PDC' | Select-Object extensionAttribute1).extensionAttribute1")
                )
                Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("Disable-Mailbox '$MailboxAlias' -DomainController '$PDC' -Confirm:0"))
                If ($ExtensionAttribute1.Length -eq 2) {
                    Script-Connect-Server -Module "ActiveDirectory" -Command ([scriptblock]::Create("Set-ADUser '$MailboxSamAccountName' -Server '$PDC' -Add @{extensionAttribute1='$ExtensionAttribute1'}"))
                }
                If ($?) {
                    Write-Host -NoNewLine ("  " + $Mailbox.Name) -ForegroundColor Yellow; Write-Host ": Mailbox successfully deleted" -ForegroundColor Green
                }
                Else {
                    Write-Host -NoNewLine ("  ERROR: An error occurred deleting the mailbox " + $Mailbox.Name + ", please investigate!") -ForegroundColor Magenta
                }
            }
        }
        "N" {
            Get-Variable -Exclude $global:StartupVariables | Remove-Variable -ErrorAction SilentlyContinue; Script-Disconnect-Server; Invoke-Expression -Command $global:MenuNameCategory
        }
    }
    Do {
        Write-Host
        Write-Host -NoNewLine "  Would you also like to delete the corresponding domain account(s)? (Y/N): "
        [string]$Choice = Read-Host
        $Input = @("Y", "N") -contains $Choice
        If (!$Input) {
            Write-Host "Please use the letters above as input" -ForegroundColor Red
        }
    } Until ($Input)
    Switch ($Choice) {
        "Y" {
            $FunctionDeleteAccount = ($global:FunctionTaskNames | Where-Object { $_ -like "*Delete-Accounts-*-SelectAccount" })
            If ($ArrayAccounts) {
                ForEach ($Mailbox in $global:Mailboxes) {
                    ($global:ArrayAccounts | Where-Object { $_.DisplayName -eq $Mailbox.Name }).msExchRecipientTypeDetails = $null
                }
            }
            Else {
                Invoke-Expression -Command ("$FunctionDeleteAccount -Objecten $global:Mailboxes")
            }
            # ----------------------------------------------------------------------
            # Go to function ActiveDirectory > Delete Accounts to delete the account
            # ----------------------------------------------------------------------
            Invoke-Expression -Command $FunctionDeleteAccount
        }
        "N" {
            Get-Variable -Exclude $global:StartupVariables | Remove-Variable -ErrorAction SilentlyContinue; Script-Disconnect-Server; Invoke-Expression -Command $global:MenuNameCategory
        }
    }
}
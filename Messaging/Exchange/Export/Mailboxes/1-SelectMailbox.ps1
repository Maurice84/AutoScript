Function Messaging-Exchange-Export-Mailboxes-1-SelectMailbox {
    $global:Titel = ($MyInvocation.MyCommand).Name
    Script-Module-SetHeaders -Name $Titel
    Script-Module-ImportMicrosoft -Name ($MyInvocation.MyCommand).Name -Module "Exchange"
    Script-Module-SetHeaders -Name $Titel
    Script-Index-Objects -CurrentTask $Task -FunctionName ($MyInvocation.MyCommand).Name -Type "MailboxExportBulk" -Functie "Markeren"
    $global:Mailboxes = $global:Array
    $global:TitelMailboxes = { Write-Host -NoNewLine "- Entered OU filter: " -ForegroundColor Gray;
        Write-Host -NoNewLine ($OUFormat + " (") -ForegroundColor Yellow;
        If ($Mailboxes.Count -ge 2) {
            Write-Host -NoNewLine ([string]$Mailboxes.Count + " mailboxes") -ForegroundColor Yellow;
        }
        Else {
            Write-Host -NoNewLine "1 mailbox" -ForegroundColor Yellow;
        };

        If ($global:SelectieSizeFormat) {
            Write-Host -NoNewLine (": " + $global:SelectieSizeFormat + ")") -ForegroundColor Yellow;
        }
        Else {
            Write-Host -NoNewLine ")" -ForegroundColor Yellow;
        };
        Write-Host;
    }
    Messaging-Exchange-Export-Mailboxes-2-SetDate
}
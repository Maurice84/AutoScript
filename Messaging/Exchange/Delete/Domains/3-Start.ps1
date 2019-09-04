Function Messaging-Exchange-Delete-Domains-3-Start {
    # =========
    # Execution
    # =========
    Script-Module-SetHeaders -Name ($MyInvocation.MyCommand).Name
    # ------------------------------------------------------------------------------------------------------------
    Write-Host -NoNewLine "Step 1/2 - Removing mail domain from all mailboxes" -ForegroundColor Magenta; Write-Host
    # ------------------------------------------------------------------------------------------------------------
    $Mailboxes = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("Get-Mailbox -DomainController '$PDC'"))
    ForEach ($Mailbox in ($Mailboxes | Where-Object { $_.DisplayName -notlike "*Discovery*" })) {
        $EmailAddresses = ($Mailbox.EmailAddresses | Where-Object { $_ -like "*$Emaildomein*" })
        If ($EmailAddresses) {
            $MailboxSamAccountName = $Mailbox.SamAccountName
            $MailboxName = $Mailbox.Name
            If (($CheckEmailAddressPolicy | Select-Object EmailAddressPolicyEnabled).EmailAddressPolicyEnabled -eq $true) {
                $AddEmailAddressPolicy = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("Set-Mailbox -Identity '$MailboxSamAccountName' -EmailAddressPolicyEnabled 0 -DomainController '$PDC'"))
                $CheckEmailAddressPolicy = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("Get-Mailbox -Identity '$MailboxSamAccountName' -DomainController '$PDC'"))
                If (($CheckEmailAddressPolicy | Select-Object EmailAddressPolicyEnabled).EmailAddressPolicyEnabled -eq $false) {
                    Write-Host " - OK: Email Address Policy on $MailboxName removed successfully"
                }
                Else {
                    Write-Host " - ERROR: An error occurred removing Email Address Policy on $MailboxName, please investigate!" -ForegroundColor Red
                    $Pause.Invoke()
                    Get-Variable -Exclude $global:StartupVariables | Remove-Variable -ErrorAction SilentlyContinue; Script-Disconnect-Server; Invoke-Expression -Command $global:MenuNameCategory
                }
            }
            Else {
                Write-Host " - INFO: Email Address Policy on $MailboxName already removed" -ForegroundColor Gray
            }
            ForEach ($EmailAddress in $EmailAddresses) {
                If ($EmailAddress.SmtpAddress) {
                    $EmailAddress = $EmailAddress.SmtpAddress
                }
                Else {
                    $EmailAddress = $EmailAddress.Split(':')[1]
                }
                $EmailAddress = $EmailAddress.Replace('smtp:', '')
                $RemoveEmailAddress = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("Set-Mailbox -Identity '$MailboxSamAccountName' -EmailAddresses @{Remove='$EmailAddress'} -DomainController '$PDC'"))
                $CheckEmailAddress = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("Get-Mailbox -Identity '$MailboxSamAccountName' -DomainController '$PDC'"))
                If ($Mailbox.EmailAddresses | Where-Object { $_.SmtpAddress -ne "$EmailAddress" }) {
                    Write-Host -NoNewLine " - OK: Email Address "; Write-Host -NoNewLine $EmailAddress -ForegroundColor Yellow; Write-Host " successfully removed"
                }
                Else {
                    Write-Host " - ERROR: An error occurred removing Email Address $EmailAddress, please investigate!" -ForegroundColor Red
                    $Pause.Invoke()
                    Get-Variable -Exclude $global:StartupVariables | Remove-Variable -ErrorAction SilentlyContinue; Script-Disconnect-Server; Invoke-Expression -Command $global:MenuNameCategory
                }
            }
        }
    }
    Write-Host
    # -----------------------------------------------------------------------------------------
    Write-Host -NoNewLine "Step 2/2 - Deleting mail domain" -ForegroundColor Magenta; Write-Host
    # -----------------------------------------------------------------------------------------
    $RemoveDomain = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("Remove-AcceptedDomain -Identity '$Emaildomein' -Confirm:0 -DomainController '$PDC'"))
    $CheckDomain = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("Get-AcceptedDomain -DomainController '$PDC'"))
    If ($CheckDomain | Where-Object { $_.DomainName -eq $Emaildomein }) {
        Write-Host -NoNewLine " - OK: Mail domain "; Write-Host -NoNewLine $Emaildomein -ForegroundColor Yellow; Write-Host " deleted successfully"
    }
    Else {
        Write-Host " - ERROR: An error occurred deleting mail domain $Emaildomein, please investigate!" -ForegroundColor Red
        $Pause.Invoke()
        Get-Variable -Exclude $global:StartupVariables | Remove-Variable -ErrorAction SilentlyContinue; Script-Disconnect-Server; Invoke-Expression -Command $global:MenuNameCategory
    }
    Write-Host
    # ==========
    # Finalizing
    # ==========
    $Pause.Invoke()
    Get-Variable -Exclude $global:StartupVariables | Remove-Variable -ErrorAction SilentlyContinue; Script-Disconnect-Server; Invoke-Expression -Command $global:MenuNameCategory
}
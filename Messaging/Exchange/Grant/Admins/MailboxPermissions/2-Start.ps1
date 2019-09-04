Function Messaging-Exchange-Grant-Admins-MailboxPermissions-2-Start {
    # =========
    # Execution
    # =========
    Script-Module-SetHeaders -Name ($MyInvocation.MyCommand).Name
    $Counter = 1
    $User = $env:userdomain + "\" + $GroupAdmins
    ForEach ($Mailbox in $Mailboxes) {
        $Name = $Mailbox.Name
        $SamAccountName = $Mailbox.SamAccountName
        $DistinguishedName = $Mailbox.DistinguishedName
        Write-Host -NoNewLine ("Granting Full Access permissions for the Administrator group on mailbox " + $Counter + " of " + $Mailboxes.Count + ": "); Write-Host -NoNewLine $Name -ForegroundColor Yellow; Write-Host -NoNewLine "..."
        $GetFullAccessDeny = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("Get-MailboxPermission -Identity '$SamAccountName'"))
        $GetFullAccessDeny = $GetFullAccessDeny | Where-Object { $_.User -like "*$env:username" -AND $_.AccessRights -like "*FullAccess*" -AND $_.Deny -eq $true } | Select-Object User, @{Name = "AccessRight"; Expression = { $_.AccessRights -join "," } }, Deny
        If ($GetFullAccessDeny | Where-Object { $_.User -eq $User }) {
            $RemoveDeny = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create(`
                        "Remove-MailboxPermission -Identity '$SamAccountName' -User '$User' -AccessRights FullAccess -DomainController '$PDC' -Confirm:0")
            )
        }
        $AddMailboxPermission = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("Add-MailboxPermission -Identity '$SamAccountName' -AccessRights FullAccess -User '$User' -InheritanceType All -AutoMapping 0 -DomainController '$PDC' -Confirm:0 -WarningAction SilentlyContinue"))
        $CheckMailboxPermission = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("Get-MailboxPermission -Identity '$SamAccountName' -User '$User' -DomainController '$PDC'"))
        If ($CheckMailboxPermission | Select-Object User, @{Name = "AccessRight"; Expression = { $_.AccessRights -join "," } }) {
            Write-Host -NoNewLine " - OK: Mailbox FullAccess permission on "; Write-Host -NoNewLine $Name -ForegroundColor Yellow; Write-Host -NoNewLine " added successfully"; Write-Host
        }
        Else {
            Write-Host " - ERROR: An error occurred adding the FullAccess permission on $Name, please investigate!" -ForegroundColor Red
            $Problem = $true
        }
        $AddADPermission = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("Add-ADPermission -Identity '$DistinguishedName' -ExtendedRights Send-As -User '$User' -DomainController '$PDC' -WarningAction SilentlyContinue"))
        $CheckSendAs = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("Get-ADPermission -Identity '$DistinguishedName' -User '$User' -DomainController '$PDC'"))
        If ($CheckSendAs | Select-Object ExtendedRights) {
            Write-Host -NoNewLine " - OK: Mailbox Send-As permission on "; Write-Host -NoNewLine $Name -ForegroundColor Yellow; Write-Host -NoNewLine " added successfully"; Write-Host
        }
        Else {
            Write-Host " - ERROR: An error occurred adding the Send-As permission on $Name, please investigate!" -ForegroundColor Red
            $Problem = $true
        }
        If ($Problem -eq $true) {
            $Pause.Invoke()
            Get-Variable -Exclude $global:StartupVariables | Remove-Variable -ErrorAction SilentlyContinue; Script-Disconnect-Server; Invoke-Expression -Command $global:MenuNameCategory
        }
        Else {
            Write-Host
            $Counter++
        }
    }
    Write-Host
    # ==========
    # Finalizing
    # ==========
    $Pause.Invoke()
    Get-Variable -Exclude $global:StartupVariables | Remove-Variable -ErrorAction SilentlyContinue; Script-Disconnect-Server; Invoke-Expression -Command $global:MenuNameCategory
}
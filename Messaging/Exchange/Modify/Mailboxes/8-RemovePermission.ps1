Function Messaging-Exchange-Modify-Mailboxes-8-RemovePermission ([string]$Soort) {
    # =========
    # Execution
    # =========
    $FunctionMailboxMenu = ($global:FunctionTaskNames | Where-Object { $_ -like "Messaging*Modify-Mailboxes-*-Menu" })
    $FunctionMailboxRemovePermission = ($global:FunctionTaskNames | Where-Object { $_ -like "Messaging*Modify-Mailboxes-*-RemovePermission" })
    Script-Index-Objects -CurrentTask $Task -FunctionName ($MyInvocation.MyCommand).Name -Type "Rechten" -Functie "Selecteren" -Objecten $global:Objects
    $global:SelectedUser = $global:Array.Name
    Do {
        Write-Host -NoNewLine "- Selected mailbox: "; Write-Host ($Mailbox.DisplayName + " (" + $Soort + " permissions)") -ForegroundColor Yellow
        Write-Host
        Write-Host -NoNewLine "  Are you sure you want to remove the "; Write-Host -NoNewLine $Soort -ForegroundColor Magenta; Write-Host -NoNewLine " permission for "; Write-Host -NoNewLine $SelectedUser -ForegroundColor Yellow; Write-Host -NoNewLine "? Use "; Write-Host -NoNewLine "N" -ForegroundColor Yellow; Write-Host -NoNewLine " to cancel (Y/N): "
        [string]$Choice = Read-Host
        $Input = @("Y","N") -contains $Choice
        If (!$Input) {
            Write-Host "Please use the letters above as input" -ForegroundColor Red
        }
    } Until ($Input)
    Switch ($Choice) {
        "Y" {
            Write-Host
            Write-Host -NoNewLine "  Removing $Soort permission for "; Write-Host -NoNewLine $SelectedUser -ForegroundColor Yellow; Write-Host -NoNewLine "..."
            If ($Soort -eq "FullAccess") {
                $RemoveMailboxPermission = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("Remove-MailboxPermission -Identity '$SamAccountName' -AccessRights FullAccess -User '$SelectedUser' -DomainController '$PDC' -Confirm:0 -WarningAction SilentlyContinue"))
                $CheckMailboxPermission = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("Get-MailboxPermission -Identity '$SamAccountName' -User '$SelectedUser' -DomainController '$PDC'"))
                If ($CheckMailboxPermission | Select-Object User,@{Name="AccessRight";Expression={$_.AccessRights -join ","}}) {
                    $global:Fout = $true
                } Else {
                    [array]$global:FullAccess = $global:FullAccess | Where-Object {$_ -ne $SelectedUser}
                    If ($global:FullAccess.Length -ge 6) {
                        $global:VolledigeToegang = [string]$global:FullAccess.Length + " domain accounts/groups"
                    }
                    If ($global:FullAccess.Length -lt 5) {
                        $global:VolledigeToegang = $global:FullAccess -join ', '
                    }
                    If (!$global:FullAccess) {
                        $global:VolledigeToegang = "n.a."
                    }
                }
            }
            If ($Soort -eq "SendAs") {
                $RemoveADPermission = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("Remove-ADPermission -Identity '$DistinguishedName' -ExtendedRights Send-As -User '$SelectedUser' -DomainController '$PDC' -Confirm:0 -WarningAction SilentlyContinue"))
                $CheckSendAs = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("Get-ADPermission -Identity '$DistinguishedName' -User '$SelectedUser' -DomainController '$PDC'"))
                If ($CheckSendAs | Select-Object ExtendedRights) {
                    $Fout = $true
                } Else {
                    [array]$global:SendAs = $global:SendAs | Where-Object {$_ -ne $SelectedUser}
                    If ($global:SendAs.Length -ge 6) {
                        $global:VerzendenAls = [string]$global:SendAs.Length + " domain accounts/groups"
                    } 
                    If ($global:SendAs.Length -lt 5) {
                        $global:VerzendenAls = $global:SendAs -join ', '
                    } 
                    If (!$global:SendAs) {
                        $global:VerzendenAls = "n.a."
                    }
                }
            }
            If ($Soort -eq "SendOnBehalf") {
                $RemoveSendOnBehalf = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("Set-Mailbox -Identity '$SamAccountName' -GrantSendOnBehalfTo @{remove='$SelectedUser'} -DomainController '$PDC' -Confirm:0 -WarningAction SilentlyContinue"))
                $CheckSendOnBehalf = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("Get-Mailbox -Identity '$SamAccountName' -DomainController '$PDC'"))
                $GetSendOnBehalf = ($CheckSendOnBehalf | Select-Object @{Name="GrantSendOnBehalfTo";Expression={$_.GrantSendOnBehalfTo}}).GrantSendOnBehalfTo
                ForEach($SendOnBehalf in $GetSendOnBehalf) {
                    If ($SendOnBehalf -like "*$SelectedUser*") {
                        $global:Fout = $true
                    }
                }
                If (!$Fout) {
                    [array]$global:SendOnBehalf = $global:SendOnBehalf | Where-Object {$_ -ne $SelectedUser}
                    If ($global:SendOnBehalf.Length -ge 6) {
                        $global:VerzendenNamens = [string]$global:SendOnBehalf.Length + " domain accounts/groups"
                    } 
                    If ($global:SendOnBehalf.Length -lt 5) {
                        $global:VerzendenNamens = $global:SendOnBehalf -join ', '
                    }
                    If (!$global:SendOnBehalf) {
                        $global:VerzendenNamens = "n.a."
                    }
                }
            }
            If (!$Fout) {
                Script-Module-ReplicateAD
                Write-Host; Write-Host -NoNewLine "  OK: $Soort permission for "; Write-Host -NoNewLine $SelectedUser -ForegroundColor Yellow; Write-Host " removed successfully"
                [array]$global:Objects = $global:Objects | Where-Object {$_ -ne $SelectedUser}
                If (!$global:Objects) {
                    Invoke-Expression -Command $FunctionMailboxMenu
                } Else {
                    Invoke-Expression -Command ("$FunctionMailboxRemovePermission -Soort $Soort")
                }
            } Else {
                Write-Host "ERROR: An error occurred removing the $Soort permission, please investigate!" -ForegroundColor Red
                $Pause.Invoke()
                Invoke-Expression -Command $FunctionMailboxMenu
            }
        }
        "N" {
            Invoke-Expression -Command ("$FunctionMailboxRemovePermission -Soort $Soort")
        }
    }
}
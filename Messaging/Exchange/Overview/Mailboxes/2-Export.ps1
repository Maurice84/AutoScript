Function Messaging-Exchange-Overview-Mailboxes-2-Export {
    # =========
    # Execution
    # =========
    Script-Module-SetHeaders -Name ($MyInvocation.MyCommand).Name
    Write-Host "Exporting mailbox(es) to CSV-file..." -ForegroundColor Magenta
    # ------------------------------------------------------------------------------------------------------------------
    # Declare the name and path of the CSV-file and remove the CSV-file if it's present (Add-Content does not overwrite)
    # ------------------------------------------------------------------------------------------------------------------
    If ($global:OUFilter -eq "*") {
        If (!$Office365Exchange) {
            $Domain = (Get-WmiObject Win32_ComputerSystem).Domain
        }
        Else {
            $Domain = (($global:Credential).UserName).Split('@')[1]
        }
        $FileCSVM = "C:\Mailboxes-" + $Domain + "_" + (Get-Date -Format "yyyyMMdd-HHmm") + ".csv"
    }
    Else {
        $FileCSVM = "C:\Mailboxes-" + $global:OUFilter + "_" + (Get-Date -Format "yyyyMMdd-HHmm") + ".csv"
    }
    Remove-Item ($FileCSVM) -ErrorAction SilentlyContinue
    # -------------------------------------------------------------------------------------
    # Iterate through each account to index the mail/forwarding address(es) and permissions
    # -------------------------------------------------------------------------------------
    $EmailAddresses = @()
    $FullAccess = @()
    $SendAs = @()
    $SendOnBehalf = @()
    $TotalColumnsEmailAddress = 0
    $TotalColumnsFullAccess = 0
    $TotalColumnsSendAs = 0
    ForEach ($Mailbox in ($global:Array | Sort-Object Name)) {
        # -------------------------------
        # Indexing extra mail address(es)
        # -------------------------------
        $Count = 0
        If ($Mailbox.EmailAddresses) {
            ForEach ($Address in $Mailbox.EmailAddresses) {
                If (!($Address -clike "SMTP:*")) {
                    If ($Address -like "smtp:*" -AND $Address -notlike "*mail.onmicrosoft.com") {
                        If ($Office365Exchange) {
                            $TenantAddress = ($Mailbox.EmailAddresses | Where-Object { $_ -notlike "SIP:*" -AND $_ -like "*.onmicrosoft.com" -AND $_ -notlike "*mail.onmicrosoft.com" } | Select-Object -First 1) #.Split(':')[1]
                            If ($TenantAddress) {
                                $TenantAddress = $TenantAddress.Split(':')[1]
                                If ($Address -notlike "*$TenantAddress") {
                                    $EmailAddress = [string]$Address.Split(':')[1].ToLower()
                                }
                                Else {
                                    $EmailAddress = $null
                                }
                            } Else {
                                $EmailAddress = $null
                            }
                        }
                        Else {
                            [string]$EmailAddress = $Address
                            $EmailAddress = $EmailAddress.Split(':')[1].ToLower()
                        }
                        If ($EmailAddress) {
                            $Properties = @{UserPrincipalName = $Mailbox.UserPrincipalName; EmailAddress = $EmailAddress }
                            $ObjectEmailAddress = New-Object PSObject -Property $Properties
                            $EmailAddresses += $ObjectEmailAddress
                            $Count++
                            If ($Count -gt $TotalColumnsEmailAddress) {
                                $TotalColumnsEmailAddress = $Count
                            }
                        }
                    }
                }
            }
        }
        # ------------------------------------
        # Indexing the Full Access permissions
        # ------------------------------------
        $Count = 0
        If ($Mailbox.FullAccess) {
            ForEach ($User in ($Mailbox.FullAccess | Sort-Object FullAccess)) {
                $Properties = @{UserPrincipalName = $Mailbox.UserPrincipalName; User = $User; Permission = "FullAccess" }
                $ObjectFullAccess = New-Object PSObject -Property $Properties
                $FullAccess += $ObjectFullAccess
                $Count++
                If ($Count -gt $TotalColumnsFullAccess) {
                    $TotalColumnsFullAccess = $Count
                }
            }
        }
        # --------------------------------
        # Indexing the Send As permissions
        # --------------------------------
        $Count = 0
        If ($Mailbox.SendAs) {
            ForEach ($User in ($Mailbox.SendAs | Sort-Object SendAs)) {
                $Properties = @{UserPrincipalName = $Mailbox.UserPrincipalName; User = $User; Permission = "SendAs" }
                $ObjectSendAs = New-Object PSObject -Property $Properties
                $SendAs += $ObjectSendAs
                $Count++
                If ($Count -gt $TotalColumnsSendAs) {
                    $TotalColumnsSendAs = $Count
                }
            }
        }
        # ---------------------------------------
        # Indexing the Send On Behalf permissions
        # ---------------------------------------
        $Count = 0
        If ($Mailbox.SendOnBehalf) {
            ForEach ($User in ($Mailbox.SendOnBehalf | Sort-Object SendOnBehalf)) {
                $Properties = @{UserPrincipalName = $Mailbox.UserPrincipalName; User = $User; Permission = "SendOnBehalf" }
                $ObjectSendOnBehalf = New-Object PSObject -Property $Properties
                $SendOnBehalf += $ObjectSendOnBehalf
                $Count++
                If ($Count -gt $TotalColumnsSendOnBehalf) {
                    $TotalColumnsSendOnBehalf = $Count
                }
            }
        }
        # -------------------------------
        # Indexing forwarding address(es)
        # -------------------------------
        $Count = 0
        If ($Mailbox.ForwardingAddress) {
            $ExtraColumnForwardingAddress = ',"ForwardingAddress"'
        }
        If ($Mailbox.DeliverToMailboxAndForward) {
            $ExtraColumnDeliverToMailboxAndForward = ',"DeliverToMailboxAndForward"'
        }
    }
    # ---------------------------------------------------
    # Declare the header and add it to the empty CSV-file
    # ---------------------------------------------------
    $ExportHeader = '"Date";"Name";'
    If (!$Office365Exchange) {
        $ExportHeader += '"SamAccountName";'
    }
    Else {    
        $ExportHeader += '"UserPrincipalName";'
    }
    $ExportHeader += '"OrganizationalUnit";"LastLogonDate";"Server";"SizeMB";"ItemCount";"Type";"Language";"Office365Guid";'
    If (!$Office365Exchange) {
        $ExportHeader += '"LegacyExchangeDN";'
    }
    Else {
        $ExportHeader += '"Office365License";"TenantAddress";'
    }
    $ExportHeader += '"PrimarySmtpAddress"'
    $ExtraColumnsEmailAddress = $null
    $ExtraColumnsFullAccess = $null
    $ExtraColumnsSendAs = $null
    $ExtraColumnsSendOnBehalf = $null
    If ($EmailAddresses) {
        For ($Count = 0; $Count -lt $TotalColumnsEmailAddress; $Count++) {
            $ExtraColumnsEmailAddress += ';"ExtraAddress' + ($Count + 1) + '"'
        }
    }
    If ($FullAccess) {
        For ($Count = 0; $Count -lt $TotalColumnsFullAccess; $Count++) {
            $ExtraColumnsFullAccess += ';"FullAccess' + ($Count + 1) + '"'
        }
    }
    If ($SendAs) {
        For ($Count = 0; $Count -lt $TotalColumnsSendAs; $Count++) {
            $ExtraColumnsSendAs += ';"SendAs' + ($Count + 1) + '"'
        }
    }
    If ($SendOnBehalf) {
        For ($Count = 0; $Count -lt $TotalColumnsSendOnBehalf; $Count++) {
            $ExtraColumnsSendOnBehalf += ';"SendOnBehalf' + ($Count + 1) + '"'
        }
    }
    Add-Content ($FileCSVM) ($ExportHeader + $ExtraColumnsEmailAddress + $ExtraColumnsFullAccess + $ExtraColumnsSendAs + $ExtraColumnsSendOnBehalf + $ExtraColumnForwardingAddress + $ExtraColumnDeliverToMailboxAndForward)
    # ----------------------------------------------------------------------------------
    # Iterate through each account with the current timestamp and add it to the CSV-file
    # ----------------------------------------------------------------------------------
    $Counter = 1
    ForEach ($Mailbox in ($global:Array | Sort-Object Name)) {
        Write-Host -NoNewLine (" - Exporting mailbox $Counter of " + $global:Array.Count + ": "); Write-Host -NoNewLine $Mailbox.Name -ForegroundColor Yellow; Write-Host "..."
        $LastCheck = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        If ($Mailbox.LastLogonDate -like "*2000*") {
            $LastLogonDate = "Never logged in"
        }
        Else {
            if ($Mailbox.LastLogonDate -ne "n.a.") {
                $LastLogonDate = Get-Date $Mailbox.LastLogonDate -Format "yyyy-MM-dd HH:mm:ss"
            }
        }
        $MailboxEmailAddresses = $EmailAddresses | Where-Object { $_.UserPrincipalName -eq $Mailbox.UserPrincipalName } | Select-Object EmailAddress | Sort-Object EmailAddress
        $MailboxFullAccess = $FullAccess | Where-Object { $_.UserPrincipalName -eq $Mailbox.UserPrincipalName } | Select-Object User | Sort-Object User
        $MailboxSendAs = $SendAs | Where-Object { $_.UserPrincipalName -eq $Mailbox.UserPrincipalName } | Select-Object User | Sort-Object User
        $MailboxSendOnBehalf = $SendOnBehalf | Where-Object { $_.UserPrincipalName -eq $Mailbox.UserPrincipalName } | Select-Object User | Sort-Object User
        If ($Office365Exchange -AND $Mailbox.EmailAddresses -like "*.onmicrosoft.com") {
            $TenantAddress = ($Mailbox.EmailAddresses | Where-Object { $_ -notlike "SIP:*" -AND $_ -like "*.onmicrosoft.com" -AND $_ -notlike "*mail.onmicrosoft.com" } | Select-Object -First 1).Split(':')[1]
        }
        Else {
            $TenantAddress = $null
        }
        If (!$Office365Exchange) {
            $MailboxServer = $Mailbox.ServerName
        }
        Else {
            $MailboxServer = "Office 365"
        }
        $ExportValues = '"' + $LastCheck + '";"' + $Mailbox.Name + '";"'
        If (!$Office365Exchange) {
            $ExportValues += $Mailbox.SamAccountName + '";"'
        }
        Else {
            $ExportValues += $Mailbox.UserPrincipalName + '";"'
        }
        $ExportValues += $Mailbox.OU + '";"' + $LastLogonDate + '";"' + $MailboxServer + '";' + $Mailbox.Size + ';' + $Mailbox.Items + ';"' + $Mailbox.RecipientTypeDetails + '";"' + $Mailbox.Language + '";"' + $Mailbox.Office365Guid + '";"'
        If (!$Office365Exchange) {
            $ExportValues += $Mailbox.LegacyExchangeDN + '";"'
        }
        Else {
            $ExportValues += $Mailbox.License + '";"' + $TenantAddress + '";"'
        }
        $ExportValues += $Mailbox.PrimarySMTPAddress.ToString() + '"'
        For ($Count = 1; $Count -le $TotalColumnsEmailAddress; $Count++) {
            If ($MailboxEmailAddresses) {
                $EmailAddress = $MailboxEmailAddresses[$Count - 1].EmailAddress
                If ($EmailAddress) {
                    $ExportValues += ';"' + $EmailAddress + '"'
                }
                Else {
                    $ExportValues += ';'
                }
            }
            Else {
                $ExportValues += ';'
            }
        }
        For ($Count = 1; $Count -le $TotalColumnsFullAccess; $Count++) {
            If ($MailboxFullAccess) {
                $User = $MailboxFullAccess[$Count - 1].User
                If ($User) {
                    $ExportValues += ';"' + $User + '"'
                }
                Else {
                    $ExportValues += ';'
                }
            }
            Else {
                $ExportValues += ';'
            }
        }
        For ($Count = 1; $Count -le $TotalColumnsSendAs; $Count++) {
            If ($MailboxSendAs) {
                $User = $MailboxSendAs[$Count - 1].User
                If ($User) {
                    $ExportValues += ';"' + $User + '"'
                }
                Else {
                    $ExportValues += ';'
                }
            }
            Else {
                $ExportValues += ';'
            }
        }
        For ($Count = 1; $Count -le $TotalColumnsSendOnBehalf; $Count++) {
            If ($MailboxSendOnBehalf) {
                $User = $MailboxSendOnBehalf[$Count - 1].User
                If ($User) {
                    $ExportValues += ';"' + $User + '"'
                }
                Else {
                    $ExportValues += ';'
                }
            }
            Else {
                $ExportValues += ';'
            }
        }
        If ($Mailbox.ForwardingAddress) {
            $ExportValues += ';"' + $Mailbox.ForwardingAddress + '"'
        }
        If ($Mailbox.DeliverToMailboxAndForward) {
            $ExportValues += ';"' + $Mailbox.DeliverToMailboxAndForward + '"'
        }
        Add-Content $FileCSVM $ExportValues
        $Counter++
    }
    If ((Test-Path $FileCSVM) -AND ($global:Array.Count -eq ($Counter - 1))) {
        Write-Host -NoNewLine " - OK: Successfully exported the mailbox(es) to CSV-file: "; Write-Host $FileCSVM -ForegroundColor Yellow
    }
    # ----------------------------------
    # Convert the CSV-file to Excel-file
    # ----------------------------------
    Script-Convert-CSV-to-Excel -File $FileCSVM -Category "Mailboxen" # -Silent $true
    $FileExcel = $FileCSVM.Substring(0, $FileCSVM.Length - 4) + ".xlsx"
    If (Test-Path $FileExcel) {
        Write-Host -NoNewLine " - OK: Successfully converted the CSV-file to Excel-file: "; Write-Host $FileExcel -ForegroundColor Yellow
    }
    Else {
        Write-Host " - ERROR: Could not convert the CSV-file to Excel-file, please investigate!" -ForegroundColor Red
    }
    # ==========
    # Finalizing
    # ==========
    $Pause.Invoke()
    Get-Variable -Exclude $global:StartupVariables | Remove-Variable -ErrorAction SilentlyContinue; Script-Disconnect-Server; Invoke-Expression -Command $global:MenuNameCategory
}
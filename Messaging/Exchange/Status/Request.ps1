Function Messaging-Exchange-Status-Request {
    param(
        [string]$Selectie,
        [string]$Type
    )
    # =========
    # Execution
    # =========
    Script-Module-Initial-5-SetSystemLocale -Locale "en-US"
    Script-Module-SetHeaders -Name ($MyInvocation.MyCommand).Name
    If ($Selectie -eq "Export") {
        $global:Actief = "Nee"
        $global:Object = "export-request"
        $global:Filter = "(Received -ge '" + $StartDate + "') -and (Received -le '" + $EndDate + "')"
        $global:FileStartDate = $StartDate.ToString("dd-MM-yyyy")
        $global:FileEndDate = $EndDate.ToString("dd-MM-yyyy")
        $global:MailboxObjects = $Mailboxes
        If (!$Office365Exchange) {
            $GetExportRequests = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("Get-MailboxExportRequest -ResultSize Unlimited -DomainController '$PDC'"))
        }
        Else {
            $GetExportRequests = Script-Connect-Server -Module "Office365" -Command ([scriptblock]::Create("Get-MailboxExportRequest -ResultSize Unlimited"))
        }
    }
    If ($Selectie -eq "Import") {
        $global:Actief = "Nee"
        $global:Object = "import-request"
        # Bij Office365 PST bestanden is de SourceRootFolder: Primary Mailbox/*
        $global:SourceRootFolder = "/*"
        $global:MailboxObjects = $Mailboxes
    }
    If ($Selectie -eq "Restore") {
        $global:Actief = "Nee"
        $global:Object = "restore-request"
        $global:MailboxObjects = $MailboxRestore
    }
    If ($global:MailboxObjects.Count) {
        [string]$global:MailboxObjectsCount = $global:MailboxObjects.Count
    }
    Else {
        [string]$global:MailboxObjectsCount = 1
    }
    $Counter = 0
    ForEach ($MailboxObject in $global:MailboxObjects) {
        $Counter++
        $global:CheckPath = $null
        $global:Status = $null
        $global:Fout = $null
        If ($CSV) {
            $MailboxObjectUPN = $MailboxObject.UserPrincipalName
            If (!$Office365Exchange) {
                $MailboxObject = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("Get-Mailbox '$MailboxObjectUPN' -DomainController '$PDC'"))
            }
            Else {
                $MailboxObject = Script-Connect-Server -Module "Office365" -Command ([scriptblock]::Create("Get-Mailbox '$MailboxObjectUPN'"))
            }
        }
        $Task = "Retrieving mailbox $Object $Counter of $global:MailboxObjectsCount"
        Script-Module-SetHeaders -CurrentTask $Task -Name ($MyInvocation.MyCommand).Name
        If ($Selectie -eq "Export") {
            $global:Name = $MailboxObject.Name
            $global:File = [string]$MailboxObject.PrimarySmtpAddress + "_" + $FileStartDate + "-tm-" + $FileEndDate + ".pst"
            $global:Path = $Drive + $File
            $global:UNC = $PathUNC + "\" + $File
            $global:Request = $global:Name + " " + $FileStartDate + "-tm-" + $FileEndDate
            $global:Email = $MailboxObject.PrimarySmtpAddress
            $global:Alias = $MailboxObject.Alias
        }
        If ($Selectie -eq "Import") {
            If (!$File) {
                $global:File = $MailboxObject.PST
                $global:Path = $MailboxObject.Name
                $global:UNC = $MailboxObject.FileUNC
                $global:Alias = $MailboxObject.Alias
                $global:Name = $MailboxObject.MailboxName
                $global:Email = $MailboxObject.PrimarySmtpAddress
            }
            $global:Email = $MailboxObject.PrimarySmtpAddress
            $global:Request = $global:Name + " " + $File
            If ($global:ExchangeVersion -like "201*" -OR $global:RemoteExchange -eq $true) {
                $global:StatusRequest = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create(`
                    "Get-MailboxImportRequest -Name '$Request' -DomainController '$PDC' | Get-MailboxImportRequestStatistics -DomainController '$PDC' -IncludeReport")
                )
            }
        }
        If ($Selectie -eq "Export" -OR $Selectie -eq "Import") {
            Write-Host -NoNewLine "> Name: "; Write-Host $global:Name -ForegroundColor Yellow
            If ($Email) {
                Write-Host -NoNewLine "> Primary Mail Address: "; Write-Host $Email -ForegroundColor Yellow
            }
            Write-Host -NoNewLine "> PST-file: "; Write-Host $Path -ForegroundColor Yellow
            If ($Selectie -eq "Export") {
                If ($global:ExchangeVersion -like "201*" -OR $global:RemoteExchange -eq $true) {
                    $CheckRequest = $GetExportRequests | Where-Object { $_.Name -like ($global:Name + " " + $FileStartDate + "-tm-*") }
                    If (($CheckRequest | Select-Object Name).Name) {
                        $ExistingRequestName = ($CheckRequest | Select-Object Name).Name
                        If (!$SkipExistingRequests) {
                            $ExistingRequestEndDate = $ExistingRequestName.Split('tm')[-1].Substring(1)
                            Write-Host
                            Do {
                                Write-Host "  An already completed mailbox export-request detected with end-date $ExistingRequestEndDate, would you like to redo?" -ForegroundColor Cyan; Write-Host -NoNewLine "Use " -ForegroundColor Cyan; Write-Host -NoNewLine "X" -ForegroundColor Yellow; Write-Host -NoNewLine " to return to the previous menu (Y/N/X): " -ForegroundColor Cyan
                                [string]$Choice = Read-Host
                                $Input = @("Y", "N", "X") -contains $Choice
                                If (!$Input) {
                                    Write-Host "  Please use the letters above as input" -ForegroundColor Red
                                }
                            } Until ($Input)
                            Switch ($Choice) {
                                "N" {
                                    $global:Request = $ExistingRequestName
                                    If ($CSV) {
                                        Do {
                                            Write-Host; Write-Host -NoNewLine "  Does this also apply for the other mailbox export-requests(s)? (Y/N): " -ForegroundColor Cyan
                                            [string]$Choice = Read-Host
                                            $Input = @("Y", "N") -contains $Choice
                                            If (!$Input) {
                                                Write-Host "  Please use the letters above as input" -ForegroundColor Red
                                            }
                                        } Until ($Input)
                                        Switch ($Choice) {
                                            "Y" {
                                                $global:SkipExistingRequests = $true
                                            }
                                        }
                                    }
                                }
                                "X" {
                                    Get-Variable -Exclude $global:StartupVariables | Remove-Variable -ErrorAction SilentlyContinue; Script-Disconnect-Server; Invoke-Expression -Command $global:MenuNameCategory
                                }
                            }
                        }
                        Else {
                            $global:Request = $ExistingRequestName
                        }
                    }
                    $global:StatusRequest = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create(`
                        "Get-MailboxExportRequest -Name '$global:Request' -DomainController '$PDC' | Get-MailboxExportRequestStatistics -DomainController '$PDC' -IncludeReport")
                    )
                }
            }
        }
        If ($Selectie -eq "Restore") {
            Write-Host -NoNewLine "> Restorable mailbox: "; Write-Host $SourceName -ForegroundColor Yellow
            Write-Host -NoNewLine "> Destination mailbox: ";
            If ($TargetFolder -ne "n.a.") {
                Write-Host ($TargetName + ": " + $TargetFolder) -ForegroundColor Yellow
            }
            Else {
                Write-Host $TargetName -ForegroundColor Yellow
            }
            $global:Request = $SourceName
        }
        Write-Host
        If (($global:ExchangeVersion -eq "2007") -OR ($StatusRequest.Status -ne "Completed" -AND $StatusRequest.Status -ne "Failed")) {
            If (($global:ExchangeVersion -eq "2007") -OR ($StatusRequest.Status -ne "Queued" -AND $StatusRequest.Status -ne "InProgress")) {
                If ($Selectie -eq "Export") {
                    Write-Host -NoNewLine "  Checking required mailbox permissions... "
                    If ($global:ExchangeVersion -like "201*" -OR $global:RemoteExchange -eq $true) {
                        $global:CheckMailbox = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("Get-MailboxPermission -Identity '$Alias'"))
                        $global:CheckMailbox = $global:CheckMailbox | Where-Object { $_.User -like "*$env:username" -AND $_.AccessRights -like "*FullAccess*" -AND $_.Deny -eq $false } | Select-Object User, @{Name = "AccessRight"; Expression = { $_.AccessRights -join "," } }, Deny
                    }
                    If ($global:ExchangeVersion -eq "2007") {
                        $global:CheckMailbox = Get-MailboxPermission -Identity $Alias | Where-Object { $_.User -like "*$env:username" -AND $_.AccessRights -like "*FullAccess*" -AND $_.Deny -eq $false }
                    }
                    If ($CheckMailbox) {
                        Write-Host -NoNewLine "OK!" -ForegroundColor Green; Write-Host
                    }
                    Else {
                        Write-Host -NoNewLine "No permissions detected" -ForegroundColor DarkGray; Write-Host
                        Write-Host -NoNewLine "  Setting Full Access permissions on the mailbox... "
                        If ($global:ExchangeVersion -like "201*" -OR $global:RemoteExchange -eq $true) {
                            $SetFullAccess = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create(`
                                "Add-MailboxPermission -Identity '$Alias' -AccessRights FullAccess -User '$env:username' -InheritanceType All -AutoMapping 0 -DomainController '$PDC' -Confirm:0 -WarningAction SilentlyContinue")
                            )
                            $GetFullAccess = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create(`
                                "Get-MailboxPermission -Identity '$Alias' -User '$env:username' -DomainController '$PDC'")
                            )
                            $GetFullAccess = $GetFullAccess | Select-Object User, @{Name = "AccessRight"; Expression = { $_.AccessRights -join "," } }
                            If (!$GetFullAccess) {
                                $SetFullAccess = "Fout"
                            }
                        }
                        If ($global:ExchangeVersion -eq "2007") {
                            Add-MailboxPermission -Identity $Alias -AccessRights FullAccess -User $env:username -InheritanceType All -WarningAction SilentlyContinue | Out-Null
                            If (!$?) {
                                $SetFullAccess = "Fout"
                            }
                        }
                        If ($SetFullAccess -eq "Fout") {
                            Write-Host
                            Write-Host "ERROR: An error occurred adding the Full Access permission on the mailbox, please investigate!" -ForegroundColor Red
                            $Pause.Invoke()
                            Get-Variable -Exclude $global:StartupVariables | Remove-Variable -ErrorAction SilentlyContinue; Invoke-Expression -Command $global:MenuNameCategory
                        }
                        Else {
                            Write-Host -NoNewLine "OK!" -ForegroundColor Green; Write-Host
                        }
                    }
                    Write-Host -NoNewLine "  Checking presence of PST-file... "
                    If (!(Test-Path $Path)) {
                        Write-Host -NoNewLine "OK!" -ForegroundColor Green; Write-Host
                    }
                    Else {
                        Do {
                            Write-Host -NoNewLine "Already exist!" -ForegroundColor Red; Write-Host
                            Write-Host -NoNewLine "  Would you like to overwrite this PST-file? Use " -ForegroundColor Magenta; Write-Host -NoNewLine "N" -ForegroundColor Yellow; Write-Host -NoNewLine " to skip this mailbox (Y/N): " -ForegroundColor Magenta
                            [string]$Choice = Read-Host
                            $Input = @("Y", "N") -contains $Choice
                            If (!$Input) { Write-Host "Please use the letters above as input" -ForegroundColor Red }
                        } Until ($Input)
                        Switch ($Choice) {
                            "N" {
                                $SkipOverwrite = $true
                                If ($global:ExchangeVersion -eq "2007") {
                                    $Status = "Completed"
                                }
                            }
                        }
                    }
                    If (!$SkipOverwrite) {
                        If ($global:ExchangeVersion -like "201*" -OR $global:RemoteExchange -eq $true) {
                            Write-Host -NoNewLine "  Initializing the mailbox $Object..."
                            If (!$RemoteExchange) {
                                $NewRequest = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create(`
                                    "New-MailboxExportRequest -Mailbox '$Alias' -Name '$Request' -DomainController '$PDC' -Filepath '$UNC' -ContentFilter {$Filter} -BadItemLimit 999 -AcceptLargeDataLoss -ErrorAction SilentlyContinue -WarningAction SilentlyContinue | Out-Null")
                                )
                            }
                            Else {
                                If (!$ExchangePSSession) {
                                    Import-PSSession -Session (Get-PSSession -Name 'Exchange') -DisableNameChecking | Out-Null
                                    $ExchangePSSession = $true
                                }
                                $NewRequest = New-MailboxExportRequest -Mailbox $Alias -Name "$Request" -DomainController $PDC -Filepath $UNC -ContentFilter { $Filter } -BadItemLimit 999 -AcceptLargeDataLoss -ErrorAction SilentlyContinue -WarningAction SilentlyContinue | Out-Null
                            }
                        }
                    }
                }
                If ($Selectie -eq "Import") {
                    If ($global:ExchangeVersion -like "201*" -OR $global:RemoteExchange -eq $true) {
                        Write-Host "  Initializing the mailbox $Object..."
                        If ($TargetFolder -ne "n.a.") {
                            If ($global:ExchangeVersion -eq "2010") {
                                $NewRequest = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create(`
                                    "New-MailboxImportRequest -Mailbox '$Alias' -Name '$Request' -DomainController '$PDC' -Filepath '$UNC' -BadItemLimit 50000 -AcceptLargeDataLoss -SourceRootFolder '$SourceRootFolder' -TargetRootFolder '$TargetFolder'") # -WarningAction SilentlyContinue")
                                )
                            }
                            Else {
                                $NewRequest = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create(`
                                    "New-MailboxImportRequest -Mailbox '$Alias' -Name '$Request' -DomainController '$PDC' -Filepath '$UNC' -BadItemLimit 50000 -LargeItemLimit 50000 -AcceptLargeDataLoss -SourceRootFolder '$SourceRootFolder' -TargetRootFolder '$TargetFolder'") # -WarningAction SilentlyContinue")
                                )
                            }
                        }
                        Else {
                            If ($global:ExchangeVersion -eq "2010") {
                                $NewRequest = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create(`
                                    "New-MailboxImportRequest -Mailbox '$Alias' -Name '$Request' -DomainController '$PDC' -Filepath '$UNC' -BadItemLimit 50000 -AcceptLargeDataLoss -SourceRootFolder '$SourceRootFolder' -WarningAction SilentlyContinue")
                                )
                            }
                            Else {
                                $NewRequest = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create(`
                                    "New-MailboxImportRequest -Mailbox '$Alias' -Name '$Request' -DomainController '$PDC' -Filepath '$UNC' -BadItemLimit 50000 -LargeItemLimit 50000 -AcceptLargeDataLoss -SourceRootFolder '$SourceRootFolder'") # -WarningAction SilentlyContinue")
                                )
                            }
                        }
                    }
                }
                If ($Selectie -eq "Restore") {
                    If ($global:ExchangeVersion -like "201*" -OR $global:RemoteExchange -eq $true) {
                        Write-Host "  Initializing the mailbox $Object..."
                        If ($global:TargetFolder -ne "n.a.") {
                            $NewRequest = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create(`
                                "New-MailboxRestoreRequest -Name '$Request' -SourceDatabase `"$SourceDB`" -SourceStoreMailbox $SourceGUID -TargetMailbox $TargetAlias -TargetRootFolder `"$TargetFolder`" -AllowLegacyDNMismatch")
                            )
                        }
                        Else {
                            $NewRequest = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create(`
                                "New-MailboxRestoreRequest -Name '$Request' -SourceDatabase `"$SourceDB`" -SourceStoreMailbox $SourceGUID -TargetMailbox $TargetAlias -AllowLegacyDNMismatch")
                            )
                        }
                    }
                }
                If (!$?) {
                    Write-Host -NoNewLine "  ERROR: An error occurred initializing the mailbox $Object." -ForegroundColor Red
                    If ($Selectie -eq "Export") {
                        Write-Host "Please check if your account is a member of the Mailbox Export Import role. Also check if the Microsoft Exchange Mailbox Replication service uses a Domain Administrator account." -ForegroundColor Red
                    }
                    Else {
                        Write-Host "Please investigate!" -ForegroundColor Red
                    }
                    $Pause.Invoke()
                    Get-Variable -Exclude $global:StartupVariables | Remove-Variable -ErrorAction SilentlyContinue; Invoke-Expression -Command $global:MenuNameCategory
                }
            }
            If ($Selectie -eq "Export" -AND ($global:ExchangeVersion -like "201*" -OR $global:RemoteExchange -eq $true) -AND !$SkipOverwrite) {
                $GetMailboxExportRequest = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("Get-MailboxExportRequest -Name '$Request' -DomainController '$PDC'"))
                $GetMailboxExportRequest = ($GetMailboxExportRequest | Select-Object Name, RequestGuid).RequestGuid
                $global:StatusRequest = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create(`
                    "Get-MailboxExportRequestStatistics '$GetMailboxExportRequest' -IncludeReport -DomainController '$PDC' -ErrorAction SilentlyContinue")
                )
            }
            If ($Selectie -eq "Import" -AND ($global:ExchangeVersion -like "201*" -OR $global:RemoteExchange -eq $true)) {
                $GetMailboxImportRequest = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("Get-MailboxImportRequest -Name '$Request' -DomainController '$PDC'"))
                $GetMailboxImportRequest = ($GetMailboxImportRequest | Select-Object Name, RequestGuid).RequestGuid
                $global:StatusRequest = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create(`
                    "Get-MailboxImportRequestStatistics '$GetMailboxImportRequest' -IncludeReport -DomainController '$PDC' -ErrorAction SilentlyContinue")
                )
            }
            If ($Selectie -eq "Restore" -AND ($global:ExchangeVersion -like "201*" -OR $global:RemoteExchange -eq $true)) {
                $GetMailboxRestoreRequest = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("Get-MailboxRestoreRequest -Name '$Request' -DomainController '$PDC'"))
                $GetMailboxRestoreRequest = ($GetMailboxRestoreRequest | Select-Object Name, RequestGuid).RequestGuid
                $global:StatusRequest = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create(`
                    "Get-MailboxRestoreRequestStatistics '$GetMailboxRestoreRequest' -IncludeReport -DomainController '$PDC' -ErrorAction SilentlyContinue")
                )
            }
        }
        # --------------------------------------------------------------------------------------------
        # As long as the mailbox-request is Queued or InProgress, then loop the output every 3 seconds
        # --------------------------------------------------------------------------------------------
        If (!$SkipOverwrite) {
            Do {
                $global:Actief = "Ja"
                [string]$CurrentStatus = ($StatusRequest.Status).Value
                If (!$CurrentStatus) {
                    [string]$CurrentStatus = $StatusRequest.Status
                }
                [string]$CurrentStatusDetail = ($StatusRequest.StatusDetail).Value
                If (!$CurrentStatusDetail) {
                    [string]$CurrentStatusDetail = $StatusRequest.StatusDetail
                }
                If ($CurrentStatus -ne $CurrentStatusDetail) {
                    $CurrentStatusDetail = " (" + $CurrentStatusDetail + ")"
                }
                Else {
                    $CurrentStatusDetail = $null
                }
                $CurrentStatusFormat = [string]$StatusRequest.PercentComplete + "% " + $CurrentStatus + $CurrentStatusDetail
                $Size = [string]$StatusRequest.EstimatedTransferSize
                $Transferred = [string]$StatusRequest.BytesTransferred
                Script-Module-SetHeaders -Name ($Titel + " Mailbox $Object " + $Counter + " of " + $global:MailboxObjectsCount)
                $Subtitel.Invoke()
                If ($global:ExchangeVersion -eq "2007") {
                    Write-Host
                    Write-Host
                    Write-Host
                }
                If ($Selectie -eq "Export" -OR $Selectie -eq "Import") {
                    Write-Host -NoNewLine "> Name: "; Write-Host $global:Name -ForegroundColor Yellow
                    If ($Email) {
                        Write-Host -NoNewLine "> Primary Mail Address: "; Write-Host $Email -ForegroundColor Yellow
                    }
                    Write-Host -NoNewLine "> PST-file: "; Write-Host $Path -ForegroundColor Yellow
                }
                If ($Selectie -eq "Restore") {
                    Write-Host -NoNewLine "> Recoverable mailbox: "; Write-Host $SourceName -ForegroundColor Yellow
                    Write-Host -NoNewLine "> Destination mailbox: ";
                    If ($TargetFolder -ne "n.a.") {
                        Write-Host ($TargetName + ": " + $TargetFolder) -ForegroundColor Yellow
                    }
                    Else {
                        Write-Host $TargetName -ForegroundColor Yellow
                    }
                }
                Write-Host 
                If ($global:ExchangeVersion -like "201*" -OR $global:RemoteExchange -eq $true) {
                    Write-Host "  Status                  :" $CurrentStatusFormat
                    If ($Selectie -eq "Export" -AND $Selectie -eq "Import") {
                        Write-Host "  Total $Object size     :" $Size
                    }
                    Else {
                        Write-Host "  Total $Object size   :" $Size
                    }
                    If ($Selectie -eq "Export") {
                        $global:TransferSize = "  Saved to PST     :"
                    }
                    If ($Selectie -eq "Import" -OR $Selectie -eq "Restore") {
                        $global:TransferSize = "  Saved to mailbox :"
                    }
                    Write-Host $TransferSize $Transferred
                }
                Write-Host
                If ($CurrentStatus -ne "Queued" -AND $CurrentStatus -ne "InProgress") {
                    $global:Status = $CurrentStatus
                }
                Else {
                    Write-Host "Please note: Do not close this window to prevent new mailbox $Object(s) not being initialized."
                    If ($global:ExchangeVersion -like "201*" -OR $global:RemoteExchange -eq $true) {
                        Write-Host -NoNewLine "Status will be refreshed in a few seconds..."
                        Start-Sleep -Seconds 2
                        If ($Selectie -eq "Export") {
                            $global:StatusRequest = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create(`
                                "Get-MailboxExportRequest -Name '$Request' -DomainController '$PDC' | Get-MailboxExportRequestStatistics -DomainController '$PDC' -IncludeReport")
                            )
                        }
                        If ($Selectie -eq "Import") {
                            $global:StatusRequest = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create(`
                                "Get-MailboxImportRequest -Name '$Request' -DomainController '$PDC' | Get-MailboxImportRequestStatistics -DomainController '$PDC' -IncludeReport")
                            )
                        }
                        If ($Selectie -eq "Restore") {
                            $global:StatusRequest = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create(`
                                "Get-MailboxRestoreRequest -Name '$Request' -DomainController '$PDC' | Get-MailboxRestoreRequestStatistics -DomainController '$PDC' -IncludeReport")
                            )
                        }
                    }
                    If ($global:ExchangeVersion -eq "2007") {
                        If ($Selectie -eq "Export") { 
                            $Mailbox | Export-Mailbox -PSTFolderPath $Path -StartDate $StartDate -BadItemLimit 999 -ReportFile ($Path + ".xml") -Confirm:$false -ErrorAction SilentlyContinue -WarningAction SilentlyContinue | Out-Null
                            If ($?) {
                                $global:Status = "Completed"
                            }
                            Else {
                                Write-Host
                                Write-Host "ERROR: An error occurred initializing mailbox $Object, please investigate!" -ForegroundColor Red
                                $Pause.Invoke()
                                EXIT
                            }
                        }
                    }
                }
            } Until ($global:Status)
            If ($Type -eq "Archive") {
                If ($global:Status -eq "Completed") {
                    Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create(`
                        "Get-MailboxExportRequest -Name '$Request' -DomainController '$PDC' | Remove-MailboxExportRequest -DomainController '$PDC' -Confirm:0")
                    )
                }
            }
            $global:File = $null
        }
    }
    If ($Type -eq "Archive") {
        If ($global:Status -eq "Completed") {
            Invoke-Expression -Command ($global:FunctionTaskNames | Where-Object { $_ -like "Messaging*Archive-Mailboxes-*-Start" })
        }
        Else {
            $global:Fout = $true
            Write-Host "ERROR: An error occurred exporting the mailbox, please investigate and remove the mailbox export-request in the next menu..." -ForegroundColor Red
            $Pause.Invoke()
        }
    }
    # ----------------------------------------------------------------
    # Go to the results to display an output of the mailbox request(s)
    # ----------------------------------------------------------------
    Script-Module-SetHeaders -Name ($MyInvocation.MyCommand).Name
    Write-Host ("Mailbox $Object " + $Counter + " of " + $global:MailboxObjectsCount + " successfully executed") -ForegroundColor Yellow
    If ($Selectie -eq "Export") {
        Script-Module-Initial-5-SetSystemLocale
        Messaging-Exchange-Status-Results -Selectie "Export" -Actief $global:Actief
    }
    If ($Selectie -eq "Import") {
        Messaging-Exchange-Status-Results -Selectie "Import" -Actief $global:Actief
    }
    If ($Selectie -eq "Restore") {
        Messaging-Exchange-Status-Results -Selectie "Restore" -Actief $global:Actief
    }
}
Function Messaging-Exchange-Move-Mailboxes-8-Start {
    # =========
    # Execution
    # =========
    Script-Module-SetHeaders -Name ($MyInvocation.MyCommand).Name
    If ($ServerDestination -eq "Office 365") {
        Write-Host -NoNewLine "Connecting to the Exchange Online session..."
        Script-Connect-Office365 -Module "ExchangeOnline" -Name ($MyInvocation.MyCommand).Name
        Import-PSSession $Office365Exchange -WarningAction SilentlyContinue | Out-Null
    }
    Else {
        Write-Host -NoNewLine "Connecting to the Exchange Server session..."
        Import-PSSession -Session (Get-PSSession -Name 'Exchange') -DisableNameChecking -WarningAction SilentlyContinue | Out-Null
    }
    Script-Module-SetHeaders -Name ($MyInvocation.MyCommand).Name
    Write-Host -NoNewLine "Indexing existing move-request(s), this can take a few moments..."
    $GetMoveRequests = Get-MoveRequest -ResultSize Unlimited -ErrorAction SilentlyContinue | SelectName
    Script-Module-SetHeaders -Name ($MyInvocation.MyCommand).Name
    $Counter = 1
    ForEach ($Mailbox in ($Mailboxes | Sort-Object Name)) {
        Write-Host ("Initiating move-request $Counter of " + $Mailboxes.Count + ": " + $Mailbox.Name + " (" + $Mailbox.UserPrincipalName + ")...") -ForegroundColor Yellow
        $MailboxName = $Mailbox.Name
        $MailboxUPN = (($Mailbox.UserPrincipalName).Replace("'", "''")).Replace("&", "`&")
        $MoveRequest = $GetMoveRequests | Where-Object { $_.DisplayName -eq "$MailboxName" -OR $_.Name -eq "$MailboxName" }
        If (!$MoveRequest) {
            If ($CSV) {
                $MailboxUPN = (Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("Get-Mailbox '$MailboxUPN' | Select-Object UserPrincipalName"))).UserPrincipalName
            }
            If ($MailboxUPN) {
                Do {
                    $MoveRequest = $null
                    $NewMoveRequest = New-MoveRequest -Identity "$MailboxUPN" -Remote -RemoteHostName $ServerSource -RemoteCredential $CredentialSource `
                        -TargetDeliveryDomain $DomainSource -Suspend -SuspendWhenReadyToComplete:$SuspendWhenReadyToComplete -BadItemLimit 999 -WarningAction SilentlyContinue
                    If ($?) {
                        Do {
                            $MoveRequest = Get-MoveRequest $MailboxUPN -ErrorAction SilentlyContinue
                            If (!$MoveRequest) {
                                Start-Sleep -Seconds 3
                            }
                        } Until ($MoveRequest)
                        Write-Host -NoNewLine " - Preparing mailbox move-request: "; Write-Host "OK!" -ForegroundColor Green
                    }
                    Else {
                        Write-Host -NoNewLine " - Preparing mailbox move-request: "; Write-Host "Failed!" -ForegroundColor Red
                        Do {
                            Write-Host -NoNewLine "   Would you like to ("; Write-Host -NoNewLine "R" -ForegroundColor Yellow; Write-Host -NoNewLine ")epeat or ("; Write-Host -NoNewLine "S" -ForegroundColor Yellow; Write-Host -NoNewLine ")kip this mailbox move-request? Use "; Write-Host -NoNewLine "X" -ForegroundColor Yellow; Write-Host -NoNewLine " to return to the previous menu: "
                            $Choice = Read-Host
                            $Input = @("R", "S", "X") -contains $Choice
                            If (!$Input) {
                                Write-Host "Please use the letters above as input" -ForegroundColor Red; Write-Host
                            }
                        } Until ($Input)
                        Switch ($Choice) {
                            "S" {
                                $MoveRequest = "Skip"
                            }
                            "X" {
                                Get-Variable -Exclude $global:StartupVariables | Remove-Variable -ErrorAction SilentlyContinue; Script-Disconnect-Server; Invoke-Expression -Command $global:MenuNameCategory
                            }
                        }
                    }
                } Until ($MoveRequest)
            }
            Else {
                Write-Host " - Mailbox does not exist anymore" -ForegroundColor Gray
            }
        }
        Else {
            Write-Host -NoNewLine " - Preparing mailbox move-request: "; Write-Host "Already exists" -ForegroundColor Gray
        }
        $Counter++
    }
    Write-Host
    # ==========
    # Finalizing
    # ==========
    $Pause.Invoke()
    Get-Variable -Exclude $global:StartupVariables | Remove-Variable -ErrorAction SilentlyContinue; Script-Disconnect-Server; $global:Exchange = $null; Invoke-Expression -Command $global:MenuNameCategory
}
Function Messaging-Exchange-Modify-Mailboxes-6-SelectPermission ([string]$Soort) {
    # =========
    # Execution
    # =========
    $FunctionMailboxMenu = ($global:FunctionTaskNames | Where-Object { $_ -like "Messaging*Modify-Mailboxes-*-Menu" })
    $FunctionMailboxAddPermission = ($global:FunctionTaskNames | Where-Object { $_ -like "Messaging*Modify-Mailboxes-*-AddPermission" })
    $FunctionMailboxRemovePermission = ($global:FunctionTaskNames | Where-Object { $_ -like "Messaging*Modify-Mailboxes-*-RemovePermission" })
    $global:Objects = @()
    If ($Soort -eq "FullAccess") {
        $global:Objects = $global:FullAccess
    }
    If ($Soort -eq "SendAs") {
        $global:Objects = $global:SendAs
    }
    If ($Soort -eq "SendOnBehalf") {
        $global:Objects = $global:SendOnBehalf
    }
    $global:Keuze = "permission"
    If ($global:Objects) {
        Do {
            Write-Host -NoNewLine "Would you like to "; Write-Host -NoNewLine "A" -ForegroundColor Yellow; Write-Host -NoNewLine "dd or "; Write-Host -NoNewLine "R" -ForegroundColor Yellow; Write-Host -NoNewline "emove "; Write-Host -NoNewLine $Soort -ForegroundColor Cyan; Write-Host -NoNewLine " permissions? Use "; Write-Host -NoNewLine "X" -ForegroundColor Yellow; Write-Host -NoNewline " to cancel: "
            $Choice = Read-Host
            $Input = @("A"; "R"; "X") -contains $Choice
            If (!$Input) {
                Write-Host "Please use the letters above as input" -ForegroundColor Red; Write-Host
            }
        } Until ($Input)
        Switch ($Choice) {
            "A" {
                Invoke-Expression -Command ("$FunctionMailboxAddPermission -Soort $Soort")
            }
            "R" {
                Invoke-Expression -Command ("$FunctionMailboxRemovePermission -Soort $Soort")
            }
            "X" {
                Invoke-Expression -Command $FunctionMailboxMenu
            }
        }
    }
    Else {
        Do {
            Write-Host -NoNewLine "There are no "; Write-Host -NoNewLine $Soort -ForegroundColor Cyan; Write-Host -NoNewLine " permissions set, would you like to add permissions? (Y/N): "
            $Choice = Read-Host
            $Input = @("Y"; "N") -contains $Choice
            If (!$Input) {
                Write-Host "Please use the letters above as input" -ForegroundColor Red; Write-Host
            }
        } Until ($Input)
        Switch ($Choice) {
            "Y" {
                Invoke-Expression -Command ("$FunctionMailboxAddPermission -Soort $Soort")
            }
            "N" {
                Invoke-Expression -Command $FunctionMailboxMenu
            }
        }
    }
}

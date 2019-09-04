Function Messaging-Exchange-Create-Mailboxes-3-SelectAddress {
    # ============
    # Declarations
    # ============
    $Task = "Select a mail address for the mailbox"
    # =========
    # Execution
    # =========
    Script-Module-SetHeaders -CurrentTask $Task -Name ($MyInvocation.MyCommand).Name
    Script-Index-Objects -CurrentTask $Task -FunctionName ($MyInvocation.MyCommand).Name -Type "GenereerEmailadres" -Functie "Selecteren"
    If ($global:Handmatig -eq "Ja") {
        Do {
            Write-Host -NoNewLine "Would you like to edit the "; Write-Host -NoNewLine -ForegroundColor Yellow "P"; Write-Host -NoNewLine "refix (before the @) or "; Write-Host -NoNewLine -ForegroundColor Yellow "D"; Write-Host -NoNewLine "omain address?: "
            $Choice = Read-Host
            $Input = @("P", "D") -contains $Choice
            If (!$Input) {
                Write-Host "Please use above as input" -ForegroundColor Red
            }
        } Until ($Input)
        Switch ($Choice) {
            "P" {
                Do {
                    $global:EmailPrefix = Read-Host "Please enter the prefix of the mail address"
                    $global:Email = $EmailPrefix + "@" + $Emaildomein
                    $CheckEmailAddress = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("Get-Mailbox -Identity '$Email' -DomainController '$PDC'"))
                    If ($CheckEmailAddress) {
                        Write-Host ("This mail address is already in use at " + $CheckEmailAddress.Name) -ForegroundColor Red
                    }
                    Else {
                        $Checked = $true
                    }
                } Until ($Checked -eq $true)
            }
            "D" {
                Do {
                    $global:EmailSuffix = Read-Host "Please enter the domain address"
                    $global:Domein = $EmailSuffix
                    If ($EmailPrefix) {
                        $global:Email = $EmailPrefix + "@" + $Emaildomein
                    }
                    Else {
                        $global:Email = $SamAccountName + "@" + $Emaildomein
                    }
                    $CheckEmailAddress = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("Get-Mailbox -Identity '$Email' -DomainController '$PDC'"))
                    If ($CheckEmailAddress) {
                        Write-Host ("This mail address is already in use at " + $CheckEmailAddress.Name) -ForegroundColor Red
                    }
                    Else {
                        $Checked = $true
                    }
                } Until ($Checked -eq $true)
            }
        }
    }
    Else {
        $global:Email = $GegenereerdeOptie
    }
    # ==========
    # Finalizing
    # ==========
    Set-Variable $global:VarHeaderName -Value $global:Email
    Set-Variable $global:VarHeaderName -Value ("Write-Host -NoNewLine ('- ' + '$Task' + ': ') -ForegroundColor Gray; Write-Host '" + (Get-Variable -Name $global:VarHeaderName).Value + "' -ForegroundColor Yellow")
    Invoke-Expression -Command $FunctionTaskNames[[int]$FunctionTaskNames.IndexOf(($MyInvocation.MyCommand).Name) + 1]
}
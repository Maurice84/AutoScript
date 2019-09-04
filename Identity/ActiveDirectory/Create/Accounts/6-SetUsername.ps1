Function Identity-ActiveDirectory-Create-Accounts-6-SetUsername {
    # ============
    # Declarations
    # ============
    $global:AccountType = $null
    $Task = "Select an username for the new account"
    # =========
    # Execution
    # =========
    Script-Module-SetHeaders -CurrentTask $Task -Name ($MyInvocation.MyCommand).Name
    If ($global:Profiel -eq "Beheerder account" -OR $global:OUPath -like "*External*") {
        $global:AccountType = "Beheer"
    }
    If ($global:Profiel -eq "Mailbox account" -OR $global:OUPath -like "*Mailbox*") {
        $global:AccountType = "MB"
    }
    If ($global:Profiel -eq "Service account" -OR $global:OUPath -like "*Service*") {
        $global:AccountType = "SA"
    }
    If ($global:Profiel -eq "Test gebruiker" -OR $global:OUPath -like "*Test*") {
        $global:AccountType = "Test"
    }
    Script-Index-Objects -CurrentTask $Task -FunctionName ($MyInvocation.MyCommand).Name -Type "GenereerGebruikersnaam" -Functie "Selecteren"
    If ($global:Handmatig -eq "Ja") {
        Do {
            $InputChoice = Read-Host "  Please enter an username (without @xxx.nl)"
            $InputKey = $InputChoice
            If (!$InputKey) {
                Write-Host "  There is no input detected" -ForegroundColor Red; Write-Host
            }
            $InputChoice = $InputChoice + "@" + $UPNSuffix
            $global:UsernameCheck = Script-Connect-Server -Module "ActiveDirectory" -Command ([scriptblock]::Create(`
                        "Get-ADUser -Filter * | Select-Object UserPrincipalName | Where-Object {`$`_.UserPrincipalName -eq '$Choice'}")
            )
            If ($UsernameCheck) {
                $InputKey = $null
                Write-Host "  This username $InputChoice already exists, please enter another username" -ForegroundColor Red; Write-Host
            }
            Else {
                $InputKey = $InputChoice
            }
        } Until ($InputKey)
        $global:UPN = $InputKey
    }
    Else {
        $global:UPN = $global:GegenereerdeOptie
    }
    If ($global:Profiel -eq "Beheerder account" -OR $global:GeenNaamIngevuld) {
        $global:GivenName = $UPN.Split('@')[0]
    }
    $global:Username = $UPN.Split('@')[0]
    # ==========
    # Finalizing
    # ==========
    Set-Variable $global:VarHeaderName -Value $global:UPN
    Set-Variable $global:VarHeaderName -Value ("Write-Host -NoNewLine ('- ' + '$Task' + ': ') -ForegroundColor Gray; Write-Host '" + (Get-Variable -Name $global:VarHeaderName).Value + "' -ForegroundColor Yellow")
    Invoke-Expression -Command $FunctionTaskNames[[int]$FunctionTaskNames.IndexOf(($MyInvocation.MyCommand).Name) + 1]
}
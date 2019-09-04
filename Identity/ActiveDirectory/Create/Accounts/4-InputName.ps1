Function Identity-ActiveDirectory-Create-Accounts-4-InputName {
    # ============
    # Declarations
    # ============
    $Task = "Entered name for the new account"
    # =========
    # Execution
    # =========
    Script-Module-SetHeaders -CurrentTask $Task -Name ($MyInvocation.MyCommand).Name
    $global:GeenNaamIngevuld = $null
    If ($global:Profiel -eq "Beheerder account" -OR $global:OUPath -like "OU=External*") {
        $global:GivenName = Read-Host "  > Please enter 1 (short) word to describe the application/function or customer for this external admin account (i.e. Microsoft)"
        $global:GivenName = $global:GivenName.Replace(' ', '')
        $global:Surname = $null
    }
    ElseIf ($global:Profiel -eq "Mailbox account" -OR $global:OUPath -like "OU=Mailbox*") {
        $global:GivenName = Read-Host "  > Please enter the displayname for recipients to see when this mailbox is mailing (i.e. Info Customername)"
        $global:Surname = $null
    }
    ElseIf ($global:Profiel -eq "Service account" -OR $global:OUPath -like "OU=Service*") {
        $global:GivenName = Read-Host "  > Please enter 1 (short) word to describe the application/function for this service account (i.e. SQL)"
        $global:GivenName = $global:GivenName.Replace(' ', '')
        $global:Surname = $null
    }
    ElseIf ($global:Profiel -eq "Test gebruiker" -OR $global:OUPath -like "OU=Test*") {
        Do {
            $InputChoice = Read-Host "  > There is no need to enter a GivenName and Surname for a test account. Would you like to enter this anyway? (Y/N)"
            $InputKey = @("Y"; "N") -contains $InputChoice
            If (!$InputKey) {
                Write-Host "    Please use the letters above as input" -ForegroundColor Red; Write-Host
            }
        } Until ($InputKey)
        Switch ($InputChoice) {
            "Y" {
                $global:GivenName = Read-Host "     - Please enter a given name for the test account"
                $global:Surname = Read-Host "     - Please enter a surname for the test account"
            }
            "N" {
                $global:GeenNaamIngevuld = $true
                $global:GivenName = $null
                $global:Surname = $null
            }
        }
    }
    Else {
        $global:GivenName = Read-Host "  Please enter a given name for the account"
        $global:Surname = Read-Host "  Please enter a surname for the test account"
    }
    # ==========
    # Finalizing
    # ==========
    If (!$global:GivenName -AND !$global:SurName) {
        Set-Variable $global:VarHeaderName -Value "n.a."
    }
    Else {
        If ($global:Profiel -eq "Profielkopie" -OR $global:Profiel -eq "Domein gebruiker" -OR $global:Profiel -eq "Mailbox account" -OR $global:Profiel -eq "Webmail gebruiker") {
            Set-Variable $global:VarHeaderName -Value ($global:GivenName + " " + $global:Surname)
        }
        Else {
            Set-Variable $global:VarHeaderName -Value "n.a."
        }
    }
    Set-Variable $global:VarHeaderName -Value ("Write-Host -NoNewLine ('- ' + '$Task' + ': ') -ForegroundColor Gray; Write-Host '" + (Get-Variable -Name $global:VarHeaderName).Value + "' -ForegroundColor Yellow")
    Invoke-Expression -Command $FunctionTaskNames[[int]$FunctionTaskNames.IndexOf(($MyInvocation.MyCommand).Name) + 1]
}
Function Identity-ActiveDirectory-Modify-Accounts-3-Menu {
    # =========
    # Execution
    # =========
    Script-Module-SetHeaders -Name ($MyInvocation.MyCommand).Name
    Write-Host "Account Info:" -ForegroundColor Magenta
    Write-Host -NoNewLine "  1. DisplayName: "; Write-Host $global:WeergaveNaam -ForegroundColor Yellow
    Write-Host -NoNewLine "  2. UserPrincipalName: "; Write-Host $global:Gebruikersnaam -ForegroundColor Yellow
    Write-Host -NoNewLine "  3. SamAccountName: "; Write-Host ($env:userdomain + "\" + $global:PreWin2000Naam) -ForegroundColor Yellow
    Write-Host -NoNewLine "  4. GivenName: "; Write-Host $global:Voornaam -ForegroundColor Yellow
    Write-Host -NoNewLine "  5. Surname: "; Write-Host $global:Achternaam -ForegroundColor Yellow
    Write-Host -NoNewLine "  6. Initials: "; Write-Host $global:Initialen -ForegroundColor Yellow
    Write-Host -NoNewLine "  7. Language: "; Write-Host $global:Taalinstelling -ForegroundColor Yellow
    Write-Host -NoNewLine "  8. Security Groups: "; Write-Host $global:Groepen -ForegroundColor Yellow
    Write-Host -NoNewLine "  9. Location (OU): "; Write-Host $global:LocatieOU -ForegroundColor Yellow
    Write-Host -NoNewLine " 10. Profile Path: "; Write-Host $global:RDPPath -ForegroundColor Yellow
    Write-Host -NoNewLine " 11. Home Folder: "; Write-Host -NoNewLine $global:HomePath -ForegroundColor Yellow
    If ($global:HomeDrive) {
        Write-Host (" (" + $global:HomeDrive + ")") -ForegroundColor Yellow
    } Else {
        Write-Host
    }
    Write-Host
    Write-Host "Organization Info:" -ForegroundColor Magenta
    Write-Host -NoNewLine " 12. Company: "; Write-Host $global:Bedrijfsnaam -ForegroundColor Yellow
    Write-Host -NoNewLine " 13. Website: "; Write-Host $global:Website -ForegroundColor Yellow
    Write-Host -NoNewLine " 14. Office: "; Write-Host $global:Kantoor -ForegroundColor Yellow
    Write-Host -NoNewLine " 15. Department: "; Write-Host $global:Afdeling -ForegroundColor Yellow
    Write-Host -NoNewLine " 16. Job: "; Write-Host $global:FunctieTitel -ForegroundColor Yellow
    Write-Host -NoNewLine " 17. Office Phone: "; Write-Host $global:TelefoonWerk -ForegroundColor Yellow
    Write-Host -NoNewLine " 18. Office Fax: "; Write-Host $global:TelefoonFax -ForegroundColor Yellow
    Write-Host
    Write-Host "Contact Info:" -ForegroundColor Magenta
    Write-Host -NoNewLine " 19. Street Address: "; Write-Host $global:Straat -ForegroundColor Yellow
    Write-Host -NoNewLine " 20. P.O. Box: "; Write-Host $global:Postbus -ForegroundColor Yellow
    Write-Host -NoNewLine " 21. Postcode: "; Write-Host $global:Postcode -ForegroundColor Yellow
    Write-Host -NoNewLine " 22. City: "; Write-Host $global:Stad -ForegroundColor Yellow
    Write-Host -NoNewLine " 23. Province: "; Write-Host $global:Provincie -ForegroundColor Yellow
    Write-Host -NoNewLine " 24. Phone: "; Write-Host $global:TelefoonPrive -ForegroundColor Yellow
    Write-Host -NoNewLine " 25. Mobile: "; Write-Host $global:TelefoonMobiel -ForegroundColor Yellow
    Write-Host
    Do {
        Write-Host -NoNewline "Please select a category or use "; Write-Host -NoNewLine "X" -ForegroundColor Yellow; Write-Host -NoNewline " to return to the previous menu: "
        $InputChoice = Read-Host
        $InputKey = @("1"; "2"; "3"; "4"; "5"; "6"; "7"; "8"; "9"; "10"; "11"; "12"; "13"; "14"; "15"; "16"; "17"; "18"; "19"; "20"; "21"; "22"; "23"; "24"; "25"; "X") -contains $Choice
        If (!$InputKey) {
            Write-Host "Please use the letters/numbers above as input" -ForegroundColor Red; Write-Host
        }
    } Until ($InputKey)
    $AccountDN = $global:Account.DistinguishedName
    Switch ($InputChoice) {
        "1" {
            $global:Soort = "weergave naam"
            $global:Command = { Script-Connect-Server -Module "ActiveDirectory" -Command ([scriptblock]::Create("Set-ADUser '$AccountDN' -Name '$Value'}")) }
        }
        "2" {
            $Task = "Select a UPN suffix for this account"
            Script-Index-Objects -CurrentTask $Task -FunctionName ($MyInvocation.MyCommand).Name -Type "UPNSuffix" -Functie "Selecteren"
            $Task = "Select a new username for this account"
            Script-Index-Objects -CurrentTask $Task -FunctionName ($MyInvocation.MyCommand).Name -Type "GenereerGebruikersnaam" -Functie "Selecteren"
            If ($global:Handmatig -eq "Ja") {
                Do {
                    $InputChoice = Read-Host "  Enter a new username (without @xxx.nl)"
                    $InputKey = $InputChoice
                    If (!$InputKey) {
                        Write-Host "  Please enter a username" -ForegroundColor Red; Write-Host
                    }
                    $InputChoice = $InputChoice + "@" + $global:UPNSuffix
                    $UsernameCheck = Script-Connect-Server -Module "ActiveDirectory" -Command ([scriptblock]::Create(`
                                "Get-ADUser -Filter * | Select-Object UserPrincipalName | Where-Object {`$`_.UserPrincipalName -eq '$InputChoice'}")
                    )
                    If ($UsernameCheck) {
                        $InputKey = $null
                        Write-Host "  The username $InputChoice already exists, please enter an unused username" -ForegroundColor Red; Write-Host
                    }
                    Else {
                        $InputKey = $InputChoice
                    }
                } Until ($InputKey)
                $global:UPN = $InputKey
            }
            Else {
                $global:UPN = $GegenereerdeOptie
            }
            Script-Connect-Server -Module "ActiveDirectory" -Command ([scriptblock]::Create("Set-ADUser '$AccountDN' -UserPrincipalName '$global:UPN'"))
            If (!$?) {
                Write-Host "ERROR: Could not modify username, please investigate!" -ForegroundColor Red
                $Pause.Invoke()
            }
            Else {
                Script-Module-ReplicateAD
            }
        }
        "3" {
            $global:Soort = "Gebruikersnaam (Pre-Win2000)"
            $global:Command = { Script-Connect-Server -Module "ActiveDirectory" -Command ([scriptblock]::Create("Set-ADUser '$AccountDN' -SamAccountName '$Value'")) }
        }
        "4" {
            $global:Soort = "Voornaam"
            $global:Command = { Script-Connect-Server -Module "ActiveDirectory" -Command ([scriptblock]::Create("Set-ADUser '$AccountDN' -GivenName '$Value'")) }
        }
        "5" {
            $global:Soort = "Achternaam"
            $global:Command = { Script-Connect-Server -Module "ActiveDirectory" -Command ([scriptblock]::Create("Set-ADUser '$AccountDN' -Surname '$Value'")) }
        }
        "6" {
            $global:Soort = "Initialen"
            $global:Command = { Script-Connect-Server -Module "ActiveDirectory" -Command ([scriptblock]::Create("Set-ADUser '$AccountDN' -Initials '$Value'")) }
        }
        "7" {
            Do {
                Write-Host -NoNewLine "Select a language: "; Write-Host -NoNewLine "D" -ForegroundColor Yellow; Write-Host -NoNewLine "utch, "; Write-Host -NoNewLine "E" -ForegroundColor Yellow; Write-Host -NoNewLine "nglish or "; Write-Host -NoNewLine "G" -ForegroundColor Yellow; Write-Host -NoNewLine "erman. Use "; Write-Host -NoNewLine "X" -ForegroundColor Yellow; Write-Host -NoNewline " to cancel: "
                $InputChoice = Read-Host
                $InputKey = @("D"; "E", "G", "X") -contains $InputChoice
                If (!$InputKey) {
                    Write-Host "Please use the letters above as input" -ForegroundColor Red; Write-Host
                }
            } Until ($InputKey)
            Switch ($InputChoice) {
                "D" {
                    $global:Taalinstelling = "Nederlands"
                    $global:Language = "NL"
                }
                "E" {
                    $global:Taalinstelling = "Engels"
                    $global:Language = "US"
                }
                "G" {
                    $global:Taalinstelling = "Duits"
                    $global:Language = "DE"
                }
                "X" {
                    Invoke-Expression -Command ($MyInvocation.MyCommand).Name
                }
            }
            Script-Connect-Server -Module "ActiveDirectory" -Command ([scriptblock]::Create("Set-ADUser '$AccountDN' -Country '$global:Language'"))
            If (!$?) {
                Write-Host "ERROR: Could not modify language, please investigate!" -ForegroundColor Red
                $Pause.Invoke()
            }
            Else {
                Script-Module-ReplicateAD
            }
        }
        "8" {
            Invoke-Expression -Command $FunctionTaskNames[[int]$FunctionTaskNames.IndexOf(($MyInvocation.MyCommand).Name) + 1]
        }
        "9" {
            $Task = "Select an OU to move the account"
            $global:OUDN = ($global:Account | Select-Object @{n = 'ParentContainer'; e = { $_.DistinguishedName -replace '^.+?,(CN|OU.+)', '$1' } }).ParentContainer
            Script-Index-OU -Path $global:OUDN
            Script-Connect-Server -Module "ActiveDirectory" -Command ([scriptblock]::Create("Move-ADObject '$AccountDN' -TargetPath '$global:OUPath'"))
            If (!$?) {
                Write-Host "ERROR: Could not move the account, please investigate!" -ForegroundColor Red
                $Pause.Invoke()
            }
            Else {
                Script-Module-ReplicateAD
            }
        }
        "10" {
            $global:Soort = "profielpad"
        }
        "11" {
            $global:Soort = "homefolder"
        }
        "12" {
            $global:Soort = "Bedrijfsnaam"
            $global:Command = { Script-Connect-Server -Module "ActiveDirectory" -Command ([scriptblock]::Create("Set-ADUser '$AccountDN' -Company '$Value'")) }
        }
        "13" {
            $global:Soort = "Website"
            $global:Command = { Script-Connect-Server -Module "ActiveDirectory" -Command ([scriptblock]::Create("Set-ADUser '$AccountDN' -HomePage '$Value'")) }
        }
        "14" {
            $global:Soort = "Kantoor"
            $global:Command = { Script-Connect-Server -Module "ActiveDirectory" -Command ([scriptblock]::Create("Set-ADUser '$AccountDN' -Office '$Value'")) }
        }
        "15" {
            $global:Soort = "Afdeling"
            $global:Command = { Script-Connect-Server -Module "ActiveDirectory" -Command ([scriptblock]::Create("Set-ADUser '$AccountDN' -Department '$Value'")) }
        }
        "16" {
            $global:Soort = "Functie"
            $global:Command = { Script-Connect-Server -Module "ActiveDirectory" -Command ([scriptblock]::Create("Set-ADUser '$AccountDN' -Title '$Value'")) }
        }
        "17" {
            $global:Soort = "Telefoonnummer (werk)"
            $global:Command = { Script-Connect-Server -Module "ActiveDirectory" -Command ([scriptblock]::Create("Set-ADUser '$AccountDN' -OfficePhone '$Value'")) }
        }
        "18" {
            $global:Soort = "Telefoonnummer (fax)"
            $global:Command = { Script-Connect-Server -Module "ActiveDirectory" -Command ([scriptblock]::Create("Set-ADUser '$AccountDN' -Fax '$Value'")) }
        }
        "19" {
            $global:Soort = "Straat en huisnummer"
            $global:Command = { Script-Connect-Server -Module "ActiveDirectory" -Command ([scriptblock]::Create("Set-ADUser '$AccountDN' -StreetAddress '$Value'")) }
        }
        "20" {
            $global:Soort = "Postbus"
            $global:Command = { Script-Connect-Server -Module "ActiveDirectory" -Command ([scriptblock]::Create("Set-ADUser '$AccountDN' -POBox '$Value'")) }
        }
        "21" {
            $global:Soort = "Postcode"
            $global:Command = { Script-Connect-Server -Module "ActiveDirectory" -Command ([scriptblock]::Create("Set-ADUser '$AccountDN' -PostalCode '$Value'")) }
        }
        "22" {
            $global:Soort = "Stad"
            $global:Command = { Script-Connect-Server -Module "ActiveDirectory" -Command ([scriptblock]::Create("Set-ADUser '$AccountDN' -City '$Value'")) }
        }
        "23" {
            $global:Soort = "Provincie"
            $global:Command = { Script-Connect-Server -Module "ActiveDirectory" -Command ([scriptblock]::Create("Set-ADUser '$AccountDN' -State '$Value'")) }
        }
        "24" {
            $global:Soort = "Telefoonnummer (privé)"
            $global:Command = { Script-Connect-Server -Module "ActiveDirectory" -Command ([scriptblock]::Create("Set-ADUser '$AccountDN' -HomePhone '$Value'")) }
        }
        "25" {
            $global:Soort = "Telefoonnummer (mobiel)"
            $global:Command = { Script-Connect-Server -Module "ActiveDirectory" -Command ([scriptblock]::Create("Set-ADUser '$AccountDN' -MobilePhone '$Value'")) }
        }
        "X" {
            Get-Variable -Exclude $global:StartupVariables | Remove-Variable -ErrorAction SilentlyContinue; Script-Disconnect-Server; Invoke-Expression -Command $global:MenuNameCategory
        }
    }
    If ($global:Soort) {
        Do {
            Write-Host -NoNewLine ("Would you like to modify " + $global:Soort + "? (Y/N): ")
            $InputChoice = Read-Host
            $InputKey = @("Y"; "N") -contains $InputChoice
            If (!$InputKey) {
                Write-Host "Please use the letters above as input" -ForegroundColor Red; Write-Host
            }
        } Until ($InputKey)
        Switch ($InputChoice) {
            "Y" {
                Do {
                    Write-Host -NoNewLine ("Please enter a new value for " + $global:Soort + ": ")
                    $global:Value = Read-Host
                    If ($Value) {
                        If ($Soort -like "*Pre-Win2000*") {
                            If ($Value.Length -le 20) {
                                $global:Check = Script-Connect-Server -Module "ActiveDirectory" -Command ([scriptblock]::Create(`
                                    "Get-ADUser -Filter * | Select-Object SamAccountName | Where-Object {`$`_.SamAccountName -eq '$Value'}")
                                )
                                If ($Check) {
                                    Write-Host "ERROR: SamAccountName already exist! Please enter an unused SamAccountName" -ForegroundColor Red
                                    $global:Value = $null
                                }
                            }
                            Else {
                                Write-Host "ERROR: SamAccountName cannot be longer than 20 characters, please shorten the input" -ForegroundColor Red
                                $global:Value = $null
                            }
                        }
                    }
                    Else {
                        Write-Host "Please enter a value" -ForegroundColor Red; Write-Host
                    }
                } Until ($Value)
                $Command.Invoke()
                If (!$?) {
                    Write-Host "ERROR: The $global:Soort could not be modified, please investigate!" -ForegroundColor Red
                    $Pause.Invoke()
                }
                Script-Module-ReplicateAD
            }
            "N" {
                Invoke-Expression -Command ($MyInvocation.MyCommand).Name
            }
        }
    }
    Invoke-Expression -Command $FunctionTaskNames[[int]$FunctionTaskNames.IndexOf(($MyInvocation.MyCommand).Name) - 1]
}
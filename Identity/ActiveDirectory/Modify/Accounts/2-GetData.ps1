Function Identity-ActiveDirectory-Modify-Accounts-2-GetData {
    # =========
    # Execution
    # =========
    Script-Module-SetHeaders -Name ($MyInvocation.MyCommand).Name
    Write-Host -NoNewLine "Loading selected domain account: "; Write-Host -NoNewLine $Name -ForegroundColor Yellow; Write-Host -NoNewLine "..."
    $global:Account = Script-Connect-Server -Module "ActiveDirectory" -Command ([scriptblock]::Create(
            "Get-ADUser '$global:SamAccountName' -Properties *")
    )
    # ------------
    # Account info
    # ------------
    $global:Voornaam = $global:Account.GivenName; If (!$global:Voornaam) { $global:Voornaam = "n.a." }
    $global:Achternaam = $global:Account.Surname; If (!$global:Achternaam) { $global:Achternaam = "n.a." }
    $global:Initialen = $global:Account.Initials; If (!$global:Initialen) { $global:Initialen = "n.a." }
    $global:Omschrijving = $global:Account.Description; If (!$global:Omschrijving) { $global:Omschrijving = "n.a." }
    $global:PreWin2000Naam = $global:Account.SamAccountName
    $global:Gebruikersnaam = $global:Account.UserPrincipalName
    $global:WeergaveNaam = $global:Account.Name
    $global:LocatieOU = ($global:Account.CanonicalName).Replace(("/" + $global:WeergaveNaam), '')
    # ------
    # Groups
    # ------
    $global:GroepenArray = Script-Connect-Server -Module "ActiveDirectory" -Command ([scriptblock]::Create(
            "Get-ADPrincipalGroupMembership '$global:SamAccountName'")
    )
    $global:GroepenArray = ($global:GroepenArray | Where-Object { $_.Name -ne "Domain Users" -AND $_.Name -ne "Domeingebruikers" -AND $_.Name -ne $null } | Select-Object Name | Sort-Object Name).Name
    If ($global:GroepenArray.Count -gt 5) {
        $global:Groepen = [string]$global:GroepenArray.Count + " groups"
    }
    Else {
        $global:Groepen = $global:GroepenArray -join ', '
    }    
    If (!$global:GroepenArray) {
        $global:Groepen = "n.a."
    }
    # --------------------------
    # Profilepath and Homefolder
    # --------------------------
    $global:ADSI = [ADSI]('LDAP://{0}' -f $global:Account.DistinguishedName)
    $global:RDPPath = Try { $ADSI.InvokeGet('TerminalServicesProfilePath') } Catch { $null }; If (!$global:RDPPath) { $global:RDPPath = "n.a." }
    $global:HomePath = Try { $ADSI.InvokeGet('TerminalServicesHomeDirectory') } Catch { $null }; If (!$global:HomePath) { $global:HomePath = "n.a." }
    $global:HomeDrive = Try { $ADSI.InvokeGet('TerminalServicesHomeDrive') } Catch { $null }
    # --------
    # Language
    # --------
    If ($global:Account.Country -eq "NL") { $global:Taalinstelling = "Dutch" }
    If ($global:Account.Country -eq "US") { $global:Taalinstelling = "English" }
    If ($global:Account.Country -eq "DE") { $global:Taalinstelling = "German" }
    If (!$global:Account.Country) { $global:Taalinstelling = "n.a." }
    # ------------
    # Company info
    # ------------
    $global:Bedrijfsnaam = $global:Account.Company; If (!$global:Bedrijfsnaam) { $global:Bedrijfsnaam = "n.a." }
    $global:Website = $global:Account.HomePage; If (!$global:Website) { $global:Website = "n.a." }  
    $global:Kantoor = $global:Account.Office; If (!$global:Kantoor) { $global:Kantoor = "n.a." }
    $global:Afdeling = $global:Account.Department; If (!$global:Afdeling) { $global:Afdeling = "n.a." }
    $global:FunctieTitel = $global:Account.Title; If (!$global:FunctieTitel) { $global:FunctieTitel = "n.a." }
    $global:TelefoonWerk = $global:Account.OfficePhone; If (!$global:TelefoonWerk) { $global:TelefoonWerk = "n.a." }
    $global:TelefoonFax = $global:Account.Fax; If (!$global:TelefoonFax) { $global:TelefoonFax = "n.a." }
    # -------------
    # Location info
    # -------------
    $global:Straat = $global:Account.StreetAddress; If (!$global:Straat) { $global:Straat = "n.a." }
    $global:Postbus = $global:Account.POBox; If (!$global:Postbus) { $global:Postbus = "n.a." }
    $global:Postcode = $global:Account.PostalCode; If (!$global:Postcode) { $global:Postcode = "n.a." }
    $global:Stad = $global:Account.City; If (!$global:Stad) { $global:Stad = "n.a." }
    $global:Provincie = $global:Account.State; If (!$global:Provincie) { $global:Provincie = "n.a." }
    $global:TelefoonPrive = $global:Account.HomePhone; If (!$global:TelefoonPrive) { $global:TelefoonPrive = "n.a." }
    $global:TelefoonMobiel = $global:Account.MobilePhone; If (!$global:TelefoonMobiel) { $global:TelefoonMobiel = "n.a." }
    # ==========
    # Finalizing
    # ==========
    Invoke-Expression -Command $FunctionTaskNames[[int]$FunctionTaskNames.IndexOf(($MyInvocation.MyCommand).Name) + 1]
}

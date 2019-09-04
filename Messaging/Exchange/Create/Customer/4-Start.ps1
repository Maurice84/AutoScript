Function Messaging-Exchange-Create-Customer-4-Start {
    # =========
    # Execution
    # =========
    Script-Module-SetHeaders -Name ($MyInvocation.MyCommand).Name
    # -----------------
    # Declare variables
    # -----------------
    $Naam = $CustomerUPNSuffix
    $NaamGAL = $Naam + " - GAL"
    $NaamAllRooms = $Naam + " - All Rooms"
    $NaamAllUsers = $Naam + " - All Users"
    $NaamAllContacts = $Naam + " - All Contacts"
    $NaamAllGroups = $Naam + " - All Groups"
    $Emailadres = "SMTP:%1g.%s@" + $CustomerUPNSuffix
    $OU = $Customer
    # ---------------------------------------------------------------------------------------------------
    Write-Host -NoNewLine "Step 1/4 - Creating accepted mail domain" -ForegroundColor Magenta; Write-Host
    # ---------------------------------------------------------------------------------------------------
    $CheckDomain = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("Get-AcceptedDomain -DomainController '$PDC'"))
    If (!($CheckDomain | Where-Object { $_.DomainName -eq $CustomerUPNSuffix }) -AND !($CheckDomain | Where-Object { $_.Name -eq $Naam })) {
        $AddDomain = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("New-AcceptedDomain -Name '$Naam' -DomainName '$CustomerUPNSuffix' -DomainType:Authoritative -DomainController '$PDC'"))
        $CheckDomain = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("Get-AcceptedDomain -DomainController '$PDC'"))
        If ($CheckDomain | Where-Object { $_.Name -eq $Naam }) {
            Write-Host -NoNewLine " - OK: Mail domain "; Write-Host -NoNewLine $CustomerUPNSuffix -ForegroundColor Yellow; Write-Host " created successfully"
        }
        Else {
            Write-Host " - ERROR: An error occurred creating mail domain $CustomerUPNSuffix, please investigate!" -ForegroundColor Red
            $Pause.Invoke()
            Get-Variable -Exclude $global:StartupVariables | Remove-Variable -ErrorAction SilentlyContinue; Script-Disconnect-Server; Invoke-Expression -Command $global:MenuNameCategory
        }
    }
    Else {
        Write-Host " - INFO: Mail domain $CustomerUPNSuffix already exists" -ForegroundColor Gray
    }
    # -------------------------------------------------------------------------------------------
    Write-Host -NoNewLine "Step 2/4 - Creating address lists" -ForegroundColor Magenta; Write-Host
    # -------------------------------------------------------------------------------------------
    $CheckGAL = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("Get-GlobalAddressList -DomainController '$PDC'"))
    If (!($CheckGAL | Where-Object { $_.Name -eq $NaamGAL })) {
        $AddGAL = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("New-GlobalAddressList -Name '$NaamGAL' -RecipientFilter `"(CustomAttribute1 -eq '$KlantPrefix')`" -RecipientContainer '$OU' -DomainController '$PDC'"))
        $CheckGAL = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("Get-GlobalAddressList -DomainController '$PDC'"))
        If ($CheckGAL | Where-Object { $_.Name -eq $NaamGAL }) {
            Write-Host -NoNewLine " - OK: Global Address List "; Write-Host -NoNewLine $NaamGAL -ForegroundColor Yellow; Write-Host " created successfully"
        }
        Else {
            Write-Host " - ERROR: An error occurred creating Global Address List $NaamGAL, please investigate!" -ForegroundColor Red
            $Pause.Invoke()
            Get-Variable -Exclude $global:StartupVariables | Remove-Variable -ErrorAction SilentlyContinue; Script-Disconnect-Server; Invoke-Expression -Command $global:MenuNameCategory
        }
    }
    Else {
        Write-Host " - INFO: Global Address List $NaamGAL already exists" -ForegroundColor Gray
    }
    $NaamAL = @()
    $NaamAL += "$NaamAllContacts"
    $NaamAL += "$NaamAllGroups"
    $NaamAL += "$NaamAllRooms"
    $NaamAL += "$NaamAllUsers"
    $CheckAL = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("Get-AddressList -DomainController '$PDC'"))
    ForEach ($Naam in $NaamAL) {
        If (!($CheckAL | Where-Object { $_.Name -eq $Naam })) {
            If ($Naam -like "*Contacts") {
                $AddAL = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create(`
                    "New-AddressList -Name '$Naam' -RecipientFilter `"(ObjectClass -eq 'Contact')`" -RecipientContainer '$OU' -DomainController '$PDC'")
                )
            }
            If ($Naam -like "*Groups") {
                $AddAL = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create(`
                    "New-AddressList -Name '$Naam' -RecipientFilter `"ObjectClass -eq 'Group'`" -RecipientContainer '$OU' -DomainController '$PDC'")
                )
            }
            If ($Naam -like "*Rooms") {
                $AddAL = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create(`
                    "New-AddressList -Name '$Naam' -RecipientFilter `"(RecipientDisplayType -eq 'ConferenceRoomMailbox')`" -RecipientContainer '$OU' -DomainController '$PDC'")
                )
            }
            If ($Naam -like "*Users") {
                $AddAL = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create(`
                    "New-AddressList -Name '$Naam' -RecipientFilter `"(ObjectClass -eq 'User')`" -RecipientContainer '$OU' -DomainController '$PDC'")
                )
            }
            $CheckAL = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("Get-AddressList -DomainController '$PDC'"))
            If ($CheckAL | Where-Object { $_.Name -eq $Naam }) {
                Write-Host -NoNewLine " - OK: Address List "; Write-Host -NoNewLine $Naam -ForegroundColor Yellow; Write-Host " created successfully"
            }
            Else {
                $Problem = $true
                BREAK;
            }
        }
        Else {
            Write-Host " - INFO: Address List $Naam already exists" -ForegroundColor Gray
        }
    }
    If ($Problem) {
        Write-Host " - ERROR: An error occurred creating Address List $Naam, please investigate!" -ForegroundColor Red
        $Pause.Invoke()
        Get-Variable -Exclude $global:StartupVariables | Remove-Variable -ErrorAction SilentlyContinue; Script-Disconnect-Server; Invoke-Expression -Command $global:MenuNameCategory
    }
    # ---------------------------------------------------------------------------------------------------
    Write-Host -NoNewLine "Step 3/4 - Creating Offline Address Book" -ForegroundColor Magenta; Write-Host
    # ---------------------------------------------------------------------------------------------------
    $CheckOAB = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("Get-OfflineAddressBook -DomainController '$PDC'"))
    If (!($CheckOAB | Where-Object { $_.Name -eq $Naam })) {
        $AddOAB = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("New-OfflineAddressBook -Name '$Naam' -AddressLists '$NaamGAL' -DomainController '$PDC'"))
        $CheckOAB = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("Get-OfflineAddressBook -DomainController '$PDC'"))
        If ($CheckOAB | Where-Object { $_.Name -eq $Naam }) {
            Write-Host -NoNewLine " - OK: Offline Address Book "; Write-Host -NoNewLine $Naam -ForegroundColor Yellow; Write-Host " created successfully"
        }
        Else {
            Write-Host " - ERROR: An error occurred creating Offline Address Book $Naam, please investigate!" -ForegroundColor Red
            $Pause.Invoke()
            Get-Variable -Exclude $global:StartupVariables | Remove-Variable -ErrorAction SilentlyContinue; Script-Disconnect-Server; Invoke-Expression -Command $global:MenuNameCategory
        }
    }
    Else {
        Write-Host " - INFO: Offline Address Book $Naam already exists" -ForegroundColor Gray
    }
    # -----------------------------------------------------------------------------------------------------------------------------
    Write-Host -NoNewLine "Step 4/4 - Creating policies for Email Addresses and Address Books" -ForegroundColor Magenta; Write-Host
    # -----------------------------------------------------------------------------------------------------------------------------
    $CheckEmailAddressPolicy = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("Get-EmailAddressPolicy -DomainController '$PDC'"))
    If (!($CheckEmailAddressPolicy | Where-Object { $_.Name -eq $Naam })) {
        $AddEmailAddressPolicy = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("New-EmailAddressPolicy -Name '$Naam' -IncludedRecipients `"AllRecipients`" -RecipientContainer '$OU' -EnabledPrimarySMTPAddressTemplate $Emailadres -DomainController '$PDC'"))
        $CheckEmailAddressPolicy = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("Get-EmailAddressPolicy -DomainController '$PDC'"))
        If ($CheckEmailAddressPolicy | Where-Object { $_.Name -eq $Naam }) {
            Write-Host -NoNewLine " - OK: Email Address Policy "; Write-Host -NoNewLine $Naam -ForegroundColor Yellow; Write-Host " created successfully"
        }
        Else {
            Write-Host " - ERROR: An error occurred creating Email Address Policy $Naam, please investigate!" -ForegroundColor Red
            $Pause.Invoke()
            Get-Variable -Exclude $global:StartupVariables | Remove-Variable -ErrorAction SilentlyContinue; Script-Disconnect-Server; Invoke-Expression -Command $global:MenuNameCategory
        }
    }
    Else {
        Write-Host " - INFO: Email Address Policy $Naam already exists" -ForegroundColor Gray
    }
    $CheckAddressBookPolicy = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("Get-AddressBookPolicy -DomainController '$PDC'"))
    If (!($CheckAddressBookPolicy | Where-Object { $_.Name -eq $Naam })) {
        $AddAddressBookPolicy = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("New-AddressBookPolicy -Name '$Naam' -AddressLists '$NaamAllUsers','$NaamAllContacts','$NaamAllGroups' -GlobalAddressList '$NaamGAL' -OfflineAddressBook '$Naam' -RoomList '$NaamAllRooms' -DomainController '$PDC'"))
        $CheckAddressBookPolicy = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("Get-EmailAddressPolicy -DomainController '$PDC'"))
        If ($CheckAddressBookPolicy | Where-Object { $_.Name -eq $Naam }) {
            Write-Host -NoNewLine " - OK: Address Book Policy "; Write-Host -NoNewLine $Naam -ForegroundColor Yellow; Write-Host " created successfully"
        }
        Else {
            Write-Host " - ERROR: An error occurred creating Address Book Policy $Naam, please investigate!" -ForegroundColor Red
        }
    }
    Else {
        Write-Host " - INFO: Address Book Policy $Naam already exists" -ForegroundColor Gray
    }
    Write-Host
    # ==========
    # Finalizing
    # ==========
    $Pause.Invoke()
    Get-Variable -Exclude $global:StartupVariables | Remove-Variable -ErrorAction SilentlyContinue; Script-Disconnect-Server; Invoke-Expression -Command $global:MenuNameCategory
}
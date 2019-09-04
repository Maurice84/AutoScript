Function Messaging-Exchange-Create-Mailboxes-8-Start {
    # =========
    # Execution
    # =========
    # -----------------
    # Declare variables
    # -----------------
    $Overslaan = $null
    $Alias = $SamAccountName
    $Admins = "Dom*dmin*"
    $Fout = $null
    If (!$CSV) {
        Script-Module-SetHeaders -Name ($MyInvocation.MyCommand).Name
    }
    # -------------------------------------------------------------------------------------------------------------
    # Checking and recovering mailbox database when excluded from provisioning (must be required to create mailbox)
    # -------------------------------------------------------------------------------------------------------------
    $IsExcludedFromProvisioning = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("Get-MailboxDatabase -Status"))
    $IsExcludedFromProvisioning = $IsExcludedFromProvisioning | Where-Object { $_.Mounted -eq $true -AND $_.IsExcludedFromProvisioning -eq $true }
    If ($IsExcludedFromProvisioning) {
        Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("Set-MailboxDatabase '$IsExcludedFromProvisioning' -IsExcludedFromProvisioning 0 | Out-Null"))
    }
    # ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Write-Host -NoNewLine "Step 1/2 - Creating mailbox" -ForegroundColor Magenta; If ($MailboxGroep) { Write-Host -NoNewLine " and permission group" -ForegroundColor Magenta }; Write-Host
    # ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    $CheckMailbox = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("Get-Mailbox '$Alias' -DomainController '$PDC' -ErrorAction SilentlyContinue"))
    If (!$CheckMailbox) {
        If ($AddressBookPolicy) {
            $EnableMailbox = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("Enable-Mailbox -Identity '$SamAccountName' -Alias '$Alias' -AddressBookPolicy '$AddressBookPolicy' -DomainController '$PDC'"))
        }
        Else {
            $EnableMailbox = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("Enable-Mailbox -Identity '$SamAccountName' -Alias '$Alias' -DomainController '$PDC'"))
        }
        $CheckMailbox = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("Get-Mailbox '$Alias' -DomainController '$PDC' -ErrorAction SilentlyContinue"))
        If ($CheckMailbox) {
            Write-Host -NoNewLine " - OK: Mailbox "; Write-Host -NoNewLine $DisplayName -ForegroundColor Yellow; Write-Host -NoNewLine " created successfully"; Write-Host
            If ($OUPath -like "*Mailboxes") {
                Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("Set-Mailbox '$SamAccountName' -Type Shared -DomainController '$PDC'"))
                $CheckMailbox = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("Get-Mailbox '$Alias' -DomainController '$PDC' -ErrorAction SilentlyContinue"))
                $CheckMailbox = ($CheckMailbox | Select-Object RecipientType).RecipientType
                If ($CheckMailbox -ne "SharedMailbox") {
                    Write-Host " - ERROR: An error occurred modifying the mailbox to a shared mailbox, please investigate!" -ForegroundColor Magenta
                    Write-Host
                    $Pause.Invoke()
                    Get-Variable -Exclude $global:StartupVariables | Remove-Variable -ErrorAction SilentlyContinue; Script-Disconnect-Server; Invoke-Expression -Command $global:MenuNameCategory
                }
            }
        }
        Else {
            Write-Host " - ERROR: An error occurred creating the mailbox, please investigate!" -ForegroundColor Magenta
            Write-Host
            $Pause.Invoke()
            Get-Variable -Exclude $global:StartupVariables | Remove-Variable -ErrorAction SilentlyContinue; Script-Disconnect-Server; Invoke-Expression -Command $global:MenuNameCategory
        }
    }
    Else {
        Write-Host " - INFO: Mailbox $DisplayName already exists" -ForegroundColor Gray
    }
    Script-Module-ReplicateAD
    $CurrentLanguage = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("Get-MailboxRegionalConfiguration -Identity '$SamAccountName' -DomainController '$PDC' -WarningAction SilentlyContinue"))
    If ((($CurrentLanguage | Select-Object Language).Language).Name -ne $Language) {
        Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("Set-MailboxRegionalConfiguration -Identity '$SamAccountName' -Language '$Language' -DomainController '$PDC' -WarningAction SilentlyContinue"))
        $CurrentLanguage = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("Get-MailboxRegionalConfiguration -Identity '$SamAccountName' -DomainController '$PDC' -WarningAction SilentlyContinue"))
        If ((($CurrentLanguage | Select-Object Language).Language).Name -eq $Language) {
            Write-Host -NoNewLine " - OK: Language successfully set to "; Write-Host -NoNewLine $Taalkeuze -ForegroundColor Yellow; Write-Host
        }
        Else {
            Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("Disable-Mailbox -Identity '$SamAccountName' -DomainController '$PDC' -Confirm:0"))
            Write-Host " - ERROR: An error occurred setting the language of the mailbox (required)! Therefore the mailbox has been deleted. Please recreate the mailbox correctly." -ForegroundColor Red
            $Pause.Invoke()
            Get-Variable -Exclude $global:StartupVariables | Remove-Variable -ErrorAction SilentlyContinue; Script-Disconnect-Server; Invoke-Expression -Command $global:MenuNameCategory
        }
    }
    Else {
        Write-Host " - INFO: Language already set to" $Taalkeuze -ForegroundColor Gray
    }
    $CheckEmailAddressPolicy = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("Get-Mailbox -Identity '$SamAccountName' -DomainController '$PDC'"))
    If (($CheckEmailAddressPolicy | Select-Object EmailAddressPolicyEnabled).EmailAddressPolicyEnabled -eq $true) {
        Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("Set-Mailbox -Identity '$SamAccountName' -EmailAddressPolicyEnabled 0 -DomainController '$PDC'"))
        $CheckEmailAddressPolicy = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("Get-Mailbox -Identity '$SamAccountName' -DomainController '$PDC'"))
        If (($CheckEmailAddressPolicy | Select-Object EmailAddressPolicyEnabled).EmailAddressPolicyEnabled -eq $false) {
            Write-Host " - OK: Email Address Policy removed from mailbox"
        }
        Else {
            Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("Disable-Mailbox -Identity '$SamAccountName' -DomainController '$PDC' -Confirm:0"))
            Write-Host " - ERROR: An error occurred removing the Email Address Policy from the selected mailbox (required)! Therefore the mailbox has been deleted. Please recreate the mailbox correctly." -ForegroundColor Red
            $Pause.Invoke()
            Get-Variable -Exclude $global:StartupVariables | Remove-Variable -ErrorAction SilentlyContinue; Script-Disconnect-Server; Invoke-Expression -Command $global:MenuNameCategory
        }
    }
    Else {
        Write-Host " - INFO: Email Address Policy already removed from mailbox" -ForegroundColor Gray
    }
    $CheckEmailAddress = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("Get-Mailbox -Identity '$SamAccountName' -DomainController '$PDC'"))
    If (($CheckEmailAddress | Select-Object PrimarySmtpAddress).PrimarySmtpAddress.ToString() -ne $Email) {
        $CurrentEmail = ($CheckEmailAddress | Select-Object PrimarySmtpAddress).PrimarySmtpAddress.ToString()
        Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("Set-Mailbox -Identity '$SamAccountName' -PrimarySmtpAddress '$Email' -DomainController '$PDC'"))
        Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("Set-Mailbox -Identity '$SamAccountName' -EmailAddresses @{remove='$CurrentEmail'} -DomainController '$PDC'"))
        $CheckEmailAddress = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("Get-Mailbox -Identity '$SamAccountName' -DomainController '$PDC'"))
        If (($CheckEmailAddress | Select-Object PrimarySmtpAddress).PrimarySmtpAddress.ToString() -eq $Email) {
            Write-Host -NoNewLine " - OK: Primary Email Address successfully set to "; Write-Host -NoNewLine $Email -ForegroundColor Yellow; Write-Host
        }
        Else {
            Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("Disable-Mailbox -Identity '$SamAccountName' -DomainController '$PDC' -Confirm:0"))
            Write-Host " - ERROR: An error occurred setting the Primary Email Address (required)! Therefore the mailbox has been deleted. Please recreate the mailbox correctly." -ForegroundColor Red
            $Pause.Invoke()
            Get-Variable -Exclude $global:StartupVariables | Remove-Variable -ErrorAction SilentlyContinue; Script-Disconnect-Server; Invoke-Expression -Command $global:MenuNameCategory
        }
    }
    Else {
        Write-Host " - INFO: Primary Email Address already set to" $Email -ForegroundColor Gray
    }
    If ($ExtraAddresses) {
        $CheckEmailAddresses = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("Get-Mailbox -Identity '$SamAccountName' -DomainController '$PDC'"))
        ForEach ($Address in $ExtraAddresses) {
            If ($CheckEmailAddresses | Where-Object { $_.EmailAddresses -notlike "*smtp:$Address*" } | Select-Object EmailAddresses) {
                $AddEmailAddress = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("Set-Mailbox -Identity '$SamAccountName' -EmailAddresses @{Add='$Address'} -DomainController '$PDC'"))
                $CheckEmailAddresses = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("Get-Mailbox -Identity '$SamAccountName' -DomainController '$PDC'"))
                If ($CheckEmailAddresses | Where-Object { $_.EmailAddresses -like "*smtp:$Address*" } | Select-Object EmailAddresses) {
                    Write-Host -NoNewLine " - OK: Extra Email Address "; Write-Host -NoNewLine $Address -ForegroundColor Yellow; Write-Host " successfully added"
                }
                Else {
                    Write-Host " - ERROR: An error occurred adding the extra Email Address $Address, please investigate!" -ForegroundColor Red
                    $Fout = $true
                    BREAK
                }
            }
            Else {
                Write-Host " - INFO: Extra Email Address $Address already set" -ForegroundColor Gray
            }
        }
        If ($Fout -eq $true) {
            $Pause.Invoke()
            Get-Variable -Exclude $global:StartupVariables | Remove-Variable -ErrorAction SilentlyContinue; Script-Disconnect-Server; Invoke-Expression -Command $global:MenuNameCategory
        }
    }
    If ($X400) {
        $CheckX400 = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("Get-Mailbox -Identity '$SamAccountName' -DomainController '$PDC'"))
        If (($CheckX400 | Select-Object @{Name = "X400"; Expression = { $_.EmailAddresses | Where-Object { $_.PrefixString -eq "X400" } } }).X400.AddressString -notlike "*$X400*") {
            Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("Set-Mailbox -Identity '$SamAccountName' -EmailAddresses -EmailAddresses @{Add=`"X400:$X400`"} -DomainController '$PDC'"))
            $CheckX400 = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("Get-Mailbox -Identity '$SamAccountName' -DomainController '$PDC'"))
            If (($CheckX400 | Select-Object @{Name = "X400"; Expression = { $_.EmailAddresses | Where-Object { $_.PrefixString -eq "X400" } } }).X400.AddressString -like "*$X400*") {
                Write-Host -NoNewLine " - OK: X400 Address successfully set to "; Write-Host $X400 -ForegroundColor Yellow
            }
            Else {
                Write-Host " - ERROR: An error occurred setting the X400 Address, please investigate!" -ForegroundColor Red
                $Pause.Invoke()
                Get-Variable -Exclude $global:StartupVariables | Remove-Variable -ErrorAction SilentlyContinue; Script-Disconnect-Server; Invoke-Expression -Command $global:MenuNameCategory
            }
        }
        Else {
            Write-Host " - INFO: X400 Address already set" -ForegroundColor Gray
        }
    }
    If ($LegacyExchangeDN) {
        $CheckLegacyExchangeDN = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("Get-Mailbox -Identity '$SamAccountName' -DomainController '$PDC'"))
        If (($CheckLegacyExchangeDN | Select-Object @{Name = "X500"; Expression = { $_.EmailAddresses | Where-Object { $_.PrefixString -eq "X500" } } }).X500.AddressString -notlike "*$LegacyExchangeDN*") {
            Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("Set-Mailbox -Identity '$SamAccountName' -EmailAddresses -EmailAddresses @{Add=`"X500:$LegacyExchangeDN`"} -DomainController '$PDC'"))
            $CheckLegacyExchangeDN = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("Get-Mailbox -Identity '$SamAccountName' -DomainController '$PDC'"))
            If (($CheckLegacyExchangeDN | Select-Object @{Name = "X500"; Expression = { $_.EmailAddresses | Where-Object { $_.PrefixString -eq "X500" } } }).X500.AddressString -like "*$LegacyExchangeDN*") {
                Write-Host -NoNewLine " - OK: X500 Address successfully set to "; Write-Host $LegacyExchangeDN -ForegroundColor Yellow
            }
            Else {
                Write-Host " - ERROR: An error occurred setting the X500 Address, please investigate!" -ForegroundColor Red
                $Pause.Invoke()
                Get-Variable -Exclude $global:StartupVariables | Remove-Variable -ErrorAction SilentlyContinue; Script-Disconnect-Server; Invoke-Expression -Command $global:MenuNameCategory
            }
        }
        Else {
            Write-Host " - INFO: X500 Address already set" -ForegroundColor Gray
        }
    }
    Write-Host
    # --------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Write-Host -NoNewLine "Step 2/2 - Granting permissions to Administrators" -ForegroundColor Magenta; If ($MailboxGroep) { Write-Host -NoNewLine " and permissions groups" -ForegroundColor Magenta }; Write-Host
    # --------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    If ($MailboxRechten -eq "Ja") {
        $CheckMailboxGroep = Script-Connect-Server -Module "ActiveDirectory" -Command ([scriptblock]::Create("Get-ADGroup -Filter * -Server '$PDC' | Where-Object {`$`_.Name -eq '$MailboxGroep'}"))
        If (!$CheckMailboxGroep) {
            Script-Connect-Server -Module "ActiveDirectory" -Command ([scriptblock]::Create(`
                        "New-ADGroup -Name '$MailboxGroep' -SamAccountName '$MailboxGroep' -GroupCategory Security -GroupScope Global -DisplayName '$MailboxGroep' -Path '$GroupOUPath' -Server '$PDC'")
            )
            If ($?) {
                Write-Host -NoNewLine " - OK: Permission group "; Write-Host -NoNewLine $MailboxGroep -ForegroundColor Yellow; Write-Host -NoNewLine " successfully created in "; Write-Host -NoNewLine $global:Subtree -ForegroundColor Yellow; Write-Host
            }
            Else {
                Write-Host " - ERROR: An error occurred creating the permission group $MailboxGroep in $global:Subtree, please investigate!" -ForegroundColor Magenta
                Write-Host
                $Pause.Invoke()
                Get-Variable -Exclude $global:StartupVariables | Remove-Variable -ErrorAction SilentlyContinue; Script-Disconnect-Server; Invoke-Expression -Command $global:MenuNameCategory
            }
        }
        Else {
            Write-Host " - INFO: Permission group $MailboxGroep already exists" -ForegroundColor Gray
        }
        ForEach ($Account in $RechtenArray) {
            $AccountSamAccountName = $Account.SamAccountName
            Script-Connect-Server -Module "ActiveDirectory" -Command ([scriptblock]::Create("Add-ADGroupMember '$MailboxGroep' '$AccountSamAccountName' -Server '$PDC'"))
            If ($?) {
                Write-Host " - OK: Account" $Account.DisplayName "successfully added to the permission group"
            }
            Else {
                Write-Host " - ERROR: An error occurred adding" $Account.DisplayName "to the permission group, please investigate!" -ForegroundColor Magenta
                $Pause.Invoke()
                Get-Variable -Exclude $global:StartupVariables | Remove-Variable -ErrorAction SilentlyContinue; Script-Disconnect-Server; Invoke-Expression -Command $global:MenuNameCategory
            }
        }
    }
    [array]$Groups = Script-Connect-Server -Module "ActiveDirectory" -Command ([scriptblock]::Create("(Get-ADGroup -Filter * -Server '$PDC' | Where-Object {`$`_.Name -like `"$Admins`"}).Name"))
    If ($MailboxRechten -eq "Ja") {
        $Groups += $MailboxGroep
    }
    $GetFullAccessDeny = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("Get-MailboxPermission -Identity '$SamAccountName'"))
    $GetFullAccessDeny = $GetFullAccessDeny | Where-Object { $_.User -like "*$env:username" -AND $_.AccessRights -like "*FullAccess*" -AND $_.Deny -eq $true } | Select-Object User, @{Name = "AccessRight"; Expression = { $_.AccessRights -join "," } }, Deny
    ForEach ($Group in $Groups) {
        $User = $env:userdomain + "\" + $Group
        If ($GetFullAccessDeny | Where-Object { $_.User -eq $User }) {
            $RemoveDeny = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create(`
                        "Remove-MailboxPermission -Identity '$SamAccountName' -User '$User' -AccessRights FullAccess -DomainController '$PDC' -Confirm:0")
            )
        }
        $AddMailboxPermission = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("Add-MailboxPermission -Identity '$SamAccountName' -AccessRights FullAccess -User '$User' -InheritanceType All -AutoMapping 0 -DomainController '$PDC' -Confirm:0 -WarningAction SilentlyContinue"))
        $CheckMailboxPermission = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("Get-MailboxPermission -Identity '$SamAccountName' -User '$User' -DomainController '$PDC'"))
        If ($CheckMailboxPermission | Select-Object User, @{Name = "AccessRight"; Expression = { $_.AccessRights -join "," } }) {
            Write-Host -NoNewLine " - OK: Mailbox FullAccess permission to "; Write-Host -NoNewLine $Group -ForegroundColor Yellow; Write-Host -NoNewLine " successfully granted"; Write-Host
        }
        Else {
            Write-Host " - ERROR: An error occurred granting FullAccess permission to $Group, please investigate!" -ForegroundColor Red
            $Fout = $true
        }
        $AddADPermission = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("Add-ADPermission -Identity '$DistinguishedName' -ExtendedRights Send-As -User '$User' -DomainController '$PDC' -WarningAction SilentlyContinue"))
        $CheckSendAs = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("Get-ADPermission -Identity '$DistinguishedName' -User '$User' -DomainController '$PDC'"))
        If ($CheckSendAs | Select-Object ExtendedRights) {
            Write-Host -NoNewLine " - OK: Mailbox Send-As permission to "; Write-Host -NoNewLine $Group -ForegroundColor Yellow; Write-Host -NoNewLine " successfully granted"; Write-Host
        }
        Else {
            Write-Host " - ERROR: An error occurred granting Send-As permission to $Group, please investigate!" -ForegroundColor Red
            $Fout = $true
        }
    }
    If ($Fout -eq $true) {
        $Pause.Invoke()
        Get-Variable -Exclude $global:StartupVariables | Remove-Variable -ErrorAction SilentlyContinue; Script-Disconnect-Server; Invoke-Expression -Command $global:MenuNameCategory
    }
    # ==========
    # Finalizing
    # ==========
    If (!$CSV) {
        Write-Host
        $Pause.Invoke()
        Get-Variable -Exclude $global:StartupVariables | Remove-Variable -ErrorAction SilentlyContinue; Script-Disconnect-Server; Invoke-Expression -Command $global:MenuNameCategory
    }
}
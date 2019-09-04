Function Messaging-Exchange-Modify-Mailboxes-2-GetData {
    # =========
    # Execution
    # =========
    Script-Module-SetHeaders -CurrentTask $Task -Name ($MyInvocation.MyCommand).Name
    Write-Host -NoNewLine "Loading selected mailbox: "; Write-Host -NoNewLine $Name -ForegroundColor Yellow; Write-Host -NoNewLine "..."
    $global:Mailbox = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("Get-Mailbox -Identity '$SamAccountName' -DomainController '$PDC' -WarningAction SilentlyContinue"))
    # -----------------
    # Declare variables
    # -----------------
    $global:Alias = $Mailbox.Alias
    $global:DistinguishedName = $Mailbox.DistinguishedName
    $global:WeergaveNaam = $Mailbox.DisplayName
    # ------------
    # Mailbox Type
    # ------------
    If ($Mailbox.RecipientTypeDetails -eq "EquipmentMailbox") { $global:SoortMailbox = "Equipment" }
    If ($Mailbox.RecipientTypeDetails -eq "RoomMailbox") { $global:SoortMailbox = "Room" }
    If ($Mailbox.RecipientTypeDetails -eq "SharedMailbox") { $global:SoortMailbox = "Shared" }
    If ($Mailbox.RecipientTypeDetails -eq "UserMailbox") { $global:SoortMailbox = "User" }
    # --------
    # Language
    # --------
    If ($Mailbox.Languages[0] -eq "nl-NL") { $global:Taalinstelling = "Dutch" }
    If ($Mailbox.Languages[0] -eq "en-US") { $global:Taalinstelling = "English" }
    If ($Mailbox.Languages[0] -eq "de-DE") { $global:Taalinstelling = "German" }
    If (!$Mailbox.Languages) { $global:Taalinstelling = "n.a." }
    # ------------------------------
    # Primary & Extra mail addresses
    # ------------------------------
    $global:Emailaddresses = @()
    If (($Mailbox.EmailAddresses).ProxyAddressString) {
        $global:MailboxEmailAddresses = ($Mailbox.EmailAddresses).ProxyAddressString
    }
    Else {
        $global:MailboxEmailAddresses = $Mailbox.EmailAddresses
    }
    ForEach ($global:Address in ($MailboxEmailAddresses | Where-Object { $_ -like "smtp:*" -AND $_ -notlike "*local" })) {
        If ($global:Address.Split(':')[0] -cmatch "SMTP") {
            $IsPrimaryAddress = $true
        }
        Else {
            $IsPrimaryAddress = $false
        }
        $global:EmailadresProperties = @{Name = $Address.Split(':')[1]; Primary = $IsPrimaryAddress }
        $global:EmailadresObject = New-Object PSObject -Property $EmailadresProperties
        $global:Emailaddresses += $EmailadresObject
    }
    $global:PrimaryEmailaddress = $global:Emailaddresses | Where-Object { $_.Primary -eq $true }
    $global:ExtraEmailaddresses = $global:Emailaddresses | Where-Object { $_.Primary -eq $false }
    $global:PrimairEmailadres = $global:PrimaryEmailaddress.Name
    If ($global:ExtraEmailaddresses.Count -ge 4) {
        $global:Emailadressen = [string]$global:ExtraEmailaddresses.Count + " mail addresses"
    }
    Else {
        $global:Emailadressen = $global:ExtraEmailaddresses.Name -join ', '
    }    
    If (!$global:ExtraEmailaddresses) {
        $global:Emailadressen = "n.a."
    }
    # -----------------
    # Emailadres beleid
    # -----------------
    If ($Mailbox.EmailAddressPolicyEnabled -eq $false) {
        $global:EmailadresBeleid = "Disabled"
    }
    Else {
        $global:EmailadresBeleid = "Enabled"
    }
    # -------------------
    # Address Book Policy
    # -------------------
    $GetAddressBookPolicy = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("Get-AddressBookPolicy -DomainController '$PDC'"))
    $GetAddressBookPolicy = $GetAddressBookPolicy | Sort-Object Name
    $global:EmailaddressBookPolicies = @()
    ForEach ($AddressBookPolicy in $GetAddressBookPolicy) {
        If ($AddressBookPolicy.Name -eq $Mailbox.AddressBookPolicy.Name) {
            $global:EnabledAddressBookPolicy = $true
        }
        Else {
            $global:EnabledAddressBookPolicy = $false
        }
        $global:EmailadresboekProperties = @{Name = $AddressBookPolicy.Name; Enabled = $EnabledAddressBookPolicy }
        $global:EmailadresboekObject = New-Object PSObject -Property $EmailadresboekProperties
        $global:EmailaddressBookPolicies += $EmailadresboekObject
    }
    If ($Mailbox.AddressBookPolicy.Name) {
        $global:AdresboekBeleid = ($global:EmailaddressBookPolicies | Where-Object { $_.Enabled -eq $true } | Select-Object Name).Name
    }
    Else {
        $global:AdresboekBeleid = "n.a."
    }
    # -------------------------------
    # Mailbox permission: Full Access
    # -------------------------------
    $global:CurrentFullAccess = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("Get-MailboxPermission -Identity '$Alias' -DomainController '$PDC'"))
    $global:CurrentFullAccess = ($global:CurrentFullAccess | Where-Object { ($_.IsInherited -eq $false) -AND ($_.User -notlike "*ELF") } | Select-Object User).User
    If ($global:CurrentFullAccess) {
        $global:FullAccess = @()
        ForEach ($global:User in $global:CurrentFullAccess) {
            If ($User.RawIdentity) {
                $global:User = ($User.RawIdentity).Split("\")[1]
            }
            Else {
                $global:User = $global:User.Split("\")[1]
            }
            If ($global:User) {
                $global:FullAccess += Script-Connect-Server -Module "ActiveDirectory" -Command ([scriptblock]::Create(`
                            "(Get-ADObject -Filter `"(SamAccountName -eq '$User')`" -Properties * | Select-Object Name).Name")
                )
            }
        }
        [array]$global:FullAccess = $global:FullAccess | Sort-Object
    }
    If ($global:FullAccess.Length -ge 6) {
        $global:VolledigeToegang = [string]$global:FullAccess.Length + " domain accounts/groups"
    }
    If ($global:FullAccess.Length -lt 5) {
        $global:VolledigeToegang = $global:FullAccess -join ', '
    }
    If (!$global:FullAccess) {
        $global:VolledigeToegang = "n.a."
    }
    # ---------------------------
    # Mailbox permission: Send-As
    # ---------------------------
    $global:CurrentSendAs = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("Get-ADPermission -Identity '$DistinguishedName' -DomainController '$PDC'"))
    $global:CurrentSendAs = ($global:CurrentSendAs | Where-Object { ($_.ExtendedRights -like "*Send-As*") -AND ($_.User -notlike "*ELF") } | Select-Object User).User
    If ($global:CurrentSendAs) {
        $global:SendAs = @()
        ForEach ($global:User in $global:CurrentSendAs) {
            If ($User.RawIdentity) {
                $global:User = ($User.RawIdentity).Split("\")[1]
            }
            Else {
                $global:User = $global:User.Split("\")[1]
            }
            If ($global:User) {
                $global:SendAs += Script-Connect-Server -Module "ActiveDirectory" -Command ([scriptblock]::Create("(Get-ADObject -Filter `"(SamAccountName -eq '$User')`" -Properties * | Select-Object Name).Name"))
            }
        }
        [array]$global:SendAs = $global:SendAs | Sort-Object
    }
    If ($global:SendAs.Length -ge 6) {
        $global:VerzendenAls = [string]$global:SendAs.Length + " domain accounts/groups"
    } 
    If ($global:SendAs.Length -lt 5) {
        $global:VerzendenAls = $global:SendAs -join ', '
    } 
    If (!$global:SendAs) {
        $global:VerzendenAls = "n.a."
    }
    # ----------------------------------
    # Mailbox permission: Send On Behalf
    # ----------------------------------
    $global:CurrentSendOnBehalf = ($Mailbox | Select-Object @{Name = "GrantSendOnBehalfTo"; Expression = { $_.GrantSendOnBehalfTo } }).GrantSendOnBehalfTo
    If ($CurrentSendOnBehalf) {
        $global:SendOnBehalf = @()
        ForEach ($global:User in $CurrentSendOnBehalf) {
            If ($User.Name) {
                $global:User = ($User.Name)
            }
            Else {
                $global:User = $User.Split('/')[-1]
            }
            $global:SendOnBehalf += $User
        }
        [array]$global:SendOnBehalf = $global:SendOnBehalf | Sort-Object
    }
    If ($global:SendOnBehalf.Length -ge 6) {
        $global:VerzendenNamens = [string]$global:SendOnBehalf.Length + " domain accounts/groups"
    } 
    If ($global:SendOnBehalf.Length -lt 5) {
        $global:VerzendenNamens = $global:SendOnBehalf -join ', '
    }
    If (!$global:SendOnBehalf) {
        $global:VerzendenNamens = "n.a."
    }
    # -------------------------------------------------------------
    # Mailbox information (size, item count, lastlogon/logoff time)
    # -------------------------------------------------------------
    $global:MailboxStat = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("Get-MailboxStatistics -Identity '$Alias' -DomainController '$PDC' -WarningAction SilentlyContinue"))
    If ($MailboxStat) {
        If ($MailboxStat.LastLogonTime) {
            $global:LastLogon = $MailboxStat.LastLogonTime.ToLongDateString() + " " + $MailboxStat.LastLogonTime.ToShortTimeString()
        }
        Else {
            $global:LastLogon = "n.a."
        }
        If ($MailboxStat.LastLogoffTime) {
            $global:LastLogoff = $MailboxStat.LastLogoffTime.ToLongDateString() + " " + $MailboxStat.LastLogoffTime.ToShortTimeString()
        }
        Else {
            $global:LastLogoff = "n.a."
            If ($LastLogon -notlike "n.a.") {
                $global:LastLogoff = "Mailbox currently in use"
            }
        }
        [string]$global:TotalItemSize = $MailboxStat.TotalItemSize
        [int]$global:TotalItemSize = "{0:0}" -f ([double]$TotalItemSize.Split(' ')[2].Replace('(', '').Replace(',', '') / 1MB)
        If (([string]$TotalItemSize).Length) {
            If (([string]$TotalItemSize).Length -ge 4) {
                [string]$global:TotalItemSize = "{0:F2}" -f ($TotalItemSize / 1024) + (" GB")
            }
            Else {
                [string]$global:TotalItemSize = [string]$TotalItemSize + " MB"
            }
        }
        Else {
            $global:TotalItemSize = "n.a."
        }
        $global:ItemCount = $MailboxStat.ItemCount
    }
    Else {
        $global:LastLogon = "n.a."
        $global:LastLogoff = "n.a."
        $global:TotalItemSize = "n.a."
        $global:ItemCount = "n.a."
    }
    If ($Mailbox.ProhibitSendQuota -ne "Unlimited") {
        [string]$global:ProhibitSendQuota = $Mailbox.ProhibitSendQuota
        [int]$global:ProhibitSendQuota = "{0:0}" -f ([double]$ProhibitSendQuota.Split(' ')[2].Replace('(', '').Replace(',', '') / 1MB)
        $global:Quota = " (" + ("{0:F0}" -f ($ProhibitSendQuota / 1024)) + " GB max)"
    }
    If ($MailboxStat.DatabaseProhibitSendReceiveQuota) {
        [string]$global:DatabaseProhibitSendReceiveQuota = $MailboxStat.DatabaseProhibitSendReceiveQuota
        [int]$global:DatabaseProhibitSendReceiveQuota = "{0:0}" -f ([double]$DatabaseProhibitSendReceiveQuota.Split(' ')[2].Replace('(', '').Replace(',', '') / 1MB)
        $global:Quota = " (" + ("{0:F0}" -f ($DatabaseProhibitSendReceiveQuota / 1024)) + " GB max)"
    }
    # ==========
    # Finalizing
    # ==========
    Invoke-Expression -Command $FunctionTaskNames[[int]$FunctionTaskNames.IndexOf(($MyInvocation.MyCommand).Name) + 1]
}
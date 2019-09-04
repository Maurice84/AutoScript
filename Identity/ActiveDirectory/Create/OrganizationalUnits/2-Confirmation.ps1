Function Identity-ActiveDirectory-Create-OrganizationalUnits-2-Confirmation {
    # ============
    # Declarations
    # ============
    $global:UPNSuffixes = @()
    # =========
    # Execution
    # =========
    Script-Module-SetHeaders -Name ($MyInvocation.MyCommand).Name
    # -----------------
    # Declare variables
    # -----------------
    $DomainDN = Script-Connect-Server -Module "ActiveDirectory" -Command ([scriptblock]::Create("(Get-ADDomain).DistinguishedName"))
    $GetUPNSuffixes = Script-Connect-Server -Module "ActiveDirectory" -Command ([scriptblock]::Create(`
        "`$UPNDN = `"cn=Partitions,cn=Configuration,$DomainDN`";
        Get-ADObject -Identity `$UPNDN -Properties UPNSuffixes | Select-Object -ExpandProperty UPNSuffixes")
    )
    ForEach ($Object in $GetUPNSuffixes) {
        $global:UPNSuffixes += $Object
    }
    $global:UPNSuffixes += (Get-WmiObject Win32_ComputerSystem).Domain
    $global:UPNSuffixes = $global:UPNSuffixes | Sort-Object
    $global:OrgUnitRoot = $DomainDN
    $global:Klant = $global:Customer
    $global:StdApp1 = "App-Office"
    $global:StdDoc1 = "Doc-Data"
    $global:GroupUsers = "RDS-Users"
    $global:OrgUnitKlant = "OU=" + $global:Customer + "," + $OrgUnitRoot
    $global:OrgUnitAccounts = "OU=Accounts," + $OrgUnitKlant
    $global:OrgUnitGroups = "OU=Groups," + $OrgUnitKlant
    $global:OrgUnitApps = "OU=Applications," + $OrgUnitKlant
    $global:OrgUnitDocs = "OU=Documents," + $OrgUnitKlant
    $global:OrgUnitUsers = "OU=Users," + $OrgUnitAccounts
    $global:Property = New-Object PSObject
    $Property | Add-Member -type NoteProperty -Name "Name" -Value $global:Customer
    $Property | Add-Member -type NoteProperty -Name "DN" -Value $OrgUnitRoot
    [array]$global:Objects += $Property
    $global:KlantObjects = "Accounts", "Applications", "Documents", "Groups", "Printers"
    ForEach ($global:KlantOU in $KlantObjects) {
        $global:Property = New-Object PSObject
        $Property | Add-Member -type NoteProperty -Name "Name" -Value $KlantOU
        $Property | Add-Member -type NoteProperty -Name "RootOU" -Value $global:Customer
        $Property | Add-Member -type NoteProperty -Name "DN" -Value $OrgUnitKlant
        [array]$global:Objects += $Property
    }
    $global:AccountObjects = "Contacts", "Disabled", "External", "Mailboxes", "Service", "Test", "Users", "Webmail"
    $global:AccountObjects += "Admins"
    $global:GroupAdmins = "Admins"
    ForEach ($global:KlantOU in $AccountObjects) {
        $global:Property = New-Object PSObject
        $Property | Add-Member -type NoteProperty -Name "Name" -Value $KlantOU
        $Property | Add-Member -type NoteProperty -Name "RootOU" -Value ($global:Customer + "\Accounts")
        $Property | Add-Member -type NoteProperty -Name "DN" -Value $OrgUnitAccounts
        [array]$global:Objects += $Property
    }
    $global:GroupObjects = "Mailboxes"
    ForEach ($global:KlantOU in $GroupObjects) {
        $global:Property = New-Object PSObject
        $Property | Add-Member -type NoteProperty -Name "Name" -Value $KlantOU
        $Property | Add-Member -type NoteProperty -Name "RootOU" -Value ($global:Customer + "\Groups")
        $Property | Add-Member -type NoteProperty -Name "DN" -Value $OrgUnitGroups
        [array]$global:Objects += $Property
    }
    Write-Host
    Write-Host "  The following Organizational Units will be created:"
    ForEach ($global:KlantOU in $Objects) {
        If ($KlantOU.RootOU) {
            $global:Subtree = "- " + $KlantOU.RootOU + "\" + $KlantOU.Name
            Write-Host ("  " + $global:Subtree) -ForegroundColor Yellow
        }
    }
    Write-Host
    Do {
        Write-Host -NoNewLine "Would you also like to create the following groups: "; Write-Host -NoNewLine $StdApp1 -ForegroundColor Yellow; Write-Host -NoNewLine ", "; Write-Host -NoNewLine $StdDoc1 -ForegroundColor Yellow; Write-Host -NoNewLine " & "; Write-Host -NoNewLine $GroupUsers -ForegroundColor Yellow; Write-Host -NoNewLine "? (Y/N): "
        [string]$InputChoice = Read-Host
        $InputKey = @("Y", "N") -contains $InputChoice
        If (!$InputKey) {
            Write-Host "Please use the letters above as input" -ForegroundColor Red
        }
    } Until ($InputKey)
    Switch ($InputChoice) {
        "Y" {
            $global:CreateGroups = $true
        }
    }
    # ==========
    # Finalizing
    # ==========
    Invoke-Expression -Command $FunctionTaskNames[[int]$FunctionTaskNames.IndexOf(($MyInvocation.MyCommand).Name) + 1]
}
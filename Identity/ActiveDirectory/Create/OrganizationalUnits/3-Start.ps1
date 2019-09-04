Function Identity-ActiveDirectory-Create-OrganizationalUnits-3-Start {
    # =========
    # Execution
    # =========
    Script-Module-SetHeaders -CurrentTask $Task -Name ($MyInvocation.MyCommand).Name
    # -------------------
    # Creating UPN Suffix
    # -------------------
    Write-Host -NoNewLine "Step 1/3 Creating UPN suffix" -ForegroundColor Magenta; Write-Host
    If ($global:UPNSuffixes -notcontains $global:CustomerUPNSuffix) {
        $AddUPNSuffix = Script-Connect-Server -Module "ActiveDirectory" -Command ([scriptblock]::Create(`
            "Get-ADForest | Set-ADForest -UPNSuffixes @{Add='$global:CustomerUPNSuffix'}")
        )
        $CheckUPNSuffix = Script-Connect-Server -Module "ActiveDirectory" -Command ([scriptblock]::Create(
            "Get-ADForest | Select-Object UPNSuffixes")
        )
        $CheckUPNSuffix = $CheckUPNSuffix | Where-Object { $_.UPNSuffixes -like "*$global:CustomerUPNSuffix*" }
        If ($CheckUPNSuffix) {
            Write-Host -NoNewLine "- OK: UPN Suffix "; Write-Host -NoNewLine $global:CustomerUPNSuffix -ForegroundColor Yellow; Write-Host " created successfully"
        }
        Else {
            Write-Host "- ERROR: An error occurred creating the UPN Suffix, please investigate!" -ForegroundColor Red
            $Problem = $true
        }
    }
    Else {
        Write-Host "- INFO: $global:CustomerUPNSuffix already exists" -ForegroundColor Gray
    }
    # --------------------------------------
    # Creating customer Organizational Units
    # --------------------------------------
    if (!$Problem) {
        Write-Host; Write-Host -NoNewLine "Step 2/3 Creating customer Organizational Units" -ForegroundColor Magenta; Write-Host
        ForEach ($KlantOU in $global:Objects) {
            $KlantOUDN = $KlantOU.DN
            $KlantOUName = $KlantOU.Name
            $KlantOURootOU = $KlantOU.RootOU
            If ($KlantOURootOU) {
                $global:Subtree = $KlantOURootOU + "\" + $KlantOUName
            }
            Else {
                $global:Subtree = $KlantOUName
            }
            $CheckOU = Script-Connect-Server -Module "ActiveDirectory" -Command ([scriptblock]::Create(`
                "Get-ADOrganizationalUnit -Filter * | Where-Object {`$`_.DistinguishedName -eq (`"OU=`"+'$KlantOUName'+`",`"+'$KlantOUDN')}")
            )
            If (!$CheckOU) {
                $NewOU = Script-Connect-Server -Module "ActiveDirectory" -Command ([scriptblock]::Create(`
                    "New-ADOrganizationalUnit -Name '$KlantOUName'`
                    -Path '$KlantOUDN'`
                    -ProtectedFromAccidentalDeletion `$False")
                )
                $CheckOU = Script-Connect-Server -Module "ActiveDirectory" -Command ([scriptblock]::Create(`
                    "Get-ADOrganizationalUnit -Filter * | Where-Object {`$`_.DistinguishedName -eq (`"OU=`"+'$KlantOUName'+`",`"+'$KlantOUDN')}")
                )
                If ($CheckOU) {
                    Write-Host -NoNewLine "- OK: Organizational Unit "; Write-Host -NoNewLine $global:Subtree -ForegroundColor Yellow; Write-Host " created successfully"
                }
                Else {
                    Write-Host "- ERROR: An error occurred creating Organizational Unit $global:Subtree, please investigate!" -ForegroundColor Red
                    $Problem = $true
                }
            }
            Else {
                Write-Host "INFO: $global:Subtree already exists" -ForegroundColor Gray
            }
        }
        Script-Connect-Server -Module "ActiveDirectory" -Command ([scriptblock]::Create(
            "Set-ADOrganizationalUnit '$global:OrgUnitKlant' -Description '$global:Customer'"))
    }
    # ------------------------
    # Creating customer groups
    # ------------------------
    if (!$Problem) {
        Write-Host; Write-Host -NoNewLine "Step 3/3 Creating customer groups" -ForegroundColor Magenta; Write-Host
        $CheckAdminGroup = Script-Connect-Server -Module "ActiveDirectory" -Command ([scriptblock]::Create(
            "(Get-ADGroup -Filter * | Where-Object {`$`_.Name -eq '$global:GroupAdmins'})")
        )
        If (!$CheckGroupAdmins) {
            $AddAdminGroup = Script-Connect-Server -Module "ActiveDirectory" -Command ([scriptblock]::Create(
                "New-ADGroup -Path '$global:OrgUnitGroups'`
                -Name '$global:GroupAdmins'`
                -SamAccountName '$global:GroupAdmins'`
                -GroupCategory Security -GroupScope Global")
            )
            $CheckAdminGroup = Script-Connect-Server -Module "ActiveDirectory" -Command ([scriptblock]::Create(
                "(Get-ADGroup -Filter * | Where-Object {`$`_.Name -eq '$global:GroupAdmins'})"))
            If ($CheckAdminGroup) {
                Write-Host -NoNewLine "- OK: Admins group "; Write-Host -NoNewLine $global:GroupAdmins -ForegroundColor Yellow; Write-Host " created successfully"
            }
            Else {
                Write-Host "- ERROR: An error occurred creating Admins group $global:GroupAdmins, please investigate!" -ForegroundColor Red
                $Problem = $true
            }
        }
        Else {
            Write-Host "- INFO: $GroupAdmins already exists" -ForegroundColor Gray
        }
    }
    If (!$Problem -AND $global:CreateGroups) {
        $CheckUsersGroup = Script-Connect-Server -Module "ActiveDirectory" -Command ([scriptblock]::Create(
            "Get-ADGroup -Filter * | Where-Object {`$`_.Name -eq '$global:GroupUsers'}"))
        If (!$CheckUsersGroup) {
            $AddUsersGroup = Script-Connect-Server -Module "ActiveDirectory" -Command ([scriptblock]::Create(
                "New-ADGroup -Path '$global:OrgUnitGroups'
                -Name '$global:GroupUsers'`
                -SamAccountName '$global:GroupUsers'`
                -GroupCategory Security -GroupScope Global")
            )
            $CheckUsersGroup = Script-Connect-Server -Module "ActiveDirectory" -Command ([scriptblock]::Create(
                "Get-ADGroup -Filter * | Where-Object {`$`_.Name -eq '$global:GroupUsers'}")
            )
            If ($CheckUsersGroup) {
                Write-Host -NoNewLine "- OK: Users group "; Write-Host -NoNewLine $global:GroupUsers -ForegroundColor Yellow; Write-Host " created successfully"
            }
            Else {
                Write-Host "- ERROR: An error occurred creating Users group $global:GroupUsers, please investigate!" -ForegroundColor Red
            }
        }
        Else {
            Write-Host "- INFO: $global:GroupUsers already exists" -ForegroundColor Gray
        }
        $CheckStdApp1 = Script-Connect-Server -Module "ActiveDirectory" -Command ([scriptblock]::Create(
            "Get-ADGroup -Filter * | Where-Object {`$`_.Name -eq '$StdApp1'}")
        )
        If (!$CheckStdApp1) {
            $AddStdApp1 = Script-Connect-Server -Module "ActiveDirectory" -Command ([scriptblock]::Create(
                "New-ADGroup -Path '$global:OrgUnitApps'`
                -Name '$StdApp1'`
                -SamAccountName '$StdApp1'`
                -GroupCategory Security`
                -GroupScope Global")
            )
            $CheckStdApp1 = Script-Connect-Server -Module "ActiveDirectory" -Command ([scriptblock]::Create(
                "Get-ADGroup -Filter * | Where-Object {`$`_.Name -eq '$StdApp1'}")
            )
            If ($CheckStdApp1) {
                Write-Host -NoNewLine "- OK: Application group "; Write-Host -NoNewLine $StdApp1 -ForegroundColor Yellow; Write-Host " created successfully"
            }
            Else {
                Write-Host "- ERROR: An error occurred creating Application group $StdApp1, please investigate!" -ForegroundColor Red
            }
        }
        Else {
            Write-Host "- INFO: $StdApp1 already exists" -ForegroundColor Gray
        }
        $CheckStdDoc1 = Script-Connect-Server -Module "ActiveDirectory" -Command ([scriptblock]::Create(
            "Get-ADGroup -Filter * | Where-Object {`$`_.Name -eq '$StdDoc1'}")
        )
        If (!$CheckStdDoc1) {
            $AddStdDoc1 = Script-Connect-Server -Module "ActiveDirectory" -Command ([scriptblock]::Create(
                "New-ADGroup -Path '$global:OrgUnitDocs'`
                -Name '$StdDoc1'`
                -SamAccountName '$StdDoc1'`
                -GroupCategory Security`
                -GroupScope Global")
            )
            $CheckStdDoc1 = Script-Connect-Server -Module "ActiveDirectory" -Command ([scriptblock]::Create(
                "Get-ADGroup -Filter * | Where-Object {`$`_.Name -eq '$StdDoc1'}")
            )
            If ($CheckStdDoc1) {
                Write-Host -NoNewLine "- OK: Document group "; Write-Host -NoNewLine $StdDoc1 -ForegroundColor Yellow; Write-Host " created successfully"
            }
            Else {
                Write-Host "- ERROR: An error occurred creating Document group $StdDoc1, please investigate!" -ForegroundColor Red
            }
        }
        Else {
            Write-Host "INFO: $StdDoc1 already exists" -ForegroundColor Gray
        }
    }
    Else {
        Write-Host "- INFO: Skipping creating groups $StdApp1, $StdDoc1 & $global:GroupUsers" -ForegroundColor Gray
    }
    # ==========
    # Finalizing
    # ==========
    $Pause.Invoke()
    Get-Variable -Exclude $global:StartupVariables | Remove-Variable -ErrorAction SilentlyContinue; Script-Disconnect-Server; Invoke-Expression -Command $global:MenuNameCategory
}
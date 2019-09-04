Function Identity-ActiveDirectory-Create-Accounts-WithAdminRights-2-Start {
    # ============
    # Declarations
    # ============
    $AdminAccounts = @()
    # =========
    # Execution
    # =========
    Script-Module-SetHeaders -CurrentTask $Task -Name ($MyInvocation.MyCommand).Name
    $AdminAccounts += New-Object PSObject -Property @{
        GivenName      = "Maurice"
        Surname        = "Heikens"
        SamAccountName = "adm-m.heikens"
        Email          = "maurice.heikens@sogeti.com"
    }
    $Groups = Script-Connect-Server -Module "ActiveDirectory" -Command ([scriptblock]::Create(`
                "Get-ADPrincipalGroupMembership Administrator | Select-Object Name | Sort-Object Name")
    )
    $GroupAdmins = Script-Connect-Server -Module "ActiveDirectory" -Command ([scriptblock]::Create(`
                "Get-ADGroup -Filter * | Where-Object {`$`_.Name -like `"*-Admins`"}")
    )
    $Filter = Script-Connect-Server -Module "ActiveDirectory" -Command ([scriptblock]::Create(`
                "(Get-WmiObject Win32_ComputerSystem).Domain + `"/`" + (Get-ADOrganizationalUnit -Filter * | Where-Object {`$`_.Name -like `"*-*`"}).Name")
    )
    $OrgUnitKlant = Script-Connect-Server -Module "ActiveDirectory" -Command ([scriptblock]::Create(`
                "Get-ADOrganizationalUnit -Filter * -Properties * | Where-Object {`$`_.CanonicalName -eq '$Filter'}")
    )
    $OrgUnitKlantDN = $OrgUnitKlant.DistinguishedName
    $OrgUnitAdmins = Script-Connect-Server -Module "ActiveDirectory" -Command ([scriptblock]::Create(`
                "Get-ADOrganizationalUnit -Filter * | Where-Object {`$`_.DistinguishedName -eq `"OU=Admins,OU=Accounts,`"+'$OrgUnitKlantDN'}")
    )
    If (!$OrgUnitAdmins) {
        $OrgUnitAdmins = Script-Connect-Server -Module "ActiveDirectory" -Command ([scriptblock]::Create(`
                    "(Get-ADUser -Identity Administrator -Properties DistinguishedName,CN | Select-Object @{n='ParentContainer';e={`$`_.DistinguishedName -replace '^.+?,(CN|OU.+)','`$1'}}).ParentContainer")
        )
    }
    $FileCSVP = "C:\Admin-Accounts-"
    If ($OrgUnitKlant.Description) {
        $Klantnaam = $OrgUnitKlant.Description
        $FileCSVP += $Klantnaam.Replace(" ", "-")
    }
    Else {
        $FileCSVP += Script-Connect-Server -Module "ActiveDirectory" -Command ([scriptblock]::Create("(Get-ADDomain).NetBIOSName"))
    }
    $FileCSVP += ".csv"
    # --------------------------------------------------------------------------------------------------------------
    # Declare the name and path of the CSV-file and remove the file if it's present (Add-Content does not overwrite)
    # --------------------------------------------------------------------------------------------------------------
    Remove-Item ($FileCSVP) -ErrorAction SilentlyContinue
    # ---------------------------------------------------
    # Declare the header and add it to the empty CSV-file
    # ---------------------------------------------------
    $ExportHeader = '"Name","UserPrincipalName","Password"'
    Add-Content ($FileCSVP) $ExportHeader
    # -------------------------------------------------------------------------------------------------------
    Write-Host "Creating Fine Grained Password Policy for Administrators (180 days)" -ForegroundColor Magenta
    # -------------------------------------------------------------------------------------------------------
    $CheckPasswordPolicy = Script-Connect-Server -Module "ActiveDirectory" -Command ([scriptblock]::Create(`
                "Try {Get-ADFineGrainedPasswordPolicy `"Password-Policy-Admins`" -Server '$PDC'} Catch {}")
    )
    If (!$CheckPasswordPolicy) {
        Script-Connect-Server -Module "ActiveDirectory" -Command ([scriptblock]::Create(`
                    "New-ADFineGrainedPasswordPolicy -Name `"Password-Policy-Admins`"`
             -Description `"Password-Policy-Admins`"`
             -Precedence 1`
             -MinPasswordLength 8`
             -PasswordHistoryCount 24`
             -ComplexityEnabled `$true`
             -ReversibleEncryptionEnabled `$true`
             -MinPasswordAge `"0.00:00:00`"`
             -MaxPasswordAge `"180.00:00:00`"`
             -Server '$PDC'")
        )
        If ($?) {
            Write-Host " - OK: Policy created"
        }
        Else {
            Write-Host " - ERROR: Policy could not be created, please investigate!"
            $Problem = $true
        }
    }
    Else {
        Write-Host " - INFO: Policy already exists" -ForegroundColor Gray
    }
    $GroupAdminName = @()
    $GroupAdminName = ($GroupAdmins | Select-Object Name).Name    
    $CheckPolicy = Script-Connect-Server -Module "ActiveDirectory" -Command ([scriptblock]::Create(`
                "Get-ADFineGrainedPasswordPolicySubject `"Password-Policy-Admins`" -Server '$PDC' | Where-Object {`$`_.Name -like `"*-Admins`"}")
    )
    if (!$CheckPolicy) {
        Script-Connect-Server -Module "ActiveDirectory" -Command ([scriptblock]::Create(`
                    "Add-ADFineGrainedPasswordPolicySubject -Identity `"Password-Policy-Admins`" -Subjects '$GroupAdminName' -Server '$PDC'")
        )
        If ($?) {
            Write-Host " - OK:" $GroupAdminName "added to the Administrator password policy"
        }
        Else {
            Write-Host " - ERROR:" $GroupAdminName "could not be added to the Administrator password policy, please investigate!"
            $Problem = $true
        }
    }
    Else {
        Write-Host " - INFO:" $GroupAdminName "already added to the Administrator password policy" -ForegroundColor Gray
    }
    if (!$Problem) {
        Write-Host
        # -----------------------------------------------------------------------------------------
        Write-Host "Creating Administrator accounts with random passwords" -ForegroundColor Magenta
        # -----------------------------------------------------------------------------------------
        ForEach ($Account in $AdminAccounts) {
            Script-Module-SetPassword
            $SamAccountName = $Account.SamAccountName
            $Email = $Account.Email
            $GivenName = $Account.GivenName
            $Surname = $Account.Surname
            $Name = $GivenName + " " + $Surname
            $UPN = $SamAccountName + '@' + (Get-WmiObject Win32_ComputerSystem).Domain
            $CheckUser = Script-Connect-Server -Module "ActiveDirectory" -Command ([scriptblock]::Create("Get-ADUser -Filter {SamAccountName -eq '$SamAccountName'}"))
            If (!$CheckUser) {
                Script-Connect-Server -Module "ActiveDirectory" -Command ([scriptblock]::Create(`
                            "New-ADUser -Name '$Name'`
                    -Path '$OrgUnitAdmins'`
                    -GivenName '$GivenName'`
                    -Surname '$Surname'`
                    -DisplayName '$Name'`
                    -UserPrincipalName '$UPN'`
                    -SamAccountName '$SamAccountName'`
                    -Email '$Email'`
                    -Description `"Admin Account`"`
                    -AccountPassword (ConvertTo-SecureString '$Password' -AsPlainText -Force)`
                    -Server '$PDC'`
                    -PasswordNeverExpires `$False`
                    -Enabled `$True")
                )
                If ($?) {
                    Write-Host -NoNewLine " - OK: Account "; Write-Host -NoNewLine $Name -ForegroundColor Yellow; Write-Host -NoNewLine " created with password "; Write-Host $Password -ForegroundColor Cyan
                    Script-Module-ReplicateAD
                    # ----------------------------------------------
                    # Adding the account to the Administrator groups
                    # ----------------------------------------------
                    If ($GroupAdmins) {
                        Script-Connect-Server -Module "ActiveDirectory" -Command ([scriptblock]::Create("Add-ADGroupMember -Identity '$GroupAdminName' -Members '$SamAccountName'"))
                    }
                    $CurrentGroups = Script-Connect-Server -Module "ActiveDirectory" -Command ([scriptblock]::Create("Get-ADPrincipalGroupMembership '$SamAccountName' | Select-Object Name | Sort-Object Name"))
                    ForEach ($Group in $Groups) {
                        If (($CurrentGroups | Select-Object Name | % { $_.Name }) -notcontains ($Group | % { $_.Name })) {
                            $GroupName = ($Group | Select-Object Name).Name
                            Script-Connect-Server -Module "ActiveDirectory" -Command ([scriptblock]::Create("Add-ADGroupMember -Identity '$GroupName' -Members '$SamAccountName'"))
                            If (!$?) {
                                Write-Host -NoNewLine " - ERROR: An error occurred adding the account to " -ForegroundColor Red; Write-Host -NoNewLine $Group.Name -ForegroundColor Red; Write-Host -NoNewLine ", please investigate!" -ForegroundColor Red; Write-Host
                                $Pause.Invoke()
                                BREAK
                            }
                        }
                    }
                    # -----------------------------
                    # Add each line to the CSV-file
                    # -----------------------------
                    $ExportValues = '"' + $Name + '","' + $UPN + '","' + $Password + '"'
                    Add-Content ($FileCSVP) $ExportValues
                }
                Else {
                    Write-Host " - ERROR: An error occurred creating account $Name, please investigate!" -ForegroundColor Red
                    $Pause.Invoke()
                    BREAK
                }
            }
            Else {
                Write-Host " - INFO: Account $Name already exists" -ForegroundColor Gray
            }
        }
        If (Test-Path $FileCSVP) {
            Write-Host -NoNewLine " - OK: Successfully exported the Administrator account(s) to CSV-file: "; Write-Host $FileCSVP -ForegroundColor Yellow
            # ----------------------------------
            # Convert the CSV-file to Excel-file
            # ----------------------------------
            Script-Convert-CSV-to-Excel -File $FileCSVP -Category "Accounts" -Silent $true
            If (Test-Path $FileExcel) {
                Write-Host -NoNewLine " - OK: Successfully converted the CSV-file to Excel-file:"; Write-Host $FileExcel.Split('\')[-1] -ForegroundColor Yellow
            }
            Else {
                Write-Host " - ERROR: Could not convert the CSV-file to Excel-file, please investigate!" -ForegroundColor Red
            }
        }
        Else {
            Write-Host " - ERROR: Could not export the Administrator accounts to CSV-file, please investigate!" -ForegroundColor Red
        }
    }
    # ==========
    # Finalizing
    # ==========
    $Pause.Invoke()
    Get-Variable -Exclude $global:StartupVariables | Remove-Variable -ErrorAction SilentlyContinue; Script-Disconnect-Server; Invoke-Expression -Command $global:MenuNameCategory
}
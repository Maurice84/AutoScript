Function Identity-ActiveDirectory-Modify-ProfilePaths-7-Start {
    # =========
    # Execution
    # =========
    Script-Module-SetHeaders -Name ($MyInvocation.MyCommand).Name
    # ----------------------------------------------------------------------------------------------------------------------------------------
    Write-Host -NoNewLine "Step 1/2 - Modifying attributes on the customer Organizational Unit:" $Subtree -ForegroundColor Magenta; Write-Host
    # ----------------------------------------------------------------------------------------------------------------------------------------
    $AttribDescription = $global:HomefolderPath
    $AttribDestinationIndicator = $global:HomefolderDrive.ToUpper() + ":"
    ForEach ($OUObject in $OUObjects) {
        $ObjectGUID = $OUObject.ObjectGUID
        $global:OU = Script-Connect-Server -Module "ActiveDirectory" -Command ([scriptblock]::Create("Get-ADOrganizationalUnit '$ObjectGUID' -Properties *"))
        $OUGUID = $OU.ObjectGUID
        $OUDN = $OU.DistinguishedName
        If ($global:ProfilePath -ne "n.a.") {
            $AttribDesktopProfile = $global:ProfilePath
            If (!$OU.desktopProfile -OR $OU.desktopProfile -ne $AttribDesktopProfile) {
                Script-Connect-Server -Module "ActiveDirectory" -Command ([scriptblock]::Create("Set-ADOrganizationalUnit 'OUGUID' -Clear desktopProfile -Server '$PDC'"))
                If ($?) {
                    Script-Connect-Server -Module "ActiveDirectory" -Command ([scriptblock]::Create("Set-ADOrganizationalUnit '$OUDN' -Add @{desktopProfile=`"$AttribDesktopProfile`"} -Server '$PDC'"))
                    If ($?) {
                        Write-Host -NoNewLine "OK: Profile folder "; Write-Host -NoNewLine $AttribDesktopProfile -ForegroundColor Yellow; Write-Host -NoNewLine " added to the desktopProfile attribute of "; Write-Host $OU.Name -ForegroundColor Yellow
                    }
                    Else {
                        Write-Host "ERROR: An error occurred adding the Profile folder $AttribDesktopProfile to the desktopProfile attribute of" $OU.Name "!" -ForegroundColor Red
                        $Pause.Invoke()
                        BREAK
                    }
                }
                Else {
                    Write-Host ("ERROR: An error occurred clearing the desktopProfile attribute of " + $OU.Name + ", please investigate!") -ForegroundColor Red
                    $Pause.Invoke()
                    BREAK
                }
            }
            Else {
                Write-Host ("INFO: Profile Folder $AttribDesktopProfile already set on the desktopProfile attribute " + $OU.Name) -ForegroundColor Gray
            }
        }
        If (!$OU.Description -OR $OU.Description -ne $AttribDescription) {
            Script-Connect-Server -Module "ActiveDirectory" -Command ([scriptblock]::Create("Set-ADOrganizationalUnit '$OUGUID' -Clear Description -Server '$PDC'"))
            If ($?) {
                Script-Connect-Server -Module "ActiveDirectory" -Command ([scriptblock]::Create("Set-ADOrganizationalUnit '$OUDN' -Description '$AttribDescription' -Server '$PDC'"))
                If ($?) {
                    Write-Host -NoNewLine "OK: Home folder "; Write-Host -NoNewLine $AttribDescription -ForegroundColor Yellow; Write-Host -NoNewLine " added to the description attribute of "; Write-Host $OU.Name -ForegroundColor Yellow
                }
                Else {
                    Write-Host "ERROR: An error occurred adding the Home folder $AttribDescription to the description attribute of" $OU.Name"!" -ForegroundColor Red
                    $Pause.Invoke()
                    BREAK
                }
            }
            Else {
                Write-Host ("ERROR: An error occurred clearing the description attribute of " + $OU.Name + ", please investigate!") -ForegroundColor Red
                $Pause.Invoke()
                BREAK
            }
        }
        Else {
            Write-Host ("INFO: Home folder $AttribDescription already set on the desktopProfile attribute of " + $OU.Name) -ForegroundColor Gray
        }
        If (!$OU.destinationIndicator -OR $OU.destinationIndicator -ne $AttribDestinationIndicator) {
            Script-Connect-Server -Module "ActiveDirectory" -Command ([scriptblock]::Create("Set-ADOrganizationalUnit '$OUGUID' -Clear destinationIndicator -Server '$PDC'"))
            If ($?) {
                Script-Connect-Server -Module "ActiveDirectory" -Command ([scriptblock]::Create("Set-ADOrganizationalUnit '$OUDN' -Add @{destinationIndicator=`"$AttribDestinationIndicator`"} -Server '$PDC'"))
                If ($?) {
                    Write-Host -NoNewLine "OK: Home drive "; Write-Host -NoNewLine $AttribDestinationIndicator -ForegroundColor Yellow; Write-Host -NoNewLine " has been added to the destinationIndicator attribute of "; Write-Host $OU.Name -ForegroundColor Yellow
                }
                Else {
                    Write-Host "ERROR: An error occurred adding the home drive $AttribDestinationIndicator to the destinationIndicator attribute of" $OU.Name"!" -ForegroundColor Red
                    $Pause.Invoke()
                    BREAK
                }
            }
            Else {
                Write-Host ("ERROR: An error occurred clearing the destinationIndicator attribute of " + $OU.Name + ", please investigate!") -ForegroundColor Red
                $Pause.Invoke()
                BREAK
            }
        }
        Else {
            Write-Host ("INFO: Home drive $AttribDestinationIndicator already added to the destinationIndicator attribute of " + $OU.Name) -ForegroundColor Gray
        }
    }
    Write-Host
    # --------------------------------------------------------------------------------------------------------------------
    Write-Host -NoNewLine "Step 2/2 - Modifying attributes on the domain account(s):" -ForegroundColor Magenta; Write-Host
    # --------------------------------------------------------------------------------------------------------------------
    ForEach ($OUObject in $OUObjects) {
        $Users = Script-Connect-Server -Module "ActiveDirectory" -Command ([scriptblock]::Create(`
                    "Get-ADUser -Filter * -SearchBase '$OUObject' -Properties extensionAttribute1 | Select-Object DistinguishedName, extensionAttribute1, GivenName, SamAccountName,`
            @{Name=`"RDPPath`";Expression={([ADSI](`"LDAP://'$PDC'/`$(`$`_.DistinguishedName)`")).PSBase.InvokeGet(`"TerminalServicesProfilePath`")}},`
            @{Name=`"HomeDrive`";Expression={([ADSI](`"LDAP://'$PDC'/`$(`$`_.DistinguishedName)`")).PSBase.InvokeGet(`"TerminalServicesHomeDrive`")}},`
            @{Name=`"HomePath`";Expression={([ADSI](`"LDAP://'$PDC'/`$(`$`_.DistinguishedName)`")).PSBase.InvokeGet(`"TerminalServicesHomeDirectory`")}}")
        )
        ForEach ($User in $Users) {
            $HomefolderUserPath = $HomefolderPath + "\" + $User.SamAccountName
            $ADSI = [ADSI]("LDAP://$PDC/" + $User.DistinguishedName)
            $UserSamAccountName = $User.SamAccountName
            If ($ProfilePath -ne "n.a.") {
                If (!$User.RDPPath -OR $User.RDPPath -ne $ProfilePath) {
                    $ADSI.InvokeSet('TerminalServicesProfilePath', $ProfilePath)
                    $ADSI.SetInfo()
                    If ($?) {
                        Write-Host -NoNewLine "OK: Remote Desktop Profile path for account" $User.GivenName "has ben set to: "; Write-Host -NoNewLine $ProfilePath -ForegroundColor Yellow; Write-Host
                    }
                    Else {
                        Write-Host "ERROR: An error occurred adding the Remote Desktop Profile path for" $User.GivenName ", please investigate!" -ForegroundColor Red
                        $Pause.Invoke()
                        BREAK
                    }
                }
                Else {
                    Write-Host "INFO: Remote Desktop Profile path for" $User.GivenName "already been set" -ForegroundColor Gray
                }
            }
            If ((!$User.HomePath -OR $User.HomePath -ne $HomefolderUserPath) -OR (!$User.HomeDrive -OR $User.HomeDrive -ne $AttribDestinationIndicator)) {
                $ADSI.InvokeSet('TerminalServicesHomeDirectory', $HomefolderUserPath)
                $ADSI.SetInfo()
                If ($?) {
                    If (!$User.HomeDrive -OR $User.HomeDrive -ne $AttribDestinationIndicator) {
                        $ADSI.InvokeSet('TerminalServicesHomeDrive', $AttribDestinationIndicator)
                        $ADSI.SetInfo()
                        If (!$?) {
                            Write-Host "ERROR: An error occurred adding the homefolder drive letter for" $User.GivenName ", please investigate!" -ForegroundColor Red
                            $Pause.Invoke()
                            BREAK
                        }
                    }
                    Write-Host -NoNewLine "OK: Homefolder path for" $User.GivenName "has been set to: "; Write-Host -NoNewLine $AttribDestinationIndicator $HomefolderUserPath -ForegroundColor Yellow; Write-Host
                }
                Else {
                    Write-Host "ERROR: An error occurred adding the Homefolder path for" $User.GivenName ", please investigate!" -ForegroundColor Red
                    $Pause.Invoke()
                    BREAK
                }
            }
            Else {
                Write-Host "INFO: Homefolder path for" $User.GivenName "already been set" -ForegroundColor Gray
            }
            If ($Prefix) {
                If (!$User.extensionAttribute1 -OR $User.extensionAttribute1 -ne $Prefix) {
                    Script-Connect-Server -Module "ActiveDirectory" -Command ([scriptblock]::Create("Set-ADUser '$UserSamAccountName' -Server '$PDC' -Add @{extensionAttribute1='$Prefix'}"))
                    If ($?) {
                        Write-Host -NoNewLine "OK: Prefix for" $User.GivenName "has been set to: "; Write-Host -NoNewLine $Prefix -ForegroundColor Yellow; Write-Host
                    }
                    Else {
                        Write-Host "ERROR: An error occurred adding the Prefix for" $User.GivenName ", please investigate!" -ForegroundColor Red
                        $Pause.Invoke()
                        BREAK
                    }
                }
                Else {
                    Write-Host "INFO: Prefix for" $User.GivenName "already set" -ForegroundColor Gray
                }
            }
            If (!$User.Country) {
                Script-Connect-Server -Module "ActiveDirectory" -Command ([scriptblock]::Create("Set-ADUser '$UserSamAccountName' -Server '$PDC' -Country '$Language'"))
                If ($?) {
                    Write-Host -NoNewLine "OK: Language for" $User.GivenName "has been set to: "; Write-Host -NoNewLine $Taalkeuze -ForegroundColor Yellow; Write-Host
                }
                Else {
                    Write-Host "ERROR: An error occurred adding the language for" $User.GivenName ", please investigate!" -ForegroundColor Red
                    $Pause.Invoke()
                    BREAK
                }
            }
            Else {
                Write-Host "INFO: Language" $User.GivenName "already set" -ForegroundColor Gray
            }
        }  
    }
    Write-Host
    # ==========
    # Finalizing
    # ==========
    $Pause.Invoke()
    Get-Variable -Exclude $global:StartupVariables | Remove-Variable -ErrorAction SilentlyContinue; Script-Disconnect-Server; Invoke-Expression -Command $global:MenuNameCategory
}
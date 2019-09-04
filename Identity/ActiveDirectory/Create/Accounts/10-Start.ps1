Function Identity-ActiveDirectory-Create-Accounts-10-Start {
    # =========
    # Execution
    # =========
    Script-Module-SetHeaders -Name ($MyInvocation.MyCommand).Name
    If (!$global:AccountHerstelCSV) {
        # -------------------------------------------------------------------------------------------------------------------------
        # Converting UPN to SamAccountName (20 characters max). When applicable: GivenName + Surname combination and optional digit
        # -------------------------------------------------------------------------------------------------------------------------
        If ($global:Username.Length -gt 20) {
            $SamAccountNameConvert = $global:Username.Substring(0, 20)
        }
        Else {
            $SamAccountNameConvert = $global:Username
        }
        $SamAccountNameCheck = Script-Connect-Server -Module "ActiveDirectory" -Command ([scriptblock]::Create(`
                    "Get-ADUser -Filter * | Select-Object SamAccountName | Where-Object {`$`_.SamAccountName -eq '$SamAccountNameConvert'}")
        )
        If ($SamAccountNameCheck) {
            $SamAccountNameConvert = $global:GivenName + $global:Surname
            $SamAccountNameConvert = $SamAccountNameConvert.Replace(' ', '')
            If ($SamAccountNameConvert.Length -gt 20) {
                $SamAccountNameConvert = $SamAccountNameConvert.Substring(0, 20)
            }
            $SamAccountNameCheck = Script-Connect-Server -Module "ActiveDirectory" -Command ([scriptblock]::Create(`
                        "Get-ADUser -Filter * | Select-Object SamAccountName | Where-Object {`$`_.SamAccountName -eq '$SamAccountNameConvert'}")
            )
            If ($SamAccountNameCheck) {
                If ($SamAccountNameConvert.Length -eq 20) {
                    $SamAccountNameConvertLength = $SamAccountNameConvert.Length - 1
                }
                Else {
                    $SamAccountNameConvertLength = $SamAccountNameConvert.Length
                }
                $LastChar = $SamAccountNameConvert.Substring($SamAccountNameConvertLength)
                If ($LastChar -eq 1) {
                    $LastChar = 2
                }
                ElseIf ($LastChar -eq 2) {
                    $LastChar = 3
                }
                Else {
                    $LastChar = 1
                }
                $SamAccountNameConvert = $SamAccountNameConvert.Substring(0, $SamAccountNameConvertLength) + $LastChar
            }
        }
    }
    # --------------
    # Path detection
    # --------------
    If ($global:Profiel -eq "Profielkopie") {
        If ($global:Subtree -notlike "*Webmail" -AND $global:Subtree -notlike "*Service") {
            $ADSI = [ADSI]('LDAP://{0}' -f $global:DistinguishedName)
            $RDPPath = $ADSI.InvokeGet('TerminalServicesProfilePath')
            $HomeDrive = $ADSI.InvokeGet('TerminalServicesHomeDrive')
            $HomePath = $ADSI.InvokeGet('TerminalServicesHomeDirectory')
            If ($HomePath -like "*$global:SamAccountName") {
                $HomePath = $HomePath -Replace ($global:SamAccountName, $SamAccountNameConvert)
            }
            Else {
                If ($HomePath -notlike "*\") {
                    $HomePath += "\"
                }
                $HomePath += $SamAccountNameConvert
            }
        }
    }
    If ($SamAccountNameConvert) {
        $global:SamAccountName = $SamAccountNameConvert
    }
    If ($AccountHerstelCSV -OR ($global:Profiel -eq "Domein gebruiker" -OR $global:Profiel -eq "Test gebruiker")) {
        $OUProperties = Script-Connect-Server -Module "ActiveDirectory" -Command ([scriptblock]::Create("Get-ADOrganizationalUnit '$global:OUPath' -Properties *"))
        If ($OUProperties.desktopProfile) {
            [string]$RDPPath = $OUProperties.desktopProfile
        }
        If ($OUProperties.description) {
            [string]$HomePath = $OUProperties.description + "\" + $SamAccountName
        }
        If ($OUProperties.destinationIndicator) {
            [string]$HomeDrive = $OUProperties.destinationIndicator
        }
        If (!$RDPPath -AND !$HomeDrive -AND !$HomePath) {
            $OUAccounts = Script-Connect-Server -Module "ActiveDirectory" -Command ([scriptblock]::Create(`
                        "Get-ADUser -Filter * | Select-Object SamAccountName,DistinguishedName | Where-Object {`$`_.DistinguishedName -like `"*$global:OUPath`"}")
            )
            ForEach ($Account in $OUAccounts) {
                $ADSI = [ADSI]('LDAP://{0}' -f $Account.DistinguishedName)
                $RDPPath = Try { $ADSI.InvokeGet('TerminalServicesProfilePath') } Catch { } # \\...\Profile
                $HomeDrive = Try { $ADSI.InvokeGet('TerminalServicesHomeDrive') } Catch { } # H:
                $HomePath = Try { $ADSI.InvokeGet('TerminalServicesHomeDirectory').Replace($Account.SamAccountName, $global:SamAccountName) } Catch { } # \\...\Users
                If ($RDPPath -AND $HomeDrive -AND $HomePath) { BREAK }
            }
        }
    }
    If (!$global:AccountHerstelCSV) {
        If (!$global:CSV) {
            Script-Module-SetHeaders -Name ($MyInvocation.MyCommand).Name
        }
        $Country = $Language
    }
    # -----------------
    # Declare variables
    # -----------------
    If ($global:Surname.Length -ne 0) {
        $Name = $global:GivenName + " " + $global:Surname
    }
    Else {
        $Name = $global:GivenName
    }
    $GroupAdmins = Script-Connect-Server -Module "ActiveDirectory" -Command ([scriptblock]::Create("(Get-ADGroup -Filter * | Where-Object {`$`_.Name -like `"Dom*dmin*`"}).Name"))
    # ---------------------------------------------------------------------------------------------------------
    Write-Host -NoNewLine "Step 1/2 - Creating account in" $global:Subtree -ForegroundColor Magenta; Write-Host
    # ---------------------------------------------------------------------------------------------------------
    $CheckUPN = Script-Connect-Server -Module "ActiveDirectory" -Command ([scriptblock]::Create("Get-ADUser -Filter {UserPrincipalName -eq '$UPN'}"))
    If (!$CheckUPN) {
        Script-Module-SetPassword
        Script-Connect-Server -Module "ActiveDirectory" -Command ([scriptblock]::Create(
                "New-ADUser -Name '$Name'`
            -Path '$global:OUPath'`
            -GivenName '$global:GivenName'`
            -DisplayName '$Name'`
            -UserPrincipalName '$global:UPN'`
            -SamAccountName '$global:SamAccountName'`
            -AccountPassword (ConvertTo-SecureString '$global:Password' -AsPlainText -Force)`
            -Server '$global:PDC'`
            -Enabled `$True")
        )
        If ($?) {
            Script-Module-ReplicateAD
            Script-Connect-Server -Module "ActiveDirectory" -Command ([scriptblock]::Create(`
                        "`$Surname = '$global:Surname'; If (`$SurName) {`
                    Set-ADUser '$global:SamAccountName' -Server '$global:PDC' -Surname '$global:Surname'`
                };
                `$Initials = '$global:Initials'; If (`$Initials) {`
                    Set-ADUser '$global:SamAccountName' -Server '$global:PDC' -Initials '$global:Initials'`
                };
                `$Description = '$global:Description'; If (`$Description) {`
                    Set-ADUser '$global:SamAccountName' -Server '$global:PDC' -Description '$global:Description'`
                };
                `$Office = '$global:Office'; If (`$Office) {`
                    Set-ADUser '$global:SamAccountName' -Server '$global:PDC' -Office '$global:Office'`
                };
                `$Title = '$global:Title'; If (`$Title) {`
                    Set-ADUser '$global:SamAccountName' -Server '$global:PDC' -Title '$global:Title'`
                };
                `$Department = '$global:Department'; If (`$Department) {`
                    Set-ADUser '$global:SamAccountName' -Server '$global:PDC' -Department '$global:Department'`
                };
                `$Company = '$global:Company'; If (`$Company) {`
                    Set-ADUser '$global:SamAccountName' -Server '$global:PDC' -Company '$global:Company'`
                };
                `$HomePage = '$global:HomePage'; If (`$HomePage) {`
                    Set-ADUser '$global:SamAccountName' -Server '$global:PDC' -HomePage '$global:HomePage'`
                };
                `$StreetAddress = '$global:StreetAddress'; If (`$StreetAddress) {`
                    Set-ADUser '$global:SamAccountName' -Server '$global:PDC' -StreetAddress '$global:StreetAddress'`
                };
                `$POBox = '$global:POBox'; If (`$POBox) {`
                    Set-ADUser '$global:SamAccountName' -Server '$global:PDC' -POBox '$global:POBox'`
                };
                `$City = '$global:City'; If (`$City) {`
                    Set-ADUser '$global:SamAccountName' -Server '$global:PDC' -City '$global:City'`
                };
                `$State = '$global:State'; If (`$State) {`
                    Set-ADUser '$global:SamAccountName' -Server '$global:PDC' -State '$global:State'`
                };
                `$PostalCode = '$global:PostalCode'; If (`$PostalCode) {`
                    Set-ADUser '$global:SamAccountName' -Server '$global:PDC' -PostalCode '$global:PostalCode'`
                };
                `$Country = '$global:Country'; If (`$Country) {`
                    Set-ADUser '$global:SamAccountName' -Server '$global:PDC' -Country '$global:Country'`
                };
                `$OfficePhone = '$global:OfficePhone'; If (`$OfficePhone) {`
                    Set-ADUser '$global:SamAccountName' -Server '$global:PDC' -OfficePhone '$global:OfficePhone'`
                };
                `$MobilePhone = '$global:MobilePhone'; If (`$MobilePhone) {`
                    Set-ADUser '$global:SamAccountName' -Server '$global:PDC' -MobilePhone '$global:MobilePhone'`
                };
                `$HomePhone = '$global:HomePhone'; If (`$HomePhone) {`
                    Set-ADUser '$global:SamAccountName' -Server '$global:PDC' -HomePhone '$global:HomePhone'`
                };
                `$Fax = '$global:Fax'; If (`$Fax) {`
                    Set-ADUser '$global:SamAccountName' -Server '$global:PDC' -Fax '$global:Fax'`
                };")
            )
            If ($?) {
                Script-Module-ReplicateAD
                Write-Host -NoNewLine " - OK: Account "; Write-Host -NoNewLine $Name -ForegroundColor Yellow; Write-Host -NoNewLine " has been created with password "; Write-Host $global:Password -ForegroundColor Cyan
            }
            Else {
                Write-Host " - ERROR: An error occurred modifying the attributes on account $Name, please investigate!" -ForegroundColor Red
                $Pause.Invoke()
                BREAK
            }
        }
        Else {
            Write-Host " - ERROR: An error occurred during creation of account $Name, please investigate!" -ForegroundColor Red
            $Pause.Invoke()
            BREAK
        }
    }
    Else {
        Write-Host " - INFO: Account $Name already exists (in $global:Subtree)" -ForegroundColor Gray
    }
    # -----------------------
    # Adding group membership
    # -----------------------
    If ($global:Groups) {
        $GroupNames = @()
        $GroupNames = ($Groups.Name | Sort-Object) -join ";"
        Script-Connect-Server -Module "ActiveDirectory" -Command ([scriptblock]::Create(`
                    "`$AccountHerstelCSV = '$global:AccountHerstelCSV';
            `$SamAccountName = '$global:SamAccountName';
            `$GroupNames = @(); ForEach(`$Temp in '$GroupNames'.Split(';')) {`$GroupNames += `$Temp;};
            `$CurrentGroups = Try {(Get-ADPrincipalGroupMembership '$global:SamAccountName' -ErrorAction SilentlyContinue | Select-Object Name | Sort-Object Name).Name} Catch {`$False};
            ForEach (`$Group in `$GroupNames) {`
                If (`$AccountHerstelCSV) {`
                    `$CheckGroup = Try {Get-ADGroup `$Group -ErrorAction SilentlyContinue | Out-Null} Catch {`$False};
                } Else {
                    `$CheckGroup = `$True;
                };
                If (`$CheckGroup -ne `$False) {`
                    If (`$CurrentGroups -eq `$false -OR `$CurrentGroups -notcontains `$Group) {`
                        Add-ADGroupMember -Identity `"`$Group`" -Members '$global:SamAccountName';
                        If (`$?) {`
	                        Write-Host -NoNewLine `" - OK: Group `"; Write-Host -NoNewLine `$Group -ForegroundColor Yellow; Write-Host `" has been added to the account`";
                        } Else {`
                            Write-Host -NoNewLine `" - ERROR: An error occurred adding group `" -ForegroundColor Red; Write-Host -NoNewLine `$Group -ForegroundColor Red; Write-Host -NoNewLine `" to the account, please investigate!`" -ForegroundColor Red; Write-Host;
                        };
                    } Else {`
                        Write-Host -NoNewLine `" - INFO: Account is already a member of `" -ForegroundColor Gray; Write-Host -NoNewLine `$Group -ForegroundColor Gray; Write-Host;
                    };
                } Else {`
                    Write-Host -NoNewLine `" - ERROR: Group `" -ForegroundColor Red; Write-Host -NoNewLine `$Group -ForegroundColor Yellow; Write-Host -NoNewLine `" does not exist anymore`" -ForegroundColor Red; Write-Host;
               };
            }")
        )
    }
    Else {
        If ($global:Profiel -eq "Profielkopie") {
            Write-Host " - INFO: No groups are detected on the domain account to duplicate" -ForegroundColor Gray
        }
        If ($global:Profiel -eq "Domein gebruiker" -OR $global:Profiel -eq "Beheerder account") {
            Write-Host " - INFO: No groups have been selected" -ForegroundColor Gray
        }
    }
    # --------------------
    # Modify Profile paths
    # --------------------
    $User = Script-Connect-Server -Module "ActiveDirectory" -Command ([scriptblock]::Create(`
                "Get-ADUser '$global:SamAccountName' | Select-Object DistinguishedName,`
        @{Name=`"RDPPath`";Expression={([ADSI](`"LDAP://'$global:PDC'/`$(`$`_.DistinguishedName)`")).PSBase.InvokeGet(`"TerminalServicesProfilePath`")}},`
        @{Name=`"HomeDrive`";Expression={([ADSI](`"LDAP://'$global:PDC'/`$(`$`_.DistinguishedName)`")).PSBase.InvokeGet(`"TerminalServicesHomeDrive`")}},`
        @{Name=`"HomePath`";Expression={([ADSI](`"LDAP://'$global:PDC'/`$(`$`_.DistinguishedName)`")).PSBase.InvokeGet(`"TerminalServicesHomeDirectory`")}}")
    )
    $ADSI = [ADSI]("LDAP://$global:PDC/" + $User.DistinguishedName)
    If ($RDPPath) {
        If ($global:AccountHerstelCSV) {
            $CheckRDPPath = Get-Item ($RDPPath + ".V2") -ErrorAction SilentlyContinue
            If (!$CheckRDPPath) {
                $CheckRDPPath = $False
            }
        }
        Else {
            $CheckRDPPath = $True
        }
        If ($CheckRDPPath -ne $False) {
            If ($User.RDPPath -ne $RDPPath) {
                $ADSI.InvokeSet('TerminalServicesProfilePath', $RDPPath)
                $ADSI.SetInfo()
                If ($?) {
                    Write-Host -NoNewLine " - OK: Remote Desktop Profile Path set to: "; Write-Host -NoNewLine $RDPPath -ForegroundColor Yellow; Write-Host
                }
                Else {
                    Write-Host " - ERROR: An error occurred adding the Remote Desktop Profile Path, please investigate!" -ForegroundColor Red
                    $Pause.Invoke()
                    BREAK
                }
            }
            Else {
                Write-Host " - INFO: Remote Desktop Profile Path already set" -ForegroundColor Gray
            }
        }
        Else {
            Write-Host -NoNewLine " - ERROR: Remote Desktop Profile Path " -ForegroundColor Red; Write-Host -NoNewLine $RDPPath -ForegroundColor Yellow; Write-Host -NoNewLine " does not exist anymore" -ForegroundColor Red; Write-Host
        }
    }
    Else {
        Write-Host " - INFO: Remote Desktop Profile Path not needed" -ForegroundColor Gray
    }
    If ($HomePath) {
        If ($global:AccountHerstelCSV) {
            $Length = $HomePath.Split('\').Count - 2
            $HomePathRoot = "\\" + ($HomePath.Split('\')[2..$Length] -join '\')
            $CheckHomePath = Get-Item $HomePathRoot -ErrorAction SilentlyContinue
            If (!$CheckHomePath) {
                $CheckHomePath = $False
            }
        }
        Else {
            $CheckHomePath = $True
        }
        If ($CheckHomePath -ne $False) {
            # --------------------------------------------------------------------------------------------------------------------------------------------
            # Checking if the detected Homefolder is correct to circumvent permission issues. Homefolder should not exist, or only contain profile folders
            # --------------------------------------------------------------------------------------------------------------------------------------------
            If (Get-ChildItem $HomePath -ErrorAction SilentlyContinue) {
                If (!(Get-ChildItem $HomePath | Where-Object { $_.Name -eq "Bureaublad" -OR $_.Name -eq "Desktop" })) {
                    Write-Host " - ERROR: Homefolder already exist, but it contains no profile data. Please investigate!" -ForegroundColor Red
                    Write-Host
                    $Pause.Invoke()
                    Get-Variable -Exclude $global:StartupVariables | Remove-Variable -ErrorAction SilentlyContinue; Invoke-Expression -Command $global:MenuNameCategory
                }
            }
            If ($User.HomePath -ne $HomePath) {
                $ADSI.InvokeSet('TerminalServicesHomeDirectory', $HomePath)
                $ADSI.SetInfo()
                If ($?) {
                    Write-Host -NoNewLine " - OK: Homefolder Path set to: "; Write-Host -NoNewLine $HomeDrive $HomePath -ForegroundColor Yellow; Write-Host
                }
                Else {
                    Write-Host " - ERROR: An error occurred adding the Homefolder Path, please investigate!" -ForegroundColor Red
                    $Pause.Invoke()
                    BREAK
                }
            }
            Else {
                Write-Host " - INFO: Homefolder Path already set" -ForegroundColor Gray
            }
            If ($HomeDrive) {
                If ($User.HomeDrive -ne $HomeDrive) {
                    $ADSI.InvokeSet('TerminalServicesHomeDrive', $HomeDrive)
                    $ADSI.SetInfo()
                    If (!$?) {
                        Write-Host " - ERROR: An error occurred adding the Homefolder drive letter, please investigate!" -ForegroundColor Red
                        $Pause.Invoke()
                        BREAK
                    }
                }
            }
        }
        Else {
            Write-Host -NoNewLine " - ERROR: Homefolder Path " -ForegroundColor Red; Write-Host -NoNewLine $HomePath -ForegroundColor Yellow; Write-Host -NoNewLine " does not exist" -ForegroundColor Red; Write-Host
        }
    }
    Else {
        Write-Host " - INFO: Homefolder Path not needed" -ForegroundColor Gray
    }
    Write-Host
    # ---------------------------------------------------------------------------------------------------
    Write-Host -NoNewLine "Step 2/2 - Creating regional homefolders" -ForegroundColor Magenta; Write-Host
    # ---------------------------------------------------------------------------------------------------
    If ($CheckHomePath -ne $False) {
        If ($HomePath) {
            $Folders = @()
            If ($global:Country -eq "DE") {
                $Folders = "Bilder", "AutoRecover", "AutoRecover\Excel", "AutoRecover\PowerPoint", "AutoRecover\Word", "Desktop", "Downloads", "Favoriten", "Links", "Musik", "Outlook", "Personal Settings", "Videos", "Windows"
            }
            If ($global:Country -eq "NL") {
                $Folders = "Afbeeldingen", "AutoRecover", "AutoRecover\Excel", "AutoRecover\PowerPoint", "AutoRecover\Word", "Bureaublad", "Downloads", "Favorieten", "Koppelingen", "Muziek", "Outlook", "Personal Settings", "Video's", "Windows"
            }
            If ($Cglobal:ountry -eq "US") {
                $Folders = "Pictures", "AutoRecover", "AutoRecover\Excel", "AutoRecover\PowerPoint", "AutoRecover\Word", "Desktop", "Downloads", "Favorites", "Links", "Music", "Outlook", "Personal Settings", "Videos", "Windows"
            }
            ForEach ($Folder in $Folders) {
                $Homefolder = ($HomePath + "\" + $Folder)
                If (!(Test-Path $Homefolder)) {
                    New-Item $Homefolder -Type Directory -ErrorAction SilentlyContinue | Out-Null
                    If ($?) {
                        Write-Host -NoNewLine " - OK: Homefolders are created at: "; Write-Host $Homefolder -ForegroundColor Yellow
                    }
                    Else {
                        Write-Host " - ERROR: An error occurred creating the homefolders, please investigate!" -ForegroundColor Red
                        $Pause.Invoke()
                        BREAK
                    }
                }
            }
            If ((Get-ACL $HomePath).Owner -notlike "*$GroupAdmins") {
                ICACLS "$HomePath" /setowner "$GroupAdmins" /T | Out-Null
                If ($?) {
                    Write-Host -NoNewLine " - OK: Owner permission set to: "; Write-Host -NoNewLine $GroupAdmins -ForegroundColor Yellow; Write-Host
                }
                Else {
                    Write-Host " - ERROR: An error occurred setting the owner permission of $GroupAdmins, please investigate!" -ForegroundColor Red
                    $Pause.Invoke()
                    BREAK
                }
            }
            Else {
                Write-Host " - INFO: $GroupAdmins is already Homefolder owner" -ForegroundColor Gray
            }
            $ACL = (Get-Item $HomePath).GetAccessControl('Access')
            $RechtenObjects = @()
            $RechtenObjects += $GroupAdmins, $global:SamAccountName
            ForEach ($RechtenObject in $RechtenObjects) {
                $GetACL = $ACL.Access | Where-Object { $_.IdentityReference -like "*$RechtenObject*" }
                If ($GetACL.FileSystemRights -ne "FullControl") {
                    $FullControl = New-Object System.Security.AccessControl.FileSystemAccessRule($RechtenObject, "FullControl", "ContainerInherit,ObjectInherit", "None", "Allow")
                    $ACL.SetAccessRule($FullControl)
                    Set-ACL -Path $HomePath -ACLObject $ACL
                    If ($?) {
                        Write-Host -NoNewLine " - OK: Permissions set to "; Write-Host $RechtenObject -ForegroundColor Yellow
                    }
                    Else {
                        Write-Host " - ERROR: An error occurred setting permissions to $RechtenObject, please investigate!" -ForegroundColor Red
                        $Pause.Invoke()
                        BREAK
                    }
                }
                Else {
                    Write-Host " - INFO: $RechtenObject already has permissions to Homefolder" -ForegroundColor Gray
                }
            }
        }
        Else {
            Write-Host " - INFO: Homefolders not needed" -ForegroundColor Gray
        }
    }
    Else {
        Write-Host " - ERROR: Skipping creation of Homefolders due to missing Users share" -ForegroundColor Red
    }
    If ($global:AccountHerstelCSV -OR !$global:CSV) {
        $Pause.Invoke()
        Get-Variable -Exclude $global:StartupVariables | Remove-Variable -ErrorAction SilentlyContinue; Script-Disconnect-Server; Invoke-Expression -Command $global:MenuNameCategory
    }
    Else {
        Write-Host
        Write-Host -NoNewLine "Proceeding to the next account in 2 seconds..." -ForegroundColor Green
        Start-Sleep -Seconds 2
    }
}
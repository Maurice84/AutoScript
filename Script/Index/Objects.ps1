Function Script-Index-Objects {
    param (
        [string]$CurrentTask,
        [string]$Filter = "*",
        [string]$Functie,
        [string]$FunctionName,
        [array]$Objecten,
        [string]$Server = $env:computername,
        [boolean]$SkipMenu = $false,
        [string]$SortObject = "Name",
        [string]$Type
    )
    # -----------------
    # Declare variables
    # -----------------
    $Objects = $null
    $Objects = @()
    $global:Aantal = $null
    $global:Array = $null
    $global:Array = @()
    $global:Groups = $null
    $InputKey = $null
    $InputSelection = $null
    $Available = $null
    If ($Objecten) {
        $Objects = $Objecten
    }
    # -----------------------------
    # Declareren variables per type
    # -----------------------------
    If ($Type -like "Account*") {
        If ($global:AzureADConnected) {
            $ObjectType = "Office 365 account(s)"
        }
        Else {
            $ObjectType = "domain account(s)"
        }
        If ($Type -eq "AccountHerstel") {
            $ObjectType = "recoverable $ObjectType"
        }
        If ($Type -eq "AccountRecycleBin") {
            $ObjectType = "$ObjectType in AD Recycle Bin"
        }
        If ($Type -eq "AccountZonderMailbox") {
            $ObjectType = "$ObjectType without mailbox"
        }
    }
    If ($Type -eq "Adresboekbeleid") {
        $ObjectType = "addressbook policy"
    }
    If ($Type -eq "Bestand") {
        $ObjectType = "$Filter-files"
    }
    If ($Type -eq "Credentials") {
        $ObjectType = "credential(s)"
    }
    If ($Type -eq "Drive") {
        $ObjectType = "(network)drives"
    }
    If ($Type -eq "Emaildomein") {
        $Emaildomein = $null
        $ObjectType = "maildomain(s)"
    }
    If ($Type -eq "GenereerEmailadres") {
        $ObjectType = "mailaddress(ses)"
        $DisplayNameDOT = $null
        $DisplayNameDASH = $null
        $DisplayNameSPACE = $null
        $global:EmailPrefix = $null
        $global:EmailSuffix = $null
        $global:Email = $null
    }
    If ($Type -eq "GenereerGebruikersnaam") {
        $ObjectType = "usernames"
        $global:Username = $null
        $Lengte = $null
    }
    If ($Type -eq "Groepen") {
        $ObjectType = "security group(s)"
    }
    If ($Type -eq "Homefolders") {
        $ObjectType = "homefolder(s)"
    }
    If ($Type -like "Mailbox*") {
        If ($global:Office365Exchange) {
            $ObjectType = "Office 365 mailboxes"
            If ($Type -eq "MailboxHerstel") {
                $ObjectType = "recoverable Office 365 mailboxes"
            }
        }
        Else {
            $ObjectType = "mailbox(es)"
            If ($Type -eq "MailboxHerstel") {
                $ObjectType = "recoverable mailboxes"
            }
        }
    }
    If ($Type -like "Office365Account*") {
        $ObjectType = "Office 365 account(s)"
    }
    If ($Type -like "Office365Licenties") {
        $ObjectType = "Office 365 license(s)"
    }
    If ($Type -eq "UPNSuffix") {
        $ObjectType = "domain suffix(es)"
    }
    $ObjectTypes = $ObjectType.Replace('(', '').Replace(')', '')
    # ------------------------------------------------
    # Request to enter a search filter to index faster
    # ------------------------------------------------
    If ($Type -like "Account*" -OR $Type -eq "Groepen" -OR $Type -like "Mailbox*") {
        Write-Host -NoNewLine ("  > Please enter the name (or a part) of the OU with " + $ObjectTypes + ", or press Enter for all " + $ObjectTypes + ": ")
        $Filter = Read-Host
        If ($FunctionName -like "*Exchange*Move*") {
            $ObjectTypes += " which are not moved yet"
        }
        If ($Filter.Length -eq 0) {
            $Filter = "*"
            $global:OUFormat = "All $ObjectTypes"
        }
        Else {
            $global:OUFormat = $Filter
        }
        $global:OUFilter = $Filter
        If ($Type -like "Mailbox*") {
            If (!$FilterQuestions) {
                Do {
                    Write-Host -NoNewline "  > Would you like to have "; Write-Host -NoNewLine "mailbox size, item count, last logon date and server" -ForegroundColor Yellow; Write-Host -NoNewLine " indexed? (Y/N): "
                    $InputChoice = Read-Host
                    $InputKey = @("Y", "N") -contains $InputChoice
                    If (!$InputKey) {
                        Write-Host "    Please use the letters above as input" -ForegroundColor Red; Write-Host
                    }
                } Until ($InputKey)
                Switch ($InputChoice) {
                    "Y" {
                        $global:DetectDateItemsServerAndSize = $true
                    }
                }
                If ($FunctionName -like "*Overview*") {
                    Do {
                        Write-Host -NoNewline "  > Would you also like to have the "; Write-Host -NoNewline "mailbox FullAccess/Send-As permissions" -ForegroundColor Yellow; Write-Host -NoNewline " indexed? (Y/N): "
                        $InputChoice = Read-Host
                        $InputKey = @("Y", "N") -contains $InputChoice
                        If (!$InputKey) {
                            Write-Host "    Please use the letters above as input" -ForegroundColor Red; Write-Host
                        }
                    } Until ($InputKey)
                    Switch ($InputChoice) {
                        "Y" {
                            $global:DetectPermissions = $true
                            $ObjectsFullAccess = @()
                            $ObjectsSendAs = @()
                        }
                    }
                    Do {
                        Write-Host -NoNewline "  > And finally, would you also like to have the "; Write-Host -NoNewline "mailbox language" -ForegroundColor Yellow; Write-Host -NoNewline " indexed? (Y/N): "
                        $InputChoice = Read-Host
                        $InputKey = @("Y", "N") -contains $InputChoice
                        If (!$InputKey) {
                            Write-Host "    Please use the letters above as input" -ForegroundColor Red; Write-Host
                        }
                    } Until ($InputKey)
                    Switch ($InputChoice) {
                        "Y" {
                            $global:DetectLanguage = $true
                            $ObjectsLanguageStatistics = @()
                        }
                    }
                }
                Write-Host
                $FilterQuestions = $true
            }
        }
        #Script-Module-SetHeaders -CurrentTask $CurrentTask -Name $FunctionName
    }
    # ---------------------------------------
    # Displaying text of indexing the objects
    # ---------------------------------------
    If ($Server -AND $Server -ne $env:computername) {
        $ObjectTypeServer = " on " + $Server.Split(".")[0]
    }
    If ($SortObject -eq "Name" -AND !$SortObjectFormat) {
        If ($Filter -eq "*" -OR $Filter -eq "Folder" -OR $Filter -eq "PST") {
            If ($Type -notlike "Account*") {
                Write-Host -NoNewLine ("  - Loading $ObjectType" + $ObjectTypeServer + "...")
            }
        }
        Else {
            Write-Host -NoNewLine ("  - Loading $ObjectType" + $ObjectTypeServer + " using filter: "); Write-Host -NoNewLine $Filter -ForegroundColor Yellow; Write-Host -NoNewLine "..."
        }
    }
    Else {
        Write-Host
        Write-Host -NoNewLine "Sorting on column $SortObjectFormat of $ObjectType..." -ForegroundColor Magenta
    }
    # -------------------
    # Loading the objects
    # -------------------
    If ($Type -like "Account*") {
        If (!$Objects) {
            If ($global:OS -ne "SBS2008") {
                If ($Filter -eq "*") {
                    Do {
                        Write-Host -NoNewLine "  > Please note:" -ForegroundColor Yellow; Write-Host -NoNewLine " You have chosen to index all OUs, would you like to include all objects like build-in (admin) accounts? (Y/N): "
                        $InputChoice = Read-Host
                        $InputKey = @("Y", "N") -contains $InputChoice
                        If (!$InputKey) {
                            Write-Host "    Please use the letters above as input" -ForegroundColor Red
                        }
                    } Until ($InputKey)
                    Write-Host -NoNewLine "  - Loading $ObjectType..."
                    Switch ($InputChoice) {
                        "Y" {
                            If ($Type -eq "AccountHerstel" -OR $Type -eq "AccountRecycleBin") {
                                $Objects = Script-Connect-Server -Module "ActiveDirectory" -Command ([scriptblock]::Create(`
                                            "Get-ADObject -Filter * -Properties * -IncludeDeletedObjects | Sort-Object '$SortObject' | Where-Object {`$`_.GivenName -ne `$null -AND `$`_.Deleted -eq `$true}")
                                )
                            }
                            ElseIf ($Type -eq "AccountZonderMailbox") {
                                $Objects = Script-Connect-Server -Module "ActiveDirectory" -Command ([scriptblock]::Create(`
                                            "Get-ADUser -Filter * -Properties * | Sort-Object '$SortObject' | Where-Object {`$`_.GivenName -ne `$null -AND `$`_.msExchRecipientTypeDetails -eq `$null -AND `$`_.EmailAddress -eq `$null}")
                                )
                            }
                            Else {
                                If ($global:AzureADConnected) {
                                    $Objects = Script-Connect-Server -Module "ActiveDirectory" -Command ([scriptblock]::Create("Get-AzureADUser -All `$true | Sort-Object '$SortObject' | Where-Object {`$`_.GivenName -ne `$null}"))
                                    $AzureADSubscribedSku = Script-Connect-Server -Module "ActiveDirectory" -Command ([scriptblock]::Create("Get-AzureADSubscribedSku"))
                                }
                                Else {
                                    $Objects = Script-Connect-Server -Module "ActiveDirectory" -Command ([scriptblock]::Create("Get-ADUser -Filter * -Properties * | Sort-Object '$SortObject' | Where-Object {`$`_.GivenName -ne `$null}"))
                                }
                            }
                        }
                        "N" {
                            $AllObjects = $false
                        }
                    }
                }
                Else {
                    $AllObjects = $false
                }
                If ($AllObjects -eq $false) {
                    If ($Type -eq "AccountHerstel" -OR $Type -eq "AccountRecycleBin") {
                        $Objects = Script-Connect-Server -Module "ActiveDirectory" -Command ([scriptblock]::Create(`
                                    "Get-ADObject -Filter * -Properties * -IncludeDeletedObjects | Sort-Object `$SortObject | Where-Object {`
                            `$`_.LastKnownParent -like `"*$Filter*`" -AND `$`_.GivenName -ne `$null -AND `$`_.Deleted -eq `$true}")
                        )
                    }
                    ElseIf ($Type -eq "AccountZonderMailbox") {
                        $Objects = Script-Connect-Server -Module "ActiveDirectory" -Command ([scriptblock]::Create(`
                                    "Get-ADUser -Filter * -Properties * | Sort-Object `$SortObject | Where-Object {`
                                `$`_.DistinguishedName -like `"*$Filter*`" -AND`
                                `$`_.DistinguishedName -notlike `"*OU=*External*`" -AND`
                                `$`_.DistinguishedName -notlike `"*OU=*Service*`" -AND`
                                `$`_.DistinguishedName -notlike `"*OU=*Admin*`" -AND`
                                `$`_.GivenName -ne `$null -AND`
                                `$`_.msExchRecipientTypeDetails -eq `$null -AND`
                                `$`_.EmailAddress -eq `$null}")
                        )
                    }
                    Else {
                        $Objects = Script-Connect-Server -Module "ActiveDirectory" -Command ([scriptblock]::Create(`
                                    "Get-ADUser -Filter * -Properties * | Sort-Object `$SortObject | Where-Object {`
                            `$`_.DistinguishedName -like `"*$Filter*`" -AND`
                            `$`_.DistinguishedName -notlike `"*OU=*External*`" -AND`
                            `$`_.DistinguishedName -notlike `"*OU=*Service*`" -AND`
                            `$`_.DistinguishedName -notlike `"*OU=*Admin*`" -AND`
                            `$`_.GivenName -ne `$null}")
                        )
                    }
                }
            }
            Else {
                # --------------------------------------------------------------------------------------------------------------------------------------------
                # Accounts inventariseren op een legacy-manier (ADSI) zodat deze ook werkt op SBS en 2003 (mits PowerShell is bijgewerkt). Dit gebeurd middels 
                # de DirectoryServices.ResultPropertyCollection, de properties moeten kleine letters blijven. Voor een output van alle properties:
                # $DirSearcher.FindAll().GetEnumerator() | Where-Object {$_.Properties.samaccountname -eq "Administrator"} | ForEach-Object {$_.Properties}
                # --------------------------------------------------------------------------------------------------------------------------------------------
                $Objects = @()
                $DirSearcher = New-Object System.DirectoryServices.DirectorySearcher([adsi]'')
                $DirSearcher.Filter = '(sAMAccountType=805306368)'
                $Objects = ($DirSearcher.FindAll().GetEnumerator() | Where-Object { $_.Properties.givenname -ne $null -AND $_.Properties.lastlogontimestamp -ne $null } | Select-Object Properties) | % { $_.Properties }
            }
            If (!$Objects) {
                Write-Host; Write-Host "    There are no $ObjectType detected with entered input, now returning to previous menu" -ForegroundColor Red
                Write-Host
                $Pause.Invoke()
                If ($FunctionName -like "*ActiveDirectory*") { Get-Variable -Exclude $global:StartupVariables | Remove-Variable -ErrorAction SilentlyContinue; Invoke-Expression -Command $global:MenuNameCategory }
                If ($FunctionName -like "*Exchange*") { Get-Variable -Exclude $global:StartupVariables | Remove-Variable -ErrorAction SilentlyContinue; Invoke-Expression -Command $global:MenuNameCategory }
            }
        }
        Else {
            $SkipMenu = $true
            $Objects = $null
            Write-Host -NoNewLine "  - Loading $ObjectType..."
            ForEach ($Mailbox in $Objecten) {
                $MailboxName = $Mailbox.Name
                $Objects += Script-Connect-Server -Module "ActiveDirectory" -Command ([scriptblock]::Create(`
                            "`$MailboxName = '$Mailbox';
                    Get-ADUser -Filter * -Properties * | Where-Object {`$`_.DisplayName -eq `$MailboxName}")
                )
            }
        }
    }
    If ($Type -eq "Adresboekbeleid") {
        $Objects = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("Get-AddressBookPolicy -DomainController '$PDC'"))
        $Objects = $Objects | Where-Object { $_.Name -ne $null } | Select-Object Name | Sort-Object Name
    }
    If ($Type -eq "Bestand") {
        If ($Filter -eq "SyncBackPro") {
            $Path = $env:LOCALAPPDATA + "\2BrightSparks\SyncBackPro\"
            $Filter = "TEMPLATE*_Settings.ini"
        }
        Else {
            $Path = $Drive
        }
        $Files = Get-ChildItem -Path $Path -Force -ErrorAction SilentlyContinue | Where-Object { $_.FullName -like "*$Filter" -AND $_.PSIsContainer -eq $false }
        If ($Filter -eq "PST") {
            If ($Functie -eq "Markeren") {
                $Mailboxes = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("Get-Mailbox -WarningAction SilentlyContinue"))
                $Mailboxes = $Mailboxes | Where-Object { $_.DisplayName -notlike "Discovery*" -AND $_.DisplayName -ne $null } | Select-Object Name, Alias, PrimarySMTPAddress, EmailAddresses | Sort-Object PrimarySMTPAddress
                ForEach ($Mailbox in $Mailboxes) {
                    $EmailAddresses = ($Mailbox | Select-Object @{Name = "EmailAddresses"; Expression = { $_.EmailAddresses | Where-Object { $_.PrefixString -eq "smtp" } | ForEach-Object { $_.SmtpAddress } } } | Select-Object EmailAddresses).EmailAddresses
                    ForEach ($EmailAddress in $EmailAddresses) {
                        $Match = $Files | Where-Object { $_.Name -like "$($EmailAddress)*.pst" -AND $_.PSIsContainer -eq $false }
                        ForEach ($File in $Match) {
                            $Property = New-Object PSObject
                            $Property | Add-Member -type NoteProperty -Name 'Name' -Value $Mailbox.Name
                            $Property | Add-Member -type NoteProperty -Name 'Alias' -Value $Mailbox.Alias
                            $Property | Add-Member -type NoteProperty -Name 'Email' -Value $Mailbox.PrimarySMTPAddress
                            $Property | Add-Member -type NoteProperty -Name 'File' -Value ($File.FullName).Substring(3)
                            $Property | Add-Member -type NoteProperty -Name 'Size' -Value $File.Length
                            $Objects += $Property
                        }
                    }
                }
                If (!$Objects) {
                    Write-Host; Write-Host "There is (or are) no mailbox(es) found to match the PST-file(s)." -ForegroundColor Red
                    Write-Host
                    $Pause.Invoke()
                    Get-Variable -Exclude $global:StartupVariables | Remove-Variable -ErrorAction SilentlyContinue; Invoke-Expression -Command $global:MenuNameCategory
                }
            }
            Else {
                ForEach ($File in $Files) {
                    $Property = New-Object PSObject
                    $Property | Add-Member -type NoteProperty -Name 'File' -Value ($File.FullName).Substring(3)
                    $Property | Add-Member -type NoteProperty -Name 'Size' -Value $File.Length
                    $Objects += $Property
                }
            }
        }
        Else {
            $Objects = $Files
        }
    }
    If ($Type -eq "Credentials") {
        $Objects = Get-StoredCredential -WarningAction SilentlyContinue -AsCredentialObject | Where-Object { $_.UserName -like "*@*" -OR $_.UserName -like "*\*" } | Sort-Object UserName
        If ($Objects) {
            $Objects = $Objects | Where-Object { $_.TargetName -like "*$Filter*" }
        }
        If (!$Objects) {
            Script-Module-SetHeaders -CurrentTask $CurrentTask -Name $FunctionName
            #$Subtitel.Invoke()
            Write-Host
            Write-Host "  There are no saved $ObjectTypes detected." -ForegroundColor Gray
            $global:Credential = "Aanmaken"
            RETURN
        }
    }
    If ($Type -eq "Drive") {
        $Disks = Get-WMIObject -Class Win32_LogicalDisk -Computername $Server | Where-Object { $_.DriveType -eq 3 -OR $_.DriveType -eq 4 }
        ForEach ($Object in $Disks) {
            If ($SelectieSize) {
                If ($SelectieSize -gt ($Object.FreeSpace / 1MB)) {
                    $Available = "Nee"
                }
                Else {
                    $Available = "Ja"
                }
            }
            Else {
                $Available = "Ja"
            }
            $Properties = @{
                Server      = $Server;
                Drive       = $Object.DeviceID;
                Name        = $Object.VolumeName;
                Network     = $Object.ProviderName;
                FreeSpace   = $Object.FreeSpace;
                Size        = $Object.Size;
                Beschikbaar = $Available
            }
            $NewObject = New-Object PSObject -Property $Properties
            $Objects += $NewObject
        }
    }
    If ($Type -eq "Emaildomein") {
        $Objects = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("Get-AcceptedDomain"))
        $Objects = $Objects | Where-Object { $_.DomainName -notlike "*.local" -AND $_.DomainName -notlike "*.int" -AND $_.DomainName -ne $null } | Sort-Object DomainName
    }
    If ($Type -like "Genereer*") {
        $GeneratedObjects = @()
        $Objects = @()
        $global:Handmatig = $null
        $GegenereerdeOptie = $null
        $FirstLetterSurNameDOT = $null
        $FirstLetterSurNameDASH = $null
        $FirstLetterSurNameSPACE = $null
        $FullNameDOT = $null
        $FullNameDASH = $null
        $FullNameSPACE = $null
        $ThreeLettersName = $null
        If ($GivenName) {
            $Count = 1
            ForEach ($Split in $GivenName.Split(' ')) {
                If ($Count -eq 1) {
                    $SplitFirst = $Split
                    $GeneratedObjects += $Split
                }
                Else {
                    $SplitInitials += $Split[0]
                }
                $Count++
            }
            $GeneratedObjects += ($SplitFirst + $SplitInitials)
            $GivenNameDOT = $GivenName.Replace(' ', '.')
            $GivenNameDASH = $GivenName.Replace(' ', '-')
            $GivenNameSPACE = $GivenName.Replace(' ', '')
            $ThreeLettersName = $GivenName.Substring(0, 3)
            $GeneratedObjects += $GivenNameDOT, $GivenNameDASH, $GivenNameSPACE
        }
        If ($SurName) {
            $SurNameDOT = $SurName.Replace(' ', '.')
            $SurNameDASH = $SurName.Replace(' ', '-')
            $SurNameSPACE = $SurName.Replace(' ', '')
            $FirstLetterSurNameDOT = ($GivenName[0] + " " + $SurNameSPACE).Replace(' ', '.')
            $FirstLetterSurNameDASH = ($GivenName[0] + " " + $SurName).Replace(' ', '-')
            $FirstLetterSurNameSPACE = ($GivenName[0] + $SurNameSPACE).Replace(' ', '')
            $FullNameDOT = ($GivenName + "." + $SurNameSPACE).Replace(' ', '.')
            $FullNameDASH = ($GivenName + "-" + $SurNameSPACE).Replace(' ', '-')
            $FullNameSPACE = ($GivenName + $SurNameSPACE).Replace(' ', '')
            $ThreeLettersName = $GivenName.Substring(0, 1) + $SurName.Substring(0, 2)
            $GeneratedObjects += $ThreeLettersName, $FirstLetterSurNameDOT, $FirstLetterSurNameDASH, $FirstLetterSurNameSPACE, $FullNameDOT, $FullNameDASH, $FullNameSPACE, $SurNameDOT, $SurNameDASH, $SurNameSPACE
        }
        Else {
            If ($GivenName) {
                $GeneratedObjects += $ThreeLettersName
            }
        }
        If ($GeneratedObjects) {
            [array]$GeneratedObjects = $GeneratedObjects.ToLower() | Select-Object -Unique | Sort-Object
        }
        Else {
            If ($Type -eq "GenereerGebruikersnaam" -AND $AccountType) {
                [array]$GeneratedObjects = $AccountType
            }
        }
        If ($Type -eq "GenereerGebruikersnaam") {
            $Objects = @()
            ForEach ($Object in $GeneratedObjects) {
                If ($AccountType) {
                    If ($Object -notlike "$AccountType*") {
                        $Object = $AccountType + "-" + $Object
                    }
                }
                $Object = $Object + "@" + $UPNSuffix
                $UserNameCheck = Script-Connect-Server -Module "ActiveDirectory" -Command ([scriptblock]::Create(`
                            "Get-ADUser -Filter * | Select-Object UserPrincipalName | Where-Object {`$`_.UserPrincipalName -eq '$Object'}")
                )
                If ($UserNameCheck) {
                    $Available = "Nee"
                }
                Else {
                    $Available = "Ja"
                }
                $Properties = @{Name = $Object; Beschikbaar = $Available }
                $NewObject = New-Object PSObject -Property $Properties
                $Objects += $NewObject
            }
        }
        If ($Type -eq "GenereerEmailadres") {
            $DisplayNameDOT = $DisplayName.Replace(' ', '.')
            $DisplayNameDASH = $DisplayName.Replace(' ', '-')
            $DisplayNameSPACE = $DisplayName.Replace(' ', '')
            $GeneratedObjects += $SamAccountName, $DisplayNameDOT, $DisplayNameDASH, $DisplayNameSPACE
            $GeneratedObjects = $GeneratedObjects.ToLower() | Select-Object -Unique | Sort-Object
            If ($Emaildomein) {
                $Objects = @()
                ForEach ($Object in $GeneratedObjects) {
                    $Object = $Object + "@" + $Emaildomein
                    $EmailadresCheck = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("Get-Mailbox '$Object' -ErrorAction SilentlyContinue"))
                    If ($EmailadresCheck) {
                        $Available = "No, already assigned to " + ($EmailadresCheck | Select-Object Name).Name
                    }
                    Else {
                        $Available = "Yes"
                    }
                    $Properties = @{Name = $Object; Beschikbaar = $Available }
                    $NewObject = New-Object PSObject -Property $Properties
                    $Objects += $NewObject
                }
            }
            $Objects = $Objects | Where-Object { $_.Beschikbaar -ne $null }
        }
    }
    If ($Type -eq "Groepen") {
        If ($Filter) {
            $Groups = Script-Connect-Server -Module "ActiveDirectory" -Command ([scriptblock]::Create(`
                        "Get-ADGroup -Filter * | Where-Object {`
                    `$`_.DistinguishedName -like `"*$Filter*`" -AND`
                    `$`_.Name -ne 'Domain Users' -AND`
                    `$`_.Name -ne 'Domeingebruikers'}")
            )
        }
        Else {
            $Groups = Script-Connect-Server -Module "ActiveDirectory" -Command ([scriptblock]::Create(`
                        "Get-ADGroup -Filter * | Where-Object {`
                    `$`_.DistinguishedName -notlike `"*CN=Builtin,DC*`" -AND`
                    `$`_.DistinguishedName -notlike `"*CN=*Computer*,DC*`" -AND`
                    `$`_.DistinguishedName -notlike `"*CN=Microsoft*,DC*`" -AND`
                    `$`_.DistinguishedName -notlike `"*OU=Microsoft*,DC*`" -AND`
                    `$`_.DistinguishedName -notlike `"*CN=*Server*,DC*`" -AND`
                    `$`_.DistinguishedName -notlike `"*CN=Users,DC*`" -AND`
                    `$`_.DistinguishedName -notlike `"*CN=*Werkstation*,DC*`" -AND`
                    `$`_.DistinguishedName -notlike `"*CN=*Workstation*,DC*`"}")
            )
        }
        $Objects = @()
        $Groups = ($Groups | Where-Object { $_.Name -ne $null } | Select-Object Name | Sort-Object Name).Name
        ForEach ($Object in $Groups) {
            If ($Objecten -contains $Object) {
                $Properties = @{Name = $Object; Markeren = 'Ja' }
            }
            Else {
                $Properties = @{Name = $Object; Markeren = 'Nee' }
            }
            $NewObject = New-Object PSObject -Property $Properties
            $Objects += $NewObject
        }
    }
    If ($Type -eq "Homefolders") {
        $Objects = $null
        $Objects = @()
        ForEach ($Account in $Objecten) {
            $HomePath = $null
            $ADSI = [ADSI]('LDAP://{0}' -f $Account.DistinguishedName)
            $CheckHomePath = Try { $ADSI.InvokeGet('TerminalServicesHomeDirectory').Length } Catch { 0 }
            If ($CheckHomePath -gt 1) {
                $HomePath = $ADSI.InvokeGet('TerminalServicesHomeDirectory')
                $Size = Invoke-Command -ComputerName $HomePath.Split('\')[2] -ArgumentList $HomePath -ScriptBlock { param($HomePath);
                    $FolderUsed = (robocopy.exe "$HomePath" C:\Temp /e /l /xj /r:1 /w:1 /nfl /ndl /nc /fp /np /njh) | Where-Object { $_ -like '*Bytes*' } | ForEach-Object { (-split $_)[2] + (-split $_)[3] };
                    If ($FolderUsed -like "*g") { $FolderUsed = ($FolderUsed -replace 'g'); $DataFolder = $FolderUsed -as [decimal]; $DataFolder = "{0:0}" -f ($DataFolder * 1024) };
                    If ($FolderUsed -like "*m") { $FolderUsed = ($FolderUsed -replace 'm'); $DataFolder = $FolderUsed -as [decimal]; $DataFolder = "{0:0}" -f ($DataFolder) };
                    If ($FolderUsed -like "*k") { $FolderUsed = ($FolderUsed -replace 'k'); $DataFolder = $FolderUsed -as [decimal]; $DataFolder = "{0:0}" -f ($DataFolder / 1024) };
                    If ($FolderUsed -eq "00") { $DataFolder = 0 }; $DataFolder }
            }
            Else {
                $HomePath = "No homefolder detected in AD"
                $Size = "0"
            }
            If ($Size -eq $null) {
                $HomePath = "No homefolder present"
                $Size = "0"
            }
            $Properties = @{Name = $Account.Name; DN = $Account.DistinguishedName; HomePath = $HomePath; Size = $Size }
            $NewObject = New-Object PSObject -Property $Properties
            $Objects += $NewObject
        }
    }
    If ($Type -like "Mailbox*") {
        If (!$Office365Exchange) {
            $Databases = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("Get-MailboxDatabase -Status"))
            $Databases = ($Databases | Where-Object { $_.Mounted -eq $true }).Name
        }
        Else {
            $Databases = 1
        }
        $TotalSize = 0
        $TotalSizeFormat = $null
        $ObjectsFolderStatistics = @()
        $ObjectsRegionalConfiguration = @()
        $ObjectsStatistics = @()
        $CounterDatabases = 0
        ForEach ($Database in $Databases) {
            $CounterDatabases++
            If ($Type -eq "MailboxHerstel") {
                If (!$Office365Exchange) {
                    $Objects += Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("Get-MailboxDatabase '$Database' | Get-MailboxStatistics -WarningAction SilentlyContinue"))
                }
                $Objects = $Objects | Where-Object { $_.DisconnectReason -like "*Disabled*" }
            }
            Else {
                If (!$Objects) {
                    If (!$Office365Exchange) {
                        $GetObjects = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("Get-Mailbox -Database '$Database' -ResultSize Unlimited -WarningAction SilentlyContinue"))
                        $Objects += $GetObjects | Sort-Object $SortObject | Where-Object { $_.DisplayName -notlike 'Discovery*' -AND $_.DisplayName -ne $null -AND $_.OrganizationalUnit -like "*$Filter*" }
                    }
                    Else {
                        $GetObjects = Script-Connect-Server -Module "Office365" -Command ([scriptblock]::Create("Get-Mailbox -ResultSize Unlimited"))
                        $Objects += $GetObjects | Sort-Object $SortObject | Where-Object { $_.DisplayName -notlike 'Discovery*' -AND $_.DisplayName -ne $null }
                    }
                }
                Else {
                    If (!$Objecten) {
                        If (!$Office365Exchange) {
                            $GetObjects = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("Get-Mailbox -Database '$Database' -WarningAction SilentlyContinue"))
                            $Objects += $GetObjects | Sort-Object $SortObject | Where-Object { $_.DisplayName -notlike 'Discovery*' -AND $_.DisplayName -ne $null -AND $_.OrganizationalUnit -like "*$Filter*" }
                        }
                        Else {
                            $GetObjects = Script-Connect-Server -Module "Office365" -Command ([scriptblock]::Create("Get-Mailbox -ResultSize Unlimited"))
                            $Objects += $GetObjects | Sort-Object $SortObject | Where-Object { $_.DisplayName -notlike 'Discovery*' -AND $_.DisplayName -ne $null }
                        }
                    }
                    Else {
                        $SkipMenu = $true
                        If (!$Office365Exchange) {
                            $AllObjects = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("Get-Mailbox -Database '$Database' -ResultSize Unlimited -WarningAction SilentlyContinue"))
                        }
                        Else {
                            $AllObjects = Script-Connect-Server -Module "Office365" -Command ([scriptblock]::Create("Get-Mailbox -ResultSize Unlimited"))
                        }
                        $Objects = $null
                        ForEach ($Account in $Objecten) {
                            $Objects += $AllObjects | Where-Object { $_.DisplayName -eq $Account.DisplayName }
                        }
                    }
                }
                If ($Objects) {
                    If ($DetectDateItemsServerAndSize) {
                        If (!$Office365Exchange) {
                            $LoadingDateItemsServerAndSize = {
                                Script-Module-SetHeaders -CurrentTask $CurrentTask -Name $FunctionName
                                Write-Host -NoNewLine ("Indexing details of $ObjectType per database, this can take a few minutes... (" + $CounterDatabases + "/" + [string]$Databases.Count + ")")
                            }
                            $LoadingDateItemsServerAndSize.Invoke()
                            $ObjectsStatistics += Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("Get-MailboxDatabase '$Database' | Get-MailboxStatistics -WarningAction SilentlyContinue | Select-Object DisplayName, LastLogonTime, TotalItemSize, ItemCount, MailboxGuid, ServerName"))
                        }
                    }
                }
            }
        }
        If ($Objects) {
            If ($FunctionName -like "*Exchange*Move*") {
                $Objects = $Objects | Where-Object { $_.MailboxMoveRemoteHostName -ne $ServerDestination }
            }
            Else {
                If ($Objects | Where-Object { $_.MailboxMoveRemoteHostName -ne $null }) {
                    Do {
                        Write-Host
                        Write-Host
                        Write-Host -NoNewline "Please note: " -ForegroundColor Red; Write-Host "There are mailboxes detected in the process of a move migration."; Write-Host -NoNewline "Would you like to ("; Write-Host -NoNewLine "E" -ForegroundColor Yellow; Write-Host -NoNewLine ")xclude or ("; Write-Host -NonewLine "A" -ForegroundColor Yellow; Write-Host -NoNewLine ")dd these?: "
                        $InputChoice = Read-Host
                        $InputKey = @("E", "A") -contains $InputChoice
                        If (!$InputKey) {
                            Write-Host "Please use the letters above as input" -ForegroundColor Red; Write-Host
                        }
                    } Until ($InputKey)
                    Switch ($InputChoice) {
                        "E" {
                            $Objects = $Objects | Where-Object { !($_.MailboxMoveRemoteHostName) }
                        }
                    }
                }
            }
            $ObjectsMailForward = @()
            ForEach ($Object in $Objects) {
                If ($Object.ForwardingAddress) {
                    If (!$Office365Exchange) {
                        [string]$ObjectForwardingAddress = $Object.ForwardingAddress
                        $ObjectForwardingAddress = $ObjectForwardingAddress.Split("/")[-1]
                        $ForwardingAddress = (Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("Get-Mailbox '$ObjectForwardingAddress' -WarningAction SilentlyContinue -ErrorAction SilentlyContinue | Select-Object PrimarySmtpAddress"))).PrimarySmtpAddress
                        If (!$ForwardingAddress) {
                            $ForwardingAddress = (Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("Get-Contact '$ObjectForwardingAddress' -WarningAction SilentlyContinue -ErrorAction SilentlyContinue | Select-Object WindowsEmailAddress"))).WindowsEmailAddress
                        }
                    }
                    Else {
                        $ObjectForwardingAddress = $Object.ForwardingAddress
                        $ForwardingAddress = ($Objects | Where-Object { $_.Name -eq $ObjectForwardingAddress } | Select-Object PrimarySmtpAddress).PrimarySmtpAddress
                    }
                    $Properties = @{Name = $Object.DisplayName; ForwardingAddress = $ForwardingAddress; DeliverToMailboxAndForward = $Object.DeliverToMailboxAndForward }
                    $ObjectMailForward = New-Object PSObject -Property $Properties
                    $ObjectsMailForward += $ObjectMailForward
                }
            }
            If ($Office365Exchange) {
                $AzureADSubscribedSku = Get-AzureADSubscribedSku
                $GetAzureADObjects = Get-AzureADUser -All $True | Select-Object DisplayName, UserPrincipalName, @{Name = "DN"; Expression = { ($_.ExtensionProperty)["onPremisesDistinguishedName"] } }, AssignedLicenses
                If ($Filter -eq "*") {
                    $AzureADObjects = $GetAzureADObjects
                }
                Else {
                    $AzureADObjects = $GetAzureADObjects | Where-Object { $_.DN -like "*OU=*$Filter*" }
                    $TempObjects = $Objects
                    $Objects = @()
                    ForEach ($AzureADObject in $AzureADObjects) {
                        $Objects += $TempObjects | Where-Object { $_.Name -eq $AzureADObject.DisplayName }
                    }
                }
            }
            If ($FunctionName -like "*Overview*") {
                If ($DetectDateItemsServerAndSize) {
                    If ($Office365Exchange) {
                        $CounterDateItemsServerAndSize = 0
                        ForEach ($Object in $Objects) {
                            $CounterDateItemsServerAndSize++
                            $LoadingDateItemsServerAndSize = {
                                Script-Module-SetHeaders -CurrentTask $CurrentTask -Name $FunctionName
                                Write-Host -NoNewLine ("Indexing details of $ObjectType, this can take a few minutes... (" + $CounterDateItemsServerAndSize + "/" + [string]$Objects.Count + ")")
                            }
                            $LoadingDateItemsServerAndSize.Invoke()
                            $ObjectAlias = $Object.Alias
                            $ObjectsStatistics += Script-Connect-Server -Module "Office365" -Command ([scriptblock]::Create("Get-Mailbox '$ObjectAlias' | Get-MailboxStatistics -WarningAction SilentlyContinue | Select-Object DisplayName, LastLogonTime, TotalItemSize, ItemCount, MailboxGuid"))
                        }
                    }
                }
                If ($DetectPermissions) {
                    $CounterPermissions = 0
                    $ObjectsFullAccess = @()
                    $ObjectsSendAs = @()
                    $ObjectsSendOnBehalf = @()
                    ForEach ($Object in $Objects) {
                        $CounterPermissions++
                        $LoadingPermissions = {
                            Script-Module-SetHeaders -CurrentTask $CurrentTask -Name $FunctionName
                            If ($LoadingDateItemsServerAndSize) { $LoadingDateItemsServerAndSize.Invoke(); Write-Host };
                            Write-Host -NoNewLine ("Indexing permissions of $ObjectType, this can take a few minutes... (" + $CounterPermissions + "/" + [string]$Objects.Count + ")")
                        }
                        $LoadingPermissions.Invoke()
                        $ObjectAlias = $Object.Alias
                        $ObjectDN = $Object.DistinguishedName
                        $ObjectUPN = $Object.UserPrincipalName
                        # -----------
                        # Full Access
                        # -----------
                        If (!$Office365Exchange) {
                            $GetFullAccess = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("Get-MailboxPermission -Identity '$ObjectUPN' -DomainController '$PDC'"))
                            $GetFullAccess = ($GetFullAccess | Where-Object { $_.IsInherited -eq $false -AND $_.User -notlike "*ELF" } | Select-Object User).User
                            If ($GetFullAccess) {
                                ForEach ($User in $GetFullAccess) {
                                    If ($User.RawIdentity) {
                                        $User = ($User.RawIdentity).Split("\")[1]
                                    }
                                    Else {
                                        $User = $User.Split("\")[1]
                                    }
                                    If ($User) {
                                        $GetUPN = Script-Connect-Server -Module "ActiveDirectory" -Command ([scriptblock]::Create(`
                                                    "(Get-ADObject -Filter `"(SamAccountName -eq '$User')`" -Properties * | Select-Object UserPrincipalName).UserPrincipalName")
                                        )
                                        If ($GetUPN) {
                                            $ObjectFullAccess = @()
                                            $Properties = @{Name = $Object.DisplayName; FullAccess = $GetUPN }
                                            $ObjectFullAccess = New-Object PSObject -Property $Properties
                                            $ObjectsFullAccess += $ObjectFullAccess
                                        }
                                    }
                                }
                            }
                        }
                        Else {
                            $GetFullAccess = Script-Connect-Server -Module "Office365" -Command ([scriptblock]::Create("Get-MailboxPermission -Identity '$ObjectUPN'"))
                            $ObjectsFullAccess += $GetFullAccess | Where-Object { $_.AccessRights -like "*FullAccess*" -AND $_.IsInherited -eq $false -AND $_.User -notlike "*\*" -AND $_.User -notlike "S-1-5-21*" -AND $_.User -ne $ObjectUPN } | Select-Object @{Name = "Name"; Expression = { $Object.Name } }, @{Name = "FullAccess"; Expression = { $_.User } }
                        }
                        # -------
                        # Send As
                        # -------
                        If (!$Office365Exchange) {
                            $GetSendAs = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("Get-ADPermission -Identity '$ObjectDN' -DomainController '$PDC'"))
                            $GetSendAs = ($GetSendAs | Where-Object { ($_.ExtendedRights -like "*Send-As*") -AND ($_.User -notlike "*ELF") } | Select-Object User).User
                            If ($GetSendAs) {
                                ForEach ($User in $GetSendAs) {
                                    If ($User.RawIdentity) {
                                        $User = ($User.RawIdentity).Split("\")[1]
                                    }
                                    Else {
                                        $User = $User.Split("\")[1]
                                    }
                                    If ($User) {
                                        $GetUPN = Script-Connect-Server -Module "ActiveDirectory" -Command ([scriptblock]::Create(`
                                                    "(Get-ADObject -Filter `"(SamAccountName -eq '$User')`" -Properties * | Select-Object UserPrincipalName).UserPrincipalName")
                                        )
                                        If ($GetUPN) {
                                            $ObjectSendAs = @()
                                            $Properties = @{Name = $Object.DisplayName; SendAs = $GetUPN }
                                            $ObjectSendAs = New-Object PSObject -Property $Properties
                                            $ObjectsSendAs += $ObjectSendAs
                                        }
                                    }
                                }
                            }
                        }
                        Else {
                            $GetSendAs = Script-Connect-Server -Module "Office365" -Command ([scriptblock]::Create("Get-RecipientPermission -Identity '$ObjectUPN'"))
                            $ObjectsSendAs += $GetSendAs | Where-Object { $_.AccessRights -eq "SendAs" -AND $_.Trustee -like "*@*" -AND $_.Trustee -ne $ObjectUPN } | Select-Object @{Name = "Name"; Expression = { $Object.Name } }, @{Name = "SendAs"; Expression = { $_.Trustee } }
                        }
                        # --------------
                        # Send On Behalf
                        # --------------
                        $GetSendOnBehalf = $Object.GrantSendOnBehalfTo
                        If ($GetSendOnBehalf) {
                            ForEach ($User in $GetSendOnBehalf) {
                                If (!$Office365Exchange) {
                                    If ($User.Name) {
                                        $User = ($User.Name)
                                    }
                                    Else {
                                        $User = $User.Split('/')[-1]
                                    }
                                    If ($User) {
                                        $GetUPN = Script-Connect-Server -Module "ActiveDirectory" -Command ([scriptblock]::Create(`
                                                    "(Get-ADObject -Filter `"(SamAccountName -eq '$User')`" -Properties * | Select-Object UserPrincipalName).UserPrincipalName")
                                        )
                                    }
                                }
                                Else {
                                    $GetUPN = ($GetAzureADObjects | Where-Object { $_.DisplayName -eq $User } | Select-Object UserPrincipalName).UserPrincipalName
                                }
                                If ($GetUPN -AND $GetUPN -ne $ObjectUPN) {
                                    $ObjectSendOnBehalf = @()
                                    $Properties = @{Name = $Object.DisplayName; SendOnBehalf = $GetUPN }
                                    $ObjectSendOnBehalf = New-Object PSObject -Property $Properties
                                    $ObjectsSendOnBehalf += $ObjectSendOnBehalf
                                }
                            }
                        }
                    }
                }
                If ($DetectLanguage) {
                    $CounterLanguage = 0
                    $ObjectLanguageStatistics = @()
                    ForEach ($Object in $Objects) {
                        $CounterLanguage++
                        $LoadingLanguage = {
                            Script-Module-SetHeaders -CurrentTask $CurrentTask -Name $FunctionName
                            If ($LoadingDateItemsServerAndSize) { $LoadingDateItemsServerAndSize.Invoke(); Write-Host };
                            If ($LoadingPermissions) { $LoadingPermissions.Invoke(); Write-Host };
                            Write-Host -NoNewLine ("Indexing language of $ObjectType, this can take a few minutes... (" + $CounterLanguage + "/" + [string]$Objects.Count + ")")
                        }
                        $LoadingLanguage.Invoke()
                        $ObjectAlias = $Object.Alias
                        If (!$Office365Exchange) {
                            $ObjectFolder = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("Get-MailboxFolderStatistics '$ObjectAlias' -WarningAction SilentlyContinue"))
                            $ObjectRegion = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("Get-MailboxRegionalConfiguration '$ObjectAlias' -WarningAction SilentlyContinue"))
                        }
                        Else {
                            $ObjectFolder = Script-Connect-Server -Module "Office365" -Command ([scriptblock]::Create("Get-MailboxFolderStatistics '$ObjectAlias' | Select-Object Identity, FolderPath, ItemsInFolderAndSubfolders"))
                            $ObjectRegion = Script-Connect-Server -Module "Office365" -Command ([scriptblock]::Create("Get-MailboxRegionalConfiguration '$ObjectAlias' -WarningAction SilentlyContinue | Select-Object Language"))
                        }
                        $ObjectsLanguageStatistics += $ObjectFolder | Where-Object { $_.Identity -like "*Boîte de réception" -OR $_.Identity -like "*Inbox" -OR $_.Identity -like "*Posteingang" -OR $_.Identity -like "*Postvak IN" } | Select-Object @{Name = "Alias"; Expression = { $ObjectAlias } }, @{Name = "Language"; Expression = { ($ObjectRegion | Select-Object Language).Language } }, FolderPath, ItemsInFolder | Sort-Object ItemsInFolder -Descending | Select-Object -First 1
                    }
                }
                Script-Disconnect-Server
            }
        }
        Else {
            Write-Host
            Write-Host "There are no mailboxes detected" -ForegroundColor Red
            Write-Host
            $Pause.Invoke()
            Get-Variable -Exclude $global:StartupVariables | Remove-Variable -ErrorAction SilentlyContinue; Invoke-Expression -Command $global:MenuNameCategory
        }
    }
    If ($Type -like "Office365*") {
        If ($Type -like "Office365Account*") {
            Write-Host; Write-Host -NoNewLine "Indexing Azure AD licenses..."
            $AzureADSubscribedSku = Get-AzureADSubscribedSku
            Write-Host; Write-Host -NoNewLine "Indexing Azure AD accounts..."
            $Objects = Get-AzureADUser -All $True
        }
        If ($Type -eq "Office365Licenties") {
            $Objects = Get-AzureADSubscribedSku
        }
        Script-Disconnect-Server
    }
    If ($Type -eq "UPNSuffix") {
        $Objects = @()
        $UPNSuffixes = Script-Connect-Server -Module "ActiveDirectory" -Command ([scriptblock]::Create(`
                    "`$DomainDN = (Get-ADDomain).DistinguishedName;
            `$UPNDN = `"cn=Partitions,cn=Configuration,`$DomainDN`";
            Get-ADObject -Identity `$UPNDN -Properties UPNSuffixes | Select-Object -ExpandProperty UPNSuffixes")
        )
        ForEach ($Object in $UPNSuffixes) {
            $Objects += $Object
        }
        $Objects += (Get-WmiObject Win32_ComputerSystem).Domain
        $Objects = $Objects | Sort-Object
    }
    # ------------------------------------------------------------------------------------------------
    # Array: Converteren van bovenstaande objecten naar een globaal array met consistente naamstelling
    # ------------------------------------------------------------------------------------------------
    $Counter = 1
    $global:Column1Length = 0
    $global:Column2Length = 0
    $global:Column3Length = 0
    $global:Column4Length = 0
    $global:Column5Length = 0
    $ObjectTotalSize = 0
    $SortObjectFormat = $null
    If ($Objects.Count) {
        [string]$ObjectsCount = $Objects.Count
    }
    Else {
        [string]$ObjectsCount = 1
    }
    ForEach ($Object in $Objects) {
        Script-Module-SetHeaders -CurrentTask $CurrentTask -Name $FunctionName
        Write-Host -NoNewLine ("  - Indexing objects... (" + $Counter + "/" + [string]$ObjectsCount + ")")
        If ($Object.Markeren) {
            $ObjectMarkeren = $Object.Markeren
        }
        Else {
            $ObjectMarkeren = "Nee"
        }
        If ($SkipMenu -eq $true) {
            $ObjectMarkeren = "Ja"
        }
        $Properties = @{Nr = $Counter; Markeren = $ObjectMarkeren }
        If ($Type -like "Account*") {
            $ObjectUPN = [string]$Object.UserPrincipalName
            If ($AzureADConnected) {
                $ObjectName = $Object.DisplayName
                $ObjectSurname = $Object.Surname
                $ObjectEmail = $Object.Mail
                $ObjectTitle = $Object.JobTitle
                $ObjectCompany = $Object.CompanyName
                $ObjectCity = $Object.City
                $ObjectState = $Object.State
                $ObjectCountry = $Object.Country
                $ObjectOfficePhone = $Object.TelephoneNumber
                $ObjectMobilePhone = $Object.Mobile
                $ObjectOffice365Guid = [System.Convert]::ToBase64String(($Object.Guid).ToByteArray())
                $ObjectDN = ($Object | Select-Object @{Name = "DN"; Expression = { ($_.ExtensionProperty)["onPremisesDistinguishedName"] } }).DN
                If (!$ObjectDN) {
                    $ObjectDN = "n.a."
                }
                If ($ObjectName -eq "On-Premises Directory Synchronization Service Account") {
                    $ObjectName = "On-Premises DirSync Account"
                }
                $ArrayLicense = @()
                $ObjectLicense = $null
                ForEach ($SkuId in (($Object.AssignedLicenses).SkuId | Sort-Object)) {
                    $Sku = ($AzureADSubscribedSku | Where-Object { $_.SkuId -eq $SkuId } | Select-Object SkuPartNumber).SkuPartNumber
                    If ($Sku -eq "ENTERPRISEPACK") { $ArrayLicense += "Enterprise E3" }
                    ElseIf ($Sku -eq "ENTERPRISEPREMIUM") { $ArrayLicense += "Enterprise E5" }
                    ElseIf ($Sku -eq "ENTERPRISEPREMIUM_NOPSTNCONF") { $ArrayLicense += "Enterprise E5 (No conf)" }
                    ElseIf ($Sku -eq "EXCHANGESTANDARD") { $ArrayLicense += "Exchange Online" }
                    ElseIf ($Sku -eq "MICROSOFT_BUSINESS_CENTER") { $ArrayLicense += "Microsoft Business Center" }
                    ElseIf ($Sku -eq "O365_BUSINESS_ESSENTIALS") { $ArrayLicense += "Business Essentials" }
                    ElseIf ($Sku -eq "O365_BUSINESS_PREMIUM") { $ArrayLicense += "Business Premium" }
                    ElseIf ($Sku -eq "POWERFLOW_P1") { $ArrayLicense += "PowerApps P1" }
                    ElseIf ($Sku -eq "POWER_BI_PRO") { $ArrayLicense += "Power BI Pro" }
                    ElseIf ($Sku -eq "POWER_BI_STANDARD") { $ArrayLicense += "Power BI" }
                    ElseIf ($Sku -eq "VISIOCLIENT") { $ArrayLicense += "Visio" }
                    ElseIf ($Sku -eq "FLOW_FREE" -OR $Sku -eq "POWERAPPS_VIRAL") { BREAK }
                    Else { $ArrayLicense += $Sku }
                }
                [string]$ObjectLicense = ($ArrayLicense | Sort-Object) -join ", "
                If (!$ObjectLicense) { $ObjectLicense = "n.a." }
                If ($Object.AccountEnabled -eq $false) { $ObjectLicense = "Geblokkeerd" }
                If ($Object.UserType -eq "Guest") {
                    $ObjectUPN = $Object.Mail
                    $ObjectLicense = "Gast (extern)"
                }
                Else {
                    $ObjectUPN = $Object.UserPrincipalName
                }
                If ($Object.LastDirSyncTime) {
                    $ObjectSyncStatus = "Sync met AD"
                }
                Else {
                    $ObjectSyncStatus = "Office 365"
                }
                If ($ObjectLicense.Length -gt $global:Column4Length) {
                    $global:Column4Length = $ObjectLicense.Length
                }
                $Properties += @{
                    Alias            = $Object.MailNickName
                    EmailAddresses   = $Object.ProxyAddresses
                    AccountEnabled   = $Object.AccountEnabled
                    DirSyncEnabled   = $Object.DirSyncEnabled
                    SyncStatus       = $ObjectSyncStatus
                    License          = $ObjectLicense
                    Office365Guid    = $ObjectOffice365Guid
                    PasswordPolicies = $Object.PasswordPolicies
                }
            }
            Else {
                If ($OS -eq "SBS2008") {
                    If ([string]$Object.sn.Length -ne 0) {
                        $ObjectName = [string]$Object.givenname + " " + [string]$Object.sn
                        $ObjectSurname = [string]$Object.sn
                    }
                    Else {
                        $ObjectName = [string]$Object.givenname
                    }
                    $ObjectLastLogonDate = ([datetime]::FromFileTime([string]$Object.lastlogontimestamp).ToString('g'))
                    $ObjectOffice = [string]$_.physicaldeliveryofficename
                    $ObjectHomePage = [string]$_.wwwhomepage
                    $ObjectPOBox = [string]$_.postofficebox
                    $ObjectCity = [string]$_.l
                    $ObjectState = [string]$_.st
                    $ObjectCountry = [string]$Object.co
                    $ObjectOfficePhone = [string]$_.telephonenumber
                    $ObjectMobilePhone = [string]$_.mobile
                    $ObjectFax = [string]$_.facsimiletelephonenumber
                }
                Else {
                    $ObjectName = $Object.Name
                    $ObjectSurname = $Object.Surname
                    $ObjectLastLogonDate = $Object.LastLogonDate
                    $ObjectOffice = $Object.Office
                    $ObjectHomePage = $Object.HomePage
                    $ObjectPOBox = $Object.POBox
                    $ObjectCity = $Object.City
                    $ObjectState = $Object.State
                    $ObjectCountry = $Object.Country
                    $ObjectOfficePhone = $Object.OfficePhone
                    $ObjectMobilePhone = $Object.MobilePhone
                    $ObjectFax = $Object.Fax
                }
                $ObjectDN = [string]$Object.distinguishedname
                $ObjectEmail = [string]$Object.mail
                $ObjectCompany = [string]$Object.company
                $ObjectTitle = [string]$Object.title
                If ($Type -eq "AccountHerstel" -OR $Type -eq "AccountRecycleBin") {
                    $ObjectName = $ObjectName.Split("`n")[0]
                    $ObjectDN = $Object.LastKnownParent
                    $ObjectGUID = $Object.ObjectGUID
                    $Properties += @{GUID = $ObjectGUID }
                    If (!$ObjectUPN) {
                        $ObjectUPN = "n.a."
                    }
                    $ObjectLastLogonDate = $Object.Modified
                }
                If ($ObjectLastLogonDate) {
                    $ObjectLastLogonDate = Get-Date $ObjectLastLogonDate
                }
                Else {
                    $ObjectLastLogonDate = Get-Date 01-01-2000
                }
                $ADSI = [ADSI]('LDAP://{0}' -f $ObjectDN)
                $CheckHomePath = Try { $ADSI.InvokeGet('TerminalServicesHomeDirectory').Length } Catch { 0 }
                $CheckHomeDrive = Try { $ADSI.InvokeGet('TerminalServicesHomeDrive').Length } Catch { 0 }
                $CheckRDPPath = Try { $ADSI.InvokeGet('TerminalServicesProfilePath').Length } Catch { 0 }
                If ($CheckHomePath -gt 1) { $HomePath = $ADSI.InvokeGet('TerminalServicesHomeDirectory') } Else { $HomePath = $null }
                If ($CheckHomeDrive -gt 1) { $HomeDrive = $ADSI.InvokeGet('TerminalServicesHomeDrive') } Else { $HomeDrive = $null }
                If ($CheckRDPPath -gt 1) { $RDPPath = $ADSI.InvokeGet('TerminalServicesProfilePath') } Else { $RDPPath = $null }
                $global:Column4Length = 20
                $Properties += @{
                    Initials                   = [string]$Object.initials
                    SamAccountName             = ([string]$Object.samaccountname)
                    LastLogonDate              = $ObjectLastLogonDate
                    Description                = [string]$Object.description
                    HomePhone                  = [string]$Object.homephone
                    Fax                        = $ObjectFax
                    IPPhone                    = [string]$Object.ipphone
                    Pager                      = [string]$Object.pager
                    msExchRecipientTypeDetails = [string]$Object.msExchRecipientTypeDetails
                    extensionAttribute1        = $Object.extensionAttribute1
                    HomePath                   = $HomePath
                    HomeDrive                  = $HomeDrive
                    RDPPath                    = $RDPPath
                }
            }
            $global:Subtree = $null
            $TempOU = $null
            $OUs = $null
            $DC = $null
            If ($ObjectDN -ne "n.a.") {
                If ($ObjectDN -like "*CN=Users,DC*") {
                    $ObjectDN = $ObjectDN.Replace('CN=Users,DC', 'OU=Users,DC')
                }
                ForEach ($Item in ($ObjectDN.replace('\,', '~').split(","))) {
                    Switch -Regex ($Item.TrimStart().Substring(0, 3)) {
                        "CN=" {
                            $CN = '/' + $Item.replace("CN=", "")
                        }
                        "OU=" {
                            $TempOU += , $Item.replace("OU=", "")
                            $TempOU += '\'
                        }
                        "DC=" {
                            $DC += $Item.replace("DC=", "")
                            $DC += '.'
                        }
                    }
                }
                If ($TempOU.Count -ge 1) {
                    For ($I = $TempOU.Count; $I -ge 0; $I --) {
                        $OUs += $TempOU[$I]
                    }
                    $global:Subtree += $OUs.Substring(1)
                }
                If ($global:Subtree -like "*DEL:*") {
                    $global:Subtree = "Bestaat niet meer"
                }
            }
            Else {
                $global:Subtree = "n.a."
            }
            If ($ObjectName.Length -gt $global:Column2Length) {
                $global:Column2Length = $ObjectName.Length
            }
            If ($ObjectUPN.Length -gt $global:Column3Length) {
                $global:Column3Length = $ObjectUPN.Length
            }
            $Properties += @{
                Name              = $ObjectName
                DisplayName       = [string]$Object.displayname
                GivenName         = [string]$Object.givenname
                Surname           = $ObjectSurname
                UserPrincipalName = $ObjectUPN
                DistinguishedName = [string]$ObjectDN
                DN                = $global:Subtree
                Email             = $ObjectEmail
                Office            = $ObjectOffice
                Title             = $ObjectTitle
                Department        = [string]$Object.department
                Company           = $ObjectCompany
                HomePage          = $ObjectHomePage
                StreetAddress     = [string]$Object.streetaddress
                POBox             = $ObjectPOBox
                City              = $ObjectCity
                State             = $ObjectState
                PostalCode        = [string]$Object.postalcode
                Country           = $ObjectCountry
                OfficePhone       = $ObjectOfficePhone
                MobilePhone       = $ObjectMobilePhone
            }
        }
        If ($Type -eq "Adresboekbeleid") {
            $ObjectName = $Object.Name
            $Properties += @{Name = $ObjectName }
        }
        If ($Type -eq "Bestand") {
            If ($Filter -eq "PST") {
                $FilePath = $Drive + $Object.File
                $FileUNC = $PathUNC + "\" + $Object.File
                $Size = $Object.Size
                $SizeFormat = ("{0:0}" -f (($Size / 1024) / 1024))
                $SizeFormat = $SizeFormat.Replace('.', '') + " MB"
                $Spaces = $null
                If ($SizeFormat.Length -ge 8) { $Spaces = " " * 1 }
                ElseIf ($SizeFormat.Length -ge 7) { $Spaces = " " * 2 }
                ElseIf ($SizeFormat.Length -ge 6) { $Spaces = " " * 3 }
                ElseIf ($SizeFormat.Length -ge 5) { $Spaces = " " * 4 }
                ElseIf ($SizeFormat.Length -ge 4) { $Spaces = " " * 5 }
                Else { $Spaces = " " * 6 }
                $SizeFormat = $Spaces + $SizeFormat
                $ObjectTotalSize += $Size
                If ($FilePath.Length -gt $global:Column2Length) {
                    $global:Column2Length = $FilePath.Length
                }
                $Properties += @{Name = $FilePath; PST = $Object.File; FileUNC = $FileUNC; Size = $Size; SizeFormat = $SizeFormat }
                If ($Functie -eq "Markeren") {
                    If ($Object.Name.Length -gt $global:Column3Length) {
                        $global:Column3Length = $Object.Name.Length
                    }
                    $Properties += @{MailboxName = $Object.Name; Alias = $Object.Alias; PrimarySmtpAddress = $Object.Email }
                }
            }
            Else {
                $ObjectName = $Path + $Object
                If ($ObjectName.Length -gt $global:Column2Length) {
                    $global:Column2Length = $ObjectName.Length
                }
                $Properties += @{Name = $ObjectName }
            }
        }
        If ($Type -eq "Credentials") {
            $ObjectName = $Object.UserName
            $ObjectTargetName = ($Object.TargetName).Split('=')[1]
            $Properties += @{Name = $ObjectName; TargetName = $ObjectTargetName }
        }
        If ($Type -eq "Drive") {
            $ObjectBeschikbaar = $Object.Beschikbaar
            $ObjectDrive = $Object.Drive
            $ObjectLabel = $Object.Name
            $ObjectName = "(" + $ObjectDrive + ")"
            If ($ObjectLabel) {
                $ObjectName += " " + $ObjectLabel
            }
            $ObjectSize = $Object.Size
            $ObjectFreeSpace = $Object.FreeSpace
            If ($Object.Network) {
                $ObjectNetwork = $Object.Network
                $ObjectName += " " + "(" + $ObjectNetwork + ")"
            }
            Else {
                $ObjectNetwork = "n.a."
            }
            $Spaces = $null
            If ($ObjectSize -ge 1024000000000) { $Spaces = " " * 3 }
            ElseIf ($ObjectSize -ge 102400000000) { $Spaces = " " * 4 }
            ElseIf ($ObjectSize -ge 10240000000) { $Spaces = " " * 5 }
            Else { $Spaces = " " * 6 }
            $ObjectSizeFormat = $Spaces + ("{0:F0}" -f ($Object.Size / 1GB)) + " GB"
            $Spaces = $null
            If ($ObjectFreeSpace -ge 1073741824000) { $Spaces = " " * 4 }
            ElseIf ($ObjectFreeSpace -ge 107374182400) { $Spaces = " " * 5 }
            ElseIf ($ObjectFreeSpace -ge 10737418240) { $Spaces = " " * 6 }
            Else { $Spaces = " " * 7 }
            $ObjectFreeSpaceFormat = $Spaces + ("{0:F2}" -f ($Object.FreeSpace / 1GB)) + " GB"
            If ($ObjectName.Length -gt $global:Column2Length) {
                $global:Column2Length = $ObjectName.Length
            }
            If ($ObjectSizeFormat.Length -gt $global:Column3Length) {
                $global:Column3Length = $ObjectSizeFormat.Length
            }
            If ($ObjectFreeSpaceFormat.Length -gt $global:Column4Length) {
                $global:Column4Length = $ObjectFreeSpaceFormat.Length
            }
            $Properties += @{
                Name            = $ObjectName
                Drive           = $ObjectDrive
                Label           = $ObjectLabel
                UNC             = $ObjectNetwork
                Size            = $ObjectSize
                SizeFormat      = $ObjectSizeFormat
                FreeSpace       = $ObjectFreeSpace
                FreeSpaceFormat = $ObjectFreeSpaceFormat
                Beschikbaar     = $ObjectBeschikbaar
            }
        }
        If ($Type -eq "Emailadressen") {
            $ObjectName = $Object.Name
            $Properties += @{Name = $ObjectName }
        }
        If ($Type -eq "Emaildomein") {
            If (([string]($Object.DomainName)).Length -gt $global:Column2Length) {
                $global:Column2Length = ([string]($Object.DomainName)).Length
            }
            If (([string]($Object.DomainType)).Length -gt $global:Column3Length) {
                $global:Column3Length = ([string]($Object.DomainType)).Length
            }
            $Properties += @{Name = [string]$Object.DomainName; DomainType = [string]$Object.DomainType; Default = $Object.Default }
        }
        If ($Type -like "Genereer*") {
            $ObjectName = $Object.Name
            $ObjectBeschikbaar = $Object.Beschikbaar
            $ObjectLengte = $Object.Lengte
            If ($ObjectName.Length -gt $global:Column2Length) {
                $global:Column2Length = $ObjectName.Length
            }
            If ($ObjectBeschikbaar.Length -gt $global:Column3Length) {
                $global:Column3Length = $ObjectBeschikbaar.Length
            }
            If ($ObjectLengte.Length -gt $global:Column4Length) {
                $global:Column4Length = $ObjectLengte.Length
            }
            $Properties += @{Name = $ObjectName; Beschikbaar = $ObjectBeschikbaar; Lengte = $ObjectLengte }
        }
        If ($Type -eq "Groepen") {
            $Properties += @{Name = $Object.Name }
        }
        If ($Type -eq "Homefolders") {
            $ObjectName = $Object.Name
            $ObjectDN = $Object.DN
            $ObjectHomePath = $Object.HomePath
            $ObjectSize = $Object.Size
            $ObjectSizeFormat = $ObjectSize + " MB"
            $Spaces = $null
            If ($ObjectSizeFormat.Length -ge 8) { $Spaces = " " * 2 }
            ElseIf ($ObjectSizeFormat.Length -ge 7) { $Spaces = " " * 3 }
            ElseIf ($ObjectSizeFormat.Length -ge 6) { $Spaces = " " * 4 }
            ElseIf ($ObjectSizeFormat.Length -ge 5) { $Spaces = " " * 5 }
            ElseIf ($ObjectSizeFormat.Length -ge 4) { $Spaces = " " * 6 }
            Else { $Spaces = " " * 7 }
            $ObjectSizeFormat = $Spaces + $ObjectSizeFormat
            $ObjectTotalSize += $ObjectSize
            If ($ObjectName.Length -gt $global:Column2Length) {
                $global:Column2Length = $ObjectName.Length
            }
            If ($ObjectHomePath.Length -gt $global:Column3Length) {
                $global:Column3Length = $ObjectHomePath.Length
            }
            $Properties += @{
                Name       = $ObjectName
                DN         = $ObjectDN
                HomePath   = $ObjectHomePath
                Size       = $ObjectSize
                SizeFormat = $ObjectSizeFormat
            }
        }
        If ($Type -like "Mailbox*") {
            If ($Type -eq "MailboxHerstel") {
                $ObjectName = $Object.DisplayName
                $ObjectLastLogoffDate = $Object.LastLogoffTime
                [string]$ObjectTotalItemSize = $Object.TotalItemSize
                $ObjectItemCount = $Object.ItemCount
                If ($ObjectLastLogoffDate) {
                    $ObjectLastLogoffDate = Get-Date $ObjectLastLogoffDate.DateTime
                }
                Else {
                    $ObjectLastLogoffDate = Get-Date 01-01-2000
                }
                $global:Column3Length = 20
                $Properties += @{`
                        Identity       = $Object.Identity; `
                        LastLogoffDate = $ObjectLastLogoffDate;
                }
            }
            Else {
                $ObjectName = $Object.DisplayName
                $ObjectDN = $Object.DistinguishedName
                $ObjectForwardingAddress = ($ObjectsMailForward | Where-Object { $_.Name -eq $ObjectName } | Select-Object ForwardingAddress).ForwardingAddress
                $ObjectDeliverToMailboxAndForward = ($ObjectsMailForward | Where-Object { $_.Name -eq $ObjectName } | Select-Object DeliverToMailboxAndForward).DeliverToMailboxAndForward
                $ObjectOffice365Guid = [System.Convert]::ToBase64String(($Object.Guid).ToByteArray())
                $ObjectGUID = ($Object.ExchangeGuid).Guid
                If ($DetectDateItemsServerAndSize) {
                    [array]$ObjectStat = $ObjectsStatistics | Where-Object { $_.MailboxGuid -like "*$ObjectGUID*" }
                    If ($ObjectStat) {
                        $ObjectStatLength = $ObjectStat.Length
                        If (!$ObjectStatLength -OR $ObjectStatLength -eq 1) {
                            $ObjectStatAantal = 0
                        }
                        Else {
                            $ObjectStatAantal = ($ObjectStatLength - 1)
                        }
                        $ObjectLastLogonDate = $ObjectStat[$ObjectStatAantal].LastLogonTime
                        If ($ObjectLastLogonDate) {
                            $ObjectLastLogonDate = Get-Date $ObjectLastLogonDate
                        }
                        Else {
                            $ObjectLastLogonDate = Get-Date 01-01-2000
                        }
                        $ObjectServerName = ($ObjectStat[$ObjectStatAantal].ServerName).ToUpper()
                        [string]$ObjectTotalItemSize = $ObjectStat[$ObjectStatAantal].TotalItemSize
                        [int]$ObjectItemCount = $ObjectStat[$ObjectStatAantal].ItemCount
                    }
                    Else {
                        $ObjectLastLogonDate = Get-Date 01-01-2000
                        $ObjectTotalItemSize = $null
                        $ObjectItemCount = $null
                        $ObjectServerName = $null
                    }
                    If (!$ObjectItemCount -OR $ObjectItemCount -eq 1) {
                        $ObjectItemCount = 0
                    }
                    $Spaces = $null
                    If ($ObjectItemCount -ge 1000000) { $Spaces = " " * 1 }
                    ElseIf ($ObjectItemCount -ge 100000) { $Spaces = " " * 2 }
                    ElseIf ($ObjectItemCount -ge 10000) { $Spaces = " " * 3 }
                    ElseIf ($ObjectItemCount -ge 1000) { $Spaces = " " * 4 }
                    ElseIf ($ObjectItemCount -ge 100) { $Spaces = " " * 5 }
                    ElseIf ($ObjectItemCount -ge 10) { $Spaces = " " * 6 }
                    Else { $Spaces = " " * 7 }
                    $ObjectItemsFormat = $Spaces + $ObjectItemCount
                    If ($ObjectTotalItemSize) {
                        [int]$ObjectSize = "{0:0}" -f ([double]$ObjectTotalItemSize.Split(' ')[2].Replace('(', '').Replace(',', '') / 1MB)
                        $ObjectTotalSize += $ObjectSize
                    }
                    Else {
                        $ObjectSize = 0
                    }
                    $Spaces = $null
                    If ($ObjectSize -ge 10000) { $Spaces = " " * 2 }
                    ElseIf ($ObjectSize -ge 1000) { $Spaces = " " * 3 }
                    ElseIf ($ObjectSize -ge 100) { $Spaces = " " * 4 }
                    ElseIf ($ObjectSize -ge 10) { $Spaces = " " * 5 }
                    Else { $Spaces = " " * 6 }
                    $ObjectSizeFormat = $Spaces + $ObjectSize + " MB"
                }
                Else {
                    $ObjectLastLogonDate = "n.a."
                    $ObjectServerName = "n.a."
                    $ObjectSize = "n.a."
                    $ObjectSizeFormat = "n.a."
                    $ObjectItemCount = "n.a."
                    $ObjectItemsFormat = "n.a."
                }
                If ($DetectPermissions) {
                    $ObjectFullAccess = ($ObjectsFullAccess | Where-Object { $_.Name -eq $ObjectName } | Select-Object FullAccess).FullAccess
                    $ObjectSendAs = ($ObjectsSendAs | Where-Object { $_.Name -eq $ObjectName } | Select-Object SendAs).SendAs
                    $ObjectSendOnBehalf = ($ObjectsSendOnBehalf | Where-Object { $_.Name -eq $ObjectName } | Select-Object SendOnBehalf).SendOnBehalf
                }
                Else {
                    $ObjectFullAccess = $null
                    $ObjectSendAs = $null
                    $ObjectSendOnBehalf = $null
                }
                If ($DetectLanguage) {
                    $ObjectLanguageFolders = ($ObjectsLanguageStatistics | Where-Object { $_.Alias -eq $Object.Alias }).FolderPath
                    If ($ObjectLanguageFolders -like "*Boîte de réception") {
                        $ObjectLanguage = "fr-FR"
                    }
                    If ($ObjectLanguageFolders -like "*Inbox") {
                        $ObjectLanguage = "en-US"
                    }
                    If ($ObjectLanguageFolders -like "*Posteingang") {
                        $ObjectLanguage = "de-DE"
                    }
                    If ($ObjectLanguageFolders -like "*Postvak IN") {
                        $ObjectLanguage = "nl-NL"
                    }
                    $ObjectLanguageRegion = ($ObjectsLanguageStatistics | Where-Object { $_.Alias -eq $Object.Alias }).Language
                    If ($ObjectLanguage -eq "de-DE") {
                        If ($ObjectLanguageRegion -like "de-*") {
                            $ObjectLanguage = $ObjectLanguageRegion
                        }
                        $ObjectLanguageFormat = "Duits"
                    }
                    ElseIf ($ObjectLanguage -eq "en-US") {
                        If ($ObjectLanguageRegion -like "en-*") {
                            $ObjectLanguage = $ObjectLanguageRegion
                        }
                        $ObjectLanguageFormat = "Engels"
                    }
                    ElseIf ($ObjectLanguage -eq "fr-FR") {
                        If ($ObjectLanguageRegion -like "fr-*") {
                            $ObjectLanguage = $ObjectLanguageRegion
                        }
                        $ObjectLanguageFormat = "Frans"
                    }
                    ElseIf ($ObjectLanguage -eq "nl-NL") {
                        If ($ObjectLanguageRegion -like "nl-*") {
                            $ObjectLanguage = $ObjectLanguageRegion
                        }
                        $ObjectLanguageFormat = "Nederlands"
                    }
                    Else {
                        $ObjectLanguageFormat = "Onbekend"
                    }
                }
                Else {
                    $ObjectLanguage = "n.a."
                    $ObjectLanguageFormat = "n.a."
                }
                If ($Office365Exchange) {
                    $ArrayLicense = @()
                    $ObjectLicense = $null
                    $AzureADObject = $AzureADObjects | Where-Object { $_.UserPrincipalName -eq $Object.UserPrincipalName }
                    $ObjectDN = $AzureADObject.DN
                    ForEach ($SkuId in (($AzureADObject.AssignedLicenses).SkuId | Sort-Object)) {
                        $Sku = ($AzureADSubscribedSku | Where-Object { $_.SkuId -eq $SkuId } | Select-Object SkuPartNumber).SkuPartNumber
                        If ($Sku -eq "ENTERPRISEPACK") { $ArrayLicense += "Enterprise E3" }
                        ElseIf ($Sku -eq "ENTERPRISEPREMIUM") { $ArrayLicense += "Enterprise E5" }
                        ElseIf ($Sku -eq "ENTERPRISEPREMIUM_NOPSTNCONF") { $ArrayLicense += "Enterprise E5 (No conf)" }
                        ElseIf ($Sku -eq "EXCHANGESTANDARD") { $ArrayLicense += "Exchange Online" }
                        ElseIf ($Sku -eq "MICROSOFT_BUSINESS_CENTER") { $ArrayLicense += "Microsoft Business Center" }
                        ElseIf ($Sku -eq "O365_BUSINESS_ESSENTIALS") { $ArrayLicense += "Business Essentials" }
                        ElseIf ($Sku -eq "O365_BUSINESS_PREMIUM") { $ArrayLicense += "Business Premium" }
                        ElseIf ($Sku -eq "POWERFLOW_P1") { $ArrayLicense += "PowerApps P1" }
                        ElseIf ($Sku -eq "POWER_BI_PRO") { $ArrayLicense += "Power BI Pro" }
                        ElseIf ($Sku -eq "POWER_BI_STANDARD") { $ArrayLicense += "Power BI" }
                        ElseIf ($Sku -eq "VISIOCLIENT") { $ArrayLicense += "Visio" }
                        ElseIf ($Sku -eq "FLOW_FREE" -OR $Sku -eq "POWERAPPS_VIRAL") { BREAK }
                        Else { $ArrayLicense += $Sku }
                    }
                    [string]$ObjectLicense = ($ArrayLicense | Sort-Object) -join ", "
                    If (!$ObjectLicense) {
                        $ObjectLicense = "n.a."
                    }
                    $Properties += @{`
                            License = $ObjectLicense;
                    }
                }
                If ($ObjectDN) {
                    $Subtree = $null
                    $TempOU = $null
                    $OUs = $null
                    $DC = $null
                    If ($ObjectDN -like "*CN=Users,DC*") {
                        $ObjectDN = $ObjectDN.Replace('CN=Users,DC', 'OU=Users,DC')
                    }
                    ForEach ($Item in ($ObjectDN.replace('\,', '~').split(","))) {
                        Switch -Regex ($Item.TrimStart().Substring(0, 3)) {
                            "CN=" {
                                $CN = '/' + $Item.replace("CN=", "")
                            }
                            "OU=" {
                                $TempOU += , $Item.replace("OU=", "")
                                $TempOU += '\'
                            }
                            "DC=" {
                                $DC += $Item.replace("DC=", "")
                                $DC += '.'
                            }
                        }
                    }
                    If ($TempOU.Count -ge 1) {
                        For ($I = $TempOU.Count; $I -ge 0; $I --) {
                            $OUs += $TempOU[$I]
                        }
                        $Subtree += $OUs.Substring(1)
                    }
                }
                Else {
                    $Subtree = $null
                }
                $Properties += @{
                    Alias                       = $Object.Alias
                    SamAccountName              = $Object.SamAccountName
                    UserPrincipalName           = $Object.UserPrincipalName
                    DistinguishedName           = $Object.DistinguishedName
                    OU                          = $Subtree
                    ServerName                  = $ObjectServerName
                    RecipientTypeDetails        = $Object.RecipientTypeDetails
                    PrimarySmtpAddress          = ($Object.PrimarySmtpAddress.ToString()).ToLower()
                    EmailAddresses              = $Object.EmailAddresses
                    LegacyExchangeDN            = $Object.LegacyExchangeDN
                    Office365Guid               = $ObjectOffice365Guid
                    ForwardingAddress           = $ObjectForwardingAddress
                    DeliverToMailboxAndForward  = $ObjectDeliverToMailboxAndForward
                    Language                    = $ObjectLanguage
                    LanguageFormat              = $ObjectLanguageFormat
                    FullAccess                  = $ObjectFullAccess
                    SendAs                      = $ObjectSendAs
                    SendOnBehalf                = $ObjectSendOnBehalf
                }
                If ($Object.Alias.Length -gt $global:Column3Length) {
                    $global:Column3Length = $Object.Alias.Length
                }
                If ($ObjectLastLogonDate.Length -gt $global:Column4Length) {
                    $global:Column4Length = $ObjectLastLogonDate.Length
                }
            }
            If ($ObjectName.Length -gt $global:Column2Length) {
                $global:Column2Length = $ObjectName.Length
            }
            If ($ObjectItemsFormat.Length -gt $global:Column5Length) {
                $global:Column5Length = $ObjectItemsFormat.Length
            }
            If ($ObjectSizeFormat.Length -gt $global:Column6Length) {
                $global:Column6Length = $ObjectSizeFormat.Length
            }
            If ($ObjectLanguageFormat.Length -gt $global:Column7Length) {
                $global:Column7Length = $ObjectLanguageFormat.Length
            }
            $Properties += @{`
                Name          = $ObjectName
                LastLogonDate = $ObjectLastLogonDate
                Items         = $ObjectItemCount
                ItemsFormat   = $ObjectItemsFormat
                Size          = $ObjectSize
                SizeFormat    = $ObjectSizeFormat
                Database      = $Object.Database
            }
        }
        If ($Type -like "Office365Account*") {
            $ObjectName = $Object.DisplayName
            If ($ObjectName -eq "On-Premises Directory Synchronization Service Account") {
                $ObjectName = "On-Premises DirSync Account"
            }
            $ArrayLicense = @()
            $ObjectLicense = $null
            ForEach ($SkuId in (($Object.AssignedLicenses).SkuId | Sort-Object)) {
                $Sku = ($AzureADSubscribedSku | Where-Object { $_.SkuId -eq $SkuId } | Select-Object SkuPartNumber).SkuPartNumber
                If ($Sku -eq "ENTERPRISEPACK") { $ArrayLicense += "Enterprise E3" }
                ElseIf ($Sku -eq "ENTERPRISEPREMIUM") { $ArrayLicense += "Enterprise E5" }
                ElseIf ($Sku -eq "ENTERPRISEPREMIUM_NOPSTNCONF") { $ArrayLicense += "Enterprise E5 (No conf)" }
                ElseIf ($Sku -eq "EXCHANGESTANDARD") { $ArrayLicense += "Exchange Online" }
                ElseIf ($Sku -eq "MICROSOFT_BUSINESS_CENTER") { $ArrayLicense += "Microsoft Business Center" }
                ElseIf ($Sku -eq "O365_BUSINESS_ESSENTIALS") { $ArrayLicense += "Business Essentials" }
                ElseIf ($Sku -eq "O365_BUSINESS_PREMIUM") { $ArrayLicense += "Business Premium" }
                ElseIf ($Sku -eq "POWERFLOW_P1") { $ArrayLicense += "PowerApps P1" }
                ElseIf ($Sku -eq "POWER_BI_PRO") { $ArrayLicense += "Power BI Pro" }
                ElseIf ($Sku -eq "POWER_BI_STANDARD") { $ArrayLicense += "Power BI" }
                ElseIf ($Sku -eq "VISIOCLIENT") { $ArrayLicense += "Visio" }
                ElseIf ($Sku -eq "FLOW_FREE" -OR $Sku -eq "POWERAPPS_VIRAL") { BREAK }
                Else { $ArrayLicense += $Sku }
            }
            [string]$ObjectLicense = ($ArrayLicense | Sort-Object) -join ", "
            If (!$ObjectLicense) { $ObjectLicense = "n.a." }
            If ($Object.AccountEnabled -eq $false) { $ObjectLicense = "Geblokkeerd" }
            If ($Object.UserType -eq "Guest") {
                $ObjectUPN = $Object.Mail
                $ObjectLicense = "Gast (extern)"
            }
            Else {
                $ObjectUPN = $Object.UserPrincipalName
            }
            If ($Object.LastDirSyncTime) {
                $ObjectSyncStatus = "Sync met AD"
            }
            Else {
                $ObjectSyncStatus = "Office 365"
            }
            If ($ObjectName.Length -gt $global:Column2Length) {
                $global:Column2Length = $ObjectName.Length
            }
            If ($ObjectUPN.Length -gt $global:Column3Length) {
                $global:Column3Length = $ObjectUPN.Length
            }
            If ($ObjectLicense.Length -gt $global:Column4Length) {
                $global:Column4Length = $ObjectLicense.Length
            }
            $Properties += @{
                Name               = $ObjectName
                DisplayName        = $ObjectName
                GivenName          = $Object.GivenName
                Surname            = $Object.Surname
                UserPrincipalName  = $ObjectUPN
                Alias              = $Object.MailNickName
                PrimarySmtpAddress = $Object.Mail
                EmailAddresses     = $Object.ProxyAddresses
                JobTitle           = $Object.JobTitle
                Department         = $Object.Department
                Company            = $Object.CompanyName
                StreetAddress      = $Object.StreetAddress
                City               = $Object.City
                State              = $Object.State
                PostalCode         = $Object.PostalCode
                Country            = $Object.Country
                OfficePhone        = $Object.TelephoneNumber
                MobilePhone        = $Object.Mobile
                AccountEnabled     = $Object.AccountEnabled
                DirSyncEnabled     = $Object.DirSyncEnabled
                SyncStatus         = $ObjectSyncStatus
                License            = $ObjectLicense
                PasswordPolicies   = $Object.PasswordPolicies
            }
        }
        If ($Type -eq "Office365Licenties") {
            $ObjectName = $Object.SkuPartNumber
            [int]$ObjectTotal = ($Object.PrepaidUnits).Enabled
            If ($ObjectTotal -ge 10000000) { $Spaces = " " * 1 }
            ElseIf ($ObjectTotal -ge 1000000) { $Spaces = " " * 2 }
            ElseIf ($ObjectTotal -ge 100000) { $Spaces = " " * 3 }
            ElseIf ($ObjectTotal -ge 10000) { $Spaces = " " * 4 }
            ElseIf ($ObjectTotal -ge 1000) { $Spaces = " " * 5 }
            ElseIf ($ObjectTotal -ge 100) { $Spaces = " " * 6 }
            ElseIf ($ObjectTotal -ge 10) { $Spaces = " " * 7 }
            Else { $Spaces = " " * 8 }
            $ObjectTotalFormat = $Spaces + $ObjectTotal
            [int]$ObjectUsed = $Object.ConsumedUnits
            If ($ObjectUsed -ge 1000000) { $Spaces = " " * 4 }
            ElseIf ($ObjectUsed -ge 100000) { $Spaces = " " * 5 }
            ElseIf ($ObjectUsed -ge 10000) { $Spaces = " " * 6 }
            ElseIf ($ObjectUsed -ge 1000) { $Spaces = " " * 7 }
            ElseIf ($ObjectUsed -ge 100) { $Spaces = " " * 8 }
            ElseIf ($ObjectUsed -ge 10) { $Spaces = " " * 9 }
            Else { $Spaces = " " * 10 }
            $ObjectUsedFormat = $Spaces + $ObjectUsed
            [int]$ObjectFree = $ObjectTotal - $ObjectUsed
            If ($ObjectFree -ge 1000000) { $Spaces = " " * 6 }
            ElseIf ($ObjectFree -ge 100000) { $Spaces = " " * 7 }
            ElseIf ($ObjectFree -ge 10000) { $Spaces = " " * 8 }
            ElseIf ($ObjectFree -ge 1000) { $Spaces = " " * 9 }
            ElseIf ($ObjectFree -ge 100) { $Spaces = " " * 10 }
            ElseIf ($ObjectFree -ge 10) { $Spaces = " " * 11 }
            Else { $Spaces = " " * 12 }
            If ($ObjectFree -lt 0) { $Spaces = " " * 11 }
            $ObjectFreeFormat = $Spaces + $ObjectFree
            If ($ObjectName.Length -gt $global:Column2Length) {
                $global:Column2Length = $ObjectName.Length
            }
            If ($ObjectTotalFormat.Length -gt $global:Column3Length) {
                $global:Column3Length = $ObjectTotalFormat.Length
            }
            If ($ObjectUsedFormat.Length -gt $global:Column4Length) {
                $global:Column4Length = $ObjectUsedFormat.Length
            }
            If ($ObjectFreeFormat.Length -gt $global:Column5Length) {
                $global:Column5Length = $ObjectFreeFormat.Length
            }
            $Properties += @{
                Name        = $ObjectName
                Total       = $ObjectTotal
                TotalFormat = $ObjectTotalFormat
                Used        = $ObjectUsed
                UsedFormat  = $ObjectUsedFormat
                Free        = $ObjectFree
                FreeFormat  = $ObjectFreeFormat
            }
        }
        If ($Type -eq "Rechten") {
            $Properties += @{Name = $Object }
        }
        If ($Type -eq "UPNSuffix") {
            $Properties += @{Name = $Object }
        }
        $Object = New-Object PSObject -Property $Properties
        $global:Array += $Object
        $Counter++
    }
    $global:Array = $global:Array | Sort-Object Name
    $Counter = 0
    $Sorting = "A-Z"
    If ($global:Array.Count) {
        $ArrayTotal = $global:Array.Count
    }
    Else {
        $ArrayTotal = 1
    }
    If ($ArrayTotal -eq 1 -AND ($Type -eq "Emaildomein" -OR $Type -eq "Homefolders" -OR $Type -eq "UPNSuffix")) {
        $SkipMenu = $true
        $global:Array[0].Markeren = "Ja"
    }
    # -----------------------------------------------------------------------------------
    # Kolom: De kolom namen worden hier samengesteld en voorzien van de juiste spatiering
    # -----------------------------------------------------------------------------------
    Do {
        If ($Type -notlike "Mailbox*" -OR ($Type -like "Mailbox*" -AND $DetectDateItemsServerAndSize)) {
            $global:SelectieSize = 0
            ForEach ($Size in ($global:Array | Where-Object { $_.Markeren -eq "Ja" })) {
                $global:SelectieSize += $Size.Size
            }
        }
        If ($SkipMenu -eq $false) {
            $Column1 = $null
            $Column2 = $null
            $Column3 = $null
            $Column4 = $null
            $Column5 = $null
            $Column6 = $null
            $Column7 = $null
            $Column8 = $null
            Script-Module-SetHeaders -CurrentTask $CurrentTask -Name $FunctionName
            if ($OUFormat) {
                Write-Host -NoNewLine "  > Entered OU filter: "; Write-Host ($global:OUFormat + " (" + $global:Array.Count + ")") -ForegroundColor Yellow
            }
            If ($AzureADConnected) {
                Write-Host -NoNewLine "   - Office 365 Account: "; Write-Host $global:AzureADConnect.Account -ForegroundColor Yellow
                Write-Host -NoNewLine "   - Office 365 Domain: "; Write-Host $global:AzureADConnect.TenantDomain -ForegroundColor Yellow
            }
            If ($Functie -eq "Markeren") {
                Write-Host; Write-Host -NoNewline "  Multi-select the $ObjectType using the numbers, these will turn to "; Write-Host -NoNewLine "green" -ForegroundColor Green; Write-Host " to indicate the selection.";
                Write-Host -NoNewLine "  You can deselect by using the same numbers or use "; Write-Host -NoNewLine "A" -ForegroundColor Yellow; Write-Host -NoNewLine " to (de)select all $ObjectType."; Write-Host
            }
            If ($Type -eq "Emailadressen" -AND !$Objects) {
                Write-Host
                Write-Host "  There are no mailaddresses detected" -ForegroundColor Gray
                Write-Host
            }
            Else {
                $Column1 = (" " * 5)
                $Column2 = "[N]aam:"
                $ColumnSpaces1 = (" " * 6)
                If ($Type -like "Account*") {
                    $Column3 = "[G]ebruikersnaam:"
                    If ($Type -eq "AccountHerstel" -OR $Type -eq "AccountRecycleBin") {
                        $Column4 = "[D]atum verwijderd: "
                    }
                    If ($AzureADConnected) {
                        $Column4 = "[L]icentie:"
                    }
                    Else {
                        $Column4 = "[L]aatst aangemeld: "
                    }
                    $Column5 = "[O]rganizational Unit:"
                }
                If ($Type -eq "Bestand" -AND $Filter -eq "PST") {
                    If ($Functie -eq "Markeren") {
                        $Column3 = "[B]ijbehorende mailbox:"
                        $Column4 = "[G]rootte:"
                    }
                    Else {
                        $Column3 = "[G]rootte:"
                    }
                }
                If ($Type -eq "Drive") {
                    $Column3 = "[G]rootte:"
                    $Column4 = "[B]eschikbaar:"
                }
                If ($Type -eq "Emaildomein") {
                    $Column3 = "[T]ype:"
                    $Column4 = "[P]rimair:"
                }
                If ($Type -like "Genereer*") {
                    $Column3 = "[B]eschikbaarheid:"
                }
                If ($Type -eq "Homefolders") {
                    $Column3 = "[B]ijbehorende homefolder:"
                    $Column4 = "[G]rootte:"
                }
                If ($Type -like "Mailbox*") {
                    If ($Type -eq "MailboxHerstel") {
                        $Column3 = "[A]fmelding:"
                        $Column4 = "[I]tems:"
                        $Column5 = "[G]rootte:"
                    }
                    Else {
                        $Column3 = "[M]ailbox alias:"
                        $Column4 = "[L]aatst aangemeld: "
                        $Column5 = "[I]tems:"
                        $Column6 = "[G]rootte:"
                        If ($FunctionName -like "*Overview*") {
                            $Column7 = "Taal/[R]egio:"
                            $Column8 = "[P]rimair E-mailadres:"
                        }
                        Else {
                            $Column7 = "[P]rimair E-mailadres:"
                        }
                    }
                }
                If ($Type -like "Office365Account*") {
                    $Column3 = "[G]ebruikersnaam:"
                    $Column4 = "[L]icentie:"
                    $Column5 = "[S]ynchronisatie:"
                }
                If ($Type -eq "Office365Licenties") {
                    $Column3 = "[A]antal:"
                    $Column4 = "[G]ebruikt:"
                    $Column5 = "[O]ngebruikt:"
                }
                $ColumnOutput = {
                    If ($Column1) { If ($Column1.Length -gt $global:Column1Length) { $global:Column1Length = $Column1.Length }; $ColumnSpaces1 = (" " * ($global:Column1Length - $Column1.Length + 2)) } Else { $ColumnSpaces1 = $null }
                    If ($Column2) { If ($Column2.Length -gt $global:Column2Length) { $global:Column2Length = $Column2.Length }; $ColumnSpaces2 = (" " * ($global:Column2Length - $Column2.Length + 2)) } Else { $ColumnSpaces2 = $null }
                    If ($Column3) { If ($Column3.Length -gt $global:Column3Length) { $global:Column3Length = $Column3.Length }; $ColumnSpaces3 = (" " * ($global:Column3Length - $Column3.Length + 2)) } Else { $ColumnSpaces3 = $null }
                    If ($Column4) { If ($Column4.Length -gt $global:Column4Length) { $global:Column4Length = $Column4.Length }; $ColumnSpaces4 = (" " * ($global:Column4Length - $Column4.Length + 2)) } Else { $ColumnSpaces4 = $null }
                    If ($Column5) { If ($Column5.Length -gt $global:Column5Length) { $global:Column5Length = $Column5.Length }; $ColumnSpaces5 = (" " * ($global:Column5Length - $Column5.Length + 2)) } Else { $ColumnSpaces5 = $null }
                    If ($Column6) { If ($Column6.Length -gt $global:Column6Length) { $global:Column6Length = $Column6.Length }; $ColumnSpaces6 = (" " * ($global:Column6Length - $Column6.Length + 2)) } Else { $ColumnSpaces6 = $null }
                    If ($Column7) { If ($Column7.Length -gt $global:Column7Length) { $global:Column7Length = $Column7.Length }; $ColumnSpaces7 = (" " * ($global:Column7Length - $Column7.Length + 2)) } Else { $ColumnSpaces7 = $null }
                    $global:ColumnRow = $Column1 + $ColumnSpaces1 + $Column2 + $ColumnSpaces2 + $Column3 + $ColumnSpaces3 + $Column4 + $ColumnSpaces4 + $Column5 + $ColumnSpaces5 + $Column6 + $ColumnSpaces6 + $Column7 + $ColumnSpaces7 + $Column8
                }
                $ColumnOutput.Invoke()
                If ($global:ColumnRow.Length -ge $global:MaxRowLength) {
                    $global:ColumnRow = $global:ColumnRow.Substring(0, ($global:MaxRowLength - 5)) + "..."
                }
                Write-Host
                Write-Host $global:ColumnRow
                # ---------------------------------------------------------------------------------------------------------------------------------
                # Kolom: Elke rij wordt hier doorlopen, samengesteld en voorzien van de juiste spatiering en vervolgens met de juiste kleur getoond
                # ---------------------------------------------------------------------------------------------------------------------------------
                $ColumnCounter = 0
                Do {
                    $Nr = $Counter + 1
                    $Column1 = $null
                    $Column2 = $null
                    $Column3 = $null
                    $Column4 = $null
                    $Column5 = $null
                    $Column6 = $null
                    $Column7 = $null
                    $Column8 = $null
                    $Markeren = $global:Array[$Counter].Markeren
                    If ($Nr -ge 100) { $Column1 = (" " * 1) + $Nr + "." } ElseIf ($Nr -ge 10) { $Column1 = (" " * 2) + $Nr + "." } Else { $Column1 = (" " * 3) + $Nr + "." }
                    $Column2 = $global:Array[$Counter].Name
                    If ($Type -like "Account*") {
                        $Column3 = $global:Array[$Counter].UserPrincipalName
                        If ($AzureADConnected) {
                            $Column4 = $global:Array[$Counter].License
                        }
                        Else {
                            If ($global:Array[$Counter].LastLogonDate -eq (Get-Date 1-1-2000)) {
                                $Column4 = "Nog nooit aangemeld"
                            }
                            Else {
                                $Column4 = [string](Get-Date $global:Array[$Counter].LastLogonDate -format "dd MMM yyyy HH:mm:ss")
                            }
                        }
                        $DN = $global:Array[$Counter].DN
                    }
                    If ($Type -eq "Bestand") {
                        If ($Filter -eq "PST") {
                            If ($Functie -eq "Markeren") {
                                $Column3 = $global:Array[$Counter].MailboxName
                                $Column4 = $global:Array[$Counter].SizeFormat
                                If ($ObjectTotalSize -ge 1000000000) {
                                    $TotalSizeFormat = ("{0:F2}" -f ((($ObjectTotalSize / 1024) / 1024) / 1024))
                                    $TotalSizeFormat = $TotalSizeFormat.Replace('.', '')
                                    $TotalSizeFormat += " GB"
                                }
                                Else {
                                    $TotalSizeFormat = ("{0:F2}" -f (($ObjectTotalSize / 1024) / 1024))
                                    $TotalSizeFormat = $TotalSizeFormat.Replace('.', '')
                                    $TotalSizeFormat += " MB"
                                }
                            }
                            Else {
                                $Column3 = $global:Array[$Counter].SizeFormat
                            }
                        }
                    }
                    If ($Type -eq "Drive") {
                        $Available = $global:Array[$Counter].Beschikbaar 
                        $Column3 = $global:Array[$Counter].SizeFormat
                        $Column4 = $global:Array[$Counter].FreeSpaceFormat
                    }
                    If ($Type -eq "Emaildomein") {
                        $Column3 = $global:Array[$Counter].DomainType
                        $Column4 = $global:Array[$Counter].Default
                    }
                    If ($Type -like "Genereer*") {
                        $Available = $global:Array[$Counter].Beschikbaar
                        If ($Available -eq "Ja") {
                            $Column3 = "Niet in gebruik"
                        }
                        ElseIf ($Available -eq "Nee") {
                            $Column3 = "In gebruik"
                        }
                        ElseIf ($Available -like "Nee,*") {
                            $Column3 = $Available
                        }
                        Else {
                            $Column3 = "n.a."
                        }
                    }
                    If ($Type -eq "Homefolders") {
                        $Column3 = $global:Array[$Counter].HomePath
                        $Column4 = $global:Array[$Counter].SizeFormat
                        If ($ObjectTotalSize -ge 1024) {
                            [string]$TotalSizeFormat = ("{0:F2}" -f ($ObjectTotalSize / 1024))
                            $TotalSizeFormat = $TotalSizeFormat.Replace('.', '')
                            $TotalSizeFormat += " GB"
                        }
                        Else {
                            [string]$TotalSizeFormat = $ObjectTotalSize
                            $TotalSizeFormat += " MB"
                        }
                    }
                    If ($Type -like "Mailbox*") {
                        If ($Type -eq "MailboxHerstel") {
                            If ($global:Array[$Counter].LastLogoffDate -eq (Get-Date 1-1-2000)) {
                                $Column3 = "Nog nooit aangemeld "
                            }
                            Else {
                                $Column3 = [string](Get-Date $global:Array[$Counter].LastLogoffDate -format "dd MMM yyyy HH:mm:ss")
                            }
                            $Column4 = $global:Array[$Counter].ItemsFormat
                            $Column5 = $global:Array[$Counter].SizeFormat
                        }
                        Else {
                            $Column3 = $global:Array[$Counter].Alias
                            If ($global:Array[$Counter].LastLogonDate -eq (Get-Date 1-1-2000)) {
                                $Column4 = "Nog nooit aangemeld "
                            }
                            Else {
                                If ($DetectDateItemsServerAndSize) {
                                    $Column4 = [string](Get-Date $global:Array[$Counter].LastLogonDate -format "dd MMM yyyy HH:mm:ss")
                                }
                                Else {
                                    $Column4 = $global:Array[$Counter].LastLogonDate
                                }
                            }
                            $Column5 = $global:Array[$Counter].ItemsFormat
                            $Column6 = $global:Array[$Counter].SizeFormat
                            If ($FunctionName -like "*Overview*") {
                                $Column7 = $global:Array[$Counter].LanguageFormat
                                $Column8 = $global:Array[$Counter].PrimarySmtpAddress
                            }
                            Else {
                                $Column7 = $global:Array[$Counter].PrimarySmtpAddress
                            }
                        }
                        If ($DetectDateItemsServerAndSize) {
                            If ($ObjectTotalSize -ge 1024) {
                                $TotalSizeFormat = ("{0:F2}" -f ($ObjectTotalSize / 1024)) + " GB"
                            }
                            Else {
                                $TotalSizeFormat = [string]$ObjectTotalSize + " MB"
                            }
                        }
                        Else {
                            $TotalSizeFormat = "n.a."
                        }
                    }
                    If ($Type -like "Office365Account*") {
                        $Column3 = $global:Array[$Counter].UserPrincipalName
                        $Column4 = $global:Array[$Counter].License
                        $Column5 = $global:Array[$Counter].SyncStatus
                    }
                    If ($Type -eq "Office365Licenties") {
                        $Column3 = $global:Array[$Counter].TotalFormat
                        $Column4 = $global:Array[$Counter].UsedFormat
                        $Column5 = $global:Array[$Counter].FreeFormat
                    }
                    $ColumnOutput.Invoke()
                    If ($Type -like "Account*") {
                        If ($Type -ne "AccountHerstel" -AND $Type -ne "AccountRecycleBin") {
                            If (($DN.Length + $ColumnRij.Length) -gt ($global:MaxRowLength - 3)) {
                                $LengteVerschil = $DN.Length - (($global:MaxRowLength - 3) - $ColumnRij.Length)
                                $global:ColumnRow += "..." + $DN.Substring($LengteVerschil + 3)
                            }
                            Else {
                                $global:ColumnRow += $DN
                            }
                        }
                        Else {
                            $global:ColumnRow += $DN
                        }
                    }
                    If ($global:ColumnRow.Length -ge $global:MaxRowLength) {
                        $global:ColumnRow = $global:ColumnRow.Substring(0, ($global:MaxRowLength - 5)) + "..."
                    }
                    If ($Markeren -eq "Ja") {
                        Write-Host $global:ColumnRow -ForegroundColor Green
                    }
                    If (!$Markeren -OR $Markeren -eq "Nee") {
                        If ($Available -eq "Ja") {
                            If ($Type -eq "Drive") {
                                Write-Host $global:ColumnRow -ForegroundColor Yellow
                            }
                            Else {
                                Write-Host $global:ColumnRow -ForegroundColor Green
                            }
                        }
                        ElseIf ($Available -like "Nee*") {
                            Write-Host $global:ColumnRow -ForegroundColor Red
                        }
                        Else {
                            Write-Host $global:ColumnRow -ForegroundColor Yellow
                        }
                    }
                    $Counter++
                    $ColumnCounter++
                } Until (($ColumnCounter -eq $global:MaxRowCount) -OR ($Counter -eq $ArrayTotal))
                Write-Host
                # -----------------------
                # Kolom: Keuze pagina('s)
                # -----------------------
                If (!$InputSelection -AND $ColumnCounter -eq $global:MaxRowCount -AND $Counter -lt $ArrayTotal -AND $Counter -le $global:MaxRowCount) {
                    Write-Host "  >> [V]olgende Pagina"
                    Write-Host
                    $Segment = "Eerste"
                }
                If (!$InputSelection -AND $ColumnCounter -eq $global:MaxRowCount -AND $Counter -lt $ArrayTotal -AND $Counter -gt $global:MaxRowCount) {
                    Write-Host "  << [T]erug naar vorige pagina  |  [V]olgende Pagina >>"
                    Write-Host
                    $Segment = "Midden"
                }
                If (!$InputSelection -AND $ColumnCounter -le $global:MaxRowCount -AND $Counter -eq $ArrayTotal -AND $Counter -gt $global:MaxRowCount) {
                    Write-Host "  << [T]erug naar vorige pagina"
                    Write-Host
                    $Segment = "Eind"
                }
                If (!$InputSelection -AND $Counter -le $global:MaxRowCount -AND $Counter -eq $ArrayTotal) {
                    $Segment = "Alles"
                }
            }
        }
        # -----------------------
        # Kolom: Selectie teksten
        # -----------------------
        Do {
            If ($Functie -eq "Markeren") {
                If ($Type -eq "Bestand") {
                    If ($global:SelectieSize.ToString().Length -ge 10) {
                        $global:SelectieSizeFormat = ("{0:F2}" -f ((($global:SelectieSize / 1024) / 1024) / 1024))
                        $global:SelectieSizeFormat = $global:SelectieSizeFormat.Replace('.', '')
                        $global:SelectieSizeFormat += " GB"
                    }
                    Else {
                        If ($global:SelectieSize -ne 0) {
                            $global:SelectieSizeFormat = ("{0:F2}" -f (($global:SelectieSize / 1024) / 1024))
                            $global:SelectieSizeFormat = $global:SelectieSizeFormat.Replace('.', '')
                            $global:SelectieSizeFormat += " MB"
                        }
                        Else {
                            $global:SelectieSizeFormat = [string]$global:SelectieSize + " MB"
                        }
                    }
                } 
                If ($Type -like "Homefolders" -OR ($Type -like "Mailbox*" -AND $DetectDateItemsServerAndSize)) {
                    If ($global:SelectieSize -ge 1024) {
                        $global:SelectieSizeFormat = ("{0:F2}" -f ($global:SelectieSize / 1024)) + " GB"
                    }
                    Else {
                        $global:SelectieSizeFormat = [string]$global:SelectieSize + " MB"
                    }
                }
            }
            If ($SkipMenu -eq $true) {
                If ($Functie -eq "Markeren") {
                    $InputChoice = "J"
                }
                If ($Functie -eq "Selecteren") {
                    [int]$InputChoice = 1
                }
            }
            Else {
                If ($Functie -eq "Markeren") {
                    If ($global:SelectieSizeFormat) {
                        Write-Host "  De selectie betreft:" $global:SelectieSizeFormat ("(totaal: " + $TotalSizeFormat + ")")
                    }
                    Write-Host -NoNewline "  When you're ready you can confirm with "; Write-Host -NoNewLine "J" -ForegroundColor Yellow; Write-Host -NoNewLine "."
                    If ($Type -eq "AccountHerstel") {
                        Write-Host -NoNewLine " Use "; Write-Host -NoNewLine "C" -ForegroundColor Yellow; Write-Host -NoNewLine " to use a CSV to recover an archived domainaccount."
                        Write-Host; Write-Host -NoNewLine " "
                    }
                    If ($Type -eq "Groepen") {
                        Write-Host -NoNewLine " Use "; Write-Host -NoNewLine "G" -ForegroundColor Yellow; Write-Host -NoNewLine " to ignore adding groups."
                        Write-Host; Write-Host -NoNewLine " "
                    }
                    Write-Host -NoNewLine " Use "; Write-Host -NoNewLine "X" -ForegroundColor Yellow; Write-Host -NoNewline " to return to the previous menu: "
                }
                ElseIf ($Type -eq "Credentials" -AND $Functie -eq "Selecteren") {
                    Write-Host -NoNewline "  Maak uw keuze uit de bovenstaande selectie of druk op "; Write-Host -NoNewLine "A" -ForegroundColor Yellow; Write-Host " om een nieuwe beheeraccount credential aan te maken."
                    Write-Host -NoNewLine "  Druk eventueel op "; Write-Host -NoNewLine "X" -ForegroundColor Yellow; Write-Host -NoNewline " om terug te keren naar het hoofdmenu: "
                }
                ElseIf ($Type -eq "Drive" -AND $Functie -eq "Selecteren" -AND $FunctionName -like "*Archive*Mailbox*") {
                    Write-Host -NoNewline "  Maak uw keuze uit de bovenstaande selectie of druk op "; Write-Host -NoNewLine "Z" -ForegroundColor Yellow; Write-Host -NoNewline " om door te gaan "; Write-Host -NoNewline "ZONDER" -ForegroundColor Red; Write-Host " export (opletten!)."
                    Write-Host -NoNewLine "  Druk eventueel op "; Write-Host -NoNewLine "X" -ForegroundColor Yellow; Write-Host -NoNewline " om terug te keren naar het hoofdmenu: "
                }
                ElseIf ($Type -eq "Emailadressen" -AND $Functie -eq "Selecteren") {
                    Write-Host -NoNewLine "  Selecteer een e-mailadres om deze primair in te stellen of te verwijderen, of druk op "; Write-Host -NoNewLine "E" -ForegroundColor Yellow; Write-Host " om een extra e-mailadres toe te voegen."; Write-Host -NoNewLine "  Druk eventueel op "; Write-Host -NoNewLine "X" -ForegroundColor Yellow; Write-Host -NoNewline " om te annuleren: "
                }
                ElseIf ($Type -like "Genereer*" -AND $Functie -eq "Selecteren") {
                    If ($global:Array.Beschikbaar -contains "Ja") {
                        Write-Host -NoNewLine "  Maak uw keuze uit bovenstaande selectie of gebruik "; Write-Host -NoNewLine "H" -ForegroundColor Yellow; Write-Host -NoNewLine " voor handmatig. Druk eventueel op "; Write-Host -NoNewLine "X" -ForegroundColor Yellow; Write-Host -NoNewline " om te annuleren: "
                    }
                    Else {
                        Write-Host -NoNewLine "  LET OP:" -ForegroundColor Red; Write-Host -NoNewLine " Er is geen keuze mogelijk uit bovenstaande selectie, druk nu op "; Write-Host -NoNewLine "H" -ForegroundColor Yellow; Write-Host -NoNewLine " voor handmatig of "; Write-Host -NoNewLine "X" -ForegroundColor Yellow; Write-Host -NoNewline " om te annuleren: "
                    }
                }
                ElseIf ($Functie -eq "Overzicht") {
                    If ($Type -like "Mailbox*") {
                        Write-Host -NoNewLine "  Totale grootte betreft:" $TotalSizeFormat; Write-Host
                    }
                    Write-Host -NoNewline "  Druk op "; Write-Host -NoNewLine "E" -ForegroundColor Yellow; Write-Host -NoNewline " om dit overzicht te exporteren naar CSV/Excel of druk op "; Write-Host -NoNewLine "X" -ForegroundColor Yellow; Write-Host -NoNewline " om te annuleren: "
                }
                Else {
                    Write-Host -NoNewline "  Maak uw keuze uit de bovenstaande selectie of druk op "; Write-Host -NoNewLine "X" -ForegroundColor Yellow; Write-Host -NoNewline " om te annuleren: "
                }
                [string]$InputChoice = Read-Host
            }
            If (($Functie -eq "Markeren" -OR $Functie -eq "Selecteren") -AND $InputChoice -as [int]) {
                If ([int]$InputChoice -eq "0" -OR [int]$InputChoice -gt $ArrayTotal) {
                    Write-Host "  Gelieve de cijfers uit bovenstaande selectie te gebruiken (of druk op X)" -ForegroundColor Red; Write-Host
                }
                Else {
                    $InputKey = $InputChoice
                    If ($Functie -eq "Markeren") {
                        If ($global:Array[$InputChoice - 1].Markeren -eq "Nee") {
                            $global:Array[$InputChoice - 1].Markeren = "Ja"
                        }
                        Else {
                            $global:Array[$InputChoice - 1].Markeren = "Nee"
                        }
                    }
                    If ($Functie -eq "Selecteren") {
                        If ($global:Array[$InputChoice - 1].Beschikbaar -like "Nee*") {
                            If ($Type -eq "Drive") {
                                Do {
                                    Write-Host -NoNewLine "  LET OP: De selectie is totaal groter dan de aanwezige beschikbare ruimte. Wenst u verder te gaan? (J/N): " -ForegroundColor Red
                                    [string]$InputChoiceWarning = Read-Host
                                    $InputKeyWarning = @("J", "N") -contains $InputChoiceWarning
                                    If (!$InputKeyWarning) {
                                        Write-Host "  Please use the letters above as input" -ForegroundColor Red
                                    }
                                } Until ($InputKeyWarning)
                                Switch ($InputChoiceWarning) {
                                    "J" {
                                        $InputSelection = "Ja"
                                    }
                                }
                            }
                            Else {
                                Write-Host "  Deze is niet meer beschikbaar, kies een andere of gebruik H voor handmatig" -ForegroundColor Red; Write-Host
                                $Pause.Invoke()
                            }
                        }
                        Else {
                            If ($Type -eq "Drive" -AND $Filter -ne "*") {
                                If (!(Get-ChildItem ($global:Array[$InputChoice - 1].Drive + "\") -Force -ErrorAction SilentlyContinue | Where-Object { $_.Name -like "*$Filter*" })) {
                                    Write-Host "  Er zijn geen $Filter bestanden gevonden! Zorg ervoor dat deze in de root van de schijf staat en selecteer opnieuw" -ForegroundColor Red
                                    Write-Host
                                    $Pause.Invoke()
                                }
                                Else {
                                    $InputSelection = "Ja"
                                }
                            }
                            Else {
                                $InputSelection = "Ja"
                            }
                        }
                    }
                    If ($Segment -eq "Alles") {
                        $Counter = 0
                    }
                    Else {
                        $Counter -= $ColumnCounter
                    }
                }
            }
            Else {
                If ($Segment -eq "Eerste" -AND $InputChoice -eq "T") {
                    $Counter -= $global:MaxRowCount
                    BREAK
                }
                If (($Segment -eq "Eerste" -OR $Segment -eq "Midden") -AND $InputChoice -eq "V") {
                    BREAK
                }
                If (($Segment -eq "Midden" -OR $Segment -eq "Eind") -AND $InputChoice -eq "T") {
                    $Counter -= ($ColumnCounter + $global:MaxRowCount)
                    BREAK
                }
                If ($Segment -eq "Eind" -AND $InputChoice -eq "V") {
                    $Counter -= $ColumnCounter
                    BREAK
                }
                If (($Type -like "*Bulk" -AND $Functie -eq "Controleren") -AND $InputChoice -eq "N") {
                    Write-Host
                    Write-Host "U keert nu terug naar het vorige menu"
                    $Pause.Invoke()
                    If ($FunctionName -like "*ActiveDirectory*") { Get-Variable -Exclude $global:StartupVariables | Remove-Variable -ErrorAction SilentlyContinue; Invoke-Expression -Command $global:MenuNameCategory }
                    If ($FunctionName -like "*Exchange*") { Get-Variable -Exclude $global:StartupVariables | Remove-Variable -ErrorAction SilentlyContinue; Invoke-Expression -Command $global:MenuNameCategory }
                    If ($FunctionName -like "*Azure*") { Get-Variable -Exclude $global:StartupVariables | Remove-Variable -ErrorAction SilentlyContinue; Invoke-Expression -Command $global:MenuNameCategory }
                }
                If ($Functie -eq "Markeren") {
                    If ($InputChoice -eq "A") {
                        ForEach ($Item in $global:Array) {
                            $Item.Markeren = "Ja"
                        }
                        If ($Segment -eq "Alles") {
                            $Counter = 0
                        }
                        Else {
                            $Counter -= $ColumnCounter
                        }
                        $SortObjectFormat = $null
                        $InputKey = $InputChoice
                    }
                    If ($InputChoice -eq "J") {
                        If ($global:Array | Where-Object { $_.Markeren -eq "Ja" }) {
                            $InputKey = $InputChoice
                            $InputSelection = "Ja"
                        }
                        Else {
                            $InputKey = $null
                            Write-Host "  Gelieve eerst een regel te markeren alvorens verder te gaan" -ForegroundColor Red
                        }
                    }
                    If ($Type -eq "AccountHerstel") {
                        If ($InputChoice -eq "C") {
                            $InputKey = $InputChoice
                            $InputSelection = "CSV"
                        }
                    }
                }
                If ($InputChoice -eq "E") {
                    If ($FunctionName -like "*ActiveDirectory*Overview*") {
                        If (!$AzureADConnected) {
                            $ExportMigratie = $null
                            Do {
                                Write-Host -NoNewLine "  Druk op "; Write-Host -NoNewLine "1" -ForegroundColor Yellow; Write-Host -NoNewLine " voor een klantmigratie sheet (Nederlandstalig) of "
                                Write-Host -NoNewLine "druk op "; Write-Host -NoNewLine "2" -ForegroundColor Yellow; Write-Host -NoNewLine " voor een volledig overzicht naar CSV: "
                                [string]$InputChoice = Read-Host
                                $InputKey = @("1", "2") -contains $InputChoice
                                If (!$InputKey) {
                                    Write-Host "Gelieve 1 of 2 aan te geven" -ForegroundColor Red
                                }
                            } Until ($InputKey)
                            Switch ($InputChoice) {
                                "1" {
                                    $ExportMigratie = $true
                                }
                            }
                        }
                        Identity-ActiveDirectory-Overview-Accounts-2-Export
                    }
                    If ($Type -eq "Emailadressen") {
                        $InputSelection = "Toevoegen"
                        $InputKey = $InputChoice
                    }
                    If ($FunctionName -like "*Exchange*Overview*Domain*") { Messaging-Exchange-Overview-Domains-2-Export }
                    If ($FunctionName -like "*Exchange*Overview*Mailbox*") { Messaging-Exchange-Overview-Mailboxes-2-Export }
                    If ($FunctionName -like "*Azure*Overview*") { Identity-AzureAD-Overview-Accounts-2-Export }
                }
                ElseIf ($InputChoice -eq "X") {
                    $InputSelection = "Exit"
                    $InputKey = $InputChoice
                }
                Else {
                    If ($Type -eq "Credentials" -AND $InputChoice -eq "A") {
                        $InputSelection = "Aanmaken"
                        $InputKey = $InputChoice
                    }
                    If ($Type -eq "Groepen" -AND $InputChoice -eq "G") {
                        $InputSelection = "Geen"
                        $InputKey = $InputChoice
                    }
                    If ($Type -like "Genereer*" -AND $InputChoice -eq "H") {
                        $InputSelection = "Handmatig"
                        $InputKey = $InputChoice
                    }
                    If ($Type -eq "Drive" -AND $InputChoice -eq "Z") {
                        $InputSelection = "ZonderExport"
                        $InputKey = $InputChoice
                    }
                    If ($ArrayTotal -ge 2) {
                        If ($Type -like "Account*") {
                            If ($InputChoice -eq "G") {
                                $SortObject = "UserPrincipalName"
                                $SortObjectFormat = "Gebruikersnaam"
                                $InputKey = $InputChoice
                            }
                            If ((($Type -eq "AccountHerstel" -OR $Type -eq "AccountRecycleBin") -AND $InputChoice -eq "D") -OR (($Type -ne "AccountHerstel" -AND $Type -ne "AccountRecycleBin") -AND $InputChoice -eq "L")) {
                                If ($AzureADConnected) {
                                    If ($InputChoice -eq "L") {
                                        $SortObject = "License"
                                        $SortObjectFormat = "Licentie"
                                    }
                                }
                                Else {
                                    $SortObject = "LastLogonDate"
                                    $SortObject = [DateTime]::ParseExact($_.LastLogonDate, "d-M-yyyy HH:mm:ss", $null)
                                    If ($Type -eq "AccountHerstel") {
                                        $SortObjectFormat = "Verwijderd op"
                                    }
                                    Else {
                                        $SortObjectFormat = "Laatst aangemeld"
                                    }
                                }
                                $InputKey = $InputChoice
                            }
                            If ($InputChoice -eq "O") {
                                $SortObject = "DN"
                                $SortObjectFormat = "Organizational Unit"
                                $InputKey = $InputChoice
                            }
                        }
                        If ($Type -eq "Bestand" -AND $Filter -eq "PST") {
                            If ($InputChoice -eq "B") {
                                $SortObject = "FilePath"
                                $SortObjectFormat = "$Filter-bestand"
                                $InputKey = $InputChoice
                            }
                            If ($InputChoice -eq "G") {
                                $SortObject = "Size"
                                $SortObjectFormat = "Grootte"
                                $InputKey = $InputChoice
                            }
                        }
                        If ($Type -eq "Emaildomein") {
                            $SortObject = "Default"
                            $SortObjectFormat = "Primair"
                            $InputKey = $InputChoice
                        }
                        If ($Type -like "Mailbox*") {
                            If ($InputChoice -eq "M") {
                                $SortObject = "Alias"
                                $SortObjectFormat = "Alias"
                                $InputKey = $InputChoice
                            }
                            If (($Type -eq "MailboxHerstel" -AND $InputChoice -eq "A") -OR ($Type -ne "MailboxHerstel" -AND $InputChoice -eq "L")) {
                                If ($Type -eq "MailboxHerstel") {
                                    $SortObject = "LastLogoffDate"
                                    $SortObject = [DateTime]::ParseExact($_.LastLogoffDate, "d-M-yyyy HH:mm:ss", $null)
                                    $SortObjectFormat = "Afmelding"
                                }
                                Else {
                                    $SortObject = "LastLogonDate"
                                    $SortObject = [DateTime]::ParseExact($_.LastLogonDate, "d-M-yyyy HH:mm:ss", $null)
                                    $SortObjectFormat = "Laatst aangemeld"
                                }
                                $InputKey = $InputChoice
                            }
                            If ($InputChoice -eq "I") {
                                $SortObject = "Items"
                                $SortObjectFormat = "Items"
                                $InputKey = $InputChoice
                            }
                            If ($InputChoice -eq "G") {
                                $SortObject = "Size"
                                $SortObjectFormat = "Grootte"
                                $InputKey = $InputChoice
                            }
                            If ($InputChoice -eq "P") {
                                $SortObject = "PrimarySmtpAddress"
                                $SortObjectFormat = "Email"
                                $InputKey = $InputChoice
                            }
                            If ($InputChoice -eq "R") {
                                $SortObject = "Language"
                                $SortObjectFormat = "Taal"
                                $InputKey = $InputChoice
                            }
                        }
                        If ($Type -like "Office365Account*") {
                            If ($InputChoice -eq "G") {
                                $SortObject = "UserPrincipalName"
                                $SortObjectFormat = "Gebruikersnaam"
                                $InputKey = $InputChoice
                            }
                            If ($InputChoice -eq "L") {
                                $SortObject = "License"
                                $SortObjectFormat = "Licentie"
                                $InputKey = $InputChoice
                            }
                            If ($InputChoice -eq "S") {
                                $SortObject = "SyncStatus"
                                $SortObjectFormat = "Synchronisatie status"
                                $InputKey = $InputChoice
                            }
                        }
                        If ($Type -eq "Office365Licenties") {
                            If ($InputChoice -eq "A") {
                                $SortObject = "Total"
                                $SortObjectFormat = "Aantal"
                                $InputKey = $InputChoice
                            }
                            If ($InputChoice -eq "G") {
                                $SortObject = "Used"
                                $SortObjectFormat = "Gebruikt"
                                $InputKey = $InputChoice
                            }
                            If ($InputChoice -eq "O") {
                                $SortObject = "Free"
                                $SortObjectFormat = "Ongebruikt"
                                $InputKey = $InputChoice
                            }
                        }
                        If ($InputChoice -eq "N") {
                            $SortObject = "Name"
                            $SortObjectFormat = "Naam"
                            $InputKey = $InputChoice
                        }
                        If ($SortObjectFormat) {
                            If ($Segment -eq "Alles") {
                                $Counter = 0
                            }
                            Else {
                                $Counter -= $ColumnCounter
                            }
                            If ($Sorting -eq "Z-A") {
                                $global:Array = $global:Array | Sort-Object $SortObject
                                $Sorting = "A-Z"
                                BREAK
                            }
                            ElseIf ($Sorting -eq "A-Z") {
                                $global:Array = $global:Array | Sort-Object $SortObject -Descending
                                $Sorting = "Z-A"
                                BREAK
                            }
                            Else {
                                $Sorting = $null
                                BREAK
                            }
                        }
                    }
                }
            }
        } Until ($InputKey)
        If (!$InputSelection -AND !$InputChoice) {
            $Counter -= $ColumnCounter
        }
    } Until ($InputSelection)
    # ------------------------------------------------------
    # Return to the appropriate menu when X has been entered
    # ------------------------------------------------------
    if ($InputSelection -eq "Exit") {
        If ($FunctionName -like "*ActiveDirectory*Modify*Account*") {
            Identity-ActiveDirectory-Modify-Accounts-2-GetData
        }
        ElseIf ($FunctionName -like "*ActiveDirectory*") {
            Get-Variable -Exclude $global:StartupVariables | Remove-Variable -ErrorAction SilentlyContinue
            Script-Menu-Subcategories-1-Identity
        }
        ElseIf ($FunctionName -like "*Azure*") {
            Get-Variable -Exclude $global:StartupVariables | Remove-Variable -ErrorAction SilentlyContinue
            Script-Menu-Subcategories-1-Identity
        }
        ElseIf ($FunctionName -like "*Exchange*Modify*Mailbox*") {
            Messaging-Exchange-Modify-Mailboxes-2-GetData
        }
        ElseIf ($FunctionName -like "*Exchange*") {
            Get-Variable -Exclude $global:StartupVariables | Remove-Variable -ErrorAction SilentlyContinue
            Script-Menu-Subcategories-3-Messaging
        }
        Else {
            Get-Variable -Exclude $global:StartupVariables | Remove-Variable -ErrorAction SilentlyContinue
            Script-Menu-Categories-1-Start
        }
    }
    If ($Type -eq "Account" -OR $Type -eq "AccountZonderMailbox") {
        $global:Name = $global:Array[$InputChoice - 1].Name
        $global:DisplayName = $global:Array[$InputChoice - 1].DisplayName
        $global:GivenName = $global:Array[$InputChoice - 1].GivenName
        $global:SurName = $global:Array[$InputChoice - 1].SurName
        $global:SamAccountName = ($global:Array[$InputChoice - 1].SamAccountName) #.ToLower()
        $global:UserPrincipalName = ($global:Array[$InputChoice - 1].UserPrincipalName) #.ToLower()
        $global:DistinguishedName = $global:Array[$InputChoice - 1].DistinguishedName
        $global:Language = $global:Array[$InputChoice - 1].Country
    }
    If ($Type -eq "Credentials") {
        If ($InputSelection -eq "Ja") {
            $global:Credential = $global:Array[$InputChoice - 1].TargetName
        }
        If ($InputSelection -eq "Aanmaken") {
            $global:Credential = "Aanmaken"
        }
    }
    If ($Type -eq "Drive") {
        If ($InputSelection -ne "ZonderExport") {
            If ($FunctionName -like "*Exchange*") {
                $Server = "localhost"
            }
            $UNC = $global:Array[$InputChoice - 1].UNC
            If ($UNC -eq "n.a.") {
                $global:PathUNC = "\\" + $Server + "\" + ($global:Array[$InputChoice - 1].Drive.Substring(0, 1).ToLower()) + "$"
            }
            Else {
                $global:PathUNC = $UNC
            }
            $global:Drive = $global:Array[$InputChoice - 1].Drive + "\"
            $global:DriveName = $global:Array[$InputChoice - 1].Name
        }
    }
    If ($Type -eq "Emailadressen") {
        If ($InputSelection -eq "Toevoegen") {
            $global:Emailadres = $null
        }
        Else {
            $global:Emailadres = $global:Array[$InputChoice - 1].Name
        }
    }
    If ($Type -eq "Emaildomein") {
        $global:Emaildomein = $global:Array[$InputChoice - 1].Name
        $global:Email = $SamAccountName + "@" + $global:Emaildomein
    }
    If ($Type -like "Genereer*") {
        If ($InputSelection -eq "Ja") {
            $global:GegenereerdeOptie = $global:Array[$InputChoice - 1].Name
        }
        If ($InputSelection -eq "Handmatig") {
            $global:Handmatig = "Ja"
        }
    }
    If ($Type -eq "Mailbox") {
        If ($InputChoice -as [int]) {
            $global:Name = $global:Array[$InputChoice - 1].Name
            $global:DisplayName = $global:Array[$InputChoice - 1].DisplayName
            $global:Alias = $global:Array[$InputChoice - 1].Alias
            $global:SamAccountName = $global:Array[$InputChoice - 1].SamAccountName
            $global:Language = $global:Array[$InputChoice - 1].Language
            $global:Email = $global:Array[$InputChoice - 1].PrimarySmtpAddress
            $global:DistinguishedName = $global:Array[$InputChoice - 1].DistinguishedName
        }
    }
    If ($Type -eq "UPNSuffix") {
        $global:UPNSuffix = $global:Array[$InputChoice - 1].Name
    }
    If ($Functie -eq "Markeren") {
        [array]$global:Array = $global:Array | Where-Object { $_.Markeren -eq "Ja" }
        $global:Aantal = $null
        If ($global:Array.Count) {
            $ArrayTotal = $global:Array.Count
        }
        Else {
            $ArrayTotal = 1
        }
        If ($Type -like "Account*") {
            $ObjectType = "domeinaccount"
            $ObjectTypes = "domeinaccounts"
            If ($InputSelection -eq "CSV") {
                $global:CSV = "Ja"
            }
        }
        If ($Type -eq "Bestand") {
            $ObjectType = $Filter + "-bestand"
            $ObjectTypes = $Filter + "-bestanden"
        }
        If ($Type -eq "Groepen") {
            $ObjectType = "groep"
            $ObjectTypes = "groepen"
            If ($InputSelection -eq "Geen") {
                $ArrayTotal = 0
            }
        }
        If ($Type -eq "Homefolders") {
            $ObjectType = "homefolder"
            $ObjectTypes = "homefolders"
            $FilterArray = $global:Array | Where-Object { $_.HomePath -like "Geen homefolder*" }
            If ($FilterArray) {
                If ($FilterArray.Count) {
                    $ArrayTotal -= $FilterArray.Count
                }
                Else {
                    $ArrayTotal -= 1
                }
            }
        }
        If ($Type -eq "MailboxBulk") {
            $ObjectType = "mailbox"
            $ObjectTypes = "mailboxen"
        }
        If ($ArrayTotal -eq 1) {
            $global:Aantal = [string]$ArrayTotal + " " + $ObjectType
        }
        ElseIf ($ArrayTotal -ge 2) {
            $global:Aantal = [string]$ArrayTotal + " " + $ObjectTypes
        }
        Else {
            $global:Aantal = "n.a."
        }
    }
    If ($Functie -eq "Selecteren" -AND ($Type -ne "Credentials" -AND $Type -ne "Emailadressen" -AND $Type -notlike "Genereer*")) {
        [array]$global:Array = $global:Array[$InputChoice - 1]
    }
}
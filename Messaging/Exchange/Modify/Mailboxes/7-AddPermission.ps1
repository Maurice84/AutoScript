Function Messaging-Exchange-Modify-Mailboxes-7-AddPermission ([string]$Soort) {
    $global:Titel = ($MyInvocation.MyCommand).Name
    Script-Module-SetHeaders -Name $Titel
    If ($Soort -eq "FullAccess") {
        $global:Huidig = $VolledigeToegang
    }
    If ($Soort -eq "SendAs") {
        $global:Huidig = $VerzendenAls
    }
    If ($Soort -eq "SendOnBehalf") {
        $global:Huidig = $VerzendenNamens
    }
    $global:Subtitel = {
        Write-Host -NoNewLine "- Selected mailbox: "; Write-Host $Mailbox.DisplayName -ForegroundColor Yellow
        Write-Host -NoNewLine "- Current $Soort permissions: "; Write-Host $Huidig -ForegroundColor Yellow
        Write-Host
    }
    $Subtitel.Invoke()
    Do {
        Write-Host -NoNewLine "  Would you like to add "; Write-Host -NoNewLine "D" -ForegroundColor Yellow; Write-Host -NoNewLine "omain account(s) or a "; Write-Host -NoNewLine "G" -ForegroundColor Yellow; Write-Host -NoNewLine "roup? Use "; Write-Host -NoNewLine "X" -ForegroundColor Yellow; Write-Host -NoNewLine " to cancel: "
        $Choice = Read-Host
        $Input = @("D"; "G"; "X") -contains $Choice
        If (!$Input) {
            Write-Host "  Please use the letters above as input" -ForegroundColor Red; Write-Host
        }
        If ($Choice -eq "G" -AND $Soort -eq "SendOnBehalf") {
            Write-Host "  Currently it's not possible to set groups on Send-As permissions, please select domain account(s) instead" -ForegroundColor Red; Write-Host
            $Input = $null
        }
    } Until ($Input)
    Switch ($Choice) {
        "D" {
            $global:ObjectType = "domain account(s)"
            $global:AlleObjecten = "domain accounts"
        }
        "G" {
            $global:ObjectType = "group(s)"
            $global:AlleObjecten = "groups"
        }
        "X" {
            Messaging-Exchange-Modify-Mailboxes-3-Menu
        }
    } 
    Do {
        If ($global:OUFilter -ne "*") {
            Write-Host -NoNewLine "  Would you like to use the previous entered OU filter: "; Write-Host -NoNewLine $global:OUFilter -ForegroundColor Yellow; Write-Host -NoNewLine " to select $ObjectType? (Y/N): "
        }
        Else {
            Write-Host -NoNewLine "  Would you like to search $ObjectType through all Organizational Units? (Y/N): "
        }
        $Choice = Read-Host
        $Input = @("Y"; "N") -contains $Choice
        If (!$Input) {
            Write-Host "  Please use the letters above as input" -ForegroundColor Red; Write-Host
        }
    } Until ($Input)
    Write-Host -NoNewLine "  "
    Switch ($Choice) {
        "N" {
            Write-Host -NoNewLine ("Please enter the name (or a part) of the OU with $ObjectType, or press Enter for all $AlleObjecten" + ": ")
            $global:OUFilter = Read-Host
            Script-Module-SetHeaders -Name $Titel
            If ($global:OUFilter.Length -eq 0) {
                $global:OUFilter = "*"
                $global:OUFormat = "All $AlleObjecten"
            }
            Else {
                $global:OUFormat = $global:OUFilter
            }
        }
    }
    $global:Subtitel = {
        Write-Host -NoNewLine "- Selected mailbox: "; Write-Host $Mailbox.DisplayName -ForegroundColor Yellow
        Write-Host -NoNewLine "- Current $Soort permissions: "; Write-Host ($global:Objects -join ', ') -ForegroundColor Yellow
        Write-Host
        Write-Host -NoNewLine "- Entered OU filter: "; Write-Host $OUFormat -ForegroundColor Yellow
    }
    If ($AlleObjecten -eq "domain accounts") {
        If ($Soort -eq "SendOnBehalf") {
            Script-Index-Objects -CurrentTask $Task -FunctionName ($MyInvocation.MyCommand).Name -Filter $global:OUFilter -Type "MailboxBulk" -Functie "Markeren"
        }
        Else {
            Script-Index-Objects -CurrentTask $Task -FunctionName ($MyInvocation.MyCommand).Name -Filter $global:OUFilter -Type "AccountBulk" -Functie "Markeren"
        }
    } 
    If ($AlleObjecten -eq "groups") {
        Script-Index-Objects -CurrentTask $Task -FunctionName ($MyInvocation.MyCommand).Name -Filter $global:OUFilter -Type "Groepen" -Functie "Markeren"
    }
    [array]$global:RechtenArray = $global:Array
    Script-Module-SetHeaders -Name $Titel
    $global:Subtitel = {
        Write-Host -NoNewLine "- Selected mailbox: "; Write-Host $Mailbox.DisplayName -ForegroundColor Yellow
        Write-Host -NoNewLine "- Current $Soort permissions: "; Write-Host ($global:Objects -join ', ') -ForegroundColor Yellow
        Write-Host
    }
    $Subtitel.Invoke()
    $global:Fout = $null
    ForEach ($global:RechtenObject in $RechtenArray) {
        If ($AlleObjecten -eq "domain accounts") {
            $global:DisplayName = $RechtenObject.Name
            $global:SelectedUser = $env:userdomain + "\" + $RechtenObject.SamAccountName
        } 
        If ($AlleObjecten -eq "groups") {
            $global:DisplayName = $RechtenObject.Name
            $global:SelectedUser = $env:userdomain + "\" + $RechtenObject.Name
        }
        If ($global:Objects -notcontains $SamAccountName) {
            If ($Soort -eq "FullAccess") {
                $AddMailboxPermission = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("Add-MailboxPermission -Identity '$SamAccountName' -AccessRights FullAccess -User '$SelectedUser' -InheritanceType All -AutoMapping 0 -DomainController '$PDC' -Confirm:0 -WarningAction SilentlyContinue"))
                $CheckMailboxPermission = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("Get-MailboxPermission -Identity '$SamAccountName' -User '$SelectedUser' -DomainController '$PDC'"))
                If (!($CheckMailboxPermission | Select-Object User, @{Name = "AccessRight"; Expression = { $_.AccessRights -join "," } })) {
                    $global:Fout = $true
                }
                Else {
                    $global:FullAccess += $DisplayName
                    [array]$global:FullAccess = $global:FullAccess | Where-Object { $_ -ne "n.a." } | Sort-Object -Unique
                    If ($global:FullAccess.Length -ge 6) {
                        $global:VolledigeToegang = [string]$global:FullAccess.Length + " domain accounts/groups"
                    }
                    If ($global:FullAccess.Length -lt 5) {
                        $global:VolledigeToegang = $global:FullAccess -join ', '
                    }
                }
            }
            If ($Soort -eq "SendAs") {
                $AddADPermission = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("Add-ADPermission -Identity '$DistinguishedName' -ExtendedRights Send-As -User '$SelectedUser' -DomainController '$PDC' -WarningAction SilentlyContinue"))
                $CheckSendAs = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("Get-ADPermission -Identity '$DistinguishedName' -User '$SelectedUser' -DomainController '$PDC'"))
                If (!($CheckSendAs | Select-Object ExtendedRights)) {
                    $global:Fout = $true
                }
                Else {
                    $global:SendAs += $DisplayName
                    [array]$global:SendAs = $global:SendAs | Where-Object { $_ -ne "n.a." } | Sort-Object -Unique
                    If ($global:SendAs.Length -ge 6) {
                        $global:VerzendenAls = [string]$global:SendAs.Length + " domain accounts/groups"
                    } 
                    If ($global:SendAs.Length -lt 5) {
                        $global:VerzendenAls = $global:SendAs -join ', '
                    }
                }
            }
            If ($Soort -eq "SendOnBehalf") {
                $AddSendOnBehalf = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("Set-Mailbox -Identity '$SamAccountName' -GrantSendOnBehalfTo @{add='$DisplayName'} -DomainController '$PDC' -WarningAction SilentlyContinue"))
                $CheckSendOnBehalf = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("Get-Mailbox -Identity '$SamAccountName' -DomainController '$PDC'"))
                $GetSendOnBehalf = ($CheckSendOnBehalf | Select-Object @{Name = "GrantSendOnBehalfTo"; Expression = { $_.GrantSendOnBehalfTo } }).GrantSendOnBehalfTo
                ForEach ($SendOnBehalf in $GetSendOnBehalf) {
                    If ($SendOnBehalf -like "*$DisplayName*") {
                        $global:Fout = $false
                    }
                }
                If ($global:Fout -eq $false) {
                    $global:SendOnBehalf += $DisplayName
                    [array]$global:SendOnBehalf = $global:SendOnBehalf | Where-Object { $_ -ne "n.a." } | Sort-Object -Unique
                    If ($global:SendOnBehalf.Length -ge 6) {
                        $global:VerzendenNamens = [string]$global:SendOnBehalf.Length + " domain accounts/groups"
                    } 
                    If ($global:SendOnBehalf.Length -lt 5) {
                        $global:VerzendenNamens = $global:SendOnBehalf -join ', '
                    }
                }
                Else {
                    $global:Fout = $true
                }
            }
            If (!$Fout -OR $Fout -eq $false) {
                Script-Module-ReplicateAD
                Write-Host -NoNewLine "OK: The $Soort permissions for "; Write-Host -NoNewLine $DisplayName -ForegroundColor Yellow; Write-Host " added successfully"
            }
            Else {
                Write-Host -NoNewLine "ERROR: An error occurred adding the $Soort permissions for $DisplayName, please investigate!" -ForegroundColor Red
                $Pause.Invoke()
                Messaging-Exchange-Modify-Mailboxes-3-Menu
            }
        }
        Else {
            Write-Host -NoNewLine "INFO: The $Soort permissions for $DisplayName already set" -ForegroundColor Gray;
        }
    }
    Write-Host
    If (!$Fout) {
        Write-Host "The $Soort permissions are added successfully to the selected $ObjectType, now returning to the previous menu" -ForegroundColor Green
        $Pause.Invoke()
        Messaging-Exchange-Modify-Mailboxes-3-Menu
    }
}
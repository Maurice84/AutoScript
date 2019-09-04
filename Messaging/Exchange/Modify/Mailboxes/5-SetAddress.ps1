Function Messaging-Exchange-Modify-Mailboxes-5-SetAddress {
    # =========
    # Execution
    # =========
    Script-Module-SetHeaders -Name ($MyInvocation.MyCommand).Name
    Write-Host -NoNewLine "- Selected mailbox: "; Write-Host $WeergaveNaam -ForegroundColor Yellow
    Write-Host -NoNewLine "- Current Primary Mail Address: "; Write-Host $PrimairEmailadres -ForegroundColor Yellow
    Script-Index-Objects -FunctionName ($MyInvocation.MyCommand).Name -Type "Emailadressen" -Functie "Selecteren" -Objecten $global:ExtraEmailaddresses
    If ($global:Emailadres) {
        Do {
            Write-Host -NoNewLine "  Would you like to have "; Write-Host -NoNewLine $global:Emailadres -ForegroundColor Magenta; Write-Host -NoNewLine " set as "; Write-Host -NoNewLine "P" -ForegroundColor Yellow; Write-Host -NoNewLine "rimary mail address or "; Write-Host -NoNewLine "D" -ForegroundColor Yellow; Write-Host -NoNewLine "elete this mail address?: "
            $Choice = Read-Host
            $Input = @("P"; "D") -contains $Choice
        } Until ($Input)
        Switch ($Choice) {
            "P" {
                $SetEmailAddress = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("Set-Mailbox -Identity '$SamAccountName' -PrimarySmtpAddress '$Emailadres' -DomainController '$PDC'"))
                $CheckEmailAddress = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("Get-Mailbox -Identity '$SamAccountName' -DomainController '$PDC'"))
                If (($CheckEmailAddress | Select-Object PrimarySmtpAddress).PrimarySmtpAddress.ToString() -ne $Emailadres) {
                    Write-Host "ERROR: An error occurred setting the Primary Mail Address, please investigate!" -ForegroundColor Red
                    $Pause.Invoke()
                }
                Else {
                    Script-Module-ReplicateAD
                    $global:Mailbox = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("Get-Mailbox -Identity '$SamAccountName' -DomainController '$PDC' -WarningAction SilentlyContinue"))
                    # --------------------------------
                    # Primary & Extra Mail Address(es)
                    # --------------------------------
                    $global:Emailaddresses = @()
                    If (($Mailbox.EmailAddresses).ProxyAddressString) {
                        $global:MailboxEmailAddresses = ($Mailbox.EmailAddresses).ProxyAddressString
                    }
                    Else {
                        $global:MailboxEmailAddresses = $Mailbox.EmailAddresses
                    }
                    ForEach ($global:Address in ($MailboxEmailAddresses | Where-Object { $_ -notlike "*local" })) {
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
                }
                Invoke-Expression -Command ($MyInvocation.MyCommand).Name
            }
            "D" {
                $RemoveEmailAddress = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create(`
                    "Set-Mailbox -Identity '$SamAccountName' -EmailAddresses @{remove='$Emailadres'} -DomainController '$PDC'")
                )
                $CheckEmailAddress = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create(`
                    "Get-Mailbox -Identity '$SamAccountName' -DomainController '$PDC'")
                )
                If ($CheckEmailAddress | Where-Object { $_.EmailAddresses -like "*smtp:$Emailadres*" } | Select-Object EmailAddresses) {
                    Write-Host "ERROR: An error occurred deleting the mail address, please investigate!" -ForegroundColor Red
                    $Pause.Invoke()
                }
                Else {
                    Script-Module-ReplicateAD
                    $global:ExtraEmailaddresses = $global:ExtraEmailaddresses | Where-Object { $_.Name -ne $global:Emailadres }
                    If ($global:ExtraEmailaddresses.Count -ge 4) {
                        $global:Emailadressen = [string]$global:ExtraEmailaddresses.Count + " mail addresses"
                    }
                    Else {
                        $global:Emailadressen = $global:ExtraEmailaddresses.Name -join ', '
                    }
                    If (!$global:ExtraEmailaddresses) {
                        $global:Emailadressen = "n.a."
                    }
                }
                Invoke-Expression -Command ($MyInvocation.MyCommand).Name
            }
        }
    }
    Else {
        Script-Index-Objects -FunctionName ($MyInvocation.MyCommand).Name -Type "Emaildomein" -Functie "Selecteren"
        Script-Module-SetHeaders -Name $Titel
        Write-Host -NoNewLine "- Selected mailbox: "; Write-Host $WeergaveNaam -ForegroundColor Yellow
        Write-Host -NoNewLine "- Current Primary Mail Address: "; Write-Host $PrimairEmailadres -ForegroundColor Yellow
        Write-Host -NoNewLine "- Select a mail domain for the new mail address: "; Write-Host $Emaildomein -ForegroundColor Yellow
        Write-Host
        Do {
            Write-Host -NoNewLine ("  Please enter the prefix (before @" + $Emaildomein + ") of the new mail address: ")
            $global:EmailPrefix = Read-Host
            $global:Emailadres = ($EmailPrefix + "@" + $Emaildomein).ToLower()
            $global:EmailadresCheck = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create(`
                "Get-Mailbox -Identity '$Emailadres' -DomainController '$PDC' -ErrorAction SilentlyContinue -WarningAction SilentlyContinue")
            )
            If ($EmailadresCheck) {
                Write-Host ("This mail address is already in use at " + $EmailadresCheck.Name) -ForegroundColor Red
                Write-Host
                $Input = $null
            }
            Else {
                $Input = $EmailPrefix
            }
        } Until ($Input)
        $Input = $null
        Do {
            Write-Host -NoNewLine "  Would you like to add "; Write-Host -NoNewLine $global:Emailadres -ForegroundColor Magenta; Write-Host -NoNewLine " to this mailbox? (Y/N): "
            $Choice = Read-Host
            $Input = @("Y"; "N") -contains $Choice
        } Until ($Input)
        Switch ($Choice) {
            "Y" {
                $AddEmailAddress = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("Set-Mailbox -Identity '$SamAccountName' -EmailAddresses @{add='$Emailadres'} -DomainController '$PDC'"))
                $CheckEmailAddress = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("Get-Mailbox -Identity '$SamAccountName' -DomainController '$PDC'"))
                $GetEmailAddresses = ($CheckEmailAddress | Select-Object EmailAddresses).EmailAddresses
                ForEach ($EmailAddress in $GetEmailAddresses) {
                    If ($EmailAddress -like "*$Emailadres*") {
                        $Problem = $true
                    }
                }
                If ($Problem) {
                    Write-Host "ERROR: An error occurred adding the mail address, please investigate!" -ForegroundColor Red
                    $Pause.Invoke()
                }
                Else {
                    Script-Module-ReplicateAD
                    $global:Mailbox = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create(`
                        "Get-Mailbox -Identity '$SamAccountName' -DomainController '$PDC' -WarningAction SilentlyContinue")
                    )
                    # --------------------------------
                    # Primary & Extra mail address(es)
                    # --------------------------------
                    $global:Emailaddresses = @()
                    ForEach ($global:Address in ($Mailbox.EmailAddresses | Where-Object { $_ -notlike "*local" })) {
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
                }
                Invoke-Expression -Command ($MyInvocation.MyCommand).Name
            }
            "N" {
                Invoke-Expression -Command ($MyInvocation.MyCommand).Name
            }
        }
    }
}
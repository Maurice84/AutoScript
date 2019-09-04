Function Messaging-Exchange-Modify-Mailboxes-3-Menu {
    # =========
    # Execution
    # =========
    $FunctionModifyMailboxPermission = ($global:FunctionTaskNames | Where-Object { $_ -like "Messaging*Modify-Mailboxes-*-SelectPermission" })
    Script-Module-SetHeaders -CurrentTask $Task -Name ($MyInvocation.MyCommand).Name
    Write-Host "Account information:" -ForegroundColor Magenta
    Write-Host -NoNewLine "  1. DisplayName: "; Write-Host $WeergaveNaam -ForegroundColor Yellow
    Write-Host -NoNewLine "  2. Alias: "; Write-Host $Alias -ForegroundColor Yellow
    Write-Host -NoNewLine "  3. Mailbox Type: "; Write-Host $SoortMailbox -ForegroundColor Yellow
    Write-Host -NoNewLine "  4. Language: "; Write-Host $Taalinstelling -ForegroundColor Yellow
    Write-Host -NoNewLine "  5. Primary Mail Address: "; Write-Host $PrimairEmailadres -ForegroundColor Yellow
    Write-Host -NoNewLine "  6. Extra Mail Address(es): "; Write-Host $Emailadressen -ForegroundColor Yellow
    Write-Host -NoNewLine "  7. Email Address Policy: "; Write-Host $EmailadresBeleid -ForegroundColor Yellow
    Write-Host -NoNewLine "  8. Address Book Policy: "; Write-Host $AdresboekBeleid -ForegroundColor Yellow
    Write-Host
    Write-Host "Set Permissions:" -ForegroundColor Magenta
    Write-Host -NoNewLine "  9. Full Access: "; Write-Host $VolledigeToegang -ForegroundColor Yellow
    Write-Host -NoNewLine " 10. Send As: "; Write-Host $VerzendenAls -ForegroundColor Yellow
    Write-Host -NoNewLine " 11. Send On Behalf: "; Write-Host $VerzendenNamens -ForegroundColor Yellow
    Write-Host
    Write-Host "Mailbox information:" -ForegroundColor Magenta
    Write-Host -NoNewLine " - Size: " -ForegroundColor Gray; Write-Host -NoNewLine $TotalItemSize -ForegroundColor Yellow; Write-Host -NoNewLine $Quota -ForegroundColor Yellow; Write-Host
    Write-Host -NoNewLine " - Item Count: " -ForegroundColor Gray; Write-Host $ItemCount -ForegroundColor Yellow
    Write-Host -NoNewLine " - Domain Account: " -ForegroundColor Gray; Write-Host $Mailbox.SamAccountName -ForegroundColor Yellow
    Write-Host -NoNewLine " - Last Logon Time: " -ForegroundColor Gray; Write-Host $LastLogon -ForegroundColor Yellow
    Write-Host -NoNewLine " - last Logoff Time: " -ForegroundColor Gray; Write-Host $LastLogoff -ForegroundColor Yellow
    Write-Host
    Do {
        Write-Host -NoNewline "Please select a category or use "; Write-Host -NoNewLine "X" -ForegroundColor Yellow; Write-Host -NoNewline " to return to the previous menu: "
        $Choice = Read-Host
        $Input = @("1";"2";"3";"4";"5";"6";"7";"8";"9";"10";"11";"X") -contains $Choice
        If (!$Input) {
            Write-Host "Please use the letters/numbers above as input" -ForegroundColor Red;Write-Host
        }
    } Until ($Input)
    Switch ($Choice) {
        "1" {
            Do {
                Write-Host -NoNewLine "Would you like to change the DisplayName? (Y/N): "
                $Choice = Read-Host
                $Input = @("Y";"N") -contains $Choice
                If (!$Input) {
                    Write-Host "Please use the letters above as input" -ForegroundColor Red;Write-Host
                }
            } Until ($Input)
            Switch ($Choice) {
                "Y" {
                    $global:WeergaveNaam = Read-Host "Please enter a new DisplayName (visible to Recipients (i.e. Info)"
                    $SetDisplayName = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("Set-Mailbox -Identity '$SamAccountName' -DisplayName '$WeergaveNaam' -DomainController '$PDC'"))
                    $CheckDisplayName = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("Get-Mailbox -Identity '$SamAccountName' -DomainController '$PDC'"))
                    If (($CheckDisplayName | Select-Object DisplayName).DisplayName -ne "$WeergaveNaam") {
                        Write-Host "ERROR: An error occurred setting the DisplayName, please investigate!" -ForegroundColor Red
                        $Pause.Invoke()
                    }
                }
                "N" {
                    Invoke-Expression -Command ($MyInvocation.MyCommand).Name
                }
            }
            Invoke-Expression -Command ($MyInvocation.MyCommand).Name
        }
        "2" {
            Do {
                Write-Host -NoNewLine "Would you like to change the Alias? (Y/N): "
                $Choice = Read-Host
                $Input = @("Y";"N") -contains $Choice
                If (!$Input) {
                    Write-Host "Please use the letters above as input" -ForegroundColor Red;Write-Host
                }
            } Until ($Input)
            Switch ($Choice) {
                "J" {
                    $global:Alias = Read-Host "Please enter a new Alias (i.e. Test)"
                    $SetAlias = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("Set-Mailbox -Identity '$SamAccountName' -Alias '$Alias' -DomainController '$PDC'"))
                    $CheckAlias = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("Get-Mailbox -Identity '$SamAccountName' -DomainController '$PDC'"))
                    If (($CheckAlias | Select-Object Alias).Alias -ne $Alias) {
                        Write-Host "ERROR: An error occurred setting the Alias, please investigate!" -ForegroundColor Red
                        $Pause.Invoke()
                    }
                }
                "N" {
                    Invoke-Expression -Command ($MyInvocation.MyCommand).Name
                }
            }
            Invoke-Expression -Command ($MyInvocation.MyCommand).Name
        }
        "3" {
            Invoke-Expression -Command $FunctionTaskNames[[int]$FunctionTaskNames.IndexOf(($MyInvocation.MyCommand).Name) + 1]
        }
        "4" {
            Do {
                Write-Host -NoNewLine "Select language for the mailbox: "; Write-Host -NoNewLine "D" -ForegroundColor Yellow; Write-Host -NoNewLine "utch, "; Write-Host -NoNewLine "E" -ForegroundColor Yellow; Write-Host -NoNewLine "nglish or "; Write-Host -NoNewLine "G" -ForegroundColor Yellow; Write-Host -NoNewLine "erman. Use "; Write-Host -NoNewLine "X" -ForegroundColor Yellow; Write-Host -NoNewline " to cancel: "
                $Choice = Read-Host
                $Input = @("D";"E","G","X") -contains $Choice
                If (!$Input) {
                    Write-Host "Please use the letters above as input" -ForegroundColor Red;Write-Host
                }
            } Until ($Input)
            Switch ($Choice) {
                "D" {
                    $global:Taalinstelling = "Nederlands"
                    $global:Language = "nl-NL"
                    $global:DateFormat = "d-M-yyyy"
                    $global:TimeFormat = "HH:mm"
                }
                "E" {
                    $global:Taalinstelling="Engels"
                    $global:Language = "en-US"
                    $global:DateFormat = "M/d/yyyy"
                    $global:TimeFormat = "h:mm tt"
                }
                "G" {
                    $global:Taalinstelling="Duits"
                    $global:Language = "de-DE"
                    $global:DateFormat = "dd.MM.yyyy"
                    $global:TimeFormat = "HH:mm"
                }
                "X" {
                    Invoke-Expression -Command ($MyInvocation.MyCommand).Name
                }
            }
            $SetRegion = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create(`
                "Set-MailboxRegionalConfiguration -Identity '$SamAccountName' -Language '$Language' -DateFormat '$DateFormat' -TimeFormat '$TimeFormat' -DomainController '$PDC'")
            )
            $CheckRegion = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create(`
                "Get-MailboxRegionalConfiguration -Identity '$SamAccountName' -DomainController '$PDC' -WarningAction SilentlyContinue")
            )
            If ((($CheckRegion | Select-Object Language).Language).Name -ne $Language) {
                Write-Host "ERROR: An error occurred setting the language, please investigate!" -ForegroundColor Red
                $Pause.Invoke()
            }
            Invoke-Expression -Command ($MyInvocation.MyCommand).Name
        }
        "5" {
            Invoke-Expression -Command $FunctionTaskNames[[int]$FunctionTaskNames.IndexOf(($MyInvocation.MyCommand).Name) + 2]
        }
        "6" {
            Invoke-Expression -Command $FunctionTaskNames[[int]$FunctionTaskNames.IndexOf(($MyInvocation.MyCommand).Name) + 2]
        }
        "7" {
            Do {
                If ($EmailadresBeleid -eq "Disabled") {
                    Write-Host -NoNewLine "Would you like to "; Write-Host -NoNewLine "E" -ForegroundColor Yellow; Write-Host -NoNewLine "nable the Email Address Policy? Warning: This will change the Mail Address! Use "; Write-Host -NoNewLine "X" -ForegroundColor Yellow; Write-Host -NoNewline " to cancel: "
                    $Choice = Read-Host
                    $Input = @("E";"X") -contains $Choice
                }
                If ($EmailadresBeleid -eq "Enabled") {
                    Write-Host -NoNewLine "Would you like to "; Write-Host -NoNewLine "D" -ForegroundColor Yellow; Write-Host -NoNewLine "isable the Email Address Policy? Use "; Write-Host -NoNewLine "X" -ForegroundColor Yellow; Write-Host -NoNewline " to cancel: "
                    $Choice = Read-Host
                    $Input = @("D";"X") -contains $Choice
                }
                If (!$Input) {
                    Write-Host "Please use the letters above as input" -ForegroundColor Red;Write-Host
                }
            } Until ($Input)
            Switch ($Choice) {
                "E" {
                    $SetEmailAddressPolicy = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("Set-Mailbox -Identity '$SamAccountName' -EmailAddressPolicyEnabled 1 -DomainController '$PDC'"))
                    $CheckEmailAddressPolicy = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("Get-Mailbox -Identity '$SamAccountName' -DomainController '$PDC'"))
                    If (($CheckEmailAddressPolicy | Select-Object EmailAddressPolicyEnabled).EmailAddressPolicyEnabled -ne $true) {
                        Write-Host "ERROR: An error occurred enabling the Email Address Policy, please investigate!" -ForegroundColor Red
                        $Pause.Invoke()
                    } Else {
                        $global:EmailadresBeleid = "Enabled"
                        Invoke-Expression -Command $FunctionTaskNames[[int]$FunctionTaskNames.IndexOf(($MyInvocation.MyCommand).Name) - 1]
                    }
                }
                "I" {
                    $SetEmailAddressPolicy = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("Set-Mailbox -Identity '$SamAccountName' -EmailAddressPolicyEnabled 0 -DomainController '$PDC'"))
                    $CheckEmailAddressPolicy = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("Get-Mailbox -Identity '$SamAccountName' -DomainController '$PDC'"))
                    If (($CheckEmailAddressPolicy | Select-Object EmailAddressPolicyEnabled).EmailAddressPolicyEnabled -ne $false) {
                        Write-Host "ERROR: An error occurred disabling the Email Address Policy, please investigate!" -ForegroundColor Red
                        $Pause.Invoke()
                    } Else {
                        $global:EmailadresBeleid = "Disabled"
                    }
                }
                "X" {
                    Invoke-Expression -Command ($MyInvocation.MyCommand).Name
                }
            }
            Invoke-Expression -Command ($MyInvocation.MyCommand).Name
        }
        "8" {
            $global:Subtitel = {
                Write-Host -NoNewLine "- Selected mailbox: "; Write-Host $WeergaveNaam -ForegroundColor Yellow
                Write-Host -NoNewLine "- Current Address Book Policy: "; Write-Host $AdresboekBeleid -ForegroundColor Yellow
            }
            Script-Index-Objects -CurrentTask $Task -FunctionName ($MyInvocation.MyCommand).Name -Type "Adresboekbeleid" -Functie "Selecteren"
            $AddressBookPolicy = $global:Array.Name
            $SetAddressBookPolicy = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("Set-Mailbox -Identity '$SamAccountName' -AddressBookPolicy '$AddressBookPolicy' -DomainController '$PDC'"))
            $CheckAddressBookPolicy = Script-Connect-Server -Module "Exchange" -Command ([scriptblock]::Create("Get-Mailbox -Identity '$SamAccountName' -DomainController '$PDC'"))
            If (($CheckAddressBookPolicy | Select-Object AddressBookPolicy).AddressBookPolicy -ne $AddressBookPolicy) {
                Write-Host "ERROR: An error occurred setting the Address Book Policy, please investigate!" -ForegroundColor Red
                $Pause.Invoke()
            } Else {
                $global:AdresboekBeleid = $global:Array.Name
            }
            Invoke-Expression -Command ($MyInvocation.MyCommand).Name
        }
        "9" {
            Invoke-Expression -Command ("$FunctionModifyMailboxPermission -Soort 'FullAccess'")
        }
        "10" {
            Invoke-Expression -Command ("$FunctionModifyMailboxPermission -Soort 'SendAs'")
        }
        "11" {
            Invoke-Expression -Command ("$FunctionModifyMailboxPermission -Soort 'SendOnBehalf'")
        }
        "X" {
            Get-Variable -Exclude $global:StartupVariables | Remove-Variable -ErrorAction SilentlyContinue; Script-Disconnect-Server; Invoke-Expression -Command $global:MenuNameCategory
        }
    }
}
Function Identity-ActiveDirectory-Delete-Accounts-6-Start {
    # ============
    # Declarations
    # ============
    $Task = "Deleting selected domain account(s)"
    # =========
    # Execution
    # =========
    Script-Module-SetHeaders -CurrentTask $Task -Name ($MyInvocation.MyCommand).Name
    Identity-ActiveDirectory-Overview-Accounts-2-Export
    ForEach ($Object in $global:ArrayAccounts) {
        $ObjectName = $Object.Name
        $Account = Script-Connect-Server -Module "ActiveDirectory" -Command ([scriptblock]::Create(
            "Get-ADUser -Filter * -Properties * | Where-Object {`$`_.Name -eq '$ObjectName'}")
        )
        $AccountSID = $Account.SID
        $ADSI = [ADSI]('LDAP://{0}' -f $Account.DistinguishedName)
        $RDPPath = Try { $ADSI.InvokeGet('TerminalServicesProfilePath') } Catch { }
        If (!$RDPPath -AND ($Account.DistinguishedName -notlike "*Mailbox*")) {
            Do {
                Write-Host -NoNewLine "Would you like to scan for User Profile Disks for this account to remove them also? (Y/N): "
                [string]$InputChoice = Read-Host
                $InputKey = @("Y", "N") -contains $InputChoice
                If (!$InputKey) {
                    Write-Host "Please use the letters above as input" -ForegroundColor Red
                }
            } Until ($InputKey)
            Switch ($InputChoice) {
                "Y" {
                    # ------------------------
                    # Index User Profile Disks
                    # ------------------------
                    Write-Host -NoNewLine "Indexing User Profile Disk path per RDSH server..."
                    $global:ArrayUPDShare = @()
                    $global:ArrayUPDFiles = @()
                    $global:RDSGroups = Script-Connect-Server -Module "ActiveDirectory" -Command ([scriptblock]::Create(`
                        "(Get-ADGroup -Filter {Name -like `"RDS*`"} | Select-Object Name).Name")
                    )
                    If ($RDSGroups) {
                        ForEach ($global:RDSGroup in $RDSGroups) {
                            $global:RDSServers = Script-Connect-Server -Module "ActiveDirectory" -Command ([scriptblock]::Create(`
                                "(Get-ADGroupMember $RDSGroup -Recursive | Where-Object {`$_.objectClass -eq `"computer`"} | Sort-Object Name | Select-Object Name).Name")
                            )
                            ForEach ($global:RDSServer in $global:RDSServers) {
                                $global:TestConnection = Test-Connection $RDSServer -Count 1 -ErrorAction SilentlyContinue | Out-Null
                                If ($?) {
                                    $global:Reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $RDSServer)
                                    $global:RegKey = $Reg.OpenSubKey("SYSTEM\\CurrentControlSet\\Control\\Terminal Server\\ClusterSettings")
                                    $global:UPDPath = $RegKey.GetValue("UvhdShareUrl")
                                    If ($global:UPDPath) {
                                        If ($global:UPDPath[-1] -ne "\") {
                                            $global:UPDPath = $global:UPDPath + "\"
                                        }
                                        $global:ArrayUPDShare += $global:UPDPath
                                    }
                                }
                            }
                        }
                        If ($global:ArrayUPDShare) {
                            $global:ArrayUPDShare = $global:ArrayUPDShare | Select-Object -Unique | Sort-Object
                            ForEach ($Path in $global:ArrayUPDShare) {
                                $global:ArrayUPDFiles += (Get-ChildItem $Path | Where-Object { $_.Name -like "UVHD-S-1-5-21-*" } | Select-Object FullName).FullName
                            }
                        }
                    }
                    If ($global:ArrayUPDFiles) {
                        $AccountUPD = $global:ArrayUPDFiles | Where-Object { $_ -like "*$AccountSID*" }
                        If ($AccountUPD) {
                            ForEach ($VHDX in $AccountUPD) {
                                Remove-Item $VHDX -ErrorAction SilentlyContinue
                                If ($?) {
                                    Write-Host -NoNewLine " - "; Write-Host -NoNewLine $Account.Name -ForegroundColor Yellow; Write-Host ": User Profile Disk successfully deleted"
                                }
                                Else {
                                    Write-Host -NoNewLine " - "; Write-Host -NoNewLine ($Account.Name + ": Could not delete User Profile Disk, please investigate!") -ForegroundColor Magenta
                                }
                            }
                            Write-Host
                        }
                    }
                }
            }
        }
        $AccountDN = $Account.DistinguishedName
        Script-Connect-Server -Module "ActiveDirectory" -Command ([scriptblock]::Create(`
            "Try {Remove-ADObject -Identity '$AccountDN' -Server '$global:PDC' -ErrorAction Stop -Confirm:`$false`
            } Catch {Remove-ADObject -Identity '$AccountDN' -Server '$global:PDC' -Recursive -Confirm:`$false}")
        )
        If ($?) {
            Write-Host -NoNewLine " - "; Write-Host -NoNewLine $Account.Name -ForegroundColor Yellow; Write-Host ": Domain account successfully deleted"
        }
        Else {
            Write-Host -NoNewLine " - "; Write-Host -NoNewLine ($Account.Name + ": Could not delete domain account, please investigate!") -ForegroundColor Magenta
        }
    }
    # ==========
    # Finalizing
    # ==========
    $Pause.Invoke()
    Get-Variable -Exclude $global:StartupVariables | Remove-Variable -ErrorAction SilentlyContinue; Script-Disconnect-Server; Invoke-Expression -Command $global:MenuNameCategory
}
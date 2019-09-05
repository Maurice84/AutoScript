Function Script-Module-ImportMicrosoft {
    param (
        [string]$Module,
        [string]$Name,
        [boolean]$Forced = $false
    )
    Script-Module-SetHeaders -Name $Name
    # ===============================
    # Loading module Active Directory
    # ===============================
    If ($Module -eq "ActiveDirectory") {
        If (!$global:ActiveDirectory) {
            If ($Forced -eq $false) {
                Do {
                    Write-Host -NoNewLine "  > Would you like to connect to Office 365 (Azure AD)? (Y/N): " -ForegroundColor Cyan
                    $Choice = Read-Host
                    $Input = @("Y", "N") -contains $Choice
                    If (!$Input) {
                        Write-Host "    Please use the letters above as input" -ForegroundColor Red; Write-Host
                    }
                } Until ($Input)
            }
            Else {
                $Choice = "N"
            }
            Switch ($Choice) {
                "Y" {
                    Script-Connect-Office365 -Module "AzureAD"
                }
                "N" {
                    # ------------------------------------------
                    # Checking if the server is Windows SBS 2008
                    # ------------------------------------------
                    $global:OS = (Get-WmiObject -class Win32_OperatingSystem | Select-Object Caption).Caption
                    If ($global:OS -like "*2008*FE*" -OR $global:OS -like "*2003*") {
                        Write-Host; Write-Host "  - Ignoring the Active Directory module for SBS 2008 due to problems..."
                        $global:OS = "SBS2008"
                        $global:ActiveDirectory = $true
                    }
                    Else {
                        # --------------------------------------------------
                        # Checking if the Active Directory module is present
                        # --------------------------------------------------
                        Write-Host -NoNewLine "  - Checking if the Active Directory module is present..."
                        $PSModulePath = (Get-ItemProperty -Path "HKLM:SYSTEM\CurrentControlSet\Control\Session Manager\Environment" -Name PSModulePath -ErrorAction SilentlyContinue).PSModulePath
                        $PSModulePath = $PSModulePath.Split(';')
                        $PSModulePath = $PSModulePath | Where-Object { $_ -like "*system32*" }
                        If (Get-ChildItem ($PSModulePath + "\ActiveDirectory") -ErrorAction SilentlyContinue) {
                            Write-Host -NoNewLine " Found!" -ForegroundColor Green
                            If (!(Get-Module ActiveDirectory -ListAvailable)) {
                                # -----------------------------------
                                # Loading the Active Directory module
                                # -----------------------------------
                                Write-Host; Write-Host -NoNewLine "  - Loading the Active Directory module..."
                                Import-Module ActiveDirectory -ErrorAction SilentlyContinue
                                If ($?) {
                                    $global:PDC = Get-ADDomain | Select-Object -ExpandProperty PDCEmulator
                                }
                            }
                            Else {
                                $global:PDC = Get-ADDomain | Select-Object -ExpandProperty PDCEmulator
                            }
                            $global:ActiveDirectory = $true
                        }
                        Else {
                            # ------------------------------------------------------------
                            # Detecting the presence of an Active Directory on the network
                            # ------------------------------------------------------------
                            Write-Host -NoNewLine " Not found" -ForegroundColor DarkGray
                            Write-Host; Write-Host -NoNewLine "  - Detecting the presence of an Active Directory on the network..."
                            $CheckDomainJoined = Try { (netdom query fsmo | Where-Object { $_ -like "*PDC*" }) } Catch { }
                            If (!$CheckDomainJoined) {
                                Do {
                                    Write-Host; Write-Host -NoNewLine "  - There is no Active Directory detected. Would you like to connect to Office 365? (Y/N): " -ForegroundColor Yellow
                                    $Choice = Read-Host
                                    $Input = @("Y"; "N") -contains $Choice
                                    If (!$Input) {
                                        Write-Host "    Please use the letters/numbers above as input" -ForegroundColor Red; Write-Host
                                    }
                                } Until ($Input)
                                Switch ($Choice) {
                                    "Y" {
                                        $global:Office365Connect = $true
                                    }
                                    "N" {
                                        Script-Menu-Categories-1-Start
                                    }
                                }
                            }
                            Else {
                                $global:PDC = $CheckDomainJoined.Split(" ")[-1]
                                Write-Host -NoNewLine (" " + $global:PDC.Split('.')[0]) -ForegroundColor Magenta; Write-Host
                                # ---------------------------------------------------
                                # Checking the possibility to use PowerShell Remoting
                                # ---------------------------------------------------
                                Write-Host -NoNewLine "  - Connecting to the Active Directory..."
                                $global:RemoteActiveDirectory = Try { Invoke-Command -ComputerName $PDC -ScriptBlock { Get-Module ActiveDirectory -ListAvailable | Out-Null; If ($?) { $true } } } Catch { $false }
                                If ($global:RemoteActiveDirectory -eq $true) {
                                    Write-Host -NoNewLine " OK!" -ForegroundColor Green
                                    $global:ActiveDirectory = $true
                                }
                                Else {
                                    Write-Host; Write-Host -NoNewLine "    ERROR: Could not connect to the Active Directory, is PowerShell Remoting enabled on the DC? You can enable this on the server using " -ForegroundColor Magenta; Write-Host -NoNewLine "winrm quickconfig" -ForegroundColor Yellow
                                    Write-Host
                                    $Pause.Invoke()
                                    Script-Menu-Categories-1-Start
                                }
                            }
                        }
                    }
                    If ($global:ActiveDirectory) {
                        # -----------------------------------------
                        # Detecteren van de Easy-Cloud Entry domein
                        # -----------------------------------------
                        #If ($env:userdnsdomain -eq "entry.easy-cloud.nl") {
                        #    $global:EasyCloud = "Ja"
                        #}
                    }
                    Write-Host
                }
            }
        }        
    }
    # =======================
    # Loading Exchange module
    # =======================
    If ($Module -eq "Exchange") {
        If (!$global:Exchange) {
            If ($Forced -eq $false) {
                Do {
                    Write-Host -NoNewLine "  > Would you like to connect to Office 365 (Exchange Online)? (Y/N): " -ForegroundColor Cyan
                    $Choice = Read-Host
                    $Input = @("Y", "N") -contains $Choice
                    If (!$Input) {
                        Write-Host "    Please use the letters above as input" -ForegroundColor Red; Write-Host
                    }
                } Until ($Input)
            }
            Else {
                $Choice = "N"
            }
            Switch ($Choice) {
                "Y" {
                    Script-Module-ImportMicrosoft -Name ($MyInvocation.MyCommand).Name -Module "CredentialManager"
                    Script-Module-ImportMicrosoft -Name ($MyInvocation.MyCommand).Name -Module "AzureAD"
                    Script-Module-ImportMicrosoft -Name ($MyInvocation.MyCommand).Name -Module "MSOnline"
                    Script-Connect-Office365 -Module "AzureAD"
                    Script-Connect-Office365 -Module "ExchangeOnline"
                }
                "N" {
                    Write-Host
                    # -------------------------------------------
                    # Checking if this is an Exchange Server 2007
                    # -------------------------------------------
                    $RegExchangeModule = "HKLM:SOFTWARE\Microsoft\PowerShell\1\PowerShellSnapIns\"
                    $Exchange2007Snapin = "Microsoft.Exchange.Management.PowerShell.Admin"
                    If (Get-Item -Path ($RegExchangeModule + $Exchange2007Snapin) -ErrorAction SilentlyContinue) {
                        $ExchangeSnapin = $Exchange2007Snapin
                        $global:ExchangeVersion = "2007"
                    }
                    Else {
                        # ---------------------------------------------------
                        # Checking the presence of the Exchange Server module
                        # ---------------------------------------------------
                        Write-Host -NoNewLine "  - Checking the presence of the Exchange Server module..."
                        $RegExchangeModule = "HKLM:SOFTWARE\Microsoft\PowerShell\1\PowerShellSnapIns\"
                        $Exchange201xSnapin = "Microsoft.Exchange.Management.PowerShell.E2010"
                        If (Get-Item -Path ($RegExchangeModule + $Exchange201xSnapin) -ErrorAction SilentlyContinue) {
                            $ExchangeSnapin = $Exchange201xSnapin
                            $global:ExchangeVersion = (Get-ItemProperty -Path ($RegExchangeModule + $Exchange201xSnapin) -ErrorAction SilentlyContinue | Select-Object Version).Version
                            If ($global:ExchangeVersion -like "14*") {
                                $global:ExchangeVersion = "2010"
                            }
                            If ($global:ExchangeVersion -like "15*") {
                                $global:ExchangeVersion = "2013"
                            }
                        }
                        Else {
                            # -------------------------------------------
                            # Detecting Exchange Server(s) on the network
                            # -------------------------------------------
                            Write-Host -NoNewLine " Not found" -ForegroundColor DarkGray
                            Write-Host; Write-Host -NoNewLine "  - Detecting Exchange Server(s) on the network..."
                            $CheckDomainJoined = Try { (netdom query fsmo | Where-Object { $_ -like "*PDC*" }) } Catch { }
                            If (!$CheckDomainJoined) {
                                Write-Host -NoNewLine " Not found" -ForegroundColor DarkGray
                            }
                            Else {
                                $global:PDC = $CheckDomainJoined.Split(" ")[-1]
                                $global:ExchangeServers = Invoke-Command -ComputerName $PDC -ScriptBlock { Import-Module ActiveDirectory; $SearchBase = "CN=Configuration," + (Get-ADDomain).DistinguishedName;
                                    (Get-ADObject -LDAPFilter "(objectClass=msExchExchangeServer)" -SearchBase $SearchBase | Where-Object { $_.ObjectClass -eq "msExchExchangeServer" } | Select-Object Name) }
                            }
                            If ($global:ExchangeServers) {
                                $global:ExchangeServers = ($global:ExchangeServers | Select-Object Name).Name
                                Write-Host -NoNewLine (" " + ($global:ExchangeServers -join ", ")) -ForegroundColor Magenta
                                ForEach ($Server in $global:ExchangeServers) {
                                    If (!$global:Exchange) {
                                        # ---------------------------------------------------
                                        # Checking the possibility to use PowerShell Remoting
                                        # ---------------------------------------------------
                                        Write-Host; Write-Host -NoNewLine "  - Connecting to an Exchange Server"
                                        If ($global:ExchangeServers.Count -ge 2) {
                                            Write-Host -NoNewLine (" " + $Server)
                                        }
                                        Write-Host -NoNewLine "..."
                                        $global:RemoteExchange = Try { Invoke-Command -ConfigurationName Microsoft.Exchange -ConnectionUri http://$Server/PowerShell/ -ScriptBlock { $true } -ErrorAction SilentlyContinue } Catch { $false }
                                        If ($global:RemoteExchange -eq $true) {
                                            Write-Host -NoNewLine " OK!" -ForegroundColor Green
                                            $global:Exchange = $true
                                            $global:ExchangeServer = $Server
                                            BREAK
                                        }
                                        Else {
                                            Write-Host -NoNewLine " Failed!" -ForegroundColor Red
                                        }
                                    }
                                }
                                If (!$global:Exchange) {
                                    Write-Host; Write-Host "    ERROR: Could not connect to an Exchange Server, is it offline?" -ForegroundColor Red
                                    Write-Host; "    If it's an Exchange Server 2007 then you need to run this script on the server."
                                    Write-Host
                                    $Pause.Invoke()
                                    Script-Menu-Categories-1-Start
                                }
                            }
                            Else {
                                If (!$Office365Connect) {
                                    Do {
                                        Write-Host; Write-Host -NoNewLine "  > There is no Exchange Server detected. Would you like to connect to Office 365? (Y/N): " -ForegroundColor Yellow
                                        $Choice = Read-Host
                                        $Input = @("N") -contains $Choice
                                        If (!$Input) {
                                            Write-Host "    Please use the letters/numbers above as input" -ForegroundColor Red; Write-Host
                                        }
                                    } Until ($Input)
                                    Switch ($Choice) {
                                        "Y" {
                                            Script-Connect-Office365
                                        }
                                        "N" {
                                            Script-Menu-Categories-1-Start
                                        }
                                    }
                                }
                            }
                        }
                    }
                    If ($global:ExchangeVersion) {
                        # -------------------------
                        # Importing Exchange module
                        # -------------------------
                        Write-Host -NoNewLine " Found!" -ForegroundColor Green
                        If (!(Get-PSSnapin $ExchangeSnapin -ErrorAction SilentlyContinue)) {
                            Write-Host; Write-Host -NoNewLine "  - Loading the Exchange Server $global:ExchangeVersion module..."
                            Add-PSSnapin -Name $ExchangeSnapin
                            If (!$?) {
                                Write-Host; Write-Host "    ERROR: A problem occurred loading the $global:ExchangeVersion Module, please investigate!" -ForegroundColor Red
                                Write-Host
                                $Pause.Invoke()
                                Script-Menu-Categories-1-Start
                            }
                            Else {
                                $global:Exchange = $true
                                $global:ExchangeServer = $env:computername
                            }
                        }
                        Else {
                            $global:Exchange = $true
                        }
                    }
                    # ---------------------------------------------------------------------------------------------------------------------
                    # Importing ActiveDirectory module (must be done after loading Exchange module due to a bug with Windows Server 2008 R2
                    # ---------------------------------------------------------------------------------------------------------------------
                    Write-Host; Script-Module-ImportMicrosoft -Name ($MyInvocation.MyCommand).Name -Module "ActiveDirectory" -Forced $true
                }
            }
        }
        Write-Host
    }
    # =================================
    # Loading Credential Manager module
    # =================================
    If ($Module -eq "CredentialManager") {
        If (!$global:CredManager) {
            If (!(Get-Module -Name CredentialManager -ErrorAction SilentlyContinue)) {
                Write-Host -NoNewLine "  - Loading the Credential Manager module..."
                Import-Module -Name CredentialManager -ErrorAction SilentlyContinue
                If ($?) {
                    Write-Host
                    $global:CredManager = $true
                }
                Else {
                    Write-Host
                    $NuGetProvider = Try { Get-PackageProvider -Name "NuGet" -ErrorAction SilentlyContinue } Catch { $NuGetProvider = $null }
                    If ($NuGetProvider -ne $null) {
                        Install-PackageProvider -Name "NuGet" -Confirm:$false -Force | Out-Null
                        If ($?) {
                            Install-Module -Name CredentialManager -ErrorAction SilentlyContinue -Force -Confirm:$false
                            If ($?) {
                                Import-Module -Name CredentialManager -ErrorAction SilentlyContinue
                                If ($?) {
                                    $global:CredManager = $true
                                }
                            }
                            Else {
                                Write-Host "   ERROR: A problem occurred during installation of the Credential Manager module, please investigate!" -ForegroundColor Red
                                Write-Host
                                $Pause.Invoke()
                                Script-Menu-Categories-1-Start
                            }
                        }
                        Else {
                            Write-Host "    ERROR: A problem occurred during installation of the NuGet package provider, please investigate!" -ForegroundColor Red
                            Write-Host
                            $Pause.Invoke()
                            Script-Menu-Categories-1-Start
                        }
                    }
                }
            }
            Else {
                $global:CredManager = $true
            }
        }
    }
    # =======================
    # Loading Azure AD module
    # =======================
    If ($Module -eq "AzureAD") {
        If (!$global:AzureAD) {
            If (!(Get-Module -Name AzureAD)) {
                If (Get-Module -Name AzureAD -ListAvailable) {
                    Write-Host -NoNewLine "  - Loading the Azure AD PowerShell module..."
                    Import-Module AzureAD -ErrorAction SilentlyContinue
                    If ($?) {
                        Write-Host
                        $global:AzureAD = $true
                    }
                }
                Else {
                    If ($PSVersionTable.PSVersion.Major -le 4) {
                        $PSVersie = $PSVersionTable.PSVersion.Major
                        Script-Module-SetHeaders -Name $global:MainTitel
                        Write-Host "    Warning: PowerShell $global:PowerShell has been detected, please upgrade to the newest version for Azure/Office 365 and restart." -ForegroundColor Red
                        Write-Host "    The most recent version is Windows PowerShell 5.1. Compatible with Windows Server 2008 R2 SP1, .NET 4.5 is pre-required:" -ForegroundColor Red
                        Write-Host "    https://www.microsoft.com/en-us/download/details.aspx?id=54616" -ForegroundColor Yellow
                        Write-Host
                        $Pause.Invoke()
                        Script-Menu-Categories-1-Start
                    }
                    Else {
                        Do {
                            Write-Host -NoNewLine "  > The Azure AD PowerShell module is not present. Would you like to install this? (Y/N): " -ForegroundColor Yellow
                            [string]$Choice = Read-Host
                            $Input = @("Y", "N") -contains $Choice
                            If (!$Input) {
                                Write-Host "    Please use the letters above as input" -ForegroundColor Red
                            }
                        } Until ($Input)
                        Switch ($Choice) {
                            "Y" {
                                Write-Host -NoNewLine "  - Installing the NuGet package provider..."
                                Install-PackageProvider -Name NuGet -ErrorAction SilentlyContinue -Force -Confirm:$false | Out-Null
                                If ($?) {
                                    Write-Host
                                    Write-Host -NoNewLine "  - Installing the Azure AD PowerShell Module..."
                                    Install-Module -Name AzureAD -ErrorAction SilentlyContinue -Force -Confirm:$false
                                    If ($?) {
                                        Write-Host
                                        $global:AzureAD = $null
                                        Script-Module-ImportMicrosoft -Name ($MyInvocation.MyCommand).Name -Module "AzureAD"
                                    }
                                    Else {
                                        Write-Host
                                        Write-Host "    ERROR: An error occurred during the installation of the Azure AD PowerShell Module, please investigate!" -ForegroundColor Red
                                        $Pause.Invoke()
                                        Script-Menu-Categories-1-Start
                                    }
                                }
                                Else {
                                    Write-Host
                                    Write-Host "    ERROR: An error occurred during the installation of the NuGet package provider, please investigate!" -ForegroundColor Red
                                    $Pause.Invoke()
                                    Script-Menu-Categories-1-Start
                                }
                            }
                        }
                    }
                }
            }
        }
    }
    # =========================
    # Laden van module MSOnline
    # =========================
    If ($Module -eq "MSOnline") {
        If (!$global:MSOnline) {
            If (!(Get-Module -Name MSOnline)) {
                If (Get-Module -Name MSOnline -ListAvailable) {
                    # ----------------------------
                    # Laden van de MSOnline module
                    # ----------------------------
                    Write-Host -NoNewLine "  - Loading the MSOnline PowerShell module..."
                    Import-Module MSOnline -ErrorAction SilentlyContinue
                    If ($?) {
                        Write-Host
                        $global:MSOnline = $true
                    }
                }
                Else {
                    If ($PSVersionTable.PSVersion.Major -le 4) {
                        $PSVersie = $PSVersionTable.PSVersion.Major
                        Script-Module-SetHeaders -Name $global:MainTitel
                        Write-Host "    Warning: PowerShell $global:PowerShell has been detected, please upgrade to the newest version for Azure/Office 365 and restart." -ForegroundColor Red
                        Write-Host "    The most recent version is Windows PowerShell 5.1. Compatible with Windows Server 2008 R2 SP1, .NET 4.5 is prerequired:" -ForegroundColor Red
                        Write-Host "    https://www.microsoft.com/en-us/download/details.aspx?id=54616" -ForegroundColor Yellow
                        Write-Host
                        $Pause.Invoke()
                        Script-Menu-Categories-1-Start
                    }
                    Else {
                        Do {
                            Write-Host -NoNewLine "  > The MSOnline PowerShell module is not present. Would you like to install this? (Y/N): " -ForegroundColor Yellow
                            [string]$Choice = Read-Host
                            $Input = @("Y", "N") -contains $Choice
                            If (!$Input) {
                                Write-Host "    Please use the letters above as input" -ForegroundColor Red
                            }
                        } Until ($Input)
                        Switch ($Choice) {
                            "Y" {
                                Write-Host -NoNewLine "  - Installing the NuGet package provider..."
                                Install-PackageProvider -Name NuGet -ErrorAction SilentlyContinue -Force -Confirm:$false | Out-Null
                                If ($?) {
                                    Write-Host
                                    Write-Host -NoNewLine "  - Installing the MSOnline PowerShell Module..."
                                    Install-Module -Name MSOnline -ErrorAction SilentlyContinue -Force -Confirm:$false
                                    If ($?) {
                                        Write-Host
                                        $global:MSOnline = $null
                                        Script-Module-ImportMicrosoft -Name ($MyInvocation.MyCommand).Name -Module "MSOnline"
                                    }
                                    Else {
                                        Write-Host
                                        Write-Host "    ERROR: An error occurred during the installation of the MSOnline PowerShell Module, please investigate!" -ForegroundColor Red
                                        $Pause.Invoke()
                                        Script-Menu-Categories-1-Start
                                    }
                                }
                                Else {
                                    Write-Host
                                    Write-Host "    ERROR: An error occurred during the installation of the NuGet package provider, please investigate!" -ForegroundColor Red
                                    $Pause.Invoke()
                                    Script-Menu-Categories-1-Start
                                }
                            }
                        }
                    }
                }
            }
        }
    }
}
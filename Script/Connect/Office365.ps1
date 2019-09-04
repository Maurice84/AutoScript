Function Script-Connect-Office365 {
    param (
        [string]$Module,
        [string]$Name,
        [string]$Task
    )
    # =========
    # Execution
    # =========
    Script-Module-SetHeaders -CurrentTask $Task -Name $Name
    If ($PowerShell -le 4) {
        Write-Host "Warning: PowerShell $global:PowerShell has been detected, please upgrade to the newest version for Azure/Office 365 and restart." -ForegroundColor Red
        Write-Host "The most recent version is Windows PowerShell 5.1. Compatible with Windows Server 2008 R2 SP1, .NET 4.5 is prerequired:" -ForegroundColor Red
        Write-Host "https://www.microsoft.com/en-us/download/details.aspx?id=54616" -ForegroundColor Yellow
        Write-Host
        $Pause.Invoke()
        Get-Variable -Exclude $global:StartupVariables | Remove-Variable -ErrorAction SilentlyContinue; Script-Disconnect-Server; Invoke-Expression -Command $global:MenuNameStart
    }
    If (!$Credential) {
        $global:Subtitle = {
            Write-Host "- Select an admin credential to connect to Office 365:"
        }
        $Subtitle.Invoke()
        Write-Host
        Do {
            Script-Index-Objects -CurrentTask $Task -FunctionName $Name -Filter "Office365" -Type "Credentials" -Functie "Selecteren"
            Script-Module-SetCredentials -Credentials $global:Credential -Type "Office365"
        } Until ($Credential -ne "Aanmaken")
    }
    If ($Credential) {
        Script-Module-SetHeaders -CurrentTask $Task -Name $Name
        If ($Module -eq "AzureAD") {
            # ----------------------
            # Connecting to Azure AD
            # ----------------------
            If (!$global:AzureADConnected) {
                Write-Host -NoNewLine "Connecting to Azure AD..."
                $global:AzureADConnect = Connect-AzureAD -Credential $global:Credential
                If ($?) {
                    $global:AzureADConnected = $true
                }
                Else {
                    Write-Host "ERROR: A problem has occurred connecting to Azure AD (password related?), please investigate!" -ForegroundColor Magenta
                    $Pause.Invoke()
                    Script-Menu-Categories-1-Start
                }
                Write-Host
            }
            # ----------------------
            # Connecting to MSOnline
            # ----------------------
            If (!$global:MSOnlineConnected) {
                Write-Host -NoNewLine "Connecting to MSOnline..."
                $global:MSOnlineConnect = Connect-MSOLService -Credential $Credential
                If ($?) {
                    $global:MSOnlineConnected = $true
                }
                Else {
                    Write-Host "A problem has occurred connecting to Azure AD (password related?), please investigate!" -ForegroundColor Magenta
                    $Pause.Invoke()
                    Script-Menu-Categories-1-Start
                }
                Write-Host
            }
        }
        If ($Module -eq "ExchangeOnline") {
            # -----------------------------
            # Connecting to Exchange Online
            # -----------------------------
            $global:Office365Exchange = Get-PSSession -Name "Office365" -ErrorAction SilentlyContinue
            If (!$global:Office365Exchange) {
                Write-Host -NoNewLine "Setting up an Exchange Online session..."
                # ----------------------------------------------------
                # Checking and enabling Basic authentication if needed
                # ----------------------------------------------------
                $Path = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WinRM\Client"
                if ((Get-ItemProperty -Path $Path -Name AllowBasic).AllowBasic -eq 0) {
                    try {
                        New-ItemProperty -Path $Path -Name AllowBasic -Value 1 -PropertyType DWORD -Force # | Out-Null
                    } catch {
                        Write-Host "ERROR: Basic Authentication could not be enabled! Please investigate $Path" -ForegroundColor Red
                        $Pause.Invoke()
                    }
                }
                $Count = 1
                Do {
                    $global:Office365Exchange = New-PSSession -Name "Office365" -ConfigurationName Microsoft.Exchange `
                        -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $global:Credential `
                        -Authentication Basic -AllowRedirection -WarningAction SilentlyContinue
                    If (!$?) {
                        If ($Count -lt 5) {
                            $Count++
                            Write-Host
                            Write-Host "Retry $Count/5 connecting to Exchange Online..." -ForegroundColor Gray
                            Start-Sleep -Seconds 10
                        }
                        Else {
                            Write-Host
                            Write-Host "A problem has occurred connecting to Exchange Online, please try again later." -ForegroundColor Red
                            $Pause.Invoke()
                            Get-Variable -Exclude $global:StartupVariables | Remove-Variable -ErrorAction SilentlyContinue; Invoke-Expression -Command $global:MenuNameStart
                        }
                    }
                } Until ($global:Office365Exchange)
            }
        }
    }
}
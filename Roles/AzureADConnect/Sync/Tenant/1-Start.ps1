Function Roles-AzureADConnect-Sync-Tenant-1-Start {
    # =========
    # Execution
    # =========
    Script-Module-SetHeaders -CurrentTask $Task -Name ($MyInvocation.MyCommand).Name
    Write-Host -NoNewLine "Please enter server name where the Azure AD Connect module is present (or press Enter for current server): "
    $global:Server = Read-Host
    If (!$Server) {
        $global:Server = $env:computername
    }
    Write-Host
    Write-Host -NoNewLine "- Connecting with $Server..." -ForegroundColor Gray
    $global:Session = New-PSSession -ComputerName $Server -ErrorAction SilentlyContinue
    If ($?) {
        Write-Host
        Write-Host -NoNewLine "- Loading Azure AD Connect module..." -ForegroundColor Gray
        $global:InvokeCommand = Invoke-Command -Session $Session -ScriptBlock {
            Import-Module ADSync -ErrorAction SilentlyContinue; If ($?) { $True } Else { $False }
        }
        If ($InvokeCommand -eq $True) {
            Write-Host
            Write-Host -NoNewLine "- Initializing Office 365 OU synchronization..." -ForegroundColor Gray
            $global:InvokeCommand = Invoke-Command -Session $Session -ScriptBlock { Start-ADSyncSyncCycle -PolicyType Delta }
            If ($InvokeCommand.Result -eq "Success") {
                Write-Host
                Write-Host "- OK: Office 365 OU synchronization successfully initiated" -ForegroundColor Green
            }
            Else {
                Write-Host
                Write-Host "- ERROR: An error occurred executing the Office 365 sync, please investigate!" -ForegroundColor Red
            }
        }
        Else {
            Write-Host
            Write-Host "- ERROR: An error occurred detecting the Azure AD Connect module on $Server, please investigate!" -ForegroundColor Red
        }
        Remove-PSSession $Session
    }
    Else {
        Write-Host
        Write-Host "- ERROR: An error occurred connecting to $Server" -ForegroundColor Red
        Write-Host
        Write-Host "  To fix this problem you need to run PowerShell (Run as Administrator) and execute the following commands:" -ForegroundColor Magenta
        Write-Host "  Set-Item WSMan:\localhost\Client\TrustedHosts -Value *" -ForegroundColor Yellow
        Write-Host "  Enable-PSRemoting -Force" -ForegroundColor Yellow
        Write-Host
        Write-Host -NoNewLine "  Don't forget to set " -ForegroundColor Magenta; Write-Host -NoNewLine "HKLM\Software\Microsoft\Windows\CurrentVersion\Policies\System\EnableLUA" -ForegroundColor Yellow; Write-Host -NoNewLine " to " -ForegroundColor Magenta; Write-Host -NoNewLine "0" -ForegroundColor Yellow; Write-Host " and restart the server" -ForegroundColor Magenta
    }
    # ==========
    # Finalizing
    # ==========
    $Pause.Invoke()
    Get-Variable -Exclude $global:StartupVariables | Remove-Variable -ErrorAction SilentlyContinue; Script-Disconnect-Server; Invoke-Expression -Command $global:MenuNameCategory
}
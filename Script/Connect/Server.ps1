Function Script-Connect-Server {
    param (
        [scriptblock]$Command,
        [string]$Module,
        [string]$Debug
    )
    if ($Debug) {
        Write-Host
        Write-Host "Module:" $Module -ForegroundColor Yellow
        Write-Host "Command:" $Command -ForegroundColor Yellow
    }
    if ($Module -eq "ActiveDirectory") {
        if ($RemoteActiveDirectory -eq $true) {
            if ($Debug) {
                Write-Host "RemoteActiveDirectory: $RemoteActiveDirectory" -ForegroundColor Yellow
            }
            if (!(Get-PSSession -Name "ActiveDirectory" -ErrorAction SilentlyContinue)) {
                $global:SessionAD = New-PSSession -Name "ActiveDirectory" -ComputerName $PDC
                Invoke-Command -Session $global:SessionAD -ScriptBlock {`
                        Import-Module ActiveDirectory;
                }
            }
            Invoke-Command -Session $global:SessionAD -ArgumentList $Command -ScriptBlock { param($Command); [Scriptblock]::Create($Command).Invoke() }
        }
        Else {
            if ($Debug) {
                Write-Host "RemoteActiveDirectory: n.a." -ForegroundColor Yellow
            }
            $Result = [Scriptblock]::Create($Command).Invoke()
            if ($Debug) {
                Write-Host "Result:" $Result -ForegroundColor Yellow
            }
        }
    }
    If ($Module -eq "Exchange") {
        If ($RemoteExchange -eq $true) {
            If (!(Get-PSSession -Name "Exchange" -ErrorAction SilentlyContinue)) {
                New-PSSession -Name "Exchange" -ConfigurationName Microsoft.Exchange -ConnectionUri http://$ExchangeServer/PowerShell/ | Out-Null
            }
            $Command = [scriptblock]::Create("Invoke-Command -Session (Get-PSSession -Name 'Exchange') -ScriptBlock {$Command}");
            $Command.Invoke()
        }
        Else {
            [Scriptblock]::Create($Command).Invoke()
        }
    }
    If ($Module -eq "Office365") {
        $Command = [scriptblock]::Create("Invoke-Command -Session (Get-PSSession -Name 'Office365') -ScriptBlock {$Command}");
        $Command.Invoke()
    }
    if ($Debug) {
        $Pause.Invoke()
    }
    return $Result
}
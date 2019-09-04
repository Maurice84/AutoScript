Function Identity-ActiveDirectory-Modify-ProfilePaths-1-SelectOU {
    # ============
    # Declarations
    # ============
    $Task = "Select an OU where the Users OU with domain accounts are located"
    # =========
    # Execution
    # =========
    Script-Module-SetHeaders -CurrentTask $Task -Name ($MyInvocation.MyCommand).Name
    Script-Module-ImportMicrosoft -Name ($MyInvocation.MyCommand).Name -Module "ActiveDirectory"
    Do {
        Script-Module-SetHeaders -CurrentTask $Task -Name ($MyInvocation.MyCommand).Name
        Script-Index-OU
        Write-Host
        $global:OUObjects = @()
        $global:CheckUsersOU = Script-Connect-Server -Module "ActiveDirectory" -Command ([scriptblock]::Create("Try {Get-ADOrganizationalUnit (`"OU=Users,`" + '$OUPath')}} Catch {`$false}"))
        If ($CheckUsersOU -eq $false) {
            Write-Host "  This OU does not contain an Users OU, please select an Accounts OU" -ForegroundColor Red
            $Pause.Invoke()
        }
        Else {
            $Choice = $true
        }
    } Until ($Choice)
    # ==========
    # Finalizing
    # ==========
    Set-Variable $global:VarHeaderName -Value $global:Subtree
    Set-Variable $global:VarHeaderName -Value ("Write-Host -NoNewLine ('- ' + '$Task' + ': ') -ForegroundColor Gray; Write-Host '" + (Get-Variable -Name $global:VarHeaderName).Value + "' -ForegroundColor Yellow")
    Invoke-Expression -Command $FunctionTaskNames[[int]$FunctionTaskNames.IndexOf(($MyInvocation.MyCommand).Name) + 1]
}
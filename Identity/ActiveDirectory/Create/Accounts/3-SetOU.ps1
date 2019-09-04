Function Identity-ActiveDirectory-Create-Accounts-3-SetOU {
    # ============
    # Declarations
    # ============
    $Task = "Select an OU where the account must be created"
    # =========
    # Execution
    # =========
    Script-Module-SetHeaders -CurrentTask $Task -Name ($MyInvocation.MyCommand).Name
    If ($global:Profiel -ne "Profielkopie") {
        Script-Index-OU
    }
    Else {
        $global:OUPath = Script-Connect-Server -Module "ActiveDirectory" -Command ([scriptblock]::Create(`
                    "(Get-ADUser '$global:SamAccountName' | Select-Object @{n='ParentContainer';e={`$`_.DistinguishedName -replace '^.+?,(CN|OU.+)','`$1'}}).ParentContainer")
        )
        # -----------------------------------------------------------------------------
        # Converting DistinguishedName (OU=..,OU=...,DC=...) to CanonicalName (\OU\OU\)
        # -----------------------------------------------------------------------------
        $global:Subtree = Script-Convert-DN-to-CN -Path $global:OUPath
    }
    # ==========
    # Finalizing
    # ==========
    Set-Variable $global:VarHeaderName -Value $global:Subtree
    Set-Variable $global:VarHeaderName -Value ("Write-Host -NoNewLine ('- ' + '$Task' + ': ') -ForegroundColor Gray; Write-Host '" + (Get-Variable -Name $global:VarHeaderName).Value + "' -ForegroundColor Yellow")
    Invoke-Expression -Command $FunctionTaskNames[[int]$FunctionTaskNames.IndexOf(($MyInvocation.MyCommand).Name) + 1]
}
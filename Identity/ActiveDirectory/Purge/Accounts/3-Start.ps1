Function Identity-ActiveDirectory-Purge-Accounts-3-Start {
    # ============
    # Declarations
    # ============
    $Task = "Purging the selected domain account(s)"
    # =========
    # Execution
    # =========
    Script-Module-SetHeaders -CurrentTask $Task -Name ($MyInvocation.MyCommand).Name
    ForEach ($Account in $global:Array) {
        $Name = $Account.Name + "*DEL:*"
        Script-Connect-Server -Module "ActiveDirectory" -Command ([scriptblock]::Create(`
            "Get-ADObject -Filter {isDeleted -eq `$true -and Name -like '$Name'} -IncludeDeletedObjects | Remove-ADObject -Confirm:`$false")
        )
        If ($?) {
            Write-Host -NoNewLine ("  " + $Account.Name) -ForegroundColor Yellow; Write-Host ": Domain account successfully purged"
        }
        Else {
            Write-Host ("  " + $Account.Name + ": Unable to purge domain account, please investigate!") -ForegroundColor Magenta
        }
    }
    Write-Host
    # ==========
    # Finalizing
    # ==========
    $Pause.Invoke()
    Get-Variable -Exclude $global:StartupVariables | Remove-Variable -ErrorAction SilentlyContinue; Script-Disconnect-Server; Invoke-Expression -Command $global:MenuNameCategory
}
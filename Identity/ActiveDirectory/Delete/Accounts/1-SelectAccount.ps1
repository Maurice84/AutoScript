Function Identity-ActiveDirectory-Delete-Accounts-1-SelectAccount {
    param (
        [array]$Objecten
    )
    # ============
    # Declarations
    # ============
    $Task = "Select domain account(s) to archive/delete"
    # =========
    # Execution
    # =========
    Script-Module-SetHeaders -CurrentTask $Task -Name ($MyInvocation.MyCommand).Name
    Script-Module-ImportMicrosoft -Name ($MyInvocation.MyCommand).Name -Module "ActiveDirectory"
    If (!$global:ArrayAccounts) {
        If (!$Objecten) {
            Script-Index-Objects -CurrentTask $Task -FunctionName ($MyInvocation.MyCommand).Name -Type "AccountBulk" -Functie "Markeren"
        }
        Else {
            Script-Index-Objects -CurrentTask $Task -FunctionName ($MyInvocation.MyCommand).Name -Objecten $Objecten -Type "AccountBulk" -Functie "Markeren"
        }
        $global:ArrayAccounts = $global:Array
        $global:AantalAccounts = $global:Aantal + " (" + ($global:ArrayAccounts.SamAccountName -join ", ") + ")"
    }
    # ==========
    # Finalizing
    # ==========
    Set-Variable $global:VarHeaderName -Value $global:AantalAccounts
    Set-Variable $global:VarHeaderName -Value ("Write-Host -NoNewLine ('- ' + '$Task' + ': ') -ForegroundColor Gray; Write-Host '" + (Get-Variable -Name $global:VarHeaderName).Value + "' -ForegroundColor Yellow")
    Invoke-Expression -Command $FunctionTaskNames[[int]$FunctionTaskNames.IndexOf(($MyInvocation.MyCommand).Name) + 1]
}
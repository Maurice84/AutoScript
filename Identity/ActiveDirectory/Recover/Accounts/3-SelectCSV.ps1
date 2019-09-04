Function Identity-ActiveDirectory-Recover-Accounts-3-SelectCSV {
    # ============
    # Declarations
    # ============
    $Task = "Select the CSV-file with domain account(s) to recover"
    # =========
    # Execution
    # =========
    Script-Module-SetHeaders -CurrentTask $Task -Name ($MyInvocation.MyCommand).Name
    Script-Index-Objects -CurrentTask $Task -FunctionName ($MyInvocation.MyCommand).Name -Type "Bestand" -Functie "Selecteren" -Filter "CSV"
    $global:File = $global:Array.Name
    # ==========
    # Finalizing
    # ==========
    Set-Variable $global:VarHeaderName -Value $File.Split('\')[-1]
    Set-Variable $global:VarHeaderName -Value ("Write-Host -NoNewLine ('- ' + '$Task' + ': ') -ForegroundColor Gray; Write-Host '" + (Get-Variable -Name $global:VarHeaderName).Value + "' -ForegroundColor Yellow")
    Invoke-Expression -Command $FunctionTaskNames[[int]$FunctionTaskNames.IndexOf(($MyInvocation.MyCommand).Name) + 1]
}
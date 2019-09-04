Function Software-SyncBackPro-Create-Profiles-ForAccounts-2-SelectCSV {
    # ============
    # Declarations
    # ============
    $Task = "Select the CSV-file(s) with domain account(s) which you want to use for the SyncBackPro profiles"
    # =========
    # Execution
    # =========
    Script-Module-SetHeaders -CurrentTask $Task -Name ($MyInvocation.MyCommand).Name
    Script-Index-Objects -CurrentTask $Task -FunctionName ($MyInvocation.MyCommand).Name -Type "Bestand" -Functie "Selecteren" -Filter "CSV"
    $global:FileCSV = $global:Array.Name
    # ==========
    # Finalizing
    # ==========
    Set-Variable $global:VarHeaderName -Value $FileCSV.Split('\')[-1]
    Set-Variable $global:VarHeaderName -Value ("Write-Host -NoNewLine ('- ' + '$Task' + ': ') -ForegroundColor Gray; Write-Host '" + (Get-Variable -Name $global:VarHeaderName).Value + "' -ForegroundColor Yellow")
    Invoke-Expression -Command $FunctionTaskNames[[int]$FunctionTaskNames.IndexOf(($MyInvocation.MyCommand).Name) + 1]
}
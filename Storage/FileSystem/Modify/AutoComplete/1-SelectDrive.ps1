Function Storage-FileSystem-Modify-AutoComplete-1-SelectDrive {
    # ============
    # Declarations
    # ============
    $Task = "Select the (network)drive where the CSV-file(s) with domain account(s) is located"
    # =========
    # Execution
    # =========
    Script-Module-SetHeaders -CurrentTask $Task -Name ($MyInvocation.MyCommand).Name
    Script-Index-Objects -CurrentTask $Task -FunctionName ($MyInvocation.MyCommand).Name -Filter "CSV" -Functie "Selecteren" -Type "Drive"
    $global:DriveCSV = $global:Drive
    # ==========
    # Finalizing
    # ==========
    Set-Variable $global:VarHeaderName -Value $global:DriveCSV
    Set-Variable $global:VarHeaderName -Value ("Write-Host -NoNewLine ('- ' + '$Task' + ': ') -ForegroundColor Gray; Write-Host '" + (Get-Variable -Name $global:VarHeaderName).Value + "' -ForegroundColor Yellow")
    Invoke-Expression -Command $FunctionTaskNames[[int]$FunctionTaskNames.IndexOf(($MyInvocation.MyCommand).Name) + 1]
}